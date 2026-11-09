"""Phase 1.6: Word footnotes and endnotes keep their place, numbers, text and type.

Import: a reference becomes its Word number as a superscript marker in the
section text (¹ for footnotes, ⁱ for endnotes); the note text travels in a
``content_rich`` note segment. LaTeX: ``\\footnote`` and a linked "Son
Notlar" list; Word: native footnotes/endnotes. Notes outside the body and
references without text are reported. Note markers never break or join
citation links.
"""
import copy
import io
import json
import re
import unittest
import zipfile

from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

import app as application
from citation_links import citation_report
from docx_export import generate_docx_from_form
from formatter import extract_form_data_from_docx, generate_latex_from_form, note_marker, roman_numeral
from test_formatter_templates import sample_article

REFERENCES = ['Kaya, B., & Demir, C. (2021). Çevrim içi görünürlük. Örnek Dergi, 5(1), 20–34.',
              'Yılmaz, A. (2020). Kırsal kalkınma ve turizm. Örnek Yayınevi.']


def _notes_part(document, kind, notes):
    """Add footnotes.xml/endnotes.xml as Word writes it; ``notes`` maps id → list of paragraph XML bodies."""
    root = kind + 's'
    xml = [f'<w:{root} {nsdecls("w")}>',
           f'<w:{kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:{kind}>',
           f'<w:{kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{kind}>']
    for note_id, paragraphs in notes.items():
        body = ''.join(
            '<w:p>' + (f'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:{kind}Ref/></w:r>' if i == 0 else '')
            + paragraph + '</w:p>' for i, paragraph in enumerate(paragraphs))
        xml.append(f'<w:{kind} w:id="{note_id}">{body}</w:{kind}>')
    xml.append(f'</w:{root}>')
    content_type = CT.WML_FOOTNOTES if kind == 'footnote' else CT.WML_ENDNOTES
    relationship = RT.FOOTNOTES if kind == 'footnote' else RT.ENDNOTES
    part = Part(PackURI(f'/word/{root}.xml'), content_type, ''.join(xml).encode('utf-8'), document.part.package)
    document.part.relate_to(part, relationship)


def t(text, extra=''):
    return f'<w:r>{extra}<w:t xml:space="preserve">{text}</w:t></w:r>'


def ref(paragraph, kind, note_id):
    paragraph._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                                  f'<w:{kind}Reference w:id="{note_id}"/></w:r>'))


def _front_matter(document, title='Dipnot Aktarımı Deneme Başlığı'):
    document.add_paragraph(title, style='Title')
    document.add_paragraph('Öz: Bu çalışma Word dipnotlarının ve son notlarının aktarımını örnek bir metin ile '
                           'değerlendirmek için hazırlanmıştır.')
    document.add_paragraph('Anahtar Kelimeler: dipnot; son not')
    document.add_heading('Giriş', level=1)


def _references(document):
    document.add_heading('Kaynakça', level=1)
    for reference in REFERENCES:
        document.add_paragraph(reference)


def _import(document):
    stream = io.BytesIO(); document.save(stream)
    return extract_form_data_from_docx(stream.getvalue())[0]


def manuscript():
    document = Document()
    _notes_part(document, 'footnote', {
        '3': [t(' Veriler 2024 saha çalışmasından alınmıştır.')],
        '5': [t(' İkinci dipnot; bkz. Yılmaz, 2020.'), t('Dipnotun ikinci paragrafı: CO')
              + t('2', '<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>') + t(' ölçüldü.')],
        '9': [t(' Kullanılmayan dipnot.')],
    })
    _notes_part(document, 'endnote', {'2': [t(' Birinci son not.')], '4': [t(' İkinci son not.')]})
    _front_matter(document)
    p = document.add_paragraph('Dipnotlu cümle')
    ref(p, 'footnote', 3)
    p.add_run(' devam eder ve son not')
    ref(p, 'endnote', 2)
    p.add_run(' ekler.')
    p = document.add_paragraph('Atıf Yılmaz, 2020')
    ref(p, 'footnote', 5)
    p.add_run(' ile biter; Kaya & Demir, 2021')
    ref(p, 'endnote', 4)
    p.add_run(' de bağlanır.')
    _references(document)
    return document


class ImportTests(unittest.TestCase):
    def setUp(self):
        self.data = _import(manuscript())
        self.section = self.data['sections'][0]

    def test_markers_keep_position_and_word_numbering(self):
        self.assertEqual(self.section['content'],
                         'Dipnotlu cümle¹ devam eder ve son notⁱ ekler.\n\n'
                         'Atıf Yılmaz, 2020² ile biter; Kaya & Demir, 2021ⁱⁱ de bağlanır.')

    def test_note_segments(self):
        notes = [s for e in self.section['content_rich']['paragraphs'] for s in e['segments'] if 'note' in s]
        self.assertEqual(notes, [
            {'t': '¹', 'note': {'type': 'footnote', 'paragraphs': ['Veriler 2024 saha çalışmasından alınmıştır.']}},
            {'t': 'ⁱ', 'note': {'type': 'endnote', 'paragraphs': ['Birinci son not.']}},
            {'t': '²', 'note': {'type': 'footnote', 'paragraphs': ['İkinci dipnot; bkz. Yılmaz, 2020.',
                                                                  'Dipnotun ikinci paragrafı: CO₂ ölçüldü.']}},
            {'t': 'ⁱⁱ', 'note': {'type': 'endnote', 'paragraphs': ['İkinci son not.']}},
        ])

    def test_warnings_and_stats(self):
        self.assertEqual(self.data['stats']['footnote'], 2)
        self.assertEqual(self.data['stats']['endnote'], 2)
        self.assertNotIn('note_outside_body', self.data['stats'])
        self.assertTrue(any(w.startswith('2 Word dipnotu aktarıldı') for w in self.data['warnings']))
        self.assertTrue(any(w.startswith('2 Word son notu aktarıldı') for w in self.data['warnings']))

    def test_citations_after_markers_still_link(self):
        data = dict(self.data, references='\n'.join(REFERENCES))
        self.assertEqual(citation_report(data)['linked_in_body'], 2)

    def test_server_accepts_the_import(self):
        application.validate_form(copy.deepcopy(self.data))


class EdgeCaseTests(unittest.TestCase):
    def test_note_outside_body_is_reported_with_its_text(self):
        document = Document()
        _notes_part(document, 'footnote', {'1': [t(' Başlık notu: proje desteği.')], '2': [t(' Gövde notu.')]})
        document.add_paragraph('Dipnot Aktarımı Deneme Başlığı', style='Title')
        ref(document.paragraphs[0], 'footnote', 1)
        document.add_paragraph('Öz: Bu çalışma Word dipnotlarının aktarımını örnek bir metin ile değerlendirmek '
                               'için hazırlanmıştır.')
        document.add_paragraph('Anahtar Kelimeler: dipnot')
        document.add_heading('Giriş', level=1)
        p = document.add_paragraph('Gövde'); ref(p, 'footnote', 2); p.add_run(' cümlesi.')
        data = _import(document)
        self.assertEqual(data['cover']['tr_title'], 'Dipnot Aktarımı Deneme Başlığı¹')
        self.assertEqual(data['stats']['note_outside_body'], 1)
        self.assertIn('Aktarılamayan not metni: “Başlık notu: proje desteği.”', data['warnings'])
        self.assertEqual(data['sections'][0]['content'], 'Gövde² cümlesi.')
        tex = generate_latex_from_form(_article(data['sections']), {}, {})
        self.assertIn(r'Gövde\footnote{Gövde notu.} cümlesi.', tex)

    def test_missing_note_text_is_reported(self):
        document = Document(); _front_matter(document)
        p = document.add_paragraph('Eksik dipnot'); ref(p, 'footnote', 7); p.add_run(' burada.')
        data = _import(document)
        self.assertEqual(data['sections'][0]['content'], 'Eksik dipnot burada.')
        self.assertEqual(data['stats']['missing_note'], 1)

    def test_hidden_reference_is_not_a_note(self):
        document = Document()
        _notes_part(document, 'footnote', {'1': [t(' Gizli dipnot.')]})
        _front_matter(document)
        p = document.add_paragraph('Görünen')
        p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:vanish/></w:rPr><w:footnoteReference w:id="1"/></w:r>'))
        p.add_run(' metin.')
        data = _import(document)
        self.assertEqual(data['sections'][0]['content'], 'Görünen metin.')
        self.assertNotIn('footnote', data['stats'])

    def test_marker_helpers(self):
        self.assertEqual([note_marker('footnote', n) for n in (1, 10, 23)], ['¹', '¹⁰', '²³'])
        self.assertEqual([roman_numeral(n) for n in (1, 4, 9, 14, 40, 90, 400, 1999)],
                         ['i', 'iv', 'ix', 'xiv', 'xl', 'xc', 'cd', 'mcmxcix'])
        self.assertEqual(note_marker('endnote', 14), 'ˣⁱᵛ')


def _article(sections):
    article = sample_article()
    article['sections'] = sections
    article['figtables'] = []
    article['references'] = '\n'.join(REFERENCES)
    return article


class LatexTests(unittest.TestCase):
    def setUp(self):
        self.data = _import(manuscript())
        self.tex = generate_latex_from_form(_article(self.data['sections']), {}, {})

    def test_footnotes_and_endnotes(self):
        self.assertIn(r'Dipnotlu cümle\footnote{Veriler 2024 saha çalışmasından alınmıştır.} devam eder ve son '
                      r'not\textsuperscript{\hyperlink{aiditor_endnote_1}{i}} ekler.', self.tex)
        self.assertIn(r'Atıf \hyperlink{aiditor_ref_1}{Yılmaz, 2020}\footnote{İkinci dipnot; bkz. '
                      r'\hyperlink{aiditor_ref_1}{Yılmaz, 2020}.\par Dipnotun ikinci paragrafı: CO\textsubscript{2} '
                      r'ölçüldü.} ile biter; \hyperlink{aiditor_ref_0}{Kaya \& Demir, 2021}'
                      r'\textsuperscript{\hyperlink{aiditor_endnote_2}{ii}} de bağlanır.', self.tex)
        self.assertIn(r'\section*{Son Notlar / Endnotes}', self.tex)
        self.assertIn(r'\item[\hypertarget{aiditor_endnote_1}{}i.] Birinci son not.', self.tex)
        self.assertIn(r'\item[\hypertarget{aiditor_endnote_2}{}ii.] İkinci son not.', self.tex)
        self.assertLess(self.tex.index('Son Notlar'), self.tex.index('Kaynakça'))
        self.assertNotIn('Kullanılmayan dipnot', self.tex)

    def test_english_only_heading(self):
        tex = generate_latex_from_form(_article(self.data['sections']), {}, {'english_only': True})
        self.assertIn(r'\section*{Endnotes}', tex)

    def test_no_notes_no_endnote_section(self):
        sections = [{k: v for k, v in s.items() if k != 'content_rich'} for s in self.data['sections']]
        tex = generate_latex_from_form(_article(sections), {}, {})
        self.assertNotIn('Endnotes', tex)
        self.assertNotIn(r'\footnote{', tex.split(r'\begin{document}')[1])
        # Without content_rich the markers stay visible as superscripts.
        self.assertIn(r'Dipnotlu cümle\textsuperscript{1} devam eder ve son not\textsuperscript{i} ekler.', tex)
        self.assertIn(r'{Kaya \& Demir, 2021}\textsuperscript{i}\textsuperscript{i} de', tex)

    def test_edited_paragraph(self):
        sections = copy.deepcopy(self.data['sections'])
        sections[0]['content'] = sections[0]['content'].replace('devam eder ve son', 'sürer ve son')
        tex = generate_latex_from_form(_article(sections), {}, {})
        # Markers and their neighbours are unchanged: both notes stay.
        self.assertIn(r'Dipnotlu cümle\footnote{Veriler 2024 saha çalışmasından alınmıştır.} sürer ve son '
                      r'not\textsuperscript{\hyperlink{aiditor_endnote_1}{i}} ekler.', tex)
        sections[0]['content'] = sections[0]['content'].replace('cümle¹ sürer', 'cümle¹, sürer')
        tex = generate_latex_from_form(_article(sections), {}, {})
        self.assertIn(r'Dipnotlu cümle\textsuperscript{1}, sürer ve son not\textsuperscript{\hyperlink{aiditor_endnote_1}{i}} ekler.', tex)


class WordTests(unittest.TestCase):
    def setUp(self):
        self.data = _import(manuscript())
        self.blob = generate_docx_from_form(_article(self.data['sections']), {}, {})
        self.archive = zipfile.ZipFile(io.BytesIO(self.blob))

    def test_native_notes(self):
        names = self.archive.namelist()
        self.assertIn('word/footnotes.xml', names)
        self.assertIn('word/endnotes.xml', names)
        types = self.archive.read('[Content_Types].xml').decode()
        self.assertIn('footnotes+xml', types)
        self.assertIn('endnotes+xml', types)
        rels = self.archive.read('word/_rels/document.xml.rels').decode()
        self.assertIn('relationships/footnotes', rels)
        self.assertIn('relationships/endnotes', rels)
        body = self.archive.read('word/document.xml').decode()
        self.assertEqual(len(re.findall(r'<w:footnoteReference w:id="\d+"/>', body)), 2)
        self.assertEqual(len(re.findall(r'<w:endnoteReference w:id="\d+"/>', body)), 2)
        self.assertNotIn('¹', body)
        self.assertNotIn('ⁱ', body)

    def test_word_output_reimports_with_the_same_notes(self):
        again = extract_form_data_from_docx(self.blob)[0]
        section = next(s for s in again['sections'] if 'Dipnotlu' in s['content'])
        self.assertTrue(section['content'].endswith(self.data['sections'][0]['content']))
        notes = lambda sec: [s['note'] for e in sec['content_rich']['paragraphs'] for s in e['segments'] if 'note' in s]
        self.assertEqual(notes(section), notes(self.data['sections'][0]))
        anchors = [h.get(qn('w:anchor')) for h in parse_xml(self.archive.read('word/footnotes.xml')).iter(qn('w:hyperlink'))]
        self.assertEqual(anchors, ['aiditor_ref_1'])

    def test_document_without_notes_has_no_note_parts(self):
        sections = [{k: v for k, v in s.items() if k != 'content_rich'} for s in self.data['sections']]
        names = zipfile.ZipFile(io.BytesIO(generate_docx_from_form(_article(sections), {}, {}))).namelist()
        self.assertNotIn('word/footnotes.xml', names)
        self.assertNotIn('word/endnotes.xml', names)


class ValidationTests(unittest.TestCase):
    def section(self, note):
        return {'id': '0', 'name': 'Giriş', 'level': '1', 'content': 'Metin¹.',
                'content_rich': {'version': 1, 'paragraphs': [{'plain': 'Metin¹.', 'segments': [
                    {'t': 'Metin'}, dict({'t': '¹'}, **note), {'t': '.'}]}]}}

    def test_valid_and_invalid_notes(self):
        application.validate_form({'sections': [self.section({'note': {'type': 'footnote', 'paragraphs': ['Not.']}})]})
        for bad in ({'note': {'type': 'sidenote', 'paragraphs': ['x']}}, {'note': {'type': 'footnote', 'paragraphs': []}},
                    {'note': {'type': 'footnote', 'paragraphs': [1]}}, {'note': {'type': 'footnote', 'paragraphs': ['x'], 'id': 1}},
                    {'note': {'type': 'footnote', 'paragraphs': ['x']}, 'b': True},
                    {'note': {'type': 'footnote', 'paragraphs': ['x'] * 51}}):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                application.validate_form({'sections': [self.section(bad)]})

    def test_tex_in_note_text_is_escaped(self):
        section = self.section({'note': {'type': 'footnote', 'paragraphs': [r'\input{/etc/passwd} % & _']}})
        tex = generate_latex_from_form(_article([section]), {}, {})
        self.assertIn(r'Metin\footnote{\textbackslash{}input\{/etc/passwd\} \% \& \_}.', tex)


if __name__ == '__main__':
    unittest.main()
