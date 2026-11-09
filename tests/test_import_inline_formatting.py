"""Phase 1.4a: partial bold/italic in Word body paragraphs survives import as data.

``section['content']`` stays the authoritative plain text and is unchanged.
A section whose body has bold or italic text also gets an optional
``content_rich`` = {'version': 1, 'paragraphs': [{'plain', 'segments'}]}, one
entry per paragraph of ``content`` as the renderers split it. Segments are
{'t': text, 'b'?: True, 'i'?: True} and always rebuild ``plain`` exactly.
Bold/italic come from ``_RunPropertyResolver`` (direct formatting, character
and paragraph styles with toggling, basedOn, defaults). Heading detection,
citations, figure anchors (``after_para``), hidden text (Phase 1.2) and
Unicode sub/superscript (Phase 1.3) are unchanged. The output engines do not
read ``content_rich`` yet, so LaTeX and Word output are unchanged too.
"""
import copy
import io
import json
from pathlib import Path
import tempfile
import unittest
import uuid

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from citation_links import citation_report
from docx_export import generate_docx_from_form
from formatter import CONTENT_RICH_VERSION, _content_paragraphs, extract_form_data_from_docx, generate_latex_from_form
import app as application
from test_accounts import headers, make_setup, register
from test_citation_sources import _fld, _run
from test_formatter_templates import sample_article
from test_regression_golden import FIGURE_PNG, docx_snapshot, legacy_project

REFERENCES = ['Kaya, B., & Demir, C. (2021). Çevrim içi görünürlük. Örnek Dergi, 5(1), 20–34.',
              'Yılmaz, A. (2020). Kırsal kalkınma ve turizm. Örnek Yayınevi.']


def _import(document):
    stream = io.BytesIO(); document.save(stream)
    return extract_form_data_from_docx(stream.getvalue())[0]


def _front_matter(document):
    document.add_paragraph('Satır İçi Biçim Aktarımı Deneme Başlığı', style='Title')
    document.add_paragraph('Öz: Bu çalışma kısmi kalın ve italik biçimlerin Word dosyasından aktarımını örnek '
                           'bir metin ile değerlendirmek için hazırlanmıştır.')
    document.add_paragraph('Anahtar Kelimeler: biçim; aktarım')
    document.add_heading('Giriş', level=1)


def _references(document):
    document.add_heading('Kaynakça', level=1)
    for reference in REFERENCES:
        document.add_paragraph(reference)


def _runs(paragraph, *pieces):
    """Add (text, flags) pieces; flags is a string containing 'b' and/or 'i'."""
    for text, flags in pieces:
        run = paragraph.add_run(text)
        if 'b' in flags:
            run.bold = True
        if 'i' in flags:
            run.italic = True
    return paragraph


def _formatted_run(text, flags=''):
    """A bare w:r (for fields and content controls) with direct b/i."""
    run = _run(text)
    if flags:
        r_pr = OxmlElement('w:rPr')
        for flag in flags:
            r_pr.append(OxmlElement('w:' + flag))
        run.insert(0, r_pr)
    return run


def _plain_twin(build):
    """The same manuscript with every direct b/i removed (the pre-1.4a view of the text)."""
    document = Document(); _front_matter(document); build(document); _references(document)
    for element in document.element.body.iter(qn('w:rPr')):
        for child in list(element):
            if child.tag in (qn('w:b'), qn('w:i')):
                element.remove(child)
    return _import(document)


def _build(build):
    document = Document(); _front_matter(document); build(document); _references(document)
    return _import(document)


def _segments_of(section):
    return [entry['segments'] for entry in section['content_rich']['paragraphs']]


class InvariantMixin:
    def assert_invariants(self, data):
        for section in data['sections']:
            rich = section.get('content_rich')
            if rich is None:
                continue
            self.assertEqual(rich['version'], CONTENT_RICH_VERSION)
            self.assertEqual(set(rich), {'version', 'paragraphs'})
            self.assertEqual([entry['plain'] for entry in rich['paragraphs']], _content_paragraphs(section['content']))
            for entry in rich['paragraphs']:
                self.assertEqual(set(entry), {'plain', 'segments'})
                self.assertEqual(''.join(segment['t'] for segment in entry['segments']), entry['plain'])
                for segment in entry['segments']:
                    self.assertTrue(segment['t'])
                    self.assertLessEqual(set(segment), {'t', 'b', 'i'})
                    self.assertTrue(all(segment[key] is True for key in ('b', 'i') if key in segment))
                for first, second in zip(entry['segments'], entry['segments'][1:]):
                    self.assertNotEqual({k: v for k, v in first.items() if k != 't'},
                                        {k: v for k, v in second.items() if k != 't'}, 'adjacent segments not merged')
            # JSON-safe for the /import_docx response.
            self.assertEqual(json.loads(json.dumps(rich, ensure_ascii=False)), rich)


class PartialFormattingTests(InvariantMixin, unittest.TestCase):
    @staticmethod
    def build(document):
        _runs(document.add_paragraph(),
              ('İnsan türü ', ''), ('Homo', 'i'), (' ', 'i'), ('sapiens', 'i'),
              (' olarak bilinir ve bu bulgu ', ''), ('çok önemli', 'b'), (' kabul edilir.', ''))
        _runs(document.add_paragraph(),
              ('Kalın-italik ', ''), ('vurgu', 'bi'), (' ile ', ''), ('kalın ', 'b'), ('ve italik', 'bi'),
              (' birlikte kullanılır.', ''))
        document.add_paragraph('Biçimsiz son paragraf burada yer alır.')

    def setUp(self):
        self.data = _build(self.build)
        self.section = self.data['sections'][0]

    def test_plain_content_is_unchanged(self):
        self.assertEqual(self.data['sections'][0]['content'], _plain_twin(self.build)['sections'][0]['content'])
        self.assertEqual(self.section['content'],
                         'İnsan türü Homo sapiens olarak bilinir ve bu bulgu çok önemli kabul edilir.\n\n'
                         'Kalın-italik vurgu ile kalın ve italik birlikte kullanılır.\n\n'
                         'Biçimsiz son paragraf burada yer alır.')

    def test_segments(self):
        self.assertEqual(_segments_of(self.section), [
            [{'t': 'İnsan türü '}, {'t': 'Homo sapiens', 'i': True}, {'t': ' olarak bilinir ve bu bulgu '},
             {'t': 'çok önemli', 'b': True}, {'t': ' kabul edilir.'}],
            [{'t': 'Kalın-italik '}, {'t': 'vurgu', 'b': True, 'i': True}, {'t': ' ile '},
             {'t': 'kalın ', 'b': True}, {'t': 've italik', 'b': True, 'i': True}, {'t': ' birlikte kullanılır.'}],
            [{'t': 'Biçimsiz son paragraf burada yer alır.'}],
        ])
        self.assert_invariants(self.data)

    def test_space_between_differently_formatted_words_is_not_formatted(self):
        data = _build(lambda d: _runs(d.add_paragraph(), ('Bu ', ''), ('önemli ', 'b'), ('bir', ''), (' sonuçtur.', '')))
        self.assertEqual(_segments_of(data['sections'][0]),
                         [[{'t': 'Bu '}, {'t': 'önemli', 'b': True}, {'t': ' bir sonuçtur.'}]])

    def test_unformatted_body_has_no_content_rich(self):
        def build(document):
            document.add_paragraph('Tamamen düz bir paragraf burada bulunur.')
            document.add_paragraph('İkinci düz paragraf da burada bulunur.')
        data = _build(build)
        self.assertNotIn('content_rich', data['sections'][0])
        self.assertEqual(data, _plain_twin(build))

    def test_whole_paragraph_italic_and_bold_are_kept(self):
        def build(document):
            _runs(document.add_paragraph(), ('Bu paragrafın tamamı italik yazılmış bir alıntıdır.', 'i'))
            _runs(document.add_paragraph(), ('Bu paragrafın tamamı kalın yazılmış uzun bir uyarı cümlesidir.', 'b'))
        data = _build(build)
        self.assertEqual(len(data['sections']), 1)
        self.assertEqual(_segments_of(data['sections'][0]), [
            [{'t': 'Bu paragrafın tamamı italik yazılmış bir alıntıdır.', 'i': True}],
            [{'t': 'Bu paragrafın tamamı kalın yazılmış uzun bir uyarı cümlesidir.', 'b': True}],
        ])

    def test_all_bold_short_line_is_still_a_heading(self):
        def build(document):
            document.add_paragraph('Giriş paragrafı burada yer alır ve sürer.')
            _runs(document.add_paragraph(), ('Yöntem ve Veri', 'b'))
            _runs(document.add_paragraph(), ('Veri ', ''), ('ankette', 'i'), (' toplandı.', ''))
        data = _build(build)
        self.assertEqual([s['name'] for s in data['sections']], ['Giriş', 'Yöntem ve Veri'])
        self.assertNotIn('content_rich', data['sections'][0])
        self.assertEqual(_segments_of(data['sections'][1]),
                         [[{'t': 'Veri '}, {'t': 'ankette', 'i': True}, {'t': ' toplandı.'}]])

    def test_turkish_unicode_and_latex_special_characters_stay_raw(self):
        data = _build(lambda d: _runs(d.add_paragraph(),
                                      ('Oran ', ''), ('%50 & $3_x #1 ~ ^ {ğüşıöçİ}', 'i'), (' değerdir — “tırnak”.', '')))
        segments = _segments_of(data['sections'][0])[0]
        self.assertEqual(segments[1], {'t': '%50 & $3_x #1 ~ ^ {ğüşıöçİ}', 'i': True})
        self.assertNotIn('\\', json.dumps(data['sections'][0]['content_rich'], ensure_ascii=False))
        self.assert_invariants(data)


class StyleResolutionTests(InvariantMixin, unittest.TestCase):
    def test_character_style_and_based_on_chain(self):
        def build(document):
            base = document.styles.add_style('Vurgu Taban', WD_STYLE_TYPE.CHARACTER)
            base.font.italic = True
            derived = document.styles.add_style('Tür Adı', WD_STYLE_TYPE.CHARACTER)
            derived.base_style = base
            p = document.add_paragraph('Örnek tür ')
            p.add_run('Canis lupus', style='Tür Adı')
            p.add_run(' gözlendi.')
        data = _build(build)
        self.assertEqual(_segments_of(data['sections'][0]),
                         [[{'t': 'Örnek tür '}, {'t': 'Canis lupus', 'i': True}, {'t': ' gözlendi.'}]])

    def test_bold_paragraph_style_toggles_bold_character_style(self):
        def build(document):
            paragraph_style = document.styles.add_style('Kalın Gövde', WD_STYLE_TYPE.PARAGRAPH)
            paragraph_style.font.bold = True
            strong = document.styles.add_style('Güçlü', WD_STYLE_TYPE.CHARACTER)
            strong.font.bold = True
            p = document.add_paragraph('Kalın stil paragrafında ', style='Kalın Gövde')
            p.add_run('bu kısım', style='Güçlü')
            p.add_run(' normal görünmez, uzun bir cümledir.')
        data = _build(build)
        self.assertEqual(_segments_of(data['sections'][0]),
                         [[{'t': 'Kalın stil paragrafında', 'b': True}, {'t': ' bu kısım '},
                           {'t': 'normal görünmez, uzun bir cümledir.', 'b': True}]])

    def test_direct_off_overrides_italic_style(self):
        def build(document):
            italic_style = document.styles.add_style('İtalik Gövde', WD_STYLE_TYPE.PARAGRAPH)
            italic_style.font.italic = True
            p = document.add_paragraph('İtalik paragrafta ', style='İtalik Gövde')
            p.add_run('düz kelime').italic = False
            p.add_run(' bulunur ve cümle sürer.')
        data = _build(build)
        self.assertEqual(_segments_of(data['sections'][0]),
                         [[{'t': 'İtalik paragrafta', 'i': True}, {'t': ' düz kelime '},
                           {'t': 'bulunur ve cümle sürer.', 'i': True}]])


class ParagraphAlignmentTests(InvariantMixin, unittest.TestCase):
    def test_identical_paragraphs_keep_their_own_formatting(self):
        def build(document):
            document.add_paragraph('Aynı metinli paragraf burada yer alır.')
            _runs(document.add_paragraph(), ('Aynı metinli ', ''), ('paragraf', 'i'), (' burada yer alır.', ''))
            document.add_paragraph('Aynı metinli paragraf burada yer alır.')
        data = _build(build)
        segments = _segments_of(data['sections'][0])
        self.assertEqual(segments[0], [{'t': 'Aynı metinli paragraf burada yer alır.'}])
        self.assertEqual(segments[1], [{'t': 'Aynı metinli '}, {'t': 'paragraf', 'i': True}, {'t': ' burada yer alır.'}])
        self.assertEqual(segments[2], segments[0])

    def test_list_marker_is_plain_and_matches_content(self):
        def build(document):
            p = document.add_paragraph(style='List Bullet')
            _runs(p, ('Birinci madde ', ''), ('vurgulu', 'b'), (' kısım içerir.', ''))
        data = _build(build)
        self.assertTrue(data['sections'][0]['content'].startswith('• '))
        self.assertEqual(_segments_of(data['sections'][0]),
                         [[{'t': '• Birinci madde '}, {'t': 'vurgulu', 'b': True}, {'t': ' kısım içerir.'}]])

    def test_blank_line_inside_word_paragraph_follows_renderer_split(self):
        def build(document):
            p = _runs(document.add_paragraph(), ('İlk satır ', ''), ('italik', 'i'))
            p.add_run().add_break(); p.add_run().add_break()
            _runs(p, ('ikinci ', 'b'), ('satır.', ''))
        data = _build(build)
        self.assertEqual(_segments_of(data['sections'][0]),
                         [[{'t': 'İlk satır '}, {'t': 'italik', 'i': True}],
                          [{'t': 'ikinci', 'b': True}, {'t': ' satır.'}]])
        self.assert_invariants(data)

    def test_single_line_break_and_tab_are_kept(self):
        def build(document):
            p = _runs(document.add_paragraph(), ('Satır ', ''), ('bir', 'i'))
            p.add_run().add_break()
            _runs(p, ('satır\tiki', 'i'), ('.', ''))
        data = _build(build)
        self.assertEqual(_segments_of(data['sections'][0]),
                         [[{'t': 'Satır '}, {'t': 'bir\nsatır\tiki', 'i': True}, {'t': '.'}]])

    def test_sections_without_formatting_get_no_field(self):
        def build(document):
            _runs(document.add_paragraph(), ('Biçimli ', 'i'), ('paragraf burada.', ''))
            document.add_heading('Yöntem', level=1)
            document.add_paragraph('Düz yöntem paragrafı burada yer alır.')
        data = _build(build)
        self.assertIn('content_rich', data['sections'][0])
        self.assertNotIn('content_rich', data['sections'][1])

    def test_unlabelled_body_without_headings(self):
        document = Document()
        document.add_paragraph('Başlıksız Gövde Deneme Başlığı', style='Title')
        document.add_paragraph('Anahtar Kelimeler: deneme; gövde')
        _runs(document.add_paragraph(), ('Bu gövde paragrafı bir ', ''), ('italik', 'i'),
              (' kelime içeren yeterince uzun bir metindir.', ''))
        data = _import(document)
        self.assertEqual(data['sections'][0]['name'], 'Makale Metni / Article Body')
        self.assertEqual(_segments_of(data['sections'][0]),
                         [[{'t': 'Bu gövde paragrafı bir '}, {'t': 'italik', 'i': True},
                           {'t': ' kelime içeren yeterince uzun bir metindir.'}]])


class EarlierPhaseInteractionTests(InvariantMixin, unittest.TestCase):
    def test_hidden_text_is_not_in_segments(self):
        def build(document):
            p = _runs(document.add_paragraph(), ('Görünen ', ''), ('italik', 'i'))
            hidden = p.add_run(' GİZLİ')
            hidden.italic = True
            hidden.font.hidden = True
            _runs(p, (' metin sürer.', ''))
        data = _build(build)
        self.assertNotIn('GİZLİ', json.dumps(data, ensure_ascii=False))
        self.assertEqual(_segments_of(data['sections'][0]),
                         [[{'t': 'Görünen '}, {'t': 'italik', 'i': True}, {'t': ' metin sürer.'}]])
        self.assertEqual(data['stats'].get('hidden_text'), 1)

    def test_scientific_notation_stays_unicode_inside_formatted_text(self):
        def build(document):
            p = document.add_paragraph('Ölçümde ')
            for text, flags, script in (('CO', 'b', None), ('2', 'b', 'sub'), (', H', '', None), ('2', '', 'sub'),
                                        ('SO', '', None), ('4', '', 'sub'), (' ve SO', '', None), ('4', 'i', 'sub'),
                                        ('2', 'i', 'sup'), ('−', 'i', 'sup'), (', 10', '', None), ('-3', '', 'sup'),
                                        (', x', 'i', None), ('n', 'i', 'sup'), (' ve 4 cm', '', None), ('3', '', 'sup'),
                                        (' bulundu.', '', None)):
                run = p.add_run(text)
                run.bold = 'b' in flags or None
                run.italic = 'i' in flags or None
                if script == 'sup':
                    run.font.superscript = True
                elif script == 'sub':
                    run.font.subscript = True
        data = _build(build)
        content = data['sections'][0]['content']
        self.assertEqual(content, 'Ölçümde CO₂, H₂SO₄ ve SO₄²⁻, 10⁻³, xⁿ ve 4 cm³ bulundu.')
        self.assertEqual(content, _plain_twin(build)['sections'][0]['content'])
        self.assertEqual(_segments_of(data['sections'][0]),
                         [[{'t': 'Ölçümde '}, {'t': 'CO₂', 'b': True}, {'t': ', H₂SO₄ ve SO'},
                           {'t': '₄²⁻', 'i': True}, {'t': ', 10⁻³'}, {'t': ', xⁿ', 'i': True},
                           {'t': ' ve 4 cm³ bulundu.'}]])
        self.assertNotIn('unconverted_script', data['stats'])


class CitationTests(InvariantMixin, unittest.TestCase):
    """Zotero, EndNote and Mendeley citations whose text crosses segment boundaries."""

    @staticmethod
    def build(document):
        p = document.add_paragraph('Zotero alanı ')
        code = ' ADDIN ZOTERO_ITEM CSL_CITATION ' + json.dumps(
            {'properties': {'formattedCitation': '(Yılmaz, 2020)'}}, ensure_ascii=False)
        for node in (_fld('begin'), _run(code, 'w:instrText'), _fld('separate'),
                     _formatted_run('(', 'b'), _formatted_run('Yılmaz', 'bi'), _formatted_run(', 2020)'), _fld('end')):
            p._p.append(node)
        _runs(p, (' ile ', ''), ('başlar', 'i'), ('.', ''))

        p = _runs(document.add_paragraph(), ('EndNote alanı ', 'i'))
        code = ' ADDIN EN.CITE <EndNote><Cite><Author>Kaya</Author><Year>2021</Year></Cite></EndNote>'
        for node in (_fld('begin'), _run(code, 'w:instrText'), _fld('separate'),
                     _formatted_run('(Kaya', 'i'), _formatted_run(' & Demir, 2021)'), _fld('end')):
            p._p.append(node)
        p.add_run(' ile sürer.')

        p = document.add_paragraph('Mendeley denetimi ')
        sdt = parse_xml(f'<w:sdt {nsdecls("w")}><w:sdtPr><w:tag w:val="MENDELEY_CITATION_v3"/></w:sdtPr>'
                        '<w:sdtContent/></w:sdt>')
        content = sdt.find(qn('w:sdtContent'))
        for node in (_formatted_run('(Kaya & ', ''), _formatted_run('Demir', 'b'), _formatted_run(', 2021; Yılmaz, 2020)', 'i')):
            content.append(node)
        p._p.append(sdt)
        p.add_run(' ile biter.')

    def setUp(self):
        self.data = _build(self.build)
        self.twin = _plain_twin(self.build)

    def test_citation_text_is_whole_and_not_duplicated(self):
        content = self.data['sections'][0]['content']
        self.assertEqual(content, self.twin['sections'][0]['content'])
        for shown in ('Zotero alanı (Yılmaz, 2020) ile başlar.', 'EndNote alanı (Kaya & Demir, 2021) ile sürer.',
                      'Mendeley denetimi (Kaya & Demir, 2021; Yılmaz, 2020) ile biter.'):
            self.assertEqual(content.count(shown), 1)
        self.assertNotIn('ADDIN', json.dumps(self.data, ensure_ascii=False))
        self.assertNotIn('EndNote>', json.dumps(self.data, ensure_ascii=False))
        self.assert_invariants(self.data)

    def test_citation_links_are_unchanged(self):
        report = citation_report(self.data)
        self.assertEqual(report, citation_report(self.twin))
        self.assertEqual(report['linked_in_body'], 4)

    def test_segments_split_inside_citations(self):
        segments = _segments_of(self.data['sections'][0])
        self.assertEqual(segments[0], [{'t': 'Zotero alanı '}, {'t': '(', 'b': True}, {'t': 'Yılmaz', 'b': True, 'i': True},
                                       {'t': ', 2020) ile '}, {'t': 'başlar', 'i': True}, {'t': '.'}])
        self.assertEqual(segments[1], [{'t': 'EndNote alanı (Kaya', 'i': True}, {'t': ' & Demir, 2021) ile sürer.'}])
        self.assertEqual(segments[2], [{'t': 'Mendeley denetimi (Kaya & '}, {'t': 'Demir', 'b': True},
                                       {'t': ', 2021; Yılmaz, 2020)', 'i': True}, {'t': ' ile biter.'}])


class OutputUnchangedTests(unittest.TestCase):
    """Figure anchors are unchanged; output differs from the plain twin only where text is formatted."""

    @staticmethod
    def build(document):
        _runs(document.add_paragraph(), ('Şekilden önceki ', ''), ('italik', 'i'), (' paragraf ', ''),
              ('(Yılmaz, 2020)', 'b'), (' ve CO₂.', ''))
        document.add_paragraph().add_run().add_picture(io.BytesIO(FIGURE_PNG))
        document.add_paragraph('Şekil 1. Örnek şekil')
        _runs(document.add_paragraph(), ('Şekilden sonraki ', 'b'), ('paragraf.', ''))

    def setUp(self):
        self.data = _build(self.build)
        self.twin = _plain_twin(self.build)

    def test_after_para_anchor_is_unchanged(self):
        self.assertEqual(len(self.data['figtables']), 1)
        self.assertEqual(self.data['figtables'], self.twin['figtables'])
        self.assertEqual(self.data['figtables'][0]['after_para'], 'Şekilden önceki italik paragraf (Yılmaz, 2020) ve CO₂.')

    def test_whole_import_differs_from_plain_twin_only_by_content_rich(self):
        stripped = copy.deepcopy(self.data)
        for section in stripped['sections']:
            section.pop('content_rich', None)
        self.assertEqual(stripped, self.twin)

    def _article(self, sections):
        article = sample_article()
        article['sections'] = sections
        article['figtables'] = []
        article['references'] = '\n'.join(REFERENCES)
        return article

    def test_output_changes_only_in_formatted_paragraphs(self):
        """Phase 1.4b applies content_rich; everything outside those paragraphs stays byte-identical."""
        with_rich = self._article(copy.deepcopy(self.data['sections']))
        without = self._article(copy.deepcopy(self.twin['sections']))
        self.assertIn('content_rich', with_rich['sections'][0])
        tex, plain_tex = generate_latex_from_form(with_rich, {}, {}), generate_latex_from_form(without, {}, {})
        changed = [(a, b) for a, b in zip(tex.splitlines(), plain_tex.splitlines()) if a != b]
        self.assertEqual(len(tex.splitlines()), len(plain_tex.splitlines()))
        self.assertEqual(changed, [
            (r'Şekilden önceki \textit{italik} paragraf \textbf{(}\hyperlink{aiditor_ref_1}{\textbf{Yılmaz, 2020}}\textbf{)} ve CO\textsubscript{2}.',
             r'Şekilden önceki italik paragraf (\hyperlink{aiditor_ref_1}{Yılmaz, 2020}) ve CO\textsubscript{2}.'),
            (r'\textbf{Şekilden sonraki} paragraf.', 'Şekilden sonraki paragraf.'),
        ])
        for raw in ('segments', '"t"', '<b>', '**'):
            self.assertNotIn(raw, tex)
        rich_doc = Document(io.BytesIO(generate_docx_from_form(with_rich, {}, {})))
        plain_doc = Document(io.BytesIO(generate_docx_from_form(without, {}, {})))
        self.assertEqual([p.text for p in rich_doc.paragraphs], [p.text for p in plain_doc.paragraphs])


class ExistingProjectStorageTests(unittest.TestCase):
    """Current (pre-1.4c) server behaviour for a version-1 project that carries content_rich.

    Documents the baseline before saving is changed: the server validates the
    unknown field and stores it losslessly. The browser's collectSections()
    (templates/index.html) rebuilds sections from the form and drops it; that
    path is handled in Phase 1.4c, not here.
    """

    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix='aiditor-rich-test-')
        self.addCleanup(temporary.cleanup)
        self.app, _ = make_setup(Path(temporary.name))

    def test_server_validates_and_round_trips_content_rich(self):
        imported = _build(PartialFormattingTests.build)
        draft = legacy_project()
        draft['data']['sections'] = copy.deepcopy(imported['sections'])
        self.assertIn('content_rich', draft['data']['sections'][0])
        application.validate_draft_project(copy.deepcopy(draft))
        client = self.app.test_client()
        user = register(client)
        identifier = str(uuid.uuid4())
        saved = client.put('/api/articles/' + identifier, json={'project': draft, 'base_revision': 0}, headers=headers(user))
        self.assertEqual(saved.status_code, 200, saved.json)
        reopened = client.get('/api/articles/' + identifier, headers=headers(user)).json['article']['project']
        self.assertEqual(reopened, draft)


if __name__ == '__main__':
    unittest.main()
