"""Ç3 baseline: scientific content that Word import loses in 2.4.2.

Each `test_gap_*` test runs the target check of the planned fix and asserts
that it currently fails with an AssertionError. Any other exception is a test
bug and surfaces as an error. When Phase 1 closes a gap, the matching test
fails with "AssertionError not raised" and must be turned into a regular test
by calling its `check_*` method directly.
"""
import io
import json
import unittest

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from citation_links import CitationIndex
from docx_export import generate_docx_from_form
from formatter import _ImportReport, _RunPropertyResolver, extract_form_data_from_docx, generate_latex_from_form
from test_formatter_templates import sample_article

M_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
INLINE_EQUATION = (f'<m:oMath {M_NS}><m:r><m:t>E=m</m:t></m:r><m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e>'
                   '<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>')
DISPLAY_EQUATION = (f'<m:oMathPara {M_NS}><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num>'
                    '<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>')
FOOTNOTE_TEXT = 'Veriler 2024 saha çalışmasından alınmıştır.'


def _front_matter(document):
    document.add_paragraph('Bilimsel İçerik Aktarımı Deneme Başlığı', style='Title')
    document.add_paragraph('Öz: Bu çalışma ölçülen değerlerin aktarımını ve sonuçlarını örnek bir metin ile '
                           'değerlendirmek için hazırlanmıştır.')
    document.add_paragraph('Anahtar Kelimeler: ölçüm; aktarım')
    document.add_heading('Giriş', level=1)


def _import(document):
    stream = io.BytesIO(); document.save(stream)
    return extract_form_data_from_docx(stream.getvalue())[0]


def _body(data):
    return '\n\n'.join(section['content'] for section in data['sections'])


def _add_footnotes_part(document, text):
    """python-docx 1.2 has no footnote API; add the part Word writes."""
    xml = (f'<w:footnotes {nsdecls("w")}>'
           '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
           '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
           f'<w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
           f'<w:r><w:t xml:space="preserve"> {text}</w:t></w:r></w:p></w:footnote></w:footnotes>')
    part = Part(PackURI('/word/footnotes.xml'), CT.WML_FOOTNOTES, xml.encode('utf-8'), document.part.package)
    document.part.relate_to(part, RT.FOOTNOTES)


class ScientificImportChecks(unittest.TestCase):
    """Target behaviour for Phase 1. Called through the gap tests below."""

    def check_hidden_text_is_not_imported(self):
        document = Document(); _front_matter(document)
        hidden_style = document.styles.add_style('Gizli Not', WD_STYLE_TYPE.CHARACTER)
        hidden_style.font.hidden = True
        p = document.add_paragraph('Görünen metin.')
        p.add_run(' DOGRUDAN-GIZLI').font.hidden = True
        p.add_run(' STIL-GIZLI', style='Gizli Not')
        body = _body(_import(document))
        self.assertIn('Görünen metin.', body)
        self.assertNotIn('DOGRUDAN-GIZLI', body)
        self.assertNotIn('STIL-GIZLI', body)

    def check_superscript_and_subscript_are_preserved(self):
        document = Document(); _front_matter(document)
        p = document.add_paragraph('Alan 5 m')
        p.add_run('2').font.superscript = True
        p.add_run(', yoğunluk 10')
        p.add_run('-3').font.superscript = True
        p.add_run(' ve CO')
        p.add_run('2').font.subscript = True
        p.add_run(' ölçüldü.')
        body = _body(_import(document))
        self.assertIn('m²', body)
        self.assertIn('10⁻³', body)
        self.assertIn('CO₂', body)

    def check_equations_are_never_silently_removed(self):
        document = Document(); _front_matter(document)
        p = document.add_paragraph('Satır içi denklem ')
        p._p.append(parse_xml(INLINE_EQUATION))
        p.add_run(' sonucu verir.')
        display = document.add_paragraph()
        display._p.append(parse_xml(DISPLAY_EQUATION))
        document.add_paragraph('Denklemden sonraki paragraf.')
        data = _import(document)
        body = _body(data)
        # Each equation leaves visible content in its own place, and the user is told.
        inline = body.split('Satır içi denklem', 1)[1].split('sonucu verir.', 1)[0]
        self.assertTrue(inline.strip(), 'satır içi denklem silindi')
        self.assertIn('E', inline)
        between = body.split('sonucu verir.', 1)[1].split('Denklemden sonraki paragraf.', 1)[0]
        self.assertTrue(between.strip(), 'ayrı satırdaki denklem silindi')
        self.assertTrue(any('denklem' in warning.lower() for warning in data['warnings']), data['warnings'])

    def check_footnote_text_and_reference_are_kept(self):
        document = Document(); _front_matter(document)
        _add_footnotes_part(document, FOOTNOTE_TEXT)
        p = document.add_paragraph('Dipnotlu cümle')
        p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                              '<w:footnoteReference w:id="1"/></w:r>'))
        p.add_run(' devam eder.')
        data = _import(document)
        body = _body(data)
        self.assertIn(FOOTNOTE_TEXT, json.dumps(data, ensure_ascii=False))
        # The reference keeps its position: something marks it between the two parts of the sentence.
        self.assertNotIn('Dipnotlu cümle devam eder.', body)
        self.assertIn('Dipnotlu cümle', body)
        self.assertTrue(any('dipnot' in warning.lower() for warning in data['warnings']), data['warnings'])

    def check_superscript_marker_after_year_keeps_citation_link(self):
        index = CitationIndex(['Yılmaz, A. (2020). Kitap. Yayınevi.'])
        self.assertEqual(len(list(index.spans('Yılmaz, 2020¹ bunu göstermiştir.'))), 1)


class KnownScientificImportGaps(ScientificImportChecks):
    """Documented 2.4.2 losses (teknik analiz Ç3). These must start failing once fixed."""

    def test_gap_footnotes_are_dropped(self):
        with self.assertRaises(AssertionError):
            self.check_footnote_text_and_reference_are_kept()


class ClosedScientificImportGaps(ScientificImportChecks):
    """Gaps closed in Phase 1; each was a `test_gap_*` test in Phase 0."""

    def test_superscript_marker_after_year_keeps_citation_link(self):  # Phase 1.1
        self.check_superscript_marker_after_year_keeps_citation_link()

    def test_hidden_text_is_not_imported(self):  # Phase 1.2
        self.check_hidden_text_is_not_imported()

    def test_superscript_and_subscript_are_preserved(self):  # Phase 1.3
        self.check_superscript_and_subscript_are_preserved()

    def test_equations_are_never_silently_removed(self):  # Phase 1.5
        self.check_equations_are_never_silently_removed()


class CitationEndTests(unittest.TestCase):
    """Phase 1.1: a footnote marker after the year ends a citation; nothing else changes."""

    REFERENCES = ['Kaya, B., & Demir, C. (2021). Çevrim içi görünürlük. Örnek Dergi, 5(1), 20–34.',
                  'Yılmaz, A. (2020). Kırsal kalkınma ve turizm. Örnek Yayınevi.']

    def _linked(self, text):
        return [text[start:end] for start, end, _ in CitationIndex(self.REFERENCES).spans(text)]

    def test_superscript_markers_after_year(self):
        self.assertEqual(self._linked('Yılmaz, 2020¹ bunu göstermiştir.'), ['Yılmaz, 2020'])
        self.assertEqual(self._linked('Yılmaz (2020)² ve (Kaya & Demir, 2021)³⁴.'),
                         ['Yılmaz (2020)', 'Kaya & Demir, 2021'])
        # ("… ve Yılmaz, 2020" is deliberately not linked: it may be a multi-author tail.)
        self.assertEqual(self._linked('Yılmaz, 2020⁰; Yılmaz, 2020⁹.'), ['Yılmaz, 2020', 'Yılmaz, 2020'])

    def test_unchanged_boundaries(self):
        # A different year, a year suffix or a longer number is still not this reference.
        self.assertEqual(self._linked('Yılmaz, 2020a ve Yılmaz, 20201 ve Yılmaz, 2020_1.'), [])
        self.assertEqual(self._linked('Yılmaz, 2020₁ alt simge değildir.'), [])
        self.assertEqual(self._linked('Yılmaz, 2020. Yılmaz, 2020; (Yılmaz, 2020)'),
                         ['Yılmaz, 2020', 'Yılmaz, 2020', 'Yılmaz, 2020'])

    def _article(self):
        data = sample_article()
        data['references'] = '\n'.join(self.REFERENCES)
        data['sections'] = [{'id': '0', 'name': 'Giriş', 'level': '1',
                             'content': 'Bu bulgu Yılmaz, 2020¹ tarafından desteklenir.'}]
        data['figtables'] = []
        return data

    def test_latex_links_citation_and_keeps_marker_outside_link(self):
        tex = generate_latex_from_form(self._article(), {}, {})
        self.assertIn(r'Bu bulgu \hyperlink{aiditor_ref_1}{Yılmaz, 2020}\textsuperscript{1} tarafından desteklenir.', tex)

    def test_word_links_citation_and_keeps_marker_outside_link(self):
        document = Document(io.BytesIO(generate_docx_from_form(self._article(), {}, {})))
        paragraph = next(p for p in document.paragraphs if p.text.startswith('Bu bulgu'))
        links = paragraph._p.findall(qn('w:hyperlink'))
        self.assertEqual([''.join(t.text for t in link.iter(qn('w:t'))) for link in links], ['Yılmaz, 2020'])
        self.assertEqual(links[0].get(qn('w:anchor')), 'aiditor_ref_1')
        self.assertEqual(paragraph.text, 'Bu bulgu Yılmaz, 2020¹ tarafından desteklenir.')


class ImportReportTests(unittest.TestCase):
    """Phase 1.0: the report adds nothing unless content was counted."""

    MESSAGES = {'hidden': '{count} gizli metin parçası aktarılmadı.',
                'equation': '{count} denklem dönüştürülemedi.'}

    def test_empty_report_adds_no_warnings_or_stats(self):
        report = _ImportReport(self.MESSAGES)
        self.assertEqual(report.warnings(), [])
        self.assertEqual(report.stats(), {})

    def test_counts_become_warnings_in_message_order(self):
        report = _ImportReport(self.MESSAGES)
        report.add('equation'); report.add('hidden', 2); report.add('equation')
        self.assertEqual(report.warnings(), ['2 gizli metin parçası aktarılmadı.', '2 denklem dönüştürülemedi.'])
        self.assertEqual(report.stats(), {'equation': 2, 'hidden': 2})

    def test_unknown_key_is_a_programming_error(self):
        with self.assertRaises(KeyError):
            _ImportReport(self.MESSAGES).add('dipnot')

    def test_import_without_scientific_content_has_no_report_entries(self):
        document = Document(); _front_matter(document)
        document.add_paragraph('Düz bir gövde paragrafı ölçüm sonuçlarını açıklamaktadır.')
        data = _import(document)
        self.assertEqual(set(data['stats']), {'sections', 'tables', 'figures', 'references'})
        self.assertFalse(any('gizli' in w.lower() or 'denklem' in w.lower() or 'dipnot' in w.lower()
                             for w in data['warnings']), data['warnings'])


class RunPropertyResolverTests(unittest.TestCase):
    """Phase 1.0: effective run properties follow direct formatting and styles."""

    def _check(self, document, run, name, expected):
        paragraph = run._parent
        self.assertIs(_RunPropertyResolver(document).is_on(run._r, paragraph._p, name), expected)

    def test_direct_formatting_is_absolute(self):
        document = Document()
        bold_style = document.styles.add_style('Kalın Paragraf', WD_STYLE_TYPE.PARAGRAPH)
        bold_style.font.bold = True
        p = document.add_paragraph(style=bold_style)
        on = p.add_run('açık'); on.bold = True
        off = p.add_run('kapalı'); off.bold = False
        plain = document.add_paragraph().add_run('düz')
        self._check(document, on, 'b', True)
        self._check(document, off, 'b', False)
        self._check(document, plain, 'b', False)

    def test_character_and_paragraph_styles_apply(self):
        document = Document()
        hidden = document.styles.add_style('Gizli Not', WD_STYLE_TYPE.CHARACTER)
        hidden.font.hidden = True
        italic_paragraph = document.styles.add_style('İtalik Paragraf', WD_STYLE_TYPE.PARAGRAPH)
        italic_paragraph.font.italic = True
        run = document.add_paragraph().add_run('gizli', style='Gizli Not')
        italic_run = document.add_paragraph(style=italic_paragraph).add_run('italik')
        self._check(document, run, 'vanish', True)
        self._check(document, italic_run, 'i', True)
        self._check(document, italic_run, 'vanish', False)

    def test_based_on_chain_is_inherited_and_nearest_definition_wins(self):
        document = Document()
        parent = document.styles.add_style('Ana Stil', WD_STYLE_TYPE.CHARACTER)
        parent.font.hidden = True
        child = document.styles.add_style('Alt Stil', WD_STYLE_TYPE.CHARACTER)
        child.base_style = parent
        visible_child = document.styles.add_style('Görünür Alt Stil', WD_STYLE_TYPE.CHARACTER)
        visible_child.base_style = parent
        visible_child.font.hidden = False
        p = document.add_paragraph()
        inherited = p.add_run('a', style='Alt Stil')
        overridden = p.add_run('b', style='Görünür Alt Stil')
        self._check(document, inherited, 'vanish', True)
        self._check(document, overridden, 'vanish', False)

    def test_toggle_property_in_paragraph_and_character_style_cancels(self):
        document = Document()
        paragraph_style = document.styles.add_style('Kalın P', WD_STYLE_TYPE.PARAGRAPH)
        paragraph_style.font.bold = True
        character_style = document.styles.add_style('Kalın K', WD_STYLE_TYPE.CHARACTER)
        character_style.font.bold = True
        run = document.add_paragraph(style=paragraph_style).add_run('x', style='Kalın K')
        self._check(document, run, 'b', False)

    def test_explicit_false_values_and_document_defaults(self):
        document = Document()
        p = document.add_paragraph()
        run = p.add_run('x')
        run._r.get_or_add_rPr().append(parse_xml(f'<w:vanish {nsdecls("w")} w:val="0"/>'))
        self._check(document, run, 'vanish', False)
        defaults = document.styles.element.find(
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docDefaults')
        defaults.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPrDefault/'
                      '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr').append(
            parse_xml(f'<w:i {nsdecls("w")}/>'))
        self._check(document, document.add_paragraph().add_run('y'), 'i', True)

    def test_style_cycle_does_not_hang(self):
        document = Document()
        first = document.styles.add_style('Döngü A', WD_STYLE_TYPE.CHARACTER)
        second = document.styles.add_style('Döngü B', WD_STYLE_TYPE.CHARACTER)
        first.base_style = second
        second.base_style = first
        run = document.add_paragraph().add_run('x', style='Döngü A')
        self._check(document, run, 'vanish', False)


if __name__ == '__main__':
    unittest.main()
