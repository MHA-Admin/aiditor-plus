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
from docx.oxml.ns import nsdecls

from citation_links import CitationIndex
from formatter import extract_form_data_from_docx

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

    def test_gap_hidden_text_is_imported(self):
        with self.assertRaises(AssertionError):
            self.check_hidden_text_is_not_imported()

    def test_gap_superscript_and_subscript_formatting_is_lost(self):
        with self.assertRaises(AssertionError):
            self.check_superscript_and_subscript_are_preserved()

    def test_gap_equations_are_silently_removed(self):
        with self.assertRaises(AssertionError):
            self.check_equations_are_never_silently_removed()

    def test_gap_footnotes_are_dropped(self):
        with self.assertRaises(AssertionError):
            self.check_footnote_text_and_reference_are_kept()

    def test_gap_superscript_marker_after_year_breaks_citation_link(self):
        with self.assertRaises(AssertionError):
            self.check_superscript_marker_after_year_keeps_citation_link()


if __name__ == '__main__':
    unittest.main()
