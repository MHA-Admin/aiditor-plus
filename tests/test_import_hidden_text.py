"""Phase 1.2: Word hidden text (w:vanish) is not imported and is counted.

Only ``w:vanish`` hides run text. ``w:webHidden`` (hidden in Web Layout,
e.g. TOC page numbers), ``w:specVanish`` and the paragraph mark's own
``w:pPr/w:rPr/w:vanish`` must not remove visible text.
"""
import io
import unittest

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from citation_links import citation_report
from formatter import _paragraph_text, _run_is_hidden, extract_form_data_from_docx, generate_latex_from_form
from test_citation_sources import content_control, simple_field, zotero_field
from test_formatter_templates import sample_article

REFERENCES = ['Kaya, B., & Demir, C. (2021). Çevrim içi görünürlük. Örnek Dergi, 5(1), 20–34.',
              'Smith, J., Brown, L., & Green, P. (2022). Digital tools. Example Journal, 12(3), 101–118.',
              'Yılmaz, A. (2020). Kırsal kalkınma ve turizm. Örnek Yayınevi.']


def _import(document):
    stream = io.BytesIO(); document.save(stream)
    return extract_form_data_from_docx(stream.getvalue())[0]


def _body(data):
    return '\n\n'.join(section['content'] for section in data['sections'])


def _hide(run, value=None):
    attribute = f' w:val="{value}"' if value is not None else ''
    run._r.get_or_add_rPr().append(parse_xml(f'<w:vanish {nsdecls("w")}{attribute}/>'))
    return run


def _front_matter(document):
    document.add_paragraph('Gizli Metin Aktarımı Deneme Başlığı', style='Title')
    document.add_paragraph('Öz: Bu çalışma gizli metnin aktarım dışında kalmasını örnek bir metin ile '
                           'değerlendirmek için hazırlanmıştır.')
    document.add_paragraph('Anahtar Kelimeler: gizli metin; aktarım')
    document.add_heading('Giriş', level=1)


class HiddenTextSourcesTests(unittest.TestCase):
    """Direct formatting, character/paragraph styles and basedOn chains."""

    def test_all_hiding_sources_are_removed_and_counted(self):
        document = Document(); _front_matter(document)
        char_style = document.styles.add_style('Gizli Karakter', WD_STYLE_TYPE.CHARACTER)
        char_style.font.hidden = True
        para_style = document.styles.add_style('Gizli Paragraf', WD_STYLE_TYPE.PARAGRAPH)
        para_style.font.hidden = True
        base = document.styles.add_style('Gizli Taban', WD_STYLE_TYPE.CHARACTER)
        base.font.hidden = True
        derived = document.styles.add_style('Türetilmiş Not', WD_STYLE_TYPE.CHARACTER)
        derived.base_style = base

        p = document.add_paragraph('Görünür başlangıç')
        _hide(p.add_run(' DOGRUDAN-GIZLI'))
        p.add_run(' ortası')
        p.add_run(' KARAKTER-STILI-GIZLI', style='Gizli Karakter')
        p.add_run(' ve')
        p.add_run(' ZINCIR-GIZLI', style='Türetilmiş Not')
        p.add_run(' sonu.')
        document.add_paragraph('PARAGRAF-STILI-GIZLI tüm paragraf.', style='Gizli Paragraf')
        document.add_paragraph('Son görünür paragraf ölçüm sonuçlarını açıklar.')

        data = _import(document)
        body = _body(data)
        self.assertIn('Görünür başlangıç ortası ve sonu.', body)
        self.assertIn('Son görünür paragraf ölçüm sonuçlarını açıklar.', body)
        for hidden in ('DOGRUDAN', 'KARAKTER-STILI', 'ZINCIR', 'PARAGRAF-STILI'):
            self.assertNotIn(hidden, body)
        self.assertEqual(data['stats']['hidden_text'], 4)
        self.assertIn("4 gizli metin parçası (Word'de gizli olarak biçimlendirilmiş) makale içeriğine aktarılmadı.",
                      data['warnings'])

    def test_consecutive_hidden_runs_are_one_piece(self):
        document = Document(); _front_matter(document)
        p = document.add_paragraph('Görünür metin')
        _hide(p.add_run(' gizli bir')); _hide(p.add_run(' gizli iki'))
        p.add_run(' devam.')
        data = _import(document)
        self.assertIn('Görünür metin devam.', _body(data))
        self.assertEqual(data['stats']['hidden_text'], 1)

    def test_explicit_visible_overrides_hidden_style(self):
        document = Document(); _front_matter(document)
        style = document.styles.add_style('Gizli Not', WD_STYLE_TYPE.CHARACTER)
        style.font.hidden = True
        p = document.add_paragraph('Başlangıç')
        _hide(p.add_run(' GORUNUR-DOGRUDAN', style='Gizli Not'), value='0')
        _hide(p.add_run(' GORUNUR-FALSE'), value='false')
        p.add_run(' son.')
        data = _import(document)
        self.assertIn('Başlangıç GORUNUR-DOGRUDAN GORUNUR-FALSE son.', _body(data))
        self.assertNotIn('hidden_text', data['stats'])

    def test_derived_style_can_unhide(self):
        document = Document(); _front_matter(document)
        base = document.styles.add_style('Gizli Taban', WD_STYLE_TYPE.CHARACTER)
        base.font.hidden = True
        derived = document.styles.add_style('Görünür Alt', WD_STYLE_TYPE.CHARACTER)
        derived.base_style = base
        derived.font.hidden = False
        p = document.add_paragraph('Başlangıç')
        p.add_run(' ALT-STIL-GORUNUR', style='Görünür Alt')
        p.add_run(' son.')
        self.assertIn('Başlangıç ALT-STIL-GORUNUR son.', _body(_import(document)))


class OtherVanishLikePropertiesTests(unittest.TestCase):
    """webHidden, specVanish and the paragraph mark are not hidden text."""

    def test_web_hidden_spec_vanish_and_paragraph_mark_keep_text(self):
        document = Document(); _front_matter(document)
        p = document.add_paragraph('Web görünümü')
        run = p.add_run(' WEBHIDDEN-METIN')
        run._r.get_or_add_rPr().append(parse_xml(f'<w:webHidden {nsdecls("w")}/>'))
        run = p.add_run(' SPECVANISH-METIN')
        run._r.get_or_add_rPr().append(parse_xml(f'<w:specVanish {nsdecls("w")}/>'))
        p.add_run(' son.')
        q = document.add_paragraph('Paragraf işareti gizli ama metin görünür.')
        q._p.get_or_add_pPr().append(parse_xml(f'<w:rPr {nsdecls("w")}><w:vanish/></w:rPr>'))
        data = _import(document)
        body = _body(data)
        self.assertIn('Web görünümü WEBHIDDEN-METIN SPECVANISH-METIN son.', body)
        self.assertIn('Paragraf işareti gizli ama metin görünür.', body)
        self.assertNotIn('hidden_text', data['stats'])
        self.assertFalse(any('gizli' in w for w in data['warnings']), data['warnings'])


class HiddenTextEverywhereTests(unittest.TestCase):
    """Title, abstract, keywords, headings, tables and references follow the same rule."""

    def setUp(self):
        document = Document()
        title = document.add_paragraph('Kırsal Turizmde Dijital Tanıtım', style='Title')
        _hide(title.add_run(' BASLIK-GIZLI'))
        p = document.add_paragraph('Öz: Bu çalışma kırsal turizmde dijital tanıtımın etkisini örnek bir metinle ele alır.')
        _hide(p.add_run(' OZ-GIZLI'))
        p = document.add_paragraph('Anahtar Kelimeler: turizm; tanıtım')
        _hide(p.add_run('; ANAHTAR-GIZLI'))
        heading = document.add_heading('Giriş', level=1)
        _hide(heading.add_run(' BOLUM-GIZLI'))
        _hide(document.add_heading('GIZLI-BASLIK', level=1).runs[0])
        document.add_paragraph('Giriş paragrafı yeterince uzun bir görünür cümledir.')
        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text = 'Değişken'; table.cell(0, 1).text = 'Değer'
        table.cell(1, 0).text = 'Ziyaretçi'
        cell = table.cell(1, 1).paragraphs[0]
        cell.add_run('120'); _hide(cell.add_run(' TABLO-GIZLI'))
        merged = document.add_table(rows=1, cols=3)
        merged_cell = merged.cell(0, 0).merge(merged.cell(0, 2)).paragraphs[0]
        merged_cell.add_run('Birleşik'); _hide(merged_cell.add_run(' BIRLESIK-GIZLI'))
        document.add_heading('Kaynakça', level=1)
        for reference in REFERENCES:
            r = document.add_paragraph(reference)
        _hide(r.add_run(' KAYNAK-GIZLI'))
        _hide(document.add_paragraph().add_run('Tamamen gizli kaynak (2099).'))
        self.data = _import(document)

    def test_hidden_text_is_absent_from_every_field(self):
        cover, abstract = self.data['cover'], self.data['abstract']
        self.assertEqual(cover['tr_title'], 'Kırsal Turizmde Dijital Tanıtım')
        self.assertNotIn('GIZLI', abstract['tr_abs'])
        self.assertTrue(abstract['tr_abs'].startswith('Bu çalışma kırsal turizmde'))
        self.assertEqual(abstract['tr_kw'], 'turizm; tanıtım')
        self.assertEqual([s['name'] for s in self.data['sections']], ['Giriş'])
        self.assertEqual(self.data['references'].split('\n'), REFERENCES)
        cells = [cell['text'] for ft in self.data['figtables'] if ft['type'] == 'table'
                 for row in ft['tbl_model']['rows'] for cell in row]
        self.assertIn('120', cells)
        self.assertIn('Birleşik', cells)
        everything = repr(self.data)
        self.assertNotIn('GIZLI', everything)
        self.assertNotIn('2099', everything)

    def test_each_piece_is_counted_once(self):
        # title, abstract, keywords, heading, hidden heading, table cell,
        # merged cell (counted once), reference, fully hidden reference.
        self.assertEqual(self.data['stats']['hidden_text'], 9)


class HiddenTextHeadingDetectionTests(unittest.TestCase):
    def test_hidden_regular_run_does_not_break_bold_heading(self):
        document = Document(); _front_matter(document)
        document.add_paragraph('Giriş paragrafı yeterince uzun bir görünür cümledir.')
        heading = document.add_paragraph()
        heading.add_run('Saha Gözlemleri').bold = True
        _hide(heading.add_run(' yazara not'))
        document.add_paragraph('Gözlem paragrafı yeterince uzun bir görünür cümledir.')
        data = _import(document)
        self.assertEqual([s['name'] for s in data['sections']], ['Giriş', 'Saha Gözlemleri'])


class ReferenceManagerCitationTests(unittest.TestCase):
    """Zotero, Mendeley and EndNote citations stay; hidden field results do not."""

    def setUp(self):
        document = Document(); _front_matter(document)
        p = document.add_paragraph('Zotero alanı '); zotero_field(p, '(Yılmaz, 2020)'); p.add_run('.')
        p = document.add_paragraph('EndNote alanı '); simple_field(p, '(Kaya ve Demir, 2021)'); p.add_run(' ile.')
        p = document.add_paragraph('Mendeley denetimi '); content_control(p, '(Smith vd., 2022)'); p.add_run(' çerçevesi.')
        p = document.add_paragraph('Gizli alan sonucu ')
        zotero_field(p, '(Yılmaz, 2099)')
        for run in p._p.findall(qn('w:r')):
            if any(t.text == '(Yılmaz, 2099)' for t in run.iter(qn('w:t'))):
                run.insert(0, parse_xml(f'<w:rPr {nsdecls("w")}><w:vanish/></w:rPr>'))
        p.add_run('burada.')
        p = document.add_paragraph('Gizli denetim ')
        content_control(p, '(Smith vd., 2099)')
        p._p.find(qn('w:sdt')).find(qn('w:sdtContent')).find(qn('w:r')).insert(
            0, parse_xml(f'<w:rPr {nsdecls("w")}><w:vanish/></w:rPr>'))
        p.add_run('burada.')
        document.add_heading('Kaynakça', level=1)
        for reference in REFERENCES:
            document.add_paragraph(reference)
        self.data = _import(document)
        self.body = _body(self.data)

    def test_visible_citations_are_kept_and_linked(self):
        for citation in ('(Yılmaz, 2020)', '(Kaya ve Demir, 2021)', '(Smith vd., 2022)'):
            self.assertIn(citation, self.body)
        self.assertNotIn('ZOTERO_ITEM', self.body)
        self.assertEqual(citation_report(self.data)['linked_in_body'], 3)
        data = sample_article()
        data.update({key: self.data[key] for key in ('sections', 'references')})
        data['figtables'] = []
        tex = generate_latex_from_form(data, {}, {})
        for linked in (r'\hyperlink{aiditor_ref_2}{Yılmaz, 2020}', r'\hyperlink{aiditor_ref_0}{Kaya ve Demir, 2021}',
                       r'\hyperlink{aiditor_ref_1}{Smith vd., 2022}'):
            self.assertIn(linked, tex)
        self.assertNotIn('2099', tex)

    def test_hidden_citation_results_are_removed(self):
        self.assertIn('Gizli alan sonucu burada.', self.body)
        self.assertIn('Gizli denetim burada.', self.body)
        self.assertNotIn('2099', self.body)
        self.assertEqual(self.data['stats']['hidden_text'], 2)


class HiddenTextScopeTests(unittest.TestCase):
    def test_rule_applies_only_during_word_import(self):
        document = Document()
        p = document.add_paragraph('Görünür')
        _hide(p.add_run(' gizli'))
        self.assertFalse(_run_is_hidden(p.runs[1]._r, p._p))
        self.assertEqual(_paragraph_text(p), 'Görünür gizli')
        _import(document)  # the import context must be reset afterwards
        self.assertEqual(_paragraph_text(p), 'Görünür gizli')

    def test_unicode_and_turkish_text_is_preserved(self):
        document = Document(); _front_matter(document)
        p = document.add_paragraph('Çğıöşü İĞÜŞÖÇ α≤β ±0,05 m² — “tırnak”')
        _hide(p.add_run(' gizli'))
        self.assertIn('Çğıöşü İĞÜŞÖÇ α≤β ±0,05 m² — “tırnak”', _body(_import(document)))


if __name__ == '__main__':
    unittest.main()
