"""Phase 1.3: Word sub/superscript (w:vertAlign) survives import.

A sub/superscript run becomes Unicode sub/superscript characters when every
character has one (digits, + - − = ( ), and n i for superscript): m², 10⁻³,
CO₂. Other script runs stay plain text, as before, and are counted. Spacing,
list punctuation and note markers (, ; * † ‡) are kept without a warning.
The effective value follows direct formatting, character and paragraph
styles, their basedOn chains and the document defaults. Hidden text
(Phase 1.2) is still removed before any conversion.
"""
import io
import unittest

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from citation_links import citation_report
from docx_export import generate_docx_from_form
from formatter import (_RunPropertyResolver, _paragraph_text, extract_form_data_from_docx,
                       generate_latex_from_form)
from test_formatter_templates import sample_article

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
REFERENCES = ['Kaya, B., & Demir, C. (2021). Çevrim içi görünürlük. Örnek Dergi, 5(1), 20–34.',
              'Yılmaz, A. (2020). Kırsal kalkınma ve turizm. Örnek Yayınevi.']


def _import(document):
    stream = io.BytesIO(); document.save(stream)
    return extract_form_data_from_docx(stream.getvalue())[0]


def _body(data):
    return '\n\n'.join(section['content'] for section in data['sections'])


def _front_matter(document):
    document.add_paragraph('Üst ve Alt Simge Aktarımı Deneme Başlığı', style='Title')
    document.add_paragraph('Öz: Bu çalışma üst ve alt simgeli ölçüm değerlerinin aktarımını örnek bir '
                           'metin ile değerlendirmek için hazırlanmıştır.')
    document.add_paragraph('Anahtar Kelimeler: ölçüm; simge')
    document.add_heading('Giriş', level=1)


def _vert_align(run, value):
    run._r.get_or_add_rPr().append(parse_xml(f'<w:vertAlign {nsdecls("w")} w:val="{value}"/>'))
    return run


def _scripted(paragraph, *pieces):
    """Add (text, 'sup'|'sub'|None) pieces as separate runs."""
    for text, kind in pieces:
        run = paragraph.add_run(text)
        if kind == 'sup':
            run.font.superscript = True
        elif kind == 'sub':
            run.font.subscript = True
    return paragraph


class ScientificNotationTests(unittest.TestCase):
    """The examples of the Phase 1.3 request, through direct formatting."""

    def setUp(self):
        document = Document(); _front_matter(document)
        _scripted(document.add_paragraph(),
                  ('Alan 25 m', None), ('2', 'sup'), (', hacim 4 cm', None), ('3', 'sup'),
                  (', yoğunluk 10', None), ('-3', 'sup'), (' ve ', None), ('x', None), ('n', 'sup'),
                  (' olarak ölçüldü.', None))
        _scripted(document.add_paragraph(),
                  ('Tepkimede CO', None), ('2', 'sub'), (' ve H', None), ('2', 'sub'), ('SO', None),
                  ('4', 'sub'), (' kullanıldı; SO', None), ('4', 'sub'), ('2', 'sup'), ('−', 'sup'),
                  (' iyonu oluştu.', None))
        self.data = _import(document)
        self.body = _body(self.data)

    def test_examples_become_unicode(self):
        for expected in ('m²', 'cm³', '10⁻³', 'xⁿ', 'CO₂', 'H₂SO₄', 'SO₄²⁻'):
            self.assertIn(expected, self.body)
        self.assertIn('Alan 25 m², hacim 4 cm³, yoğunluk 10⁻³ ve xⁿ olarak ölçüldü.', self.body)

    def test_convertible_scripts_add_no_warning(self):
        self.assertNotIn('unconverted_script', self.data['stats'])
        self.assertFalse(any('simge' in warning for warning in self.data['warnings']), self.data['warnings'])

    def test_latex_output_uses_textsuperscript_and_textsubscript(self):
        article = sample_article()
        article['sections'] = self.data['sections']
        article['figtables'] = []
        tex = generate_latex_from_form(article, {}, {})
        self.assertIn(r'25 m\textsuperscript{2}', tex)
        self.assertIn(r'4 cm\textsuperscript{3}', tex)
        self.assertIn(r'CO\textsubscript{2}', tex)
        self.assertIn(r'H\textsubscript{2}SO\textsubscript{4}', tex)

    def test_word_output_keeps_unicode(self):
        article = sample_article()
        article['sections'] = self.data['sections']
        article['figtables'] = []
        document = Document(io.BytesIO(generate_docx_from_form(article, {}, {})))
        text = '\n'.join(p.text for p in document.paragraphs)
        for expected in ('m²', 'cm³', '10⁻³', 'xⁿ', 'CO₂', 'H₂SO₄', 'SO₄²⁻'):
            self.assertIn(expected, text)


class ScriptSourcesTests(unittest.TestCase):
    """Direct formatting, character/paragraph styles, basedOn and defaults."""

    def _text(self, build):
        document = Document(); _front_matter(document)
        build(document)
        data = _import(document)
        return _body(data), data

    def test_character_style_and_based_on_chain(self):
        def build(document):
            base = document.styles.add_style('Üst Taban', WD_STYLE_TYPE.CHARACTER)
            base.font.superscript = True
            derived = document.styles.add_style('Birim Üssü', WD_STYLE_TYPE.CHARACTER)
            derived.base_style = base
            lower = document.styles.add_style('Kimyasal Alt', WD_STYLE_TYPE.CHARACTER)
            lower.font.subscript = True
            p = document.add_paragraph('Yüzey 3 m')
            p.add_run('2', style='Birim Üssü')
            p.add_run(' ve O')
            p.add_run('2', style='Kimyasal Alt')
            p.add_run(' ölçüldü.')
        body, _ = self._text(build)
        self.assertIn('Yüzey 3 m² ve O₂ ölçüldü.', body)

    def test_direct_baseline_overrides_style(self):
        def build(document):
            upper = document.styles.add_style('Üst Stil', WD_STYLE_TYPE.CHARACTER)
            upper.font.superscript = True
            p = document.add_paragraph('Sayfa 12')
            _vert_align(p.add_run('34', style='Üst Stil'), 'baseline')
            p.add_run(' numarasıdır.')
        body, _ = self._text(build)
        self.assertIn('Sayfa 1234 numarasıdır.', body)

    def test_paragraph_style_and_document_defaults(self):
        def build(document):
            style = document.styles.add_style('Üs Paragraf', WD_STYLE_TYPE.PARAGRAPH)
            style.font.superscript = True
            document.add_paragraph('Normal paragraf yeterince uzun bir cümledir.')
            p = document.add_paragraph(style=style)
            p.add_run('12')
            _vert_align(p.add_run('34'), 'baseline')
        body, _ = self._text(build)
        self.assertIn('¹²34', body)

    def test_position_raised_text_is_not_a_superscript(self):
        def build(document):
            p = document.add_paragraph('Yükseltilmiş 5')
            run = p.add_run('2')
            run._r.get_or_add_rPr().append(parse_xml(f'<w:position {nsdecls("w")} w:val="6"/>'))
            p.add_run(' metin.')
        body, _ = self._text(build)
        self.assertIn('Yükseltilmiş 52 metin.', body)

    def test_existing_unicode_and_turkish_text_is_unchanged(self):
        def build(document):
            p = document.add_paragraph('Değer m² ve CO₂; ')
            p.add_run('²').font.superscript = True
            p.add_run(' çğıöşü ÇĞİÖŞÜ α β ≤ 0,05.')
        body, data = self._text(build)
        self.assertIn('Değer m² ve CO₂; ² çğıöşü ÇĞİÖŞÜ α β ≤ 0,05.', body)
        self.assertNotIn('unconverted_script', data['stats'])


class UnconvertedScriptTests(unittest.TestCase):
    """Script runs without a Unicode form stay plain text and are reported."""

    def setUp(self):
        document = Document(); _front_matter(document)
        _scripted(document.add_paragraph(),
                  ('Bu 21', None), ('st', 'sup'), (' yüzyıl ve x', None), ('2', 'sup'), ('a', 'sup'),
                  (' ile v', None), ('max', 'sub'), (' değeri.', None))
        _scripted(document.add_paragraph(),
                  ('Yazar', None), ('1,2', 'sup'), (' ve not', None), ('*', 'sup'), (' ile ', None),
                  ('†', 'sup'), (' işaretleri.', None))
        _scripted(document.add_paragraph(), ('Karışık 5', None), ('2a', 'sup'), (' biçim.', None))
        self.data = _import(document)
        self.body = _body(self.data)

    def test_plain_text_is_kept(self):
        self.assertIn('Bu 21st yüzyıl ve x²a ile vmax değeri.', self.body)
        self.assertIn('Karışık 52a biçim.', self.body)  # never a misleading "²a"

    def test_neutral_marks_are_kept_without_warning(self):
        self.assertIn('Yazar¹,² ve not* ile † işaretleri.', self.body)

    def test_each_piece_is_counted_once_and_warned(self):
        # "st", "a" (a separate run after the converted "2"), "max", "2a".
        self.assertEqual(self.data['stats']['unconverted_script'], 4)
        self.assertIn('4 üst/alt simge parçası Unicode karşılığı olmadığı için düz metin olarak aktarıldı; '
                      'bu yerleri kontrol edin.', self.data['warnings'])

    def test_consecutive_unconverted_runs_are_one_piece(self):
        document = Document(); _front_matter(document)
        _scripted(document.add_paragraph(), ('Değer v', None), ('m', 'sub'), ('ax', 'sub'), (' olarak.', None))
        self.assertEqual(_import(document)['stats']['unconverted_script'], 1)


class HiddenAndScriptTests(unittest.TestCase):
    """Phase 1.2 behaviour is unchanged: hidden script runs disappear."""

    def test_hidden_superscript_is_removed_and_counted_as_hidden(self):
        document = Document(); _front_matter(document)
        p = document.add_paragraph('Alan 5 m')
        p.add_run('2').font.superscript = True
        hidden = p.add_run('99'); hidden.font.superscript = True; hidden.font.hidden = True
        hidden_word = p.add_run('xyz'); hidden_word.font.superscript = True; hidden_word.font.hidden = True
        p.add_run(' ölçüldü.')
        data = _import(document)
        self.assertIn('Alan 5 m² ölçüldü.', _body(data))
        self.assertNotIn('⁹', repr(data))
        self.assertNotIn('xyz', repr(data))
        self.assertEqual(data['stats']['hidden_text'], 1)
        self.assertNotIn('unconverted_script', data['stats'])


class ScriptInOtherFieldsTests(unittest.TestCase):
    """Tables and references use the same rule; citations stay linked."""

    def setUp(self):
        document = Document(); _front_matter(document)
        p = document.add_paragraph('Bu bulgu Yılmaz, 2020')
        p.add_run('1').font.superscript = True
        p.add_run(' ve (Kaya & Demir, 2021) tarafından desteklenir.')
        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text = 'Değişken'; table.cell(0, 1).text = 'Birim'
        table.cell(1, 0).text = 'Alan'
        _scripted(table.cell(1, 1).paragraphs[0], ('m', None), ('2', 'sup'))
        document.add_heading('Kaynakça', level=1)
        for reference in REFERENCES:
            document.add_paragraph(reference)
        self.data = _import(document)

    def test_table_cell_is_converted(self):
        cells = [cell['text'] for ft in self.data['figtables'] if ft['type'] == 'table'
                 for row in ft['tbl_model']['rows'] for cell in row]
        self.assertIn('m²', cells)

    def test_references_are_unchanged(self):
        self.assertEqual(self.data['references'].split('\n'), REFERENCES)

    def test_citation_before_superscript_marker_is_linked(self):
        body = _body(self.data)
        self.assertIn('Bu bulgu Yılmaz, 2020¹ ve (Kaya & Demir, 2021) tarafından desteklenir.', body)
        article = sample_article()
        article.update({key: self.data[key] for key in ('sections', 'references')})
        self.assertEqual(citation_report(article)['linked_in_body'], 2)
        article['figtables'] = []
        tex = generate_latex_from_form(article, {}, {})
        self.assertIn(r'\hyperlink{aiditor_ref_1}{Yılmaz, 2020}\textsuperscript{1}', tex)
        document = Document(io.BytesIO(generate_docx_from_form(article, {}, {})))
        paragraph = next(p for p in document.paragraphs if p.text.startswith('Bu bulgu'))
        self.assertEqual(paragraph.text, 'Bu bulgu Yılmaz, 2020¹ ve (Kaya & Demir, 2021) tarafından desteklenir.')
        linked = [''.join(t.text for t in link.iter(qn('w:t'))) for link in paragraph._p.findall(qn('w:hyperlink'))]
        self.assertIn('Yılmaz, 2020', linked)


class ScopeTests(unittest.TestCase):
    def test_rule_applies_only_during_word_import(self):
        document = Document()
        p = document.add_paragraph('m')
        p.add_run('2').font.superscript = True
        self.assertEqual(_paragraph_text(p), 'm2')

    def test_no_script_content_adds_no_report_entries(self):
        document = Document(); _front_matter(document)
        document.add_paragraph('Düz bir gövde paragrafı ölçüm sonuçlarını açıklamaktadır.')
        data = _import(document)
        self.assertEqual(set(data['stats']), {'sections', 'tables', 'figures', 'references'})


class RunPropertyValueTests(unittest.TestCase):
    """_RunPropertyResolver.value() for non-toggle properties such as vertAlign."""

    def _value(self, document, run):
        return _RunPropertyResolver(document).value(run._r, run._parent._p, 'vertAlign')

    def test_precedence(self):
        document = Document()
        char_style = document.styles.add_style('K Üst', WD_STYLE_TYPE.CHARACTER)
        char_style.font.superscript = True
        para_style = document.styles.add_style('P Alt', WD_STYLE_TYPE.PARAGRAPH)
        para_style.font.subscript = True
        p = document.add_paragraph(style=para_style)
        self.assertEqual(self._value(document, p.add_run('a')), 'subscript')
        self.assertEqual(self._value(document, p.add_run('b', style='K Üst')), 'superscript')
        self.assertEqual(self._value(document, _vert_align(p.add_run('c', style='K Üst'), 'baseline')), 'baseline')
        self.assertIsNone(self._value(document, document.add_paragraph().add_run('d')))

    def test_document_defaults_and_style_cycle(self):
        document = Document()
        first = document.styles.add_style('Döngü A', WD_STYLE_TYPE.CHARACTER)
        second = document.styles.add_style('Döngü B', WD_STYLE_TYPE.CHARACTER)
        first.base_style = second; second.base_style = first
        run = document.add_paragraph().add_run('x', style='Döngü A')
        self.assertIsNone(self._value(document, run))
        defaults = document.styles.element.find(f'{{{W}}}docDefaults/{{{W}}}rPrDefault/{{{W}}}rPr')
        defaults.append(parse_xml(f'<w:vertAlign {nsdecls("w")} w:val="superscript"/>'))
        self.assertEqual(self._value(document, run), 'superscript')

    def test_toggle_properties_are_unchanged(self):
        document = Document()
        run = document.add_paragraph().add_run('x')
        run.bold = True
        self.assertIs(_RunPropertyResolver(document).is_on(run._r, run._parent._p, 'b'), True)


if __name__ == '__main__':
    unittest.main()
