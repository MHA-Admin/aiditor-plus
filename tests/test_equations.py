"""Phase 1.5: Word equations (OMML) are imported, kept and written to LaTeX/PDF and Word.

Import: each equation becomes its linear Unicode text in ``content`` and an
equation segment in ``content_rich`` that keeps the original OMML. LaTeX is
regenerated from the OMML through a whitelist (never read from stored data);
Word output writes the original OMML back. Unsupported equations fall back to
linear text with a warning; legacy OLE equations leave a placeholder and a
warning. Nothing is removed silently.
"""
import copy
import io
import json
import unittest
from xml.sax.saxutils import escape as xml_escape

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from lxml import etree

import app as application
from docx_export import generate_docx_from_form
from equations import M, omml_display_latex, omml_latex, omml_linear, safe_omml
from formatter import extract_form_data_from_docx, generate_latex_from_form, section_paragraph_attrs
from test_formatter_templates import sample_article

M_NS = f'xmlns:m="{M}"'


def r(text, normal=False):
    props = '<m:rPr><m:nor/></m:rPr>' if normal else ''
    return f'<m:r>{props}<m:t xml:space="preserve">{xml_escape(text)}</m:t></m:r>'


def omath(*parts):
    return f'<m:oMath {M_NS}>' + ''.join(parts) + '</m:oMath>'


def para(*maths):
    return f'<m:oMathPara {M_NS}>' + ''.join(x.replace(f' {M_NS}', '') for x in maths) + '</m:oMathPara>'


def node(xml):
    return etree.fromstring(xml)


FRAC = f'<m:f><m:num>{r("a+b")}</m:num><m:den>{r("2")}</m:den></m:f>'
EMC2 = r('E=m') + f'<m:sSup><m:e>{r("c")}</m:e><m:sup>{r("2")}</m:sup></m:sSup>'
SUM = (f'<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>{r("i=1")}</m:sub><m:sup>{r("n")}</m:sup>'
       f'<m:e><m:sSub><m:e>{r("x")}</m:e><m:sub>{r("i")}</m:sub></m:sSub></m:e></m:nary>')


class ConverterTests(unittest.TestCase):
    CASES = [
        ('fraction', FRAC, r'\frac{a+b}{2}', '(a+b)/2'),
        ('superscript', EMC2, r'E=m{c}^{2}', 'E=mc²'),
        ('subscript', f'<m:sSub><m:e>{r("H")}</m:e><m:sub>{r("2")}</m:sub></m:sSub>{r("O")}', r'{H}_{2}O', 'H₂O'),
        ('sub-superscript', f'<m:sSubSup><m:e>{r("x")}</m:e><m:sub>{r("i")}</m:sub><m:sup>{r("2")}</m:sup></m:sSubSup>',
         r'{x}_{i}^{2}', 'x_i²'),
        ('square root', f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{r("x")}</m:e></m:rad>',
         r'\sqrt{x}', '√x'),
        ('cube root', f'<m:rad><m:deg>{r("3")}</m:deg><m:e>{r("y")}</m:e></m:rad>', r'\sqrt[3]{y}', '∛y'),
        ('sum', SUM, r'\sum_{i=1}^{n}{{x}_{i}}', '∑_(i=1)ⁿ x_i'),
        ('integral', f'<m:nary><m:sub>{r("0")}</m:sub><m:sup>{r("1")}</m:sup><m:e>{r("f(x)dx")}</m:e></m:nary>',
         r'\int_{0}^{1}{f(x)dx}', '∫₀¹ f(x)dx'),
        ('brackets', f'<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>{r("a")}</m:e><m:e>{r("b")}</m:e></m:d>',
         r'\left[ a \middle| b \right]', '[a|b]'),
        ('absolute value', f'<m:d><m:dPr><m:begChr m:val="|"/><m:endChr m:val="|"/></m:dPr><m:e>{r("x")}</m:e></m:d>',
         r'\left| x \right|', '|x|'),
        ('accent', f'<m:acc><m:accPr><m:chr m:val="&#770;"/></m:accPr><m:e>{r("β")}</m:e></m:acc>', r'\hat{β}', 'β̂'),
        ('overbar', f'<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>{r("x")}</m:e></m:bar>', r'\overline{x}', 'x̅'),
        ('function', f'<m:func><m:fName>{r("sin")}</m:fName><m:e>{r("θ")}</m:e></m:func>', r'\sin {θ}', 'sin θ'),
        ('limit', f'<m:func><m:fName><m:limLow><m:e>{r("lim")}</m:e><m:lim>{r("n→∞")}</m:lim></m:limLow></m:fName>'
                  f'<m:e>{r("a")}</m:e></m:func>', r'\mathop{\lim }\limits_{n→∞}{a}', 'lim_(n→∞) a'),
        ('matrix', f'<m:d><m:e><m:m><m:mr><m:e>{r("1")}</m:e><m:e>{r("0")}</m:e></m:mr>'
                   f'<m:mr><m:e>{r("0")}</m:e><m:e>{r("1")}</m:e></m:mr></m:m></m:e></m:d>',
         r'\left( \begin{matrix}1 & 0 \\ 0 & 1\end{matrix} \right)', '([1 0; 0 1])'),
        ('normal text and special characters', r('p < 0,05 ve %5 & a_b', normal=True) + r('x'),
         r'\text{p < 0,05 ve \%5 \& a\_b}x', 'p < 0,05 ve %5 & a_bx'),
        ('decimal comma', r('0,05+1,5x'), '0{,}05+1{,}5x', '0,05+1,5x'),
        ('upright word', f'<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>max</m:t></m:r>', r'\mathrm{max}', 'max'),
        ('minus and greek', r('α−β≤γ'), 'α−β≤γ', 'α−β≤γ'),
    ]

    def test_supported_constructs(self):
        for name, body, latex, linear in self.CASES:
            with self.subTest(name):
                element = node(omath(body))
                self.assertEqual(omml_latex(element), latex)
                self.assertEqual(omml_linear(element), linear)

    def test_unsupported_constructs_return_none_but_keep_linear_text(self):
        for name, body in (
                ('unknown element', f'<m:foo>{r("x")}</m:foo>'),
                ('unknown n-ary operator', f'<m:nary><m:naryPr><m:chr m:val="⨁"/></m:naryPr><m:e>{r("x")}</m:e></m:nary>'),
                ('unknown accent', f'<m:acc><m:accPr><m:chr m:val="☆"/></m:accPr><m:e>{r("x")}</m:e></m:acc>'),
                ('unknown delimiter', f'<m:d><m:dPr><m:begChr m:val="«"/></m:dPr><m:e>{r("x")}</m:e></m:d>'),
                ('foreign element', f'<m:r><w:t {nsdecls("w")}>x</w:t></m:r>')):
            with self.subTest(name):
                element = node(omath(r('a+'), body))
                self.assertIsNone(omml_latex(element))
                self.assertTrue(omml_linear(element).startswith('a+'))

    def test_display_paragraph(self):
        self.assertEqual(omml_display_latex(node(para(omath(FRAC)))), r'\[\frac{a+b}{2}\]')
        self.assertEqual(omml_display_latex(node(para(omath(EMC2), omath(FRAC)))),
                         r'\begin{gather*}E=m{c}^{2} \\ \frac{a+b}{2}\end{gather*}')

    def test_tex_injection_in_run_text_is_escaped(self):
        latex = omml_latex(node(omath(r(r'\input{/etc/passwd}\write18{rm}'))))
        self.assertEqual(latex, r'\backslash input\{/etc/passwd\}\backslash write18\{rm\}')

    def test_safe_omml(self):
        self.assertIsNotNone(safe_omml(omath(EMC2)))
        for bad in ('', '<x/>', '<m:oMath', f'<w:p {nsdecls("w")}/>',
                    f'<m:oMath {M_NS} xmlns:v="urn:schemas-microsoft-com:vml"><v:shape/></m:oMath>',
                    f'<!DOCTYPE m [<!ENTITY e SYSTEM "file:///etc/passwd">]><m:oMath {M_NS}><m:r><m:t>&e;</m:t></m:r></m:oMath>',
                    omath(r('x' * 300000))):
            with self.subTest(bad=bad[:60]):
                self.assertIsNone(safe_omml(bad))


def _front_matter(document):
    document.add_paragraph('Denklem Aktarımı Deneme Başlığı', style='Title')
    document.add_paragraph('Öz: Bu çalışma Word denklemlerinin aktarımını örnek bir metin ile değerlendirmek için '
                           'hazırlanmıştır.')
    document.add_paragraph('Anahtar Kelimeler: denklem; aktarım')
    document.add_heading('Giriş', level=1)


def _import(document):
    stream = io.BytesIO(); document.save(stream)
    return extract_form_data_from_docx(stream.getvalue())


def _math_paragraph(document, before, xml, after):
    p = document.add_paragraph(before)
    p._p.append(parse_xml(xml))
    p.add_run(after)
    return p


class ImportTests(unittest.TestCase):
    def setUp(self):
        document = Document(); _front_matter(document)
        _math_paragraph(document, 'Enerji ', omath(EMC2), ' ilişkisi bilinir.')
        display = document.add_paragraph(); display._p.append(parse_xml(para(omath(FRAC))))
        upper = document.add_paragraph(); upper._p.append(parse_xml(para(omath(r('E=MC')))))
        _math_paragraph(document, 'Toplam ', omath(SUM), ' olarak yazılır.')
        _math_paragraph(document, 'Desteklenmeyen ', omath(f'<m:foo>{r("q")}</m:foo>'), ' ifade.')
        document.add_heading('Kaynakça', level=1)
        document.add_paragraph('Yılmaz, A. (2020). Kırsal kalkınma ve turizm. Örnek Yayınevi.')
        self.data, self.images = _import(document)
        self.section = self.data['sections'][0]

    def test_content_has_linear_text_in_place(self):
        self.assertEqual([s['name'] for s in self.data['sections']], ['Giriş'])
        self.assertEqual(self.section['content'], 'Enerji E=mc² ilişkisi bilinir.\n\n(a+b)/2\n\nE=MC\n\n'
                                                  'Toplam ∑_(i=1)ⁿ x_i olarak yazılır.\n\nDesteklenmeyen q ifade.')

    def test_content_rich_keeps_original_omml(self):
        paragraphs = self.section['content_rich']['paragraphs']
        inline = paragraphs[0]['segments'][1]
        self.assertEqual(inline['t'], 'E=mc²')
        self.assertFalse(inline['math']['display'])
        self.assertEqual(etree.tostring(node(inline['math']['omml']), method='c14n'),
                         etree.tostring(node(omath(EMC2)), method='c14n'))
        self.assertTrue(paragraphs[1]['segments'][0]['math']['display'])
        self.assertTrue(node(paragraphs[1]['segments'][0]['math']['omml']).tag.endswith('oMathPara'))
        self.assertNotIn('latex', json.dumps(self.section['content_rich']))
        self.assertNotIn('xmlns:w14', paragraphs[0]['segments'][1]['math']['omml'])

    def test_equation_paragraph_is_not_a_heading_or_title(self):
        self.assertEqual(self.data['cover']['tr_title'], 'Denklem Aktarımı Deneme Başlığı')
        self.assertNotIn('E=MC', [s['name'] for s in self.data['sections']])

    def test_warnings(self):
        self.assertEqual(self.data['stats']['equation'], 5)
        self.assertEqual(self.data['stats']['unconverted_equation'], 1)
        self.assertNotIn('equation_outside_body', self.data['stats'])
        self.assertTrue(any(w.startswith('5 Word denklemi aktarıldı') for w in self.data['warnings']), self.data['warnings'])
        self.assertTrue(any('1 Word denklemi LaTeX' in w for w in self.data['warnings']))

    def test_server_accepts_the_import(self):
        application.validate_form(copy.deepcopy(self.data))


class PlacementTests(unittest.TestCase):
    def test_equation_outside_body_is_linear_text_with_warning(self):
        document = Document()
        document.add_paragraph('Denklem Aktarımı Deneme Başlığı', style='Title')
        _math_paragraph(document, 'Öz: Bu çalışmada ', omath(EMC2), ' ilişkisi örnek bir metinle incelenmektedir.')
        document.add_paragraph('Anahtar Kelimeler: denklem')
        document.add_heading('Giriş', level=1)
        document.add_paragraph('Gövde paragrafı.')
        data, _ = _import(document)
        self.assertIn('E=mc²', data['abstract']['tr_abs'])
        self.assertEqual(data['stats']['equation_outside_body'], 1)

    def test_ole_equation_leaves_placeholder_and_no_figure(self):
        document = Document(); _front_matter(document)
        p = document.add_paragraph('Eski denklem ')
        p._p.append(parse_xml(
            f'<w:r {nsdecls("w")}><w:object><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="s1"/>'
            '<o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" Type="Embed" ProgID="Equation.3" '
            'ShapeID="s1" DrawAspect="Content" ObjectID="_1"/></w:object></w:r>'))
        p.add_run(' burada.')
        data, images = _import(document)
        self.assertIn('Eski denklem [Denklem: Word nesnesi dönüştürülemedi] burada.', data['sections'][0]['content'])
        self.assertEqual(data['stats']['ole_equation'], 1)
        self.assertEqual(images, [])
        self.assertEqual(data['figtables'], [])

    def test_formatting_around_equation(self):
        document = Document(); _front_matter(document)
        p = document.add_paragraph()
        p.add_run('Kalın ').bold = True
        p._p.append(parse_xml(omath(EMC2)))
        p.add_run(' ve ').bold = True
        p.add_run('italik').italic = True
        data, _ = _import(document)
        segments = data['sections'][0]['content_rich']['paragraphs'][0]['segments']
        self.assertEqual([(s['t'], s.get('b'), s.get('i'), 'math' in s) for s in segments],
                         [('Kalın', True, None, False), (' ', None, None, False), ('E=mc²', None, None, True),
                          (' ', None, None, False), ('ve', True, None, False), (' ', None, None, False),
                          ('italik', None, True, False)])


def _article(sections):
    article = sample_article()
    article['sections'] = sections
    article['figtables'] = []
    article['references'] = 'Yılmaz, A. (2020). Kırsal kalkınma ve turizm. Örnek Yayınevi.'
    return article


class OutputTests(unittest.TestCase):
    def setUp(self):
        document = Document(); _front_matter(document)
        _math_paragraph(document, 'Enerji ', omath(EMC2), ' ilişkisi (Yılmaz, 2020) bilinir.')
        display = document.add_paragraph(); display._p.append(parse_xml(para(omath(FRAC))))
        _math_paragraph(document, 'Desteklenmeyen ', omath(f'<m:foo>{r("q_1 & %")}</m:foo>'), ' ifade.')
        document.add_heading('Kaynakça', level=1)
        document.add_paragraph('Yılmaz, A. (2020). Kırsal kalkınma ve turizm. Örnek Yayınevi.')
        self.data, _ = _import(document)

    def test_latex(self):
        tex = generate_latex_from_form(_article(self.data['sections']), {}, {})
        self.assertIn(r'Enerji \(E=m{c}^{2}\) ilişkisi (\hyperlink{aiditor_ref_0}{Yılmaz, 2020}) bilinir.', tex)
        self.assertIn('\n' + r'\[\frac{a+b}{2}\]' + '\n', tex)
        self.assertIn(r'Desteklenmeyen q\_1 \& \% ifade.', tex)

    def test_word_keeps_original_equation(self):
        document = Document(io.BytesIO(generate_docx_from_form(_article(self.data['sections']), {}, {})))
        maths = document.element.body.findall('.//' + qn('m:oMath'))
        paras = document.element.body.findall('.//' + qn('m:oMathPara'))
        self.assertEqual(len(paras), 1)
        self.assertEqual(len(maths), 3)
        stored = node(self.data['sections'][0]['content_rich']['paragraphs'][0]['segments'][1]['math']['omml'])
        self.assertEqual(etree.tostring(maths[0], method='c14n', exclusive=True),
                         etree.tostring(stored, method='c14n', exclusive=True))
        texts = [p.text for p in document.paragraphs]
        self.assertIn('Enerji  ilişkisi (Yılmaz, 2020) bilinir.', texts)  # python-docx skips math text
        anchors = [h.get(qn('w:anchor')) for h in document.element.body.iter(qn('w:hyperlink'))]
        self.assertIn('aiditor_ref_0', anchors)

    def test_word_output_reimports_with_the_same_equations(self):
        blob = generate_docx_from_form(_article(self.data['sections']), {}, {})
        again, _ = extract_form_data_from_docx(blob)
        body = '\n\n'.join(section['content'] for section in again['sections'])
        for paragraph in self.data['sections'][0]['content'].split('\n\n'):
            self.assertIn(paragraph, body)
        stored = [s['math']['omml'] for e in self.data['sections'][0]['content_rich']['paragraphs'] for s in e['segments'] if 'math' in s]
        found = [s['math']['omml'] for sec in again['sections'] for e in sec.get('content_rich', {}).get('paragraphs', [])
                 for s in e['segments'] if 'math' in s]
        canonical = lambda xml: etree.tostring(node(xml), method='c14n', exclusive=True)
        self.assertEqual([canonical(x) for x in found], [canonical(x) for x in stored])

    def test_edited_paragraph_keeps_unambiguous_equation_only(self):
        sections = copy.deepcopy(self.data['sections'])
        sections[0]['content'] = sections[0]['content'].replace('ilişkisi (Yılmaz, 2020) bilinir.', 'ilişkisi iyi bilinir.')
        tex = generate_latex_from_form(_article(sections), {}, {})
        self.assertIn(r'Enerji \(E=m{c}^{2}\) ilişkisi iyi bilinir.', tex)
        # Text changed right next to the equation: it is not guessed.
        sections[0]['content'] = sections[0]['content'].replace('Enerji E=mc² ilişkisi', 'Enerji: E=mc²; ilişkisi')
        tex = generate_latex_from_form(_article(sections), {}, {})
        self.assertIn(r'Enerji: E=mc\textsuperscript{2}; ilişkisi iyi bilinir.', tex)
        self.assertIsNone(section_paragraph_attrs(sections[0])[0])

    def test_tampered_omml_falls_back_to_linear_text(self):
        sections = copy.deepcopy(self.data['sections'])
        segment = sections[0]['content_rich']['paragraphs'][0]['segments'][1]
        segment['math']['omml'] = '<m:oMath xmlns:m="%s" xmlns:x="urn:x"><x:evil/></m:oMath>' % M
        tex = generate_latex_from_form(_article(sections), {}, {})
        self.assertIn(r'Enerji E=mc\textsuperscript{2} ilişkisi', tex)
        document = Document(io.BytesIO(generate_docx_from_form(_article(sections), {}, {})))
        self.assertNotIn('urn:x', document.element.xml)
        with self.assertRaises(ValueError):
            application.validate_form({'sections': sections})

    def test_stored_latex_is_never_used(self):
        sections = copy.deepcopy(self.data['sections'])
        sections[0]['content_rich']['paragraphs'][0]['segments'][1]['math']['latex'] = r'\input{/etc/passwd}'
        tex = generate_latex_from_form(_article(sections), {}, {})
        self.assertNotIn('passwd', tex)
        with self.assertRaises(ValueError):
            application.validate_form({'sections': sections})


if __name__ == '__main__':
    unittest.main()
