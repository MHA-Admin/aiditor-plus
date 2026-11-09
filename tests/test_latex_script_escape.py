"""Phase 1.3b: Unicode sub/superscript signs and letters survive LaTeX/PDF.

The journal fonts (TeX Gyre, Latin Modern) have no glyphs in the Unicode
"Superscripts and Subscripts" block (U+2070–U+209F) and Carlito only some,
so XeLaTeX printed "10⁻³" as "10³". ``escape()`` now writes these
characters with the \\textsuperscript/\\textsubscript commands it already
used for digits; minus signs become \\textminus (U+2212) so the sign keeps
its meaning. Word output is unchanged: it keeps the Unicode characters.
"""
import io
import re
import shutil
import subprocess
import tempfile
import unicodedata
import unittest
from pathlib import Path

from docx import Document

from citation_links import citation_report
from docx_export import generate_docx_from_form
from formatter import escape, extract_form_data_from_docx, generate_latex_from_form
from test_citation_sources import content_control, simple_field, zotero_field
from test_formatter_templates import sample_article

SUP, SUB = r'\textsuperscript', r'\textsubscript'
SCRIPT_BLOCK_RE = re.compile('[\u2070-\u209f]')
REFERENCES = ['Kaya, B., & Demir, C. (2021). Çevrim içi görünürlük. Örnek Dergi, 5(1), 20–34.',
              'Smith, J., Brown, L., & Green, P. (2022). Digital tools. Example Journal, 12(3), 101–118.',
              'Yılmaz, A. (2020). Kırsal kalkınma ve turizm. Örnek Yayınevi.']


class EscapeScriptCharacterTests(unittest.TestCase):

    def test_scientific_examples_keep_their_meaning(self):
        cases = {
            '10⁻³': '10' + SUP + r'{\textminus}' + SUP + '{3}',
            'CO₂': 'CO' + SUB + '{2}',
            'H₂SO₄': 'H' + SUB + '{2}SO' + SUB + '{4}',
            'SO₄²⁻': 'SO' + SUB + '{4}' + SUP + '{2}' + SUP + r'{\textminus}',
            'NH₄⁺': 'NH' + SUB + '{4}' + SUP + '{+}',
            'xⁿ': 'x' + SUP + '{n}',
            'cm³': 'cm' + SUP + '{3}',
            'm²': 'm' + SUP + '{2}',
            'yⁱ': 'y' + SUP + '{i}',
            '(n+1)⁽ⁿ⁾': '(n+1)' + SUP + '{(}' + SUP + '{n}' + SUP + '{)}',
            'x⁼': 'x' + SUP + '{=}',
            'K₍ₑ₎': 'K' + SUB + '{(}' + SUB + '{e}' + SUB + '{)}',
            'v₋₊₌': 'v' + SUB + r'{\textminus}' + SUB + '{+}' + SUB + '{=}',
            'vₘₐₓ': 'v' + SUB + '{m}' + SUB + '{a}' + SUB + '{x}',
            'Tₜ Iₒ Pₚ Sₛ Hₕ Kₖ Lₗ Nₙ eₔ': ('T' + SUB + '{t} I' + SUB + '{o} P' + SUB + '{p} S' + SUB + '{s} H'
                                           + SUB + '{h} K' + SUB + '{k} L' + SUB + '{l} N' + SUB + '{n} e'
                                           + SUB + '{ə}'),
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(escape(text), expected)

    def test_every_script_character_of_the_block_is_converted(self):
        # A character left in the output would be dropped by the journal fonts.
        script_characters = [chr(code) for code in range(0x2070, 0x20A0)
                             if 'SCRIPT' in unicodedata.name(chr(code), '')]
        self.assertEqual(len(script_characters), 42)
        for character in script_characters:
            with self.subTest(character=f'U+{ord(character):04X}'):
                self.assertIsNone(SCRIPT_BLOCK_RE.search(escape('a' + character + 'b')))

    def test_minus_signs_become_textminus_not_a_hyphen(self):
        self.assertEqual(escape('⁻'), SUP + r'{\textminus}')
        self.assertEqual(escape('₋'), SUB + r'{\textminus}')
        self.assertNotIn('{-}', escape('10⁻³ ₋'))

    def test_previous_behaviour_is_unchanged(self):
        self.assertEqual(escape('H₂ ve H₁₀'), r'H\textsubscript{2} ve H\textsubscript{1}\textsubscript{0}')
        self.assertEqual(escape('¹²³⁰'), SUP + '{1}' + SUP + '{2}' + SUP + '{3}' + SUP + '{0}')
        self.assertEqual(escape('& % $ # _ { } ~ ^ \\'),
                         r'\& \% \$ \# \_ \{ \} \textasciitilde{} \textasciicircum{} \textbackslash{}')
        self.assertEqual(escape(''), '')
        self.assertEqual(escape(None), '')

    def test_other_unicode_is_left_alone(self):
        # Baseline minus, Turkish, Greek, math signs and IPA/modifier letters
        # (linguistics: aspiration ʰ) are outside the block and keep their form.
        text = '− ÇĞİÖŞÜ çğıöşü α β µ ≤ ≥ ± × ° ʰ ʲ ᵃ ᵢ x^2 a_1'
        self.assertEqual(escape(text),
                         '− ÇĞİÖŞÜ çğıöşü α β µ ≤ ≥ ± × ° ʰ ʲ ᵃ ᵢ x\\textasciicircum{}2 a\\_1')


class ScriptOutputTests(unittest.TestCase):
    """End to end: Word import with reference-manager citations, then each output."""

    @classmethod
    def setUpClass(cls):
        document = Document()
        document.add_paragraph('Üst ve Alt Simgelerin LaTeX Çıktısı Deneme Başlığı', style='Title')
        document.add_paragraph('Öz: Bu çalışma bilimsel gösterimlerin PDF çıktısında korunmasını örnek bir '
                               'metin ile değerlendirmek için hazırlanmıştır.')
        document.add_paragraph('Anahtar Kelimeler: simge; çıktı')
        document.add_heading('Giriş', level=1)
        p = document.add_paragraph('Zotero ')
        zotero_field(p, '(Yılmaz, 2020)')
        p.add_run(' yoğunluğu 10')
        p.add_run('-3').font.superscript = True
        p.add_run(' g/cm')
        p.add_run('3').font.superscript = True
        p.add_run(' buldu; EndNote ')
        simple_field(p, '(Kaya ve Demir, 2021)')
        p.add_run(' SO')
        p.add_run('4').font.subscript = True
        p.add_run('2−').font.superscript = True
        p.add_run(' iyonunu, Mendeley ')
        content_control(p, '(Smith vd., 2022)')
        p.add_run(' x')
        p.add_run('n').font.superscript = True
        p.add_run(' ilişkisini raporladı.')
        document.add_heading('Kaynakça', level=1)
        for reference in REFERENCES:
            document.add_paragraph(reference)
        stream = io.BytesIO(); document.save(stream)
        data = extract_form_data_from_docx(stream.getvalue())[0]
        cls.article = sample_article()
        cls.article.update({key: data[key] for key in ('sections', 'references')})
        cls.article['figtables'] = []
        cls.body = '\n'.join(section['content'] for section in data['sections'])

    def test_import_keeps_unicode(self):
        for expected in ('10⁻³', 'cm³', 'SO₄²⁻', 'xⁿ'):
            self.assertIn(expected, self.body)

    def test_latex_converts_and_keeps_citation_links(self):
        tex = generate_latex_from_form(self.article, {}, {})
        self.assertIn('10' + SUP + r'{\textminus}' + SUP + '{3}', tex)
        self.assertIn('SO' + SUB + '{4}' + SUP + '{2}' + SUP + r'{\textminus}', tex)
        self.assertIn('x' + SUP + '{n}', tex)
        self.assertIsNone(SCRIPT_BLOCK_RE.search(tex))
        for anchor, citation in ((2, 'Yılmaz, 2020'), (0, 'Kaya ve Demir, 2021'), (1, 'Smith vd., 2022')):
            self.assertIn(r'\hyperlink{aiditor_ref_' + str(anchor) + '}{' + citation + '}', tex)
        self.assertEqual(citation_report(self.article)['linked_in_body'], 3)

    def test_word_output_keeps_unicode_and_links(self):
        document = Document(io.BytesIO(generate_docx_from_form(self.article, {}, {})))
        paragraph = next(p for p in document.paragraphs if p.text.startswith('Zotero'))
        for expected in ('10⁻³', 'cm³', 'SO₄²⁻', 'xⁿ'):
            self.assertIn(expected, paragraph.text)
        self.assertEqual(len(paragraph._p.findall(
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink')), 3)

    def test_titles_and_headings_use_the_same_conversion(self):
        article = dict(self.article)
        article['sections'] = [{'id': '0', 'name': 'CO₂ ve 10⁻³ Ölçümleri', 'level': '1', 'content': 'Metin.'}]
        tex = generate_latex_from_form(article, {}, {})
        self.assertIn(r'\section{CO' + SUB + '{2} ve 10' + SUP + r'{\textminus}' + SUP + '{3} Ölçümleri}', tex)


def _font_available(name):
    return bool(shutil.which('kpsewhich')) and bool(subprocess.run(
        ['kpsewhich', name], capture_output=True, text=True).stdout.strip())


@unittest.skipUnless(shutil.which('xelatex') and _font_available('texgyrepagella-regular.otf'),
                     'XeLaTeX ve TeX Gyre yazı tipleri yok')
class XeLaTeXGlyphTests(unittest.TestCase):
    """Runs only where XeLaTeX is installed: no glyph of the line is missing."""

    def test_journal_fonts_have_every_glyph_of_escaped_scripts(self):
        line = escape('10⁻³ cm³ CO₂ H₂SO₄ SO₄²⁻ NH₄⁺ xⁿ yⁱ K₍ₑ₎ vₘₐₓ Tₜ x⁼ v₌')
        fonts = [font for font in ('texgyrepagella-regular.otf', 'texgyretermes-regular.otf',
                                   'texgyreheros-regular.otf', 'lmroman10-regular.otf')
                 if _font_available(font)]
        for font in fonts:
            with self.subTest(font=font), tempfile.TemporaryDirectory() as directory:
                source = Path(directory) / 'glyphs.tex'
                source.write_text('\\documentclass{article}\\usepackage{fontspec}\\setmainfont{' + font + '}\n'
                                  '\\begin{document}\n' + line + '\n\\end{document}\n', encoding='utf-8')
                subprocess.run(['xelatex', '-interaction=nonstopmode', '-halt-on-error', source.name],
                               cwd=directory, capture_output=True, timeout=180)
                log = (Path(directory) / 'glyphs.log').read_text(encoding='utf-8', errors='replace')
                self.assertTrue((Path(directory) / 'glyphs.pdf').exists(), log[-2000:])
                self.assertNotIn('Missing character', log)


if __name__ == '__main__':
    unittest.main()
