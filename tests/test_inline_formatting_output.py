"""Phase 1.4b: content_rich bold/italic in LaTeX/PDF and Word output.

Formatting is applied only to a paragraph whose current text equals a stored
``plain`` exactly. Edited paragraphs render plain; inserted, deleted, moved
and identical paragraphs never receive another paragraph's formatting.
Citation links (Zotero, EndNote, Mendeley text) wrap the formatted citation
text, Unicode sub/superscript keeps its LaTeX commands, LaTeX special
characters stay escaped and malformed ``content_rich`` is ignored.
"""
import copy
import io
import re
import unittest

from docx import Document
from docx.oxml.ns import qn

from docx_export import generate_docx_from_form
from formatter import CONTENT_RICH_VERSION, generate_latex_from_form, section_paragraph_flags
from test_formatter_templates import sample_article
from test_import_inline_formatting import REFERENCES, CitationTests, _build

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _segments(*pieces):
    """(text, flags) → segments; flags contains 'b' and/or 'i'."""
    segments = []
    for text, flags in pieces:
        segment = {'t': text}
        if 'b' in flags:
            segment['b'] = True
        if 'i' in flags:
            segment['i'] = True
        segments.append(segment)
    return segments


def _entry(*pieces):
    segments = _segments(*pieces)
    return {'plain': ''.join(s['t'] for s in segments), 'segments': segments}


def _section(entries, content=None, **extra):
    paragraphs = [entry['plain'] for entry in entries]
    section = {'id': '0', 'name': 'Giriş', 'level': '1',
               'content': content if content is not None else '\n\n'.join(paragraphs),
               'content_rich': {'version': CONTENT_RICH_VERSION, 'paragraphs': entries}}
    section.update(extra)
    return section


def _article(sections, references=REFERENCES, figtables=()):
    article = sample_article()
    article['sections'] = sections
    article['figtables'] = list(figtables)
    article['references'] = '\n'.join(references)
    return article


def _body_tex(article):
    tex = generate_latex_from_form(article, {}, {})
    start = tex.index(r'\section{Giriş}')
    return tex[start:tex.index(r'\begin{thebibliography}', start) if r'\begin{thebibliography}' in tex[start:] else len(tex)]


def _docx_body(article):
    document = Document(io.BytesIO(generate_docx_from_form(article, {}, {})))
    paragraphs = document.paragraphs
    start = next(i for i, p in enumerate(paragraphs) if p.text == 'Giriş')
    end = next((i for i, p in enumerate(paragraphs) if i > start and p.style.name.startswith('Heading')), len(paragraphs))
    return paragraphs[start + 1:end]


def _runs_of(paragraph):
    """(text, bold, italic, in_hyperlink) for every text run, including runs inside hyperlinks."""
    result = []
    for run in paragraph._p.iter(qn('w:r')):
        text = ''.join(t.text or '' for t in run.iter(qn('w:t')))
        if not text:
            continue
        r_pr = run.find(qn('w:rPr'))
        def on(tag):
            element = r_pr.find(qn('w:' + tag)) if r_pr is not None else None
            return element is not None and element.get(qn('w:val')) not in ('0', 'false')
        result.append((text, on('b'), on('i'), run.getparent().tag == qn('w:hyperlink')))
    return result


def _formatted_text(paragraph):
    """Collapse runs to (text, bold, italic) pieces, ignoring run boundaries of equal formatting."""
    pieces = []
    for text, bold, italic, _ in _runs_of(paragraph):
        if pieces and pieces[-1][1:] == (bold, italic):
            pieces[-1] = (pieces[-1][0] + text, bold, italic)
        else:
            pieces.append((text, bold, italic))
    return pieces


class LatexFormattingTests(unittest.TestCase):
    def test_bold_italic_and_bold_italic(self):
        tex = _body_tex(_article([_section([_entry(('Tür ', ''), ('Homo sapiens', 'i'), (' ve ', ''),
                                                   ('önemli', 'b'), (' ile ', ''), ('ikisi', 'bi'), ('.', ''))])]))
        self.assertIn(r'Tür \textit{Homo sapiens} ve \textbf{önemli} ile \textbf{\textit{ikisi}}.', tex)

    def test_special_characters_are_escaped_inside_formatting(self):
        tex = _body_tex(_article([_section([_entry(('Oran ', ''), ('%50 & $3_x #1 ~ ^ {ğİ} \\', 'i'), (' sonuç.', ''))])]))
        self.assertIn(r'Oran \textit{\%50 \& \$3\_x \#1 \textasciitilde{} \textasciicircum{} \{ğİ\} \textbackslash{}} sonuç.', tex)

    def test_scientific_notation_inside_formatting(self):
        tex = _body_tex(_article([_section([_entry(('Ölçümde ', ''), ('CO₂', 'b'), (', H₂SO₄ ve SO', ''),
                                                   ('₄²⁻', 'i'), (', 10⁻³, ', ''), ('xⁿ', 'i'), (' ve 4 cm³.', ''))])]))
        self.assertIn(r'Ölçümde \textbf{CO\textsubscript{2}}, H\textsubscript{2}SO\textsubscript{4} ve SO'
                      r'\textit{\textsubscript{4}\textsuperscript{2}\textsuperscript{\textminus}}, '
                      r'10\textsuperscript{\textminus}\textsuperscript{3}, \textit{x\textsuperscript{n}} '
                      r've 4 cm\textsuperscript{3}.', tex)

    def test_url_crossing_a_segment_boundary_keeps_one_target(self):
        url = 'https://doi.org/10.5555/abc.2021'
        tex = _body_tex(_article([_section([_entry(('Bkz. ', ''), (url[:20], 'i'), (url[20:], ''), (' adresi.', ''))])]))
        hrefs = re.findall(r'\\href\{([^}]*)\}', tex)
        self.assertEqual(hrefs, [url])
        self.assertIn(r'\textit{\href{' + url + '}', tex)

    def test_no_content_rich_is_byte_identical(self):
        section = _section([_entry(('Düz ', ''), ('italik', 'i'), (' metin.', ''))])
        plain = copy.deepcopy(section); del plain['content_rich']
        all_plain = copy.deepcopy(section)
        all_plain['content_rich']['paragraphs'] = [{'plain': 'Düz italik metin.', 'segments': [{'t': 'Düz italik metin.'}]}]
        reference = generate_latex_from_form(_article([plain]), {}, {})
        self.assertEqual(generate_latex_from_form(_article([all_plain]), {}, {}), reference)
        self.assertNotEqual(generate_latex_from_form(_article([section]), {}, {}), reference)


class CitationLinkTests(unittest.TestCase):
    """Imported Zotero, EndNote and Mendeley citations whose formatting crosses the citation text."""

    def setUp(self):
        self.data = _build(CitationTests.build)
        self.article = _article(copy.deepcopy(self.data['sections']))
        self.plain_article = _article([{k: v for k, v in s.items() if k != 'content_rich'} for s in self.data['sections']])

    def test_latex_links_match_plain_output(self):
        tex, plain = _body_tex(self.article), _body_tex(self.plain_article)
        self.assertEqual(re.findall(r'\\hyperlink\{(aiditor_ref_\d+)\}', tex), re.findall(r'\\hyperlink\{(aiditor_ref_\d+)\}', plain))
        self.assertEqual(len(re.findall(r'\\hyperlink', tex)), 4)
        self.assertIn(r'Zotero alanı \textbf{(}\hyperlink{aiditor_ref_1}{\textbf{\textit{Yılmaz}}, 2020}) ile \textit{başlar}.', tex)
        self.assertIn(r'\textit{EndNote alanı (}\hyperlink{aiditor_ref_0}{\textit{Kaya} \& Demir, 2021}) ile sürer.', tex)
        self.assertIn(r'Mendeley denetimi (\hyperlink{aiditor_ref_0}{Kaya \& \textbf{Demir}\textit{, 2021}}'
                      r'\textit{; }\hyperlink{aiditor_ref_1}{\textit{Yılmaz, 2020}}\textit{)} ile biter.', tex)

    def test_word_links_match_plain_output(self):
        rich, plain = _docx_body(self.article), _docx_body(self.plain_article)
        self.assertEqual([p.text for p in rich], [p.text for p in plain])
        anchors = lambda ps: [h.get(qn('w:anchor')) for p in ps for h in p._p.iter(qn('w:hyperlink'))]
        self.assertEqual(anchors(rich), anchors(plain))
        self.assertEqual(len(anchors(rich)), 4)
        linked = lambda p: ''.join(text for text, _, _, in_link in _runs_of(p) if in_link)
        self.assertEqual([linked(p) for p in rich], [linked(p) for p in plain])
        self.assertEqual(_formatted_text(rich[0]), [
            ('Zotero alanı ', False, False), ('(', True, False), ('Yılmaz', True, True),
            (', 2020) ile ', False, False), ('başlar', False, True), ('.', False, False)])


class WordFormattingTests(unittest.TestCase):
    def test_runs_carry_bold_and_italic(self):
        paragraphs = _docx_body(_article([_section([_entry(('Tür ', ''), ('Homo sapiens', 'i'), (' ve ', ''),
                                                           ('önemli', 'b'), (' ile ', ''), ('ikisi', 'bi'), ('.', ''))])]))
        self.assertEqual(_formatted_text(paragraphs[0]), [
            ('Tür ', False, False), ('Homo sapiens', False, True), (' ve ', False, False), ('önemli', True, False),
            (' ile ', False, False), ('ikisi', True, True), ('.', False, False)])

    def test_scientific_unicode_and_turkish_text_are_kept(self):
        text = 'Ölçümde CO₂, H₂SO₄, SO₄²⁻, 10⁻³, xⁿ ve 4 cm³; ğüşıöçİ & % $ _.'
        paragraphs = _docx_body(_article([_section([_entry(('Ölçümde ', ''), ('CO₂', 'b'), (', H₂SO₄, ', ''),
                                                           ('SO₄²⁻', 'i'), (', 10⁻³, xⁿ ve 4 cm³; ', ''),
                                                           ('ğüşıöçİ', 'bi'), (' & % $ _.', ''))])]))
        self.assertEqual(paragraphs[0].text, text)
        self.assertIn(('SO₄²⁻', False, True), _formatted_text(paragraphs[0]))

    def test_url_crossing_a_segment_boundary_keeps_one_link(self):
        url = 'https://doi.org/10.5555/abc.2021'
        paragraphs = _docx_body(_article([_section([_entry(('Bkz. ', ''), (url[:20], 'i'), (url[20:], ''), (' adresi.', ''))])]))
        links = list(paragraphs[0]._p.iter(qn('w:hyperlink')))
        self.assertEqual(len(links), 1)
        self.assertEqual(''.join(t.text for t in links[0].iter(qn('w:t'))), url)
        target = paragraphs[0].part.rels[links[0].get(qn('r:id'))].target_ref
        self.assertEqual(target, url)


class ParagraphMatchingTests(unittest.TestCase):
    """Formatting follows the paragraph text, never its position alone."""

    def setUp(self):
        self.a = _entry(('Birinci ', ''), ('alfa', 'i'), (' paragrafı.', ''))
        self.b = _entry(('İkinci ', ''), ('beta', 'b'), (' paragrafı.', ''))
        self.c = _entry(('Üçüncü ', ''), ('gama', 'bi'), (' paragrafı.', ''))
        self.entries = [self.a, self.b, self.c]

    def flags(self, paragraphs, entries=None):
        section = _section(copy.deepcopy(entries or self.entries), content='\n\n'.join(paragraphs))
        return [None if f is None else [(t, b, i) for t, b, i in _collapse(p, f)]
                for p, f in zip(paragraphs, section_paragraph_flags(section))]

    def test_unchanged(self):
        result = self.flags([self.a['plain'], self.b['plain'], self.c['plain']])
        self.assertEqual(result[0], [('Birinci ', False, False), ('alfa', False, True), (' paragrafı.', False, False)])
        self.assertTrue(all(result))

    def test_edited_paragraph_loses_only_its_formatting(self):
        result = self.flags([self.a['plain'], 'İkinci beta paragrafı değişti.', self.c['plain']])
        self.assertIsNotNone(result[0]); self.assertIsNone(result[1]); self.assertIsNotNone(result[2])
        self.assertIn(('gama', True, True), result[2])

    def test_insert_and_delete(self):
        result = self.flags(['Yeni ilk paragraf.', self.a['plain'], self.c['plain'], 'Yeni son paragraf.'])
        self.assertIsNone(result[0]); self.assertIsNone(result[3])
        self.assertIn(('alfa', False, True), result[1])
        self.assertIn(('gama', True, True), result[2])

    def test_reorder(self):
        result = self.flags([self.c['plain'], self.a['plain'], self.b['plain']])
        self.assertIn(('gama', True, True), result[0])
        self.assertIn(('alfa', False, True), result[1])
        self.assertIn(('beta', True, False), result[2])

    def test_identical_paragraphs_keep_their_own_formatting(self):
        plain = _entry(('Aynı metin paragrafı.', ''))
        formatted = _entry(('Aynı ', ''), ('metin', 'i'), (' paragrafı.', ''))
        entries = [plain, formatted, plain]
        text = 'Aynı metin paragrafı.'
        self.assertEqual([r is None for r in self.flags([text, text, text], entries)], [True, False, True])
        # The formatted copy is deleted: no remaining copy may take its formatting.
        self.assertEqual(self.flags([text, text], [plain, formatted, plain])[0], None)
        result = self.flags([text, text], entries)
        self.assertEqual(sum(r is not None for r in result), 1)

    def test_ambiguous_moved_duplicate_is_not_guessed(self):
        x = _entry(('Tekrar ', ''), ('eden', 'b'), (' paragraf.', ''))
        y = _entry(('Tekrar eden paragraf.', ''))
        # Two stored copies with different formatting; both moved after a new block.
        result = self.flags(['Yeni.', x['plain'], 'Ara.', x['plain']], [x, y])
        self.assertTrue(result[1] is None or result[3] is None)

    def test_content_with_extra_blank_lines_and_spaces(self):
        section = _section(copy.deepcopy(self.entries),
                           content='\n\n  ' + self.a['plain'] + '  \n \n\n' + self.b['plain'] + '\n\n\n' + self.c['plain'] + '\n')
        self.assertTrue(all(section_paragraph_flags(section)))


def _collapse(text, flags):
    pieces = []
    for char, (bold, italic) in zip(text, flags):
        if pieces and pieces[-1][1:] == (bold, italic):
            pieces[-1] = (pieces[-1][0] + char, bold, italic)
        else:
            pieces.append((char, bold, italic))
    return pieces


class MalformedContentRichTests(unittest.TestCase):
    def assert_ignored(self, content_rich):
        section = {'id': '0', 'name': 'Giriş', 'level': '1', 'content': 'Düz italik metin.', 'content_rich': content_rich}
        plain = {'id': '0', 'name': 'Giriş', 'level': '1', 'content': 'Düz italik metin.'}
        self.assertEqual(section_paragraph_flags(section), [None])
        self.assertEqual(generate_latex_from_form(_article([section]), {}, {}),
                         generate_latex_from_form(_article([plain]), {}, {}))

    def test_unknown_version(self):
        self.assert_ignored({'version': 99, 'paragraphs': [_entry(('Düz ', ''), ('italik', 'i'), (' metin.', ''))]})

    def test_segments_do_not_rebuild_plain(self):
        self.assert_ignored({'version': 1, 'paragraphs': [{'plain': 'Düz italik metin.',
                                                          'segments': [{'t': 'Düz '}, {'t': 'ITALIK', 'i': True}, {'t': ' metin.'}]}]})

    def test_wrong_types(self):
        for rich in (None, 'x', [], {'version': 1}, {'version': 1, 'paragraphs': 'x'},
                     {'version': 1, 'paragraphs': [{'plain': 'Düz italik metin.', 'segments': [{'t': 'Düz italik metin.', 'i': 'evet'}]}]},
                     {'version': 1, 'paragraphs': [{'plain': 'Düz italik metin.', 'segments': [{'t': 'Düz italik metin.', 'u': True}]}]},
                     {'version': 1, 'paragraphs': [{'plain': 'Düz italik metin.', 'segments': [{'t': 'Düz '}, {'t': ''}, {'t': 'italik metin.'}]}]},
                     {'version': 1, 'paragraphs': [5]}):
            with self.subTest(rich=rich):
                self.assert_ignored(rich)

    def test_raw_markup_in_text_is_escaped_not_interpreted(self):
        section = _section([_entry(('Ham ', ''), ('**kalın** <b>x</b> \\textbf{y}', 'i'), ('.', ''))])
        tex = _body_tex(_article([section]))
        self.assertIn(r'\textit{**kalın** <b>x</b> \textbackslash{}textbf\{y\}}', tex)


class FigurePlacementTests(unittest.TestCase):
    def test_after_para_placement_is_unchanged(self):
        figure = {'type': 'table', 'number': '1', 'section_id': '0', 'section': 'Giriş', 'tr_cap': 'Tablo', 'en_cap': 'Table',
                  'tbl_data': 'A | B\n1 | 2', 'after_para': 'İkinci beta paragrafı.'}
        entries = [_entry(('Birinci ', ''), ('alfa', 'i'), (' paragrafı.', '')), _entry(('İkinci ', ''), ('beta', 'b'), (' paragrafı.', '')),
                   _entry(('Üçüncü paragraf.', ''))]
        rich = _section(entries)
        plain = {k: v for k, v in rich.items() if k != 'content_rich'}
        for article in (_article([rich], figtables=[figure]), _article([plain], figtables=[figure])):
            tex = generate_latex_from_form(article, {}, {})
            self.assertLess(tex.index('beta'), tex.index(r'\begin{table'))
            self.assertLess(tex.index(r'\begin{table'), tex.index('Üçüncü paragraf.'))
            texts = [p.text for p in Document(io.BytesIO(generate_docx_from_form(article, {}, {}))).paragraphs]
            beta = texts.index('İkinci beta paragrafı.')
            self.assertLess(beta, texts.index('Üçüncü paragraf.'))
        tex_rich = _body_tex(_article([rich], figtables=[figure]))
        tex_plain = _body_tex(_article([plain], figtables=[figure]))
        changed = [(a, b) for a, b in zip(tex_rich.splitlines(), tex_plain.splitlines()) if a != b]
        self.assertEqual(len(tex_rich.splitlines()), len(tex_plain.splitlines()))
        self.assertEqual(changed, [(r'Birinci \textit{alfa} paragrafı.', 'Birinci alfa paragrafı.'),
                                   (r'İkinci \textbf{beta} paragrafı.', 'İkinci beta paragrafı.')])


if __name__ == '__main__':
    unittest.main()
