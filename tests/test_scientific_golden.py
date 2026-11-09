"""Phase 1 baseline for scientific content: import, LaTeX and Word output must not drift.

New golden files only (tests/golden/scientific/*); the Phase 0 golden files
are untouched and still checked by test_regression_golden.py.
"""
import io
import json
import unittest

from docx_export import generate_docx_from_form
from formatter import extract_form_data_from_docx, generate_latex_from_form
from scientific_fixture import manuscript
from test_formatter_templates import sample_article
from test_regression_golden import assert_golden, docx_snapshot


def scientific_import():
    stream = io.BytesIO(); manuscript().save(stream)
    return extract_form_data_from_docx(stream.getvalue())


def scientific_article(data):
    article = sample_article()
    article['sections'] = data['sections']
    article['figtables'] = []
    article['references'] = data['references']
    return article


class ScientificGoldenTests(unittest.TestCase):
    def setUp(self):
        self.data, self.images = scientific_import()

    def test_import_matches_golden(self):
        assert_golden(self, 'scientific/import.json',
                      json.dumps({'data': self.data, 'images': len(self.images)}, ensure_ascii=False, indent=2, sort_keys=True) + '\n')

    def test_latex_matches_golden(self):
        for template in ('classic', 'scholarly'):
            with self.subTest(template=template):
                assert_golden(self, f'scientific/{template}.tex',
                              generate_latex_from_form(scientific_article(self.data), {}, {'template_id': template}))

    def test_word_matches_golden(self):
        assert_golden(self, 'scientific/classic_docx.txt',
                      docx_snapshot(generate_docx_from_form(scientific_article(self.data), {}, {'template_id': 'classic'})))


if __name__ == '__main__':
    unittest.main()
