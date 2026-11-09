"""Phase 1.4c: content_rich survives saving, opening, exporting and generating.

A project whose sections carry content_rich is saved as ``version: 2``;
earlier releases accept only version 1 and refuse it instead of silently
dropping the formatting. Version-1 projects keep opening unchanged. The
server validates the structure and size of content_rich and rejects unknown
formatting versions; stale paragraphs are allowed (output ignores them).
"""
import copy
import io
import json
from pathlib import Path
import tempfile
import unittest
import uuid
import zipfile

from docx import Document
from docx.oxml.ns import qn

import app as application
from account_store import AccountStore
from test_accounts import headers, make_setup, register
from test_import_inline_formatting import PartialFormattingTests, _build
from test_regression_golden import legacy_project


def rich_project():
    project = legacy_project()
    project['version'] = 2
    sections = copy.deepcopy(_build(PartialFormattingTests.build)['sections'])
    sections[0]['id'] = '0'
    sections.append(copy.deepcopy(project['data']['sections'][1]))
    project['data']['sections'] = sections
    project['data']['figtables'][0]['after_para'] = 'Biçimsiz son paragraf burada yer alır.'
    return project


def _rich(**changes):
    rich = {'version': 1, 'paragraphs': [{'plain': 'Düz italik.', 'segments': [{'t': 'Düz '}, {'t': 'italik', 'i': True}, {'t': '.'}]}]}
    rich.update(changes)
    return rich


def _form(rich):
    return {'cover': {}, 'abstract': {}, 'extra': {}, 'authors': [], 'figtables': [], 'references': '',
            'sections': [{'id': '0', 'name': 'Giriş', 'level': '1', 'content': 'Düz italik.', 'content_rich': rich}]}


class ValidationTests(unittest.TestCase):
    def test_valid_content_rich_is_kept(self):
        data = _form(_rich())
        application.validate_form(data)
        self.assertEqual(data['sections'][0]['content_rich'], _rich())

    def test_stale_paragraphs_are_allowed(self):
        rich = _rich(paragraphs=[{'plain': 'Eski metin.', 'segments': [{'t': 'Eski ', 'b': True}, {'t': 'metin.'}]}])
        application.validate_form(_form(rich))

    def test_invalid_structures_are_rejected(self):
        bad = [
            None, 'x', [], {'paragraphs': []}, _rich(version=2), _rich(version='1'), _rich(version=True),
            _rich(extra=1), _rich(paragraphs='x'), _rich(paragraphs=[5]),
            _rich(paragraphs=[{'plain': 1, 'segments': []}]),
            _rich(paragraphs=[{'plain': 'a', 'segments': 'a'}]),
            _rich(paragraphs=[{'plain': 'a', 'segments': [{'t': 'a', 'u': True}]}]),
            _rich(paragraphs=[{'plain': 'a', 'segments': [{'t': 'a', 'b': 1}]}]),
            _rich(paragraphs=[{'plain': 'a', 'segments': [{'t': 1}]}]),
            _rich(paragraphs=[{'plain': 'a', 'segments': [], 'note': 'x'}]),
        ]
        for rich in bad:
            with self.subTest(rich=rich):
                with self.assertRaises(ValueError):
                    application.validate_form(_form(rich))

    def test_size_limits(self):
        limits = application.CONTENT_RICH_LIMITS
        too_many = _rich(paragraphs=[{'plain': 'a', 'segments': [{'t': 'a'}]}] * (limits['paragraphs'] + 1))
        too_long = _rich(paragraphs=[{'plain': 'a' * (limits['plain'] + 1), 'segments': [{'t': 'a'}]}])
        too_fragmented = _rich(paragraphs=[{'plain': 'a', 'segments': [{'t': 'a'}] * (limits['segments'] + 1)}])
        for rich in (too_many, too_long, too_fragmented):
            with self.assertRaises(ValueError):
                application.validate_form(_form(rich))

    def test_project_versions(self):
        for version in (1, 2):
            project = rich_project(); project['version'] = version
            application.validate_draft_project(project)
        for version in (0, 3, '2', True, None):
            project = rich_project(); project['version'] = version
            with self.subTest(version=version), self.assertRaises(ValueError):
                application.validate_draft_project(project)

    def test_legacy_version_one_project_still_validates(self):
        application.validate_draft_project(legacy_project())


class StorageAndOutputTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix='aiditor-rich-compat-')
        self.addCleanup(temporary.cleanup)
        self.app, self.directory = make_setup(Path(temporary.name))
        self.client = self.app.test_client()
        self.user = register(self.client)

    def test_library_save_and_reopen_is_lossless(self):
        project = rich_project()
        identifier = str(uuid.uuid4())
        saved = self.client.put('/api/articles/' + identifier, json={'project': project, 'base_revision': 0}, headers=headers(self.user))
        self.assertEqual(saved.status_code, 200, saved.json)
        reopened = self.client.get('/api/articles/' + identifier, headers=headers(self.user)).json['article']['project']
        self.assertEqual(reopened, project)
        self.assertEqual(AccountStore(self.directory).article(self.user['id'], identifier)['project'], project)
        # Second revision on top of the first keeps it too.
        again = self.client.put('/api/articles/' + identifier, json={'project': reopened, 'base_revision': saved.json['article']['revision']},
                                headers=headers(self.user))
        self.assertEqual(again.status_code, 200, again.json)
        self.assertEqual(self.client.get('/api/articles/' + identifier, headers=headers(self.user)).json['article']['project'], project)

    def test_project_file_validation_returns_content_rich(self):
        project = rich_project()
        response = self.client.post('/validate_form', data={'data': json.dumps(project['data'], ensure_ascii=False), 'draft': '1'},
                                    headers=headers(self.user))
        self.assertEqual(response.status_code, 200, response.json)
        self.assertEqual(response.json['data']['sections'][0]['content_rich'], project['data']['sections'][0]['content_rich'])

    def test_unknown_rich_version_is_refused_with_message(self):
        project = rich_project()
        project['data']['sections'][0]['content_rich']['version'] = 7
        identifier = str(uuid.uuid4())
        response = self.client.put('/api/articles/' + identifier, json={'project': project, 'base_revision': 0}, headers=headers(self.user))
        self.assertEqual(response.status_code, 400)
        self.assertIn('güncelleyin', response.json['error'])

    def test_import_response_carries_content_rich(self):
        document = Document()
        from test_import_inline_formatting import _front_matter, _references
        _front_matter(document); PartialFormattingTests.build(document); _references(document)
        stream = io.BytesIO(); document.save(stream)
        response = self.client.post('/import_docx', data={'article': (io.BytesIO(stream.getvalue()), 'makale.docx')},
                                    headers=headers(self.user), content_type='multipart/form-data')
        self.assertEqual(response.status_code, 200, response.json)
        rich = response.json['data']['sections'][0]['content_rich']
        self.assertEqual(rich['paragraphs'][0]['segments'][1], {'t': 'Homo sapiens', 'i': True})

    def _generate(self, project):
        form = {'data': json.dumps(project['data'], ensure_ascii=False), 'journal_settings': json.dumps({}),
                'fig_legacy1': (io.BytesIO(application_png()), 'fig_legacy1.png')}
        response = self.client.post('/process_form', data=dict(form), headers=headers(self.user), content_type='multipart/form-data')
        self.assertEqual(response.status_code, 200, response.json)
        with zipfile.ZipFile(io.BytesIO(self.client.get('/download/' + response.json['key'], headers=headers(self.user)).data)) as archive:
            tex = archive.read('main.tex').decode('utf-8')
        form['fig_legacy1'] = (io.BytesIO(application_png()), 'fig_legacy1.png')
        response = self.client.post('/process_docx', data=form, headers=headers(self.user), content_type='multipart/form-data')
        self.assertEqual(response.status_code, 200, response.json)
        return tex, Document(io.BytesIO(self.client.get('/download_docx/' + response.json['key'], headers=headers(self.user)).data))

    def test_generated_outputs_are_formatted(self):
        tex, document = self._generate(rich_project())
        self.assertIn(r'İnsan türü \textit{Homo sapiens} olarak bilinir ve bu bulgu \textbf{çok önemli} kabul edilir.', tex)
        self.assertIn(r'Kalın-italik \textbf{\textit{vurgu}}', tex)
        italic = [run.text for p in document.paragraphs for run in p.runs if run.italic and run.text.strip()]
        bold = [run.text for p in document.paragraphs for run in p.runs if run.bold and 'önemli' in run.text]
        self.assertIn('Homo sapiens', italic)
        self.assertEqual(bold, ['çok önemli'])
        # The table anchored to the unformatted last paragraph is still placed after it.
        self.assertLess(tex.index('Biçimsiz son paragraf'), tex.index(r'\begin{table'))

    def test_edited_paragraph_is_plain_and_others_keep_formatting(self):
        project = rich_project()
        section = project['data']['sections'][0]
        section['content'] = section['content'].replace('çok önemli kabul edilir.', 'önemli bulunmuştur.')
        tex, _ = self._generate(project)
        self.assertIn('İnsan türü Homo sapiens olarak bilinir ve bu bulgu önemli bulunmuştur.', tex)
        self.assertNotIn(r'\textit{Homo sapiens}', tex)
        self.assertIn(r'Kalın-italik \textbf{\textit{vurgu}}', tex)


def application_png():
    from test_regression_golden import FIGURE_PNG
    return FIGURE_PNG


class FrontendContractTests(unittest.TestCase):
    """The browser code that writes and reads project files (checked as source; behaviour in browser_rich_text.py)."""

    def setUp(self):
        root = Path(__file__).resolve().parent.parent
        self.projects = (root / 'static' / 'projects.js').read_text(encoding='utf-8')
        self.index = (root / 'templates' / 'index.html').read_text(encoding='utf-8')

    def test_version_two_only_with_formatting(self):
        self.assertIn("const version = data.sections.some(section => section.content_rich) ? 2 : 1;", self.projects)
        self.assertIn("![1, 2].includes(project.version)", self.projects)

    def test_sections_carry_content_rich(self):
        self.assertIn('div.contentRich = normalizeContentRich(d.content_rich);', self.index)
        self.assertIn('if (rich) section.content_rich = rich;', self.index)


if __name__ == '__main__':
    unittest.main()
