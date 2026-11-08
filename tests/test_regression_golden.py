"""Phase 0 regression baseline: current LaTeX, Word and import output must not drift.

The golden files under tests/golden record the exact output of version 2.4.2.
Any change to them must be intentional and reviewed. To regenerate after an
approved output change, run:

    AIDITOR_UPDATE_GOLDEN=1 python -m unittest discover -s tests -p test_regression_golden.py

(Windows PowerShell: $env:AIDITOR_UPDATE_GOLDEN="1"; python -m unittest discover -s tests -p test_regression_golden.py)
and review the resulting diff before committing.

Comparison ignores line-ending style and trailing spaces at line ends (Git and
editors may rewrite both; TeX treats them as insignificant). python-docx's
unchanged template parts are listed by name only, so a dependency upgrade does
not fail these tests unless the generated document parts change.
"""
import base64
import copy
import hashlib
import io
import json
import os
from pathlib import Path
import tempfile
import unittest
import uuid
import zipfile

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm

from docx_export import generate_docx_from_form
from formatter import extract_form_data_from_docx, generate_latex_from_form
from journal_templates import TEMPLATES, normalize_settings
from test_accounts import PNG, headers, make_setup, register
from test_citation_sources import content_control, simple_field, zotero_field, _fld, _run
from test_formatter_templates import sample_article

import app as application
from account_store import AccountStore

GOLDEN = Path(__file__).resolve().parent / 'golden'
UPDATE = os.environ.get('AIDITOR_UPDATE_GOLDEN') == '1'
# A 2×1 RGB PNG; fixed bytes keep figure sizes and import digests stable.
FIGURE_PNG = base64.b64decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAD0lEQVR4nGM8oaHBwMAAAAYmARoIqcpQAAAAAElFTkSuQmCC')
JOURNAL = {'journal_name_tr': 'Örnek Bilim Dergisi', 'journal_name_en': 'Sample Science Journal',
           'issn_print': '1234-5678', 'issn_online': '2345-6789', 'journal_url': 'https://journal.example.org',
           'footer_text': 'Bu makale CC BY 4.0 lisansı ile yayımlanmıştır.'}


def _normalize_text(value):
    # Git may convert line endings on Windows checkouts and strip trailing
    # spaces (apply.whitespace=fix, editor settings); content is what matters.
    return '\n'.join(line.rstrip() for line in value.replace('\r\n', '\n').split('\n'))


def assert_golden(test, name, actual):
    path = GOLDEN / name
    actual = _normalize_text(actual)
    if UPDATE:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(actual, encoding='utf-8', newline='\n')
        return
    test.assertTrue(path.exists(), f'Altın dosya eksik: {path}. AIDITOR_UPDATE_GOLDEN=1 ile oluşturup gözden geçirin.')
    expected = _normalize_text(path.read_text(encoding='utf-8'))
    if actual != expected:
        expected_lines, actual_lines = expected.splitlines(), actual.splitlines()
        first = next((i for i, (a, b) in enumerate(zip(expected_lines, actual_lines)) if a != b),
                     min(len(expected_lines), len(actual_lines)))
        test.fail(f'{name} değişti (ilk fark satır {first + 1}).\n'
                  f'  beklenen: {expected_lines[first] if first < len(expected_lines) else "<dosya sonu>"!r}\n'
                  f'  oluşan  : {actual_lines[first] if first < len(actual_lines) else "<dosya sonu>"!r}')


FULL_PARTS = ('word/document.xml', 'word/settings.xml')
# Styles touched by docx_export._setup; the rest of styles.xml is python-docx's template.
# styles.xml stores built-in names in lower case ('heading 1', 'caption').
SNAPSHOT_STYLES = {'normal', 'title', 'subtitle', 'heading 1', 'heading 2', 'heading 3', 'caption', 'header', 'footer'}


def _xml_lines(raw):
    return raw.decode('utf-8').replace('><', '>\n<')


def docx_snapshot(blob):
    """Readable, deterministic view of the Word parts that define layout and text.

    Generated parts (body, headers, footers, settings, relationships and the
    styles changed by the exporter) are recorded in full. Unchanged template
    parts come from python-docx's bundled template and are listed by name only.
    Media bytes depend on the installed Pillow encoder, so only their names and
    the XML that places them are recorded.
    """
    from lxml import etree
    lines = []
    with zipfile.ZipFile(io.BytesIO(blob)) as archive:
        names = sorted(archive.namelist())
        lines.append('### files')
        lines.extend(names)
        for name in names:
            raw = archive.read(name)
            if name.startswith('word/media/') or not name.startswith('word/'):
                continue
            if (name in FULL_PARTS or name.endswith('.rels')
                    or name.startswith(('word/header', 'word/footer'))):
                lines.append('### ' + name)
                lines.append(_xml_lines(raw))
            elif name == 'word/styles.xml':
                root = etree.fromstring(raw)
                w = root.nsmap['w']
                lines.append('### word/styles.xml (edited styles)')
                defaults = root.find(f'{{{w}}}docDefaults')
                if defaults is not None:
                    lines.append(_xml_lines(etree.tostring(defaults)))
                for style in root.findall(f'{{{w}}}style'):
                    style_name = style.find(f'{{{w}}}name')
                    if style_name is not None and style_name.get(f'{{{w}}}val', '').lower() in SNAPSHOT_STYLES:
                        lines.append(_xml_lines(etree.tostring(style)))
            # Other word/ parts (fontTable, numbering, theme, webSettings,
            # stylesWithEffects) are untouched template data: names only.
    core = Document(io.BytesIO(blob)).core_properties
    lines.append('### core properties')
    lines.extend([f'title={core.title}', f'author={core.author}', f'subject={core.subject}'])
    return '\n'.join(lines) + '\n'


def golden_article():
    """sample_article plus the content kinds most likely to be affected by import/output work."""
    data = sample_article()
    cover = data['cover']
    cover.update({'short_title': 'Tekrarlanabilir İş Akışları', 'author_short': 'Örnek & Smith',
                  'revisions': '01.05.2026\n15.06.2026'})
    data['authors'][0].update({'title': 'Doç. Dr.', 'orcid': '0000-0002-1825-0097'})
    data['sections'][0]['content'] += (
        '\n\nÖrnek (2026) ve (Smith, 2024) çalışmaları H₂O, CO₂, 5 m² ve 10⁻³ gibi değerleri raporlamıştır. '
        'Özel karakterler: & % $ # _ { } ~ ^ \\ ve Türkçe ÇĞİÖŞÜ çğıöşü; Yunanca α β; p ≤ 0,05.'
        '\n\nAyrıntılar için https://example.org/kaynak_(ek) adresine bakınız.')
    data['sections'].append({'id': 'method', 'name': '2.1. Yöntem / Method', 'level': '2',
                             'content': 'Alt bölüm metni.\nTek satır sonu korunur mu?'})
    data['figtables'].append({'type': 'figure', 'number': '1', 'file_key': 'fig1', 'section_id': 'introduction',
                              'section': 'Giriş / Introduction', 'tr_cap': 'Örnek şekil', 'en_cap': 'Sample figure',
                              'after_para': 'Editoryal çalışmalar, kaynak belgelerin düzenli ve izlenebilir biçimde işlenmesini gerektirir.'})
    data['figtables'].append({'type': 'table', 'number': '2', 'section_id': 'method', 'section': '2.1. Yöntem / Method',
                              'tr_cap': 'Eski tablo', 'en_cap': 'Legacy table',
                              'tbl_data': 'Değişken | Değer\nn | 120\nα | 0,05'})
    data['extra'] = {'ack': 'Destekleri için teşekkür ederiz.', 'contrib': 'Yazarlar eşit katkı sağlamıştır.',
                     'conflict': 'Çıkar çatışması yoktur.'}
    data['references'] = '\n'.join([
        'Örnek, D. (2026). Akademik yayıncılıkta iş akışları. Örnek Yayınları.',
        'Smith, A. (2024). Reproducible journals. Example Journal, 3(1), 1–10. https://doi.org/10.1234/ex.2024_1',
        'Çelik, B. (2020). Dergi editörlüğü. Ankara: Örnek.',
    ])
    return data


def golden_cases():
    cases = []
    for template in TEMPLATES:
        cases.append((template['id'], {**template['settings'], **JOURNAL}, golden_article()))
    classic = next(t['settings'] for t in TEMPLATES if t['id'] == 'classic')
    english = golden_article()
    cases.append(('classic_english_only', {**classic, **JOURNAL, 'english_only': True}, english))
    auto_end = golden_article()
    auto_end['cover']['end_page'] = ''
    cases.append(('classic_auto_end_page', {**classic, **JOURNAL, 'header_mode': 'odd_even', 'header_rule': 'line',
                                            'header_first_mode': 'custom', 'header_first_center': '{dergi} {sayfa_araligi}',
                                            'footer_first_mode': 'inherit', 'citation_position': 'below_authors',
                                            'footer_text_align': 'right', 'link_citations': False}, auto_end))
    return cases


class OutputGoldenTests(unittest.TestCase):
    """LaTeX and Word output for every template must stay byte-identical."""

    def test_latex_output_matches_golden(self):
        for name, settings, data in golden_cases():
            with self.subTest(case=name):
                tex = generate_latex_from_form(copy.deepcopy(data), {'fig1': ('fig_fig1.png', FIGURE_PNG)}, settings)
                assert_golden(self, f'latex/{name}.tex', tex)

    def test_docx_output_matches_golden(self):
        assets = {'logo': ('journal_logo.png', FIGURE_PNG), 'license': ('journal_license.png', FIGURE_PNG)}
        for name, settings, data in golden_cases():
            with self.subTest(case=name):
                blob = generate_docx_from_form(copy.deepcopy(data), {'fig1': ('fig_fig1.png', FIGURE_PNG)}, settings, assets)
                assert_golden(self, f'docx/{name}.txt', docx_snapshot(blob))

    def test_generation_is_repeatable(self):
        name, settings, data = golden_cases()[0]
        first = generate_latex_from_form(copy.deepcopy(data), {'fig1': ('fig_fig1.png', FIGURE_PNG)}, settings)
        second = generate_latex_from_form(copy.deepcopy(data), {'fig1': ('fig_fig1.png', FIGURE_PNG)}, settings)
        self.assertEqual(first, second)


# ── Word import baseline ─────────────────────────────────────────────────────

def endnote_field(paragraph, shown):
    """EndNote Cite While You Write: field data and a nested EN.CITE.DATA field."""
    def begin_with_data():
        run = OxmlElement('w:r')
        run.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin">'
                             '<w:fldData xml:space="preserve">PEVuZE5vdGU+PC9FbmROb3RlPg==</w:fldData></w:fldChar>'))
        return run
    code = ' ADDIN EN.CITE <EndNote><Cite><Author>Kaya</Author><Year>2021</Year><RecNum>12</RecNum></Cite></EndNote>'
    for node in (begin_with_data(), _run(code, 'w:instrText'),
                 begin_with_data(), _run(' ADDIN EN.CITE.DATA ', 'w:instrText'), _fld('end'),
                 _fld('separate'), _run(shown), _fld('end')):
        paragraph._p.append(node)


def _merged_table(document):
    table = document.add_table(rows=3, cols=3)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header = table.rows[0].cells
    header[0].text, header[1].text, header[2].text = 'Grup', 'Ölçüm', 'Not'
    for cell in header:
        for run in cell.paragraphs[0].runs:
            run.bold = True
    merged = table.cell(1, 0).merge(table.cell(2, 0))
    merged.text = 'A grubu'
    table.cell(1, 1).text, table.cell(1, 2).text = '12,5', 'ilk'
    table.cell(2, 1).text = '13,1'
    table.cell(2, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    wide = table.cell(2, 2)
    wide.text = 'son'
    shading = OxmlElement('w:shd'); shading.set(qn('w:fill'), 'D9E2F3'); shading.set(qn('w:val'), 'clear')
    wide._tc.get_or_add_tcPr().append(shading)
    return table


def import_fixture_bytes():
    """A representative manuscript without the Ç3 constructs (hidden text, sub/superscript runs,
    equations, footnotes). Phase 1 must keep its import result unchanged."""
    document = Document()
    document.add_paragraph('Araştırma Makalesi / Research Article')
    document.add_paragraph('Kırsal Turizmde Dijital Tanıtım ve Yerel Kalkınma', style='Title')
    document.add_paragraph('Digital Promotion and Local Development in Rural Tourism')
    document.add_paragraph('Ayşe Yılmaz, Örnek Üniversitesi, ayse@example.org')
    document.add_paragraph('Öz: Bu çalışma kırsal turizmde dijital tanıtım araçlarının yerel kalkınmaya etkisini '
                           'nitel bir yaklaşımla incelemektedir ve bulgular tartışılmaktadır.')
    document.add_paragraph('Anahtar Kelimeler: kırsal turizm; dijital tanıtım; kalkınma')
    document.add_paragraph('Abstract')
    document.add_paragraph('This study examines how digital promotion tools affect local development in rural '
                           'tourism using a qualitative approach and discusses the findings.')
    document.add_paragraph('Keywords: rural tourism; digital promotion; development')
    document.add_heading('1. Giriş', level=1)
    p = document.add_paragraph('Zotero alanı '); zotero_field(p, '(Yılmaz, 2020)'); p.add_run(' ile başlar.')
    p = document.add_paragraph('EndNote alanı '); endnote_field(p, '(Kaya & Demir, 2021)'); p.add_run(' ile sürer.')
    p = document.add_paragraph('Mendeley denetimi '); content_control(p, '(Smith vd., 2022)'); p.add_run(' ve basit alan ')
    simple_field(p, '(Yılmaz, 2020)'); p.add_run('.')
    p = document.add_paragraph('İzlenen değişiklik ')
    p._p.append(parse_xml(f'<w:del {nsdecls("w")} w:id="1" w:author="x"><w:r><w:delText>silinen </w:delText></w:r></w:del>'))
    p._p.append(parse_xml(f'<w:ins {nsdecls("w")} w:id="2" w:author="x"><w:r><w:t xml:space="preserve">eklenen metin</w:t></w:r></w:ins>'))
    p.add_run(' korunur. Özel karakterler & % $ # _ ~ ^ ve H₂O, m² Unicode olarak yazılmıştır.')
    document.add_paragraph('Birinci madde', style='List Bullet')
    document.add_paragraph('İkinci madde', style='List Bullet')
    document.add_heading('1.1. Kuramsal Çerçeve', level=2)
    document.add_paragraph('Bu bölümde kavramlar tanımlanmaktadır; İngilizce terimler korunur.')
    document.add_paragraph('Tablo 1. Grupların ölçüm değerleri')
    _merged_table(document)
    document.add_paragraph('Tablo verileri yukarıda özetlenmiştir.')
    document.add_paragraph().add_run().add_picture(io.BytesIO(FIGURE_PNG), width=Cm(4))
    document.add_paragraph('Şekil 1. Araştırma modeli')
    document.add_heading('2. Bulgular', level=1)
    document.add_paragraph('Bulgular (Kaya & Demir, 2021) ile uyumludur.')
    document.add_heading('Teşekkür', level=1)
    document.add_paragraph('Katılımcılara teşekkür ederiz.')
    document.add_heading('Çıkar Çatışması', level=1)
    document.add_paragraph('Yazarlar çıkar çatışması bildirmemiştir.')
    document.add_heading('Kaynakça', level=1)
    for reference in ['Kaya, B., & Demir, C. (2021). Çevrim içi görünürlük. Örnek Dergi, 5(1), 20–34.',
                      'Smith, J., Brown, L., & Green, P. (2022). Digital tools. Example Journal, 12(3), 101–118.',
                      'Yılmaz, A. (2020). Kırsal kalkınma ve turizm. Örnek Yayınevi. https://doi.org/10.1234/ky.2020']:
        document.add_paragraph(reference)
    stream = io.BytesIO(); document.save(stream)
    return stream.getvalue()


class ImportGoldenTests(unittest.TestCase):
    def test_representative_manuscript_import_matches_golden(self):
        data, images = extract_form_data_from_docx(import_fixture_bytes())
        snapshot = {
            'data': data,
            'images': [{'filename': image['filename'], 'mimetype': image['mimetype'],
                        'sha256': hashlib.sha256(image['bytes']).hexdigest()} for image in images],
        }
        assert_golden(self, 'import/representative_manuscript.json',
                      json.dumps(snapshot, ensure_ascii=False, indent=2, sort_keys=True) + '\n')

    def test_reference_manager_citations_are_kept_in_order(self):
        data, _ = extract_form_data_from_docx(import_fixture_bytes())
        content = data['sections'][0]['content']
        for shown in ('Zotero alanı (Yılmaz, 2020) ile başlar.', 'EndNote alanı (Kaya & Demir, 2021) ile sürer.',
                      'Mendeley denetimi (Smith vd., 2022) ve basit alan (Yılmaz, 2020).',
                      'İzlenen değişiklik eklenen metin korunur.'):
            self.assertIn(shown, content)
        self.assertNotIn('ADDIN', json.dumps(data, ensure_ascii=False))
        self.assertNotIn('silinen', content)


# ── Saved projects from earlier versions ─────────────────────────────────────

def legacy_project():
    """A version-1 project as saved by 2.x clients, including legacy fields."""
    figure = {'name': 'sekil.png', 'data': 'data:image/png;base64,' + base64.b64encode(FIGURE_PNG).decode()}
    return {
        'format': 'aiditor-project', 'version': 1,
        'data': {
            'cover': {'tr_title': 'Eski Proje Başlığı', 'en_title': 'Legacy Project Title', 'year': '2025',
                      'volume': '7', 'issue': '3', 'start_page': '120', 'end_page': '133',
                      'doi': 'https://doi.org/10.5555/legacy.2025.7', 'article_type': 'Derleme Makalesi -- Review Article',
                      'received': '01.02.2025', 'accepted': '03.04.2025', 'published': '05.06.2025',
                      'author_short': 'Eski & Yeni, 2025', 'short_title': '', 'revisions': '', 'editor': '',
                      'ethics': 'Etik kurul onayı gerekmemektedir.', 'title_note': '', 'first_page_fit': 'auto'},
            'abstract': {'tr_abs': 'Eski özet metni.', 'tr_kw': 'eski; proje', 'en_abs': 'Legacy abstract.', 'en_kw': 'legacy'},
            'authors': [{'title': 'Prof. Dr.', 'name': 'Ali Eski', 'affiliation': 'Eski Üniversitesi',
                         'email': 'ali@example.org', 'orcid': 'https://orcid.org/0000-0001-2345-6789', 'corresponding': True},
                        {'name': 'Ece Yeni', 'affiliation': 'Yeni Üniversitesi', 'email': '', 'orcid': ''}],
            'sections': [{'id': '0', 'name': 'Giriş', 'level': '1',
                          'content': 'Giriş paragrafı (Eski, 2020).\n\nİkinci paragraf.'},
                         {'id': '1', 'name': 'Sonuç', 'level': '1', 'content': 'Sonuç metni.'}],
            'figtables': [
                {'type': 'table', 'number': '1', 'section': 'Giriş', 'section_id': '0', 'tr_cap': 'Boru tablo',
                 'en_cap': 'Pipe table', 'tbl_data': 'A | B\n1 | 2', 'after_para': 'Giriş paragrafı (Eski, 2020).'},
                {'type': 'figure', 'number': '1', 'section': 'Sonuç', 'section_id': '1', 'tr_cap': 'Şekil',
                 'en_cap': 'Figure', 'file_key': 'legacy1', 'placement': 'section_start', 'after_para': ''},
            ],
            'extra': {'ack': '', 'contrib': 'Eşit katkı.', 'conflict': ''},
            'references': 'Eski, A. (2020). Eski kaynak. Eski Yayınevi.',
        },
        'figures': {'legacy1': figure},
        'ui': {'ack_enabled': False, 'ack_text': ''},
        # Older exports used `font` and `header_layout` without `template_id`.
        'journal_settings': {'font': 'Palatino', 'header_layout': 'centered', 'journal_name_tr': 'Eski Dergi'},
    }


class SavedProjectCompatibilityTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix='aiditor-golden-test-')
        self.addCleanup(self.temporary.cleanup)
        self.app, self.directory = make_setup(Path(self.temporary.name))

    def test_legacy_project_validates_and_round_trips_unchanged(self):
        draft = legacy_project()
        application.validate_draft_project(copy.deepcopy(draft))
        client = self.app.test_client()
        user = register(client)
        identifier = str(uuid.uuid4())
        saved = client.put('/api/articles/' + identifier, json={'project': draft, 'base_revision': 0}, headers=headers(user))
        self.assertEqual(saved.status_code, 200, saved.json)
        reopened = client.get('/api/articles/' + identifier, headers=headers(user)).json['article']['project']
        self.assertEqual(reopened, draft)
        self.assertEqual(AccountStore(self.directory).article(user['id'], identifier)['project'], draft)

    def test_legacy_project_generates_same_outputs_through_app(self):
        draft = legacy_project()
        client = self.app.test_client()
        user = register(client)
        settings = normalize_settings(draft['journal_settings'])
        form = {'data': json.dumps(draft['data'], ensure_ascii=False), 'journal_settings': json.dumps(settings),
                'fig_legacy1': (io.BytesIO(FIGURE_PNG), 'fig_legacy1.png')}
        response = client.post('/process_form', data=copy.copy(form), headers=headers(user), content_type='multipart/form-data')
        self.assertEqual(response.status_code, 200, response.json)
        with zipfile.ZipFile(io.BytesIO(client.get('/download/' + response.json['key']).data)) as archive:
            names = sorted(archive.namelist())
            tex = archive.read('main.tex').decode('utf-8')
            stored_settings = archive.read('journal_settings.json').decode('utf-8')
        assert_golden(self, 'project/legacy_project.tex', tex)
        assert_golden(self, 'project/legacy_project_zip.txt',
                      '\n'.join(names) + '\n### journal_settings.json\n' + stored_settings + '\n')
        form['fig_legacy1'] = (io.BytesIO(FIGURE_PNG), 'fig_legacy1.png')
        response = client.post('/process_docx', data=form, headers=headers(user), content_type='multipart/form-data')
        self.assertEqual(response.status_code, 200, response.json)
        blob = client.get('/download_docx/' + response.json['key']).data
        assert_golden(self, 'project/legacy_project_docx.txt', docx_snapshot(blob))


if __name__ == '__main__':
    unittest.main()
