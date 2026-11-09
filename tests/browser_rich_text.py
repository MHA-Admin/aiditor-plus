"""Phase 1.4: partial bold/italic from Word through the real editor, library, project file and outputs."""
import json
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from playwright.sync_api import sync_playwright, expect

from browser_workflow import QA, ROOT, Server, api, login, ready, saved

sys.path.insert(0, str(ROOT))  # the import tests below use the application modules
from test_import_inline_formatting import PartialFormattingTests, _front_matter, _references

RAW_MARKUP = ('**', '<b>', '<i>', '\\textit', '\\textbf', '"segments"', '{"t"')


def manuscript(path):
    document = Document()
    _front_matter(document)
    PartialFormattingTests.build(document)
    document.add_heading('Yöntem', level=1)
    document.add_paragraph('Biçimsiz yöntem paragrafı burada yer alır.')
    _references(document)
    document.save(path)


def note(page, index):
    return page.locator('#sections-list .sec-block').nth(index).locator('[data-rich-note]')


def generate_tex(page, data, name):
    page.locator('#btn-gen').click()
    page.wait_for_selector('#result-panel', state='visible')
    expect(page.locator('#dl-link')).to_be_visible()
    with page.expect_download() as download:
        page.locator('#dl-link').click()
    target = Path(data) / name
    download.value.save_as(target)
    with zipfile.ZipFile(target) as archive:
        return archive.read('main.tex').decode('utf-8')


def run():
    with tempfile.TemporaryDirectory(prefix='aiditor-rich-browser-') as data, sync_playwright() as p:
        server = Server(data); browser = None
        try:
            url = server.start(); browser = p.chromium.launch(headless=True)
            page = browser.new_page(viewport={'width': 1360, 'height': 1000})
            errors = []; page.on('pageerror', lambda e: errors.append(str(e)))
            page.on('dialog', lambda dialog: dialog.accept())
            login(page, url, 'rich_test', True, 'Biçim Test Dergisi'); ready(page)
            page.locator('#articles-tab').click()

            docfile = Path(data) / 'bicimli.docx'; manuscript(docfile)
            page.locator('#docx-file-input').set_input_files(docfile)
            page.wait_for_function('!docxImporting && collectSections().some(s => s.content.includes("Homo sapiens"))')
            saved(page)
            # The textarea keeps plain text; the note explains what is kept.
            content = page.locator('#sections-list .sec-block').first.locator('[data-f="content"]').input_value()
            assert 'Homo sapiens' in content and not any(raw in content for raw in RAW_MARKUP), content
            expect(note(page, 0)).to_be_visible()
            expect(note(page, 0)).to_contain_text('2 paragrafta korunuyor')
            expect(note(page, 1)).to_be_hidden()
            page.locator('#sections-list .sec-block').first.screenshot(path=str(QA / 'rich-text-note.png'))

            # Library autosave stores version 2 with content_rich.
            current = page.evaluate('articleLibrary.currentId')
            stored = api(page, '/api/articles/' + current)['article']['project']
            assert stored['version'] == 2, stored['version']
            assert stored['data']['sections'][0]['content_rich']['paragraphs'][0]['segments'][1] == {'t': 'Homo sapiens', 'i': True}
            assert 'content_rich' not in stored['data']['sections'][1]

            # Portable project file.
            with page.expect_download() as download:
                page.locator('button[onclick="saveProject()"]').click()
            project_path = Path(data) / 'bicimli.json'; download.value.save_as(project_path)
            exported = json.loads(project_path.read_text(encoding='utf-8'))
            assert exported['version'] == 2 and 'content_rich' in exported['data']['sections'][0]

            # Outputs carry the formatting; raw markup never appears as text.
            page.locator('#c-tr-title').fill('Satır İçi Biçim Aktarımı Deneme Başlığı'); saved(page)
            tex = generate_tex(page, data, 'bicimli.zip')
            assert r'\textit{Homo sapiens}' in tex and r'\textbf{çok önemli}' in tex and r'\textbf{\textit{vurgu}}' in tex, tex[-3000:]
            page.locator('#btn-docx').click()
            expect(page.locator('#dl-docx')).to_be_visible()
            with page.expect_download() as download:
                page.locator('#dl-docx').click()
            word_path = Path(data) / 'bicimli.docx.out.docx'; download.value.save_as(word_path)
            word = Document(word_path)
            italic = [r.text for para in word.paragraphs for r in para.runs if r.italic and r.text.strip()]
            assert 'Homo sapiens' in italic, italic
            assert not any(raw in para.text for para in word.paragraphs for raw in RAW_MARKUP)

            # Editing one paragraph invalidates only its formatting.
            page.evaluate("""() => {
                const field = document.querySelector('#sections-list .sec-block [data-f=content]');
                field.value = field.value.replace('çok önemli kabul edilir.', 'önemli bulunmuştur.');
                field.dispatchEvent(new Event('input', {bubbles: true}));
            }""")
            expect(note(page, 0)).to_have_attribute('data-state', 'partial')
            expect(note(page, 0)).to_contain_text('1 paragrafta korunuyor')
            expect(note(page, 0)).to_contain_text('1 paragraf değiştirildiği için')
            saved(page)
            tex = generate_tex(page, data, 'duzenlenmis.zip')
            assert 'İnsan türü Homo sapiens olarak bilinir ve bu bulgu önemli bulunmuştur.' in tex
            assert r'\textit{Homo sapiens}' not in tex and r'\textbf{\textit{vurgu}}' in tex
            stored = api(page, '/api/articles/' + current)['article']['project']
            plains = [e['plain'] for e in stored['data']['sections'][0]['content_rich']['paragraphs']]
            assert not any('çok önemli' in plain for plain in plains), plains  # stale entry dropped on save

            # Reopening the version-2 project file restores formatting and the note.
            page.locator('#project-file').set_input_files(project_path)
            page.wait_for_function("""() => collectSections().some(s => s.content.includes('çok önemli kabul edilir.') && s.content_rich)""")
            expect(note(page, 0)).to_contain_text('2 paragrafta korunuyor')
            saved(page)

            # Reload: the library article reopens with formatting.
            page.reload(); ready(page); page.locator('#articles-tab').click()
            page.wait_for_function('collectSections().some(s => s.content_rich)')
            expect(note(page, 0)).to_contain_text('2 paragrafta korunuyor')

            # An unknown future project version is refused, not half-loaded.
            future = dict(exported, version=3); future_path = Path(data) / 'gelecek.json'
            future_path.write_text(json.dumps(future, ensure_ascii=False), encoding='utf-8')
            page.locator('#project-file').set_input_files(future_path)
            expect(page.locator('.app-toast')).to_contain_text('geçerli bir AI-ditor Plus projesi değil')
            assert not errors, errors
            print('PASS: Word bold/italic import, editor note, edited-paragraph invalidation, library v2, project file v2 round trip, LaTeX/Word output, future-version refusal')
        finally:
            if browser: browser.close()
            server.stop()


if __name__ == '__main__':
    run()
