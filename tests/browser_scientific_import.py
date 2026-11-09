"""Phase 1 end to end: a scientific Word manuscript through the real editor to LaTeX ZIP and Word.

Partial bold/italic, Unicode sub/superscript, an inline and a display
equation, a footnote, an endnote and Zotero/EndNote/Mendeley citations go
through import (warnings shown), the library and a project file, and come out
in LaTeX and Word with their structure.
"""
import io
import json
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from playwright.sync_api import sync_playwright, expect

from browser_workflow import QA, ROOT, Server, api, login, ready, saved

sys.path.insert(0, str(ROOT))  # the fixture below uses the application modules
from scientific_fixture import manuscript


def download(page, button, link, target):
    page.locator(button).click()
    expect(page.locator(link)).to_be_visible()
    with page.expect_download() as downloaded:
        page.locator(link).click()
    downloaded.value.save_as(target)
    return target


def run():
    with tempfile.TemporaryDirectory(prefix='aiditor-scientific-browser-') as data, sync_playwright() as p:
        server = Server(data); browser = None
        try:
            url = server.start(); browser = p.chromium.launch(headless=True)
            page = browser.new_page(viewport={'width': 1360, 'height': 1000})
            errors = []; page.on('pageerror', lambda e: errors.append(str(e)))
            page.on('dialog', lambda dialog: dialog.accept())
            login(page, url, 'science_test', True, 'Bilim Test Dergisi'); ready(page)
            page.locator('#articles-tab').click()

            docfile = Path(data) / 'bilimsel.docx'; manuscript(docfile)
            page.locator('#docx-file-input').set_input_files(docfile)
            page.wait_for_function('!docxImporting && collectSections().some(s => s.content.includes("Enerji"))')
            saved(page)
            status = page.locator('#docx-import-status')
            for text in ('2 Word denklemi aktarıldı', '1 Word dipnotu aktarıldı', '1 Word son notu aktarıldı'):
                expect(status).to_contain_text(text)
            content = page.locator('#sections-list .sec-block').first.locator('[data-f="content"]').input_value()
            for expected in ('Homo sapiens', '(Yılmaz, 2020) ile desteklenir¹.', 'Enerji E=mc² ilişkisi ve CO₂ ile 10⁻³',
                             '(Kaya & Demir, 2021) ile karşılaştırıldıⁱ.', '(a+b)/2'):
                assert expected in content, (expected, content)
            expect(page.locator('#sections-list .sec-block').first.locator('[data-rich-note]')).to_contain_text('denklemler ve dipnotlar')
            page.locator('#docx-import-status').screenshot(path=str(QA / 'scientific-import-status.png'))

            current = page.evaluate('articleLibrary.currentId')
            stored = api(page, '/api/articles/' + current)['article']['project']
            assert stored['version'] == 2
            kinds = [key for e in stored['data']['sections'][0]['content_rich']['paragraphs'] for s in e['segments']
                     for key in ('math', 'note', 'b', 'i') if key in s]
            assert kinds.count('math') == 2 and kinds.count('note') == 2, kinds

            page.locator('#c-tr-title').fill('Bilimsel İçerik Uçtan Uca Deneme Başlığı'); saved(page)
            with zipfile.ZipFile(download(page, '#btn-gen', '#dl-link', Path(data) / 'bilimsel.zip')) as archive:
                tex = archive.read('main.tex').decode('utf-8')
            for expected in (r'\textit{Homo sapiens}', r'\textbf{önemli}', r'\footnote{Veriler 2024 saha çalışmasından alınmıştır.}',
                             r'\(E=m{c}^{2}\)', r'CO\textsubscript{2}', r'10\textsuperscript{\textminus}\textsuperscript{3}',
                             r'\[\frac{a+b}{2}\]', r'\section*{Son Notlar / Endnotes}', 'Ölçüm protokolü ekte verilmiştir.',
                             r'\hyperlink{aiditor_ref_1}{Yılmaz, 2020}', r'\hyperlink{aiditor_ref_0}{Kaya \& Demir, 2021}'):
                assert expected in tex, expected
            assert tex.count(r'\hyperlink{aiditor_ref_') == 4, tex.count(r'\hyperlink{aiditor_ref_')
            assert 'ADDIN' not in tex and 'EN.CITE' not in tex

            word_path = download(page, '#btn-docx', '#dl-docx', Path(data) / 'bilimsel-cikti.docx')
            with zipfile.ZipFile(word_path) as archive:
                body = archive.read('word/document.xml').decode('utf-8')
                assert 'word/footnotes.xml' in archive.namelist() and 'word/endnotes.xml' in archive.namelist()
                assert 'Veriler 2024 saha' in archive.read('word/footnotes.xml').decode('utf-8')
            assert body.count('<m:oMath>') + body.count('<m:oMath ') == 2 and '<m:oMathPara' in body
            assert 'CO₂' in body and '10⁻³' in body
            word = Document(word_path)
            assert any(r.italic and r.text == 'Homo sapiens' for para_ in word.paragraphs for r in para_.runs)

            # Portable project file keeps everything and reopens.
            with page.expect_download() as downloaded:
                page.locator('button[onclick="saveProject()"]').click()
            project_path = Path(data) / 'bilimsel.json'; downloaded.value.save_as(project_path)
            exported = json.loads(project_path.read_text(encoding='utf-8'))
            assert exported['version'] == 2
            page.locator('#project-file').set_input_files(project_path)
            page.wait_for_function("() => collectSections().some(s => s.content_rich && s.content.includes('Enerji'))")
            saved(page)
            assert not errors, errors
            print('PASS: scientific Word import (formatting, sub/superscript, equations, footnote, endnote, '
                  'Zotero/EndNote/Mendeley citations) → warnings, library v2, LaTeX ZIP, Word, project file round trip')
        finally:
            if browser: browser.close()
            server.stop()


if __name__ == '__main__':
    run()
