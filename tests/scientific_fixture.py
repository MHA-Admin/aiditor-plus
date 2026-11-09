"""A scientific Word manuscript for end-to-end and golden tests (Phase 1).

Partial bold/italic, Unicode sub/superscript, an inline and a display
equation, a footnote, an endnote and Zotero, Mendeley (content control),
simple-field and EndNote citations. Deterministic: same bytes in, same
import result out.
"""
from docx import Document
from docx.oxml import parse_xml

from test_citation_sources import content_control, simple_field, zotero_field
from test_equations import EMC2, FRAC, omath, para
from test_import_inline_formatting import _references, _runs
from test_notes import _notes_part, ref, t
from test_regression_golden import endnote_field


def manuscript(path=None):
    document = Document()
    _notes_part(document, 'footnote', {'1': [t(' Veriler 2024 saha çalışmasından alınmıştır.')]})
    _notes_part(document, 'endnote', {'1': [t(' Ölçüm protokolü ekte verilmiştir.')]})
    document.add_paragraph('Bilimsel İçerik Uçtan Uca Deneme Başlığı', style='Title')
    document.add_paragraph('Öz: Bu çalışma bilimsel içeriğin Word dosyasından dergi çıktısına kayıpsız aktarımını '
                           'örnek bir metin ile değerlendirmek için hazırlanmıştır.')
    document.add_paragraph('Anahtar Kelimeler: denklem; dipnot; biçim')
    document.add_heading('Giriş', level=1)
    p = _runs(document.add_paragraph(), ('Tür ', ''), ('Homo sapiens', 'i'), (' için ', ''), ('önemli', 'b'), (' bulgu ', ''))
    zotero_field(p, '(Yılmaz, 2020)'); p.add_run(' ile desteklenir')
    ref(p, 'footnote', 1); p.add_run('.')
    p = document.add_paragraph('Enerji ')
    p._p.append(parse_xml(omath(EMC2)))
    p.add_run(' ilişkisi ve CO')
    p.add_run('2').font.subscript = True
    p.add_run(' ile 10')
    p.add_run('-3').font.superscript = True
    p.add_run(' değerleri ')
    content_control(p, '(Kaya &amp; Demir, 2021)')  # helper inserts raw XML
    p.add_run(' ile karşılaştırıldı')
    ref(p, 'endnote', 1); p.add_run('.')
    display = document.add_paragraph(); display._p.append(parse_xml(para(omath(FRAC))))
    p = document.add_paragraph('Basit alan ')
    simple_field(p, '(Yılmaz, 2020)'); p.add_run(' ve EndNote alanı ')
    endnote_field(p, '(Kaya & Demir, 2021)'); p.add_run(' son paragrafta yer alır.')
    _references(document)
    if path is not None:
        document.save(path)
    return document
