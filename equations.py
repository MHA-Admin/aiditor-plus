"""Word equations (OMML) for import and output (Phase 1.5).

``omml_linear``  readable Unicode text of an equation; it is what the section
                 text box shows and the safe fallback when no other form fits.
``omml_latex``   LaTeX math for a whitelisted subset of OMML; None when any
                 part of the equation is outside the subset, so a partial
                 conversion can never change the meaning silently.
``safe_omml``    parses stored OMML back for Word output; only the math and
                 WordprocessingML namespaces are accepted (no entities, no
                 network, size-limited).

The LaTeX engine is XeLaTeX with ``unicode-math`` and ``amsmath`` (every
journal template loads both), so Unicode math symbols are written as they are
and only TeX special characters need escaping.
"""
import re

from lxml import etree

M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
MAX_OMML_BYTES = 200_000
_XML = 'http://www.w3.org/XML/1998/namespace'
_MAX_DEPTH = 60


def m(tag: str) -> str:
    return '{' + M + '}' + tag


def _local(node) -> str:
    return etree.QName(node).localname if isinstance(node.tag, str) else ''


def _ns(node) -> str:
    return etree.QName(node).namespace if isinstance(node.tag, str) else ''


def _child(node, tag):
    return node.find(m(tag)) if node is not None else None


def _prop(node, props: str, name: str, default=None):
    """``m:val`` of ``m:<props>/m:<name>``; ``default`` when absent."""
    pr = _child(node, props)
    element = _child(pr, name)
    if element is None:
        return default
    return element.get(m('val'), default if default is not None else '')


def _on(node, props: str, name: str) -> bool:
    value = _prop(node, props, name)
    return value is not None and value.lower() not in ('0', 'off', 'false')


class _Unsupported(Exception):
    """Raised for any OMML construct outside the LaTeX whitelist."""


# ── Linear (plain text) form ────────────────────────────────────────────────
_SUP = dict(zip('0123456789+-−=()ni', '⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁻⁼⁽⁾ⁿⁱ'))
_SUB = dict(zip('0123456789+-−=()', '₀₁₂₃₄₅₆₇₈₉₊₋₋₌₍₎'))


def _script(text: str, table: dict, mark: str) -> str:
    if text and all(char in table for char in text):
        return ''.join(table[char] for char in text)
    return mark + (text if len(text) == 1 else '(' + text + ')')


def _group(text: str) -> str:
    return text if len(text) <= 1 or (text.startswith('(') and text.endswith(')')) else '(' + text + ')'


def _linear(node, depth=0) -> str:
    if depth > _MAX_DEPTH:
        return ''
    name = _local(node)
    if _ns(node) == M and name.endswith('Pr'):
        return ''
    if name == 't':
        return node.text or ''
    if _ns(node) != M:
        return ''.join(_linear(child, depth + 1) for child in node)

    def part(tag):
        element = _child(node, tag)
        return _linear(element, depth + 1) if element is not None else ''

    if name == 'f':
        return _group(part('num')) + '/' + _group(part('den'))
    if name == 'sSup':
        return part('e') + _script(part('sup'), _SUP, '^')
    if name == 'sSub':
        return part('e') + _script(part('sub'), _SUB, '_')
    if name == 'sSubSup':
        return part('e') + _script(part('sub'), _SUB, '_') + _script(part('sup'), _SUP, '^')
    if name == 'sPre':
        return _script(part('sub'), _SUB, '_') + _script(part('sup'), _SUP, '^') + part('e')
    if name == 'rad':
        degree = part('deg')
        root = {'': '√', '2': '√', '3': '∛', '4': '∜'}.get(degree)
        return (root or degree + '√') + _group(part('e'))
    if name == 'nary':
        symbol = _prop(node, 'naryPr', 'chr', '∫') or '∫'
        lower = '' if _on(node, 'naryPr', 'subHide') else part('sub')
        upper = '' if _on(node, 'naryPr', 'supHide') else part('sup')
        return (symbol + (_script(lower, _SUB, '_') if lower else '') + (_script(upper, _SUP, '^') if upper else '')
                + ' ' + part('e')).rstrip()
    if name == 'd':
        begin = _prop(node, 'dPr', 'begChr', '(')
        end = _prop(node, 'dPr', 'endChr', ')')
        separator = _prop(node, 'dPr', 'sepChr', '|')
        return begin + separator.join(_linear(e, depth + 1) for e in node.findall(m('e'))) + end
    if name == 'acc':
        return part('e') + (_prop(node, 'accPr', 'chr', '̂') or '̂')
    if name == 'bar':
        return part('e') + ('̲' if _prop(node, 'barPr', 'pos', 'bot') == 'bot' else '̅')
    if name == 'func':
        return part('fName') + ' ' + part('e')
    if name in ('limLow', 'limUpp'):
        return part('e') + ('_' if name == 'limLow' else '^') + _group(part('lim'))
    if name == 'm':
        rows = ['  '.join(_linear(e, depth + 1) for e in row.findall(m('e'))) for row in node.findall(m('mr'))]
        return '[' + '; '.join(rows) + ']'
    if name == 'eqArr':
        return '; '.join(_linear(e, depth + 1) for e in node.findall(m('e')))
    if name == 'oMathPara':
        return '   '.join(_linear(child, depth + 1) for child in node.findall(m('oMath')))
    return ''.join(_linear(child, depth + 1) for child in node)


def omml_linear(node) -> str:
    """Single-line Unicode text of an ``m:oMath``/``m:oMathPara`` (never empty for visible math)."""
    text = ' '.join(_linear(node).split())
    return text


# ── LaTeX form ──────────────────────────────────────────────────────────────
_TEX_SPECIAL = {
    '\\': r'\backslash ', '{': r'\{', '}': r'\}', '#': r'\#', '$': r'\$', '%': r'\%', '&': r'\&',
    '_': r'\_', '~': r'\sim ', '^': r'\hat{}',
}
_NARY = {'∑': r'\sum', '∏': r'\prod', '∐': r'\coprod', '∫': r'\int', '∬': r'\iint', '∭': r'\iiint',
         '∮': r'\oint', '⋃': r'\bigcup', '⋂': r'\bigcap', '⋁': r'\bigvee', '⋀': r'\bigwedge'}
_DELIMITERS = {'(': '(', ')': ')', '[': '[', ']': ']', '{': r'\{', '}': r'\}', '|': '|', '‖': r'\|',
               '⟨': r'\langle', '⟩': r'\rangle', '〈': r'\langle', '〉': r'\rangle', '⌊': r'\lfloor',
               '⌋': r'\rfloor', '⌈': r'\lceil', '⌉': r'\rceil', '': '.'}
_ACCENTS = {'̂': r'\hat', '̃': r'\tilde', '̄': r'\bar', '̅': r'\bar', '̇': r'\dot',
            '̈': r'\ddot', '⃗': r'\vec', '́': r'\acute', '̀': r'\grave', '̆': r'\breve',
            '̌': r'\check'}
_FUNCTIONS = {'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'sinh', 'cosh', 'tanh', 'coth', 'arcsin', 'arccos',
              'arctan', 'log', 'ln', 'lg', 'exp', 'lim', 'liminf', 'limsup', 'max', 'min', 'sup', 'inf',
              'det', 'dim', 'ker', 'deg', 'arg', 'gcd', 'Pr', 'hom'}
_IGNORED = {'ctrlPr', 'rPr', 'argPr'}


_TEXT_MODE_SPECIAL = {
    '\\': r'\textbackslash{}', '{': r'\{', '}': r'\}', '#': r'\#', '$': r'\$', '%': r'\%', '&': r'\&',
    '_': r'\_', '~': r'\textasciitilde{}', '^': r'\textasciicircum{}',
}
_DECIMAL_COMMA = re.compile(r'(?<=\d),(?=\d)')


def _tex_text(text: str) -> str:
    """Math-mode text; a decimal comma (0,05) is not followed by punctuation space."""
    return _DECIMAL_COMMA.sub('{,}', ''.join(_TEX_SPECIAL.get(char, char) for char in text))


def _tex_upright(text: str) -> str:
    """Upright text inside math (Word "normal text"); \\text keeps its spaces."""
    if ' ' in text:
        return r'\text{' + ''.join(_TEXT_MODE_SPECIAL.get(char, char) for char in text) + '}'
    return r'\mathrm{' + _tex_text(text) + '}'


def _latex(node, depth=0) -> str:
    if depth > _MAX_DEPTH:
        raise _Unsupported('too deep')
    name = _local(node)
    if _ns(node) == W and name in ('rPr', 'bookmarkStart', 'bookmarkEnd', 'proofErr'):
        return ''
    if _ns(node) != M:
        raise _Unsupported(name)
    if name.endswith('Pr') or name in _IGNORED:
        return ''

    def part(tag, required=True):
        element = _child(node, tag)
        if element is None:
            if required:
                raise _Unsupported('missing ' + tag)
            return ''
        return _latex(element, depth + 1)

    def args():
        return ''.join(_latex(child, depth + 1) for child in node)

    if name in ('oMath', 'e', 'num', 'den', 'sup', 'sub', 'deg', 'lim', 'fName', 'box'):
        return args()
    if name == 'r':
        text = ''.join(t.text or '' for t in node.findall(m('t')))
        for child in node:
            if not ((_local(child) == 't' and _ns(child) == M) or (_local(child) == 'rPr' and _ns(child) in (M, W))):
                raise _Unsupported('run child ' + _local(child))
        plain = _on(node, 'rPr', 'nor') or (_prop(node, 'rPr', 'sty') == 'p' and len(text.strip()) > 1)
        if plain and text.strip():
            return _tex_upright(text)
        if text.strip() in _FUNCTIONS:
            return '\\' + text.strip() + ' '
        return _tex_text(text)
    if name == 'f':
        kind = _prop(node, 'fPr', 'type', 'bar')
        num, den = part('num'), part('den')
        if kind == 'lin':
            return '{' + num + '}/{' + den + '}'
        if kind == 'noBar':
            return r'\genfrac{}{}{0pt}{}{' + num + '}{' + den + '}'
        if kind == 'skw':
            return r'{}^{' + num + r'}\!/_{' + den + '}'
        return r'\frac{' + num + '}{' + den + '}'
    if name == 'sSup':
        return '{' + part('e') + '}^{' + part('sup') + '}'
    if name == 'sSub':
        return '{' + part('e') + '}_{' + part('sub') + '}'
    if name == 'sSubSup':
        return '{' + part('e') + '}_{' + part('sub') + '}^{' + part('sup') + '}'
    if name == 'sPre':
        return '{}_{' + part('sub') + '}^{' + part('sup') + '}{' + part('e') + '}'
    if name == 'rad':
        degree = '' if _on(node, 'radPr', 'degHide') else part('deg', required=False)
        return r'\sqrt' + ('[' + degree + ']' if degree.strip() else '') + '{' + part('e') + '}'
    if name == 'nary':
        symbol = _prop(node, 'naryPr', 'chr', '∫') or '∫'
        command = _NARY.get(symbol)
        if command is None:
            raise _Unsupported('n-ary ' + symbol)
        lower = '' if _on(node, 'naryPr', 'subHide') else part('sub', required=False)
        upper = '' if _on(node, 'naryPr', 'supHide') else part('sup', required=False)
        return (command + ('_{' + lower + '}' if lower else '') + ('^{' + upper + '}' if upper else '')
                + '{' + part('e') + '}')
    if name == 'd':
        begin = _prop(node, 'dPr', 'begChr', '(')
        end = _prop(node, 'dPr', 'endChr', ')')
        separator = _prop(node, 'dPr', 'sepChr', '|')
        if begin not in _DELIMITERS or end not in _DELIMITERS or separator not in _DELIMITERS or separator == '':
            raise _Unsupported('delimiter')
        inner = (r' \middle' + _DELIMITERS[separator] + ' ').join(_latex(e, depth + 1) for e in node.findall(m('e')))
        return r'\left' + _DELIMITERS[begin] + ' ' + inner + r' \right' + _DELIMITERS[end]
    if name == 'acc':
        accent = _ACCENTS.get(_prop(node, 'accPr', 'chr', '̂') or '̂')
        if accent is None:
            raise _Unsupported('accent')
        return accent + '{' + part('e') + '}'
    if name == 'bar':
        return (r'\underline{' if _prop(node, 'barPr', 'pos', 'bot') == 'bot' else r'\overline{') + part('e') + '}'
    if name == 'func':
        function = part('fName')
        return function + ('' if function.endswith(' ') or function.endswith('}') else ' ') + '{' + part('e') + '}'
    if name == 'limLow':
        return r'\mathop{' + part('e') + r'}\limits_{' + part('lim') + '}'
    if name == 'limUpp':
        return r'\mathop{' + part('e') + r'}\limits^{' + part('lim') + '}'
    if name == 'borderBox':
        return r'\boxed{' + part('e') + '}'
    if name == 'groupChr':
        char = _prop(node, 'groupChrPr', 'chr', '⏟') or '⏟'
        command = {'⏟': r'\underbrace', '⏞': r'\overbrace'}.get(char)
        if command is None:
            raise _Unsupported('group character')
        return command + '{' + part('e') + '}'
    if name == 'm':
        rows = []
        for row in node.findall(m('mr')):
            rows.append(' & '.join(_latex(e, depth + 1) for e in row.findall(m('e'))))
        if not rows:
            raise _Unsupported('empty matrix')
        return r'\begin{matrix}' + r' \\ '.join(rows) + r'\end{matrix}'
    if name == 'eqArr':
        return r'\begin{gathered}' + r' \\ '.join(_latex(e, depth + 1) for e in node.findall(m('e'))) + r'\end{gathered}'
    if name == 'phant':
        return r'\phantom{' + part('e') + '}'
    raise _Unsupported(name)


def omml_latex(node) -> str | None:
    """LaTeX (without math delimiters) for one ``m:oMath``; None when unsupported."""
    try:
        latex = _latex(node).strip()
    except _Unsupported:
        return None
    return latex or None


def omml_display_latex(node) -> str | None:
    """Display LaTeX for an ``m:oMathPara`` (one or more equations), with delimiters."""
    parts = [omml_latex(child) for child in node.findall(m('oMath'))]
    if not parts or any(part is None for part in parts):
        return None
    if len(parts) == 1:
        return r'\[' + parts[0] + r'\]'
    return r'\begin{gather*}' + r' \\ '.join(parts) + r'\end{gather*}'


def to_xml(node) -> str:
    return etree.tostring(node, encoding='unicode')


_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False, remove_comments=True,
                          remove_pis=True)


def safe_omml(xml: str):
    """Parse stored OMML for Word output; None unless it is plain math/Word markup."""
    if not isinstance(xml, str) or not xml or len(xml.encode('utf-8')) > MAX_OMML_BYTES:
        return None
    try:
        node = etree.fromstring(xml.encode('utf-8'), _PARSER)
    except etree.XMLSyntaxError:
        return None
    if node.tag not in (m('oMath'), m('oMathPara')):
        return None
    for element in node.iter():
        if not isinstance(element.tag, str) or _ns(element) not in (M, W):
            return None
        if any(etree.QName(attribute).namespace not in (M, W, _XML, None) for attribute in element.attrib):
            return None
    return node
