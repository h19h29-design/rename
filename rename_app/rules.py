"""Conservative local detectors. Findings are candidates, not an anonymity guarantee."""
from __future__ import annotations
from dataclasses import dataclass
from datetime import date
import re

LABELS = {'name': '이름', 'rrn': '주민등록번호', 'birth': '생년월일', 'amount': '금액', 'date': '날짜'}
PRIORITY = {'rrn': 0, 'birth': 1, 'date': 2, 'amount': 3, 'name': 4}
NAME_LABEL = r'(?:성\s*명|이\s*름|학생명|신청인|신청자|담당자|대상자|수령인|계약자|대표자|보호자|근로자|소유자|예금주|직원명|성명\(한글\))'
BIRTH_LABEL = r'(?:생\s*년\s*월\s*일|생일|출생일|출생일자|생년월일\(6자리\)|DOB|birth\s*date)'
AMOUNT_LABEL = r'(?:금액|총액|합계|단가|급여|연봉|월급|수당|지급액|입금|출금|보수|수익|비용|예산|공급가액|부가세|세액|대금|금\s*액)'
DATE_LABEL = r'(?:계약일|작성일|시행일|날짜|일자|기간|등록일|지급일|시작일|종료일)'
RRN_RE = re.compile(r'(?<![\d])\d{6}[\s\-‐‑–−]?[1-8][\d*xX●•]{6}(?![\d*xX])')
DATE_RE = re.compile(r'(?<!\d)(?:(?:19|20)\d{2}\s*(?:[-./]\s*\d{1,2}\s*[-./]\s*\d{1,2}\.?|년\s*\d{1,2}\s*월\s*\d{1,2}\s*일)|(?:19|20)\d{6})(?!\d)')
SHORT_DATE_RE = re.compile(r'(?<!\d)\d{2}(?:[-./]\s*\d{1,2}\s*[-./]\s*\d{1,2}\.?|\d{4})(?!\d)')
AMOUNT_RE = re.compile(r'(?<![\w])(?:[₩$€]\s*-?\d[\d,]*(?:\.\d+)?|-?\d[\d,]*(?:\.\d+)?\s*(?:백만원|천만원|만원|천원|억원|원|USD|KRW|달러))(?![가-힣A-Za-z])')
NUMBER_RE = re.compile(r'^\s*-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?\s*$')

@dataclass(frozen=True)
class Options:
    names: bool = True
    rrn: bool = True
    birth: bool = True
    amounts: bool = False
    dates: bool = False
    mode: str = 'token'
    manual_names: tuple[str, ...] = ()

@dataclass(frozen=True)
class Hit:
    kind: str
    value: str
    start: int
    end: int


def header_kind(text: str) -> str:
    text = text.strip(' :：\t\n()[]')
    for kind, pattern in [('birth', BIRTH_LABEL), ('rrn', r'주민\s*등록\s*번호|주민번호|주민등록번호\(13자리\)'), ('name', NAME_LABEL), ('amount', AMOUNT_LABEL), ('date', DATE_LABEL)]:
        if re.fullmatch(pattern + r'(?:\s*\([^)]{0,15}\))?', text, re.I):
            return kind
    return ''


def valid_date(text: str) -> bool:
    numbers = re.findall(r'\d+', text)
    try:
        if len(numbers) == 1 and len(numbers[0]) in (6, 8):
            raw = numbers[0]
            y, m, d = int(raw[:-4]), int(raw[-4:-2]), int(raw[-2:])
        elif len(numbers) == 3:
            y, m, d = map(int, numbers)
        else:
            return False
        if y < 100: y += 2000 if y < 30 else 1900
        date(y, m, d)
        return True
    except (ValueError, OverflowError):
        return False


def _birth_context(text: str, start: int, hint: str) -> bool:
    if header_kind(hint) == 'birth': return True
    preceding = text[max(0, start - 48):start]
    labels = list(re.finditer(f'{BIRTH_LABEL}|{DATE_LABEL}', preceding, re.I))
    return bool(labels and re.fullmatch(BIRTH_LABEL, labels[-1].group(), re.I))


def detect(text: str, hint: str, options: Options) -> list[Hit]:
    hits: list[Hit] = []
    rrns = list(RRN_RE.finditer(text))
    if options.rrn:
        hits.extend(Hit('rrn', m.group(), *m.span()) for m in rrns)
    if options.birth or options.dates:
        matches = list(DATE_RE.finditer(text))
        if header_kind(hint) == 'birth' or re.search(BIRTH_LABEL, text, re.I):
            matches += list(SHORT_DATE_RE.finditer(text))
        for m in matches:
            if any(m.start() < r.end() and r.start() < m.end() for r in rrns): continue
            if not valid_date(m.group()): continue
            birthday = _birth_context(text, m.start(), hint)
            if birthday and options.birth:
                hits.append(Hit('birth', m.group(), *m.span()))
            elif not birthday and options.dates:
                hits.append(Hit('date', m.group(), *m.span()))
    if options.amounts:
        hits.extend(Hit('amount', m.group(), *m.span()) for m in AMOUNT_RE.finditer(text))
        if header_kind(hint) == 'amount' and NUMBER_RE.fullmatch(text):
            a = len(text) - len(text.lstrip())
            hits.append(Hit('amount', text.strip(), a, len(text.rstrip())))
    if options.names:
        for m in re.finditer(NAME_LABEL + r'\s*[:：]?\s*(?P<value>[가-힣]{2,5}|[가-힣](?:[ \t]+[가-힣]){1,4})(?![가-힣])', text):
            val = m.group('value')
            if not header_kind(val) and val not in {'없음','미입력','확인','해당없음','미정'}:
                hits.append(Hit('name', val, *m.span('value')))
        if header_kind(hint) == 'name':
            value = text.strip()
            if re.fullmatch(r'[가-힣](?:[ \t]*[가-힣]){1,4}|[A-Za-z][A-Za-z .\-\']{1,59}', value) and not header_kind(value):
                a = text.find(value)
                hits.append(Hit('name', value, a, a + len(value)))
        for value in sorted(set(options.manual_names), key=len, reverse=True):
            if not value.strip(): continue
            compact = re.sub(r'\s+', '', value)
            pattern = r'[ \t]*'.join(map(re.escape, compact)) if re.fullmatch(r'[가-힣]{2,5}',compact) else re.escape(value)
            hits.extend(Hit('name', m.group(), *m.span()) for m in re.finditer(pattern, text))
    # Full RRN wins over dates; longer matches win within the same category.
    accepted: list[Hit] = []
    for h in sorted(hits, key=lambda h: (PRIORITY[h.kind], -(h.end-h.start), h.start)):
        if not any(h.start < a.end and a.start < h.end for a in accepted):
            accepted.append(h)
    return sorted(accepted, key=lambda h: h.start)


def canonical(kind: str, value: str) -> str:
    if kind == 'name' and re.fullmatch(r'[가-힣\s]{2,12}',value): return re.sub(r'\s+','',value)
    if kind == 'rrn': return re.sub(r'[\s\-‐‑–−]', '', value)
    return value


def replacement(kind: str, index: int, mode: str) -> str:
    if mode == 'delete': return ''
    if mode == 'mask':
        return {'name': '***', 'rrn': '******-*******', 'birth': '****-**-**', 'date': '****-**-**', 'amount': '***'}[kind]
    if kind == 'name': return f'대상자{index:03d}'
    return f'[{LABELS[kind]}{index:03d}]'


def apply_text(text: str, edits: list[tuple[Hit, str]]) -> str:
    end = len(text)
    for hit, new in sorted(edits, key=lambda e: e[0].start, reverse=True):
        if not 0 <= hit.start < hit.end <= end or text[hit.start:hit.end] != hit.value:
            raise ValueError('Replacement positions are inconsistent; rescan the document.')
        text = text[:hit.start] + new + text[hit.end:]
        end = hit.start
    return text
