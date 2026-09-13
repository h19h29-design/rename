"""Guarded XLSX / HWPX package editing, without starting Office or accessing a network."""
from __future__ import annotations
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import PurePosixPath
from copy import deepcopy
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED, ZIP_STORED, BadZipFile
import base64
import posixpath
import re
from lxml import etree as E
from .rules import Hit, apply_text, header_kind, NUMBER_RE

MAX_FILE = 64 * 1024 * 1024
MAX_EXPANDED = 256 * 1024 * 1024
MAX_PARTS = 10000
MAX_UNITS = 250000
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
BLANK_PNG = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a0uoAAAAASUVORK5CYII=')

class DocumentError(Exception):
    """A user-safe message: never include original text, document paths or parser errors."""


def local(node) -> str:
    return E.QName(node).localname if isinstance(node.tag, str) else ''


def children(node, name):
    return [c for c in node if local(c) == name]


def descendants(node, name):
    return [c for c in node.iter() if local(c) == name]


def ancestor(node, name):
    return next((p for p in node.iterancestors() if local(p) == name), None)


def remove(node):
    if node.getparent() is not None: node.getparent().remove(node)


def parse_xml(data: bytes):
    try:
        parser = E.XMLParser(resolve_entities=False, load_dtd=False, no_network=True, huge_tree=False, remove_comments=True, remove_pis=True)
        tree = E.fromstring(data, parser)
        if tree.getroottree().docinfo.doctype:
            raise DocumentError('외부 엔터티 또는 DTD가 있는 문서는 처리하지 않습니다.')
        return tree
    except E.XMLSyntaxError:
        raise DocumentError('문서 XML이 손상되었거나 지원 범위를 벗어났습니다.') from None


def read_package(raw: bytes) -> dict[str, bytes]:
    if len(raw) > MAX_FILE: raise DocumentError('파일 하나의 크기는 64 MB 이하로 제한됩니다.')
    try:
        with ZipFile(BytesIO(raw)) as z:
            infos = z.infolist()
            if len(infos) > MAX_PARTS or sum(i.file_size for i in infos) > MAX_EXPANDED:
                raise DocumentError('압축 해제 크기 또는 내부 파일 수가 안전 한도를 초과했습니다.')
            seen = set()
            for i in infos:
                p = PurePosixPath(i.filename)
                if i.filename in seen or p.is_absolute() or '..' in p.parts or '\\' in i.filename or ':' in i.filename:
                    raise DocumentError('안전하지 않은 내부 파일 경로 또는 중복 항목이 있습니다.')
                if i.flag_bits & 1 or (i.external_attr >> 16) & 0o170000 == 0o120000:
                    raise DocumentError('암호화되었거나 비정상적인 문서입니다. 암호를 해제한 복사본을 사용해 주세요.')
                seen.add(i.filename)
            return {i.filename: z.read(i) for i in infos if not i.is_dir()}
    except (BadZipFile, RuntimeError, NotImplementedError, ValueError, OSError):
        raise DocumentError('열 수 없는 문서입니다. 암호를 해제하고 XLSX 또는 HWPX로 다시 저장해 주세요.') from None

@dataclass
class Slot:
    node: object
    attribute: str | None = None
    tail: bool = False

    def get(self):
        return self.node.get(self.attribute, '') if self.attribute else (self.node.tail if self.tail else self.node.text) or ''

    def set(self, value):
        if self.attribute: self.node.set(self.attribute, value)
        elif self.tail: self.node.tail = value
        else:
            self.node.text = value
            if local(self.node) == 't': self.node.set(XML_SPACE, 'preserve')

@dataclass
class Unit:
    key: int
    location: str
    text: str
    hint: str = ''
    slots: list[Slot] = field(default_factory=list)
    cell: object = None
    numeric: bool = False

    def replace(self, edits: list[tuple[Hit, str]]):
        if not edits: return
        new_text = apply_text(self.text, edits)
        if self.cell is not None and self.numeric:
            ns = E.QName(self.cell).namespace
            for item in list(self.cell):
                if local(item) in {'v', 'is', 'f'}: self.cell.remove(item)
            if NUMBER_RE.fullmatch(new_text) and ',' not in new_text and len(new_text.strip()) < 16:
                self.cell.attrib.pop('t', None)
                E.SubElement(self.cell, f'{{{ns}}}v').text = new_text.strip()
            else:
                self.cell.set('t', 'inlineStr')
                node = E.SubElement(E.SubElement(self.cell, f'{{{ns}}}is'), f'{{{ns}}}t')
                node.set(XML_SPACE, 'preserve')
                node.text = new_text
        else:
            # Apply original offsets backwards; replacement adopts the first run's style.
            for hit, new in sorted(edits, key=lambda x: x[0].start, reverse=True):
                position = 0
                spans = []
                for slot in self.slots:
                    value = slot.get()
                    spans.append((slot, position, position + len(value), value))
                    position += len(value)
                started = False
                for slot, start, end, value in spans:
                    if start < hit.end and hit.start < end:
                        a, b = max(0, hit.start-start), min(len(value), hit.end-start)
                        slot.set(value[:a] + (new if not started else '') + value[b:])
                        started = True
        self.text = new_text

class Document:
    def __init__(self, raw: bytes, suffix: str):
        self.raw, self.suffix = raw, suffix.lower()
        self.parts = read_package(raw)
        self.trees = {}
        self.units: list[Unit] = []
        self.warnings: list[str] = []
        self.has_images = False
        self._handled: set[tuple[object, str | None, bool]] = set()
        self._guard_parts()
        for name, data in self.parts.items():
            if name.lower().endswith(('.xml', '.rels', '.hpf')):
                self.trees[name] = parse_xml(data)
        if self.suffix == '.xlsx': self._xlsx()
        elif self.suffix == '.hwpx': self._hwpx()
        else: raise DocumentError('지원 형식은 XLSX, HWPX입니다.')
        self._generic_units()
        if len(self.units) > MAX_UNITS:
            raise DocumentError('문서가 너무 큽니다. 시트 또는 문서를 나누어 처리해 주세요.')
        if self.has_images:
            self.warnings.append('이미지·도장·서명 안의 글자는 탐지하지 않습니다. 원본 이미지는 유지되므로 반드시 직접 확인하세요.')

    def _guard_parts(self):
        blocked = ('/embeddings/', '/activex/', '/pivot', '/externallinks/', '/model/', '/revisions/', '/charts/', 'customxml/', '_xmlsignatures/', 'vbaproject', 'digital-signature', 'docHistory'.lower(),
                   'xl/connections.xml', 'xl/querytables/', 'xl/dde/', 'xl/ctrlprops/', 'xl/macrosheets/', 'xl/dialogsheets/')
        image_extensions = {'.png','.jpg','.jpeg','.gif','.bmp','.emf','.wmf','.tif','.tiff','.svg'}
        for name in self.parts:
            low = name.lower()
            if any(x in low for x in blocked):
                raise DocumentError('차트·피벗·외부 데이터·매크로·변경 이력·첨부 개체는 지원하지 않습니다. 복사본에서 제거한 후 다시 넣어 주세요.')
            ext = PurePosixPath(low).suffix
            ordinary = ext in image_extensions | {'.xml','.rels','.hpf','.vml'} or low == 'mimetype' or low.endswith('.rels')
            special = (low.startswith('preview/') or low.startswith('scripts/') or low.startswith('xl/printersettings/'))
            if not ordinary and not special:
                raise DocumentError('검사할 수 없는 내부 파일 형식이 있습니다. 새 XLSX/HWPX 복사본으로 저장해 주세요.')
            if ext in image_extensions and not low.startswith(('preview/', 'docprops/thumbnail')):
                self.has_images = True
            if ext in {'.exe','.dll','.ole','.pdf','.doc','.docx','.xlsx','.hwp','.zip'}:
                raise DocumentError('문서 안의 첨부 파일은 안전하게 검사할 수 없어 처리를 중단했습니다.')
            if ext == '.bin' and not low.startswith('xl/printersettings/'):
                raise DocumentError('검사할 수 없는 이진 데이터가 문서 안에 있습니다.')
            if self.suffix == '.hwpx' and not (
                low in {'mimetype','version.xml','settings.xml'} or low.startswith(('contents/','meta-inf/','preview/','bindata/','scripts/'))
            ):
                raise DocumentError('알 수 없는 HWPX 내부 파일이 있습니다. 한글에서 새 복사본으로 저장해 주세요.')

    def add(self, location, slots, hint='', cell=None, numeric=False, text=None):
        if text is None: text = ''.join(s.get() for s in slots)
        if not text: return None
        u = Unit(len(self.units), location, text, hint, slots, cell, numeric)
        self.units.append(u)
        self._handled.update((s.node, s.attribute, s.tail) for s in slots)
        return u

    def _drop_parts(self, names):
        for name in names:
            self.parts.pop(name, None)
            self.trees.pop(name, None)
        # Remove content-type entries and relationships for deleted parts; external links are never copied.
        for name, tree in self.trees.items():
            if name == '[Content_Types].xml':
                for child in list(tree):
                    if child.get('PartName', '').lstrip('/') in names: remove(child)
            if name.endswith('.rels'):
                base = posixpath.dirname(posixpath.dirname(name))
                removed_ids = set()
                for child in list(tree):
                    target = child.get('Target','')
                    resolved = posixpath.normpath(posixpath.join(base, target)).lstrip('/')
                    if resolved in names or child.get('TargetMode') == 'External':
                        removed_ids.add(child.get('Id'))
                        remove(child)
                parent_name = name.replace('/_rels/', '/').removesuffix('.rels')
                parent = self.trees.get(parent_name)
                if parent is not None:
                    for el in list(parent.iter()):
                        for attr, value in list(el.attrib.items()):
                            if E.QName(attr).localname in {'id','embed','link'} and value in removed_ids:
                                if local(el) in {'hyperlink','legacyDrawing','legacyDrawingHF','hlinkClick','hlinkHover'}:
                                    remove(el)
                                else: el.attrib.pop(attr, None)

    def _xlsx(self):
        workbook = self.trees.get('xl/workbook.xml')
        if workbook is None: raise DocumentError('유효한 XLSX 통합문서가 아닙니다.')
        self._drop_parts({n for n in self.parts if n.startswith(('docProps/','xl/comments','xl/threadedComments','xl/persons/','xl/printerSettings/')) or n == 'xl/calcChain.xml' or n.lower().endswith('.vml')})
        # Comments / validation lists / names / formatting formulas can retain raw values.
        for tree in self.trees.values():
            for node in list(tree.iter()):
                if local(node) in {'definedNames','dataValidations','conditionalFormatting','extLst','autoFilter','sortState','calculatedColumnFormula','totalsRowFormula','hyperlinks'}:
                    remove(node)
        shared = self.trees.get('xl/sharedStrings.xml')
        strings = children(shared, 'si') if shared is not None else []
        styles = self.trees.get('xl/styles.xml')
        formats, xfs = {}, []
        if styles is not None:
            formats = {int(n.get('numFmtId')): n.get('formatCode','') for n in descendants(styles, 'numFmt')}
            xfl = children(styles, 'cellXfs')
            if xfl: xfs = [int(n.get('numFmtId','0')) for n in xfl[0]]
        props = children(workbook, 'workbookPr')
        epoch1904 = bool(props and props[0].get('date1904') in {'1','true'})
        relations = self.trees.get('xl/_rels/workbook.xml.rels')
        sheet_names = {}
        if relations is not None:
            targets = {x.get('Id'): posixpath.normpath(posixpath.join('xl',x.get('Target',''))).lstrip('/') for x in relations}
            for sheet in descendants(workbook,'sheet'):
                rid = next((v for a,v in sheet.attrib.items() if E.QName(a).localname == 'id'), '')
                sheet_names[targets.get(rid,'')] = sheet.get('name','시트')
        formulas = missing = 0
        for name, tree in self.trees.items():
            if not name.startswith('xl/worksheets/') or not name.endswith('.xml'): continue
            cells = {}
            for cell in descendants(tree,'c'):
                ns = E.QName(cell).namespace
                addr = cell.get('r','')
                f, v = children(cell,'f'), children(cell,'v')
                if f:
                    formulas += 1
                    if not v or v[0].text is None: missing += 1
                    for n in f: remove(n)
                if cell.get('t') == 's':
                    try:
                        index = int(v[0].text)
                        if not 0 <= index < len(strings): raise IndexError
                        si = deepcopy(strings[index])
                    except (IndexError, TypeError, ValueError):
                        raise DocumentError('엑셀 공유 문자열 인덱스가 올바르지 않습니다.') from None
                    for n in list(cell):
                        if local(n) in {'v','is'}: remove(n)
                    si.tag = f'{{{ns}}}is'
                    cell.set('t','inlineStr')
                    cell.append(si)
                for n in list(cell.iter()):
                    if local(n) in {'rPh','phoneticPr'}: remove(n)
                slots = [Slot(t) for t in descendants(cell,'t')]
                numeric = not slots and cell.get('t','n') in {'n','d'}
                if not slots: slots = [Slot(n) for n in children(cell,'v')]
                text = ''.join(s.get() for s in slots)
                hint = ''
                try:
                    fmt = xfs[int(cell.get('s','0'))] if xfs else 0
                    code = formats.get(fmt,'')
                    date_style = fmt in {14,15,16,17,22,27,28,29,30,31,32,33,34,35,36,50,51,52,53,54,55,56,57,58} or bool(re.search('[yd]', re.sub(r'"[^"]*"|\[[^]]*\]', '', code), re.I))
                    if numeric and text and date_style:
                        serial = float(text)
                        if not 0 <= serial <= 2958465: raise ValueError
                        adjusted = serial + 1 if not epoch1904 and serial < 60 else serial
                        base = datetime(1904,1,1) if epoch1904 else datetime(1899,12,30)
                        text = (base + timedelta(days=adjusted)).strftime('%Y-%m-%d')
                        hint = '날짜'
                    elif re.search(r'[₩$€]|원',code): hint = '금액'
                except (ValueError, OverflowError, IndexError):
                    pass
                u = self.add(f'{sheet_names.get(name,"시트")}!{addr}', slots, hint, cell, numeric, text)
                if u and re.fullmatch('[A-Z]+[0-9]+',addr):
                    match = re.fullmatch('([A-Z]+)([0-9]+)',addr)
                    col = 0
                    for c in match[1]: col = col*26 + ord(c)-64
                    cells[(int(match[2]),col)] = u
            self._table_hints(cells)
            for unit in cells.values():
                if not unit.numeric: continue
                try:
                    number = Decimal(unit.text)
                    if not number.is_finite() or abs(number.adjusted()) > 30: continue
                    unit.text = format(number,'f')
                    kind = header_kind(unit.hint)
                    if number >= 0 and number == number.to_integral_value():
                        digits = str(int(number))
                        if kind == 'rrn' and len(digits) <= 13: unit.text = digits.zfill(13)
                        elif kind == 'birth' and len(digits) <= 6: unit.text = digits.zfill(6)
                except (InvalidOperation,ValueError,OverflowError):
                    pass
        self._drop_parts({'xl/sharedStrings.xml'})
        if formulas:
            self.warnings.append(f'수식 {formulas}개는 저장된 결과값으로 고정됩니다. 계산 기능은 유지되지 않습니다.')
        if missing:
            self.warnings.append(f'저장된 결과값이 없는 수식 {missing}개는 빈칸이 됩니다. 원본을 Excel에서 계산·저장한 뒤 다시 넣으세요.')
        self.warnings.append('메모·작성자 속성·외부 링크·이름 정의·조건부 서식·데이터 유효성 규칙은 공유 복사본에서 제거됩니다.')

    @staticmethod
    def _table_hints(cells):
        last_headers = {}
        for (row,col), unit in sorted(cells.items()):
            kind = header_kind(unit.text)
            if kind:
                last_headers[col] = unit.text
                continue
            left = cells.get((row,col-1))
            if left and header_kind(left.text): unit.hint = left.text
            elif col in last_headers: unit.hint = last_headers[col]

    def _hwpx(self):
        if self.parts.get('mimetype',b'').strip() != b'application/hwp+zip':
            raise DocumentError('유효한 HWPX 문서가 아닙니다.')
        if not any(re.fullmatch(r'Contents/section\d+\.xml',n) for n in self.parts):
            raise DocumentError('HWPX 본문을 찾을 수 없습니다.')
        for name,tree in self.trees.items():
            if local(tree) in {'encryption','encrypted-package'} or any(local(n).lower() == 'encrypted-data' for n in tree.iter()):
                raise DocumentError('암호화된 문서는 처리할 수 없습니다.')
            for node in tree.iter():
                if local(node) == 'metadata':
                    for sub in node.iter():
                        if sub is not node:
                            sub.text = ''
                            for a in ('content','value'):
                                if a in sub.attrib: sub.set(a,'')
            if not name.startswith('Contents/section') or not name.endswith('.xml'): continue
            if any(local(n).lower() in {'trackchange','insertbegin','deletebegin','memogroup','memo'} for n in tree.iter()):
                raise DocumentError('메모 또는 변경 이력이 있는 한글 문서는 먼저 이를 제거해 주세요.')
            table_cells = {}
            for index,p in enumerate(descendants(tree,'p'),1):
                slots = []
                for t in descendants(p,'t'):
                    if ancestor(t,'p') is not p: continue
                    if t.text: slots.append(Slot(t))
                    for sub in t.iterdescendants():
                        if sub.text: slots.append(Slot(sub))
                        if sub.tail: slots.append(Slot(sub,tail=True))
                unit = self.add(f'{name.rsplit("/",1)[-1]} · 문단 {index}',slots)
                if unit:
                    tc, tbl = ancestor(p,'tc'), ancestor(p,'tbl')
                    if tc is not None and tbl is not None:
                        addrs = children(tc,'cellAddr')
                        if addrs:
                            a = addrs[0]
                            key = (int(a.get('rowAddr','0')),int(a.get('colAddr','0')))
                            table_cells.setdefault(tbl,{}).setdefault(key,[]).append(unit)
            for table in table_cells.values():
                firsts = {k: units[0] for k,units in table.items()}
                self._table_hints(firsts)
                for k,units in table.items():
                    for u in units[1:]: u.hint = units[0].hint
        for name in list(self.parts):
            if name.lower().startswith('scripts/'): self.parts[name] = b''
        self.warnings.append('한글 미리보기·작성자 속성을 정리합니다. 치환 길이에 따라 줄바꿈과 쪽 배치가 달라질 수 있습니다.')

    def _generic_units(self):
        # Inspect remaining text and human-readable attributes, including headers and drawing text.
        allowed = {'name','title','descr','description','tooltip','formatCode','text','author','creator','subject','label'}
        for name,tree in self.trees.items():
            for index,node in enumerate(tree.iter()):
                if not isinstance(node.tag,str): continue
                if node.text and node.text.strip() and (node,None,False) not in self._handled:
                    self.add(f'{name} · 부가 텍스트 {index}',[Slot(node)])
                for attr,value in node.attrib.items():
                    if E.QName(attr).localname in allowed and value and (node,attr,False) not in self._handled:
                        self.add(f'{name} · 속성 {index}',[Slot(node,attr)])

    def render(self, edits: dict[int,list[tuple[Hit,str]]]) -> bytes:
        for key, replacements in edits.items(): self.units[key].replace(replacements)
        if self.suffix == '.xlsx':
            sheets = descendants(self.trees['xl/workbook.xml'],'sheet')
            titles = [s.get('name','') for s in sheets]
            if len(set(t.casefold() for t in titles)) != len(titles) or any(not t or len(t)>31 or re.search(r'[\\/*?:\[\]]',t) for t in titles):
                raise DocumentError('변환된 시트 이름이 Excel 규칙에 맞지 않습니다. 이름 대체값을 조정해 주세요.')
        if self.suffix == '.hwpx':
            preview = '\n'.join(u.text for u in self.units if u.location.startswith('section'))
            for name in list(self.parts):
                if name.lower().startswith('preview/'):
                    if name.lower().endswith('.txt'): self.parts[name] = preview.encode('utf-8')
                    elif name.lower().endswith('.png'): self.parts[name] = BLANK_PNG
                    else: self.parts[name] = b''
        output = BytesIO()
        with ZipFile(output,'w',ZIP_DEFLATED,compresslevel=6) as z:
            names = sorted(self.parts, key=lambda n:(n != 'mimetype',n))
            for name in names:
                if name in self.trees:
                    data = E.tostring(self.trees[name],encoding='UTF-8',xml_declaration=True,standalone=True)
                else: data = self.parts[name]
                z.writestr(name,data,compress_type=ZIP_STORED if name=='mimetype' else ZIP_DEFLATED)
        return output.getvalue()
