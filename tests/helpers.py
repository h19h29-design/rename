from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape

X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'


def package(path, parts):
    with ZipFile(path, 'w', ZIP_DEFLATED) as z:
        for name, value in parts.items():
            z.writestr(name, value)
    return path


def xlsx(path, extra=None):
    strings = ['성명', '주민등록번호', '생년월일', '금액', '계약일', '홍길동', '900101-1234567', '1990-01-01', '미사용원문']
    ss = ''.join(f'<si><t>{escape(t)}</t></si>' for t in strings)
    cells = '<row r="1">' + ''.join(f'<c r="{chr(65+i)}1" t="s"><v>{i}</v></c>' for i in range(5)) + '</row>'
    cells += '<row r="2"><c r="A2" t="s"><v>5</v></c><c r="B2" t="s"><v>6</v></c><c r="C2" t="s"><v>7</v></c><c r="D2"><v>1234500</v></c><c r="E2" s="1"><v>46278</v></c></row>'
    cells += '<row r="3" hidden="1"><c r="A3" t="s"><v>5</v></c><c r="F3" t="str"><f>"홍길동"</f><v>홍길동</v></c></row>'
    parts = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
        '_rels/.rels': f'<Relationships xmlns="{P}"><Relationship Id="r1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': f'<workbook xmlns="{X}" xmlns:r="{R}"><sheets><sheet name="명단" sheetId="1" r:id="r1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="{P}"><Relationship Id="r1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Type="{R}/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
        'xl/sharedStrings.xml': f'<sst xmlns="{X}" count="10" uniqueCount="9">{ss}</sst>',
        'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{X}"><sheetData>{cells}</sheetData><mergeCells><mergeCell ref="A5:B5"/></mergeCells></worksheet>',
        'xl/styles.xml': f'<styleSheet xmlns="{X}"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>',
        'docProps/core.xml': '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>비공개작성자</dc:creator></cp:coreProperties>',
    }
    parts.update(extra or {})
    return package(path, parts)


def hwpx(path, extra=None):
    hp = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
    texts = ['성명: 홍', '길동 / 주민등록번호: 900101-', '1234567 / 생년월일: 1990.01.01.']
    runs = ''.join(f'<hp:run charPrIDRef="{i}"><hp:t>{escape(t)}</hp:t></hp:run>' for i,t in enumerate(texts))
    parts = {
        'mimetype': 'application/hwp+zip',
        'Contents/section0.xml': f'<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="{hp}"><hp:p id="1">{runs}</hp:p><hp:p id="2"><hp:run><hp:t>홍길동 확인</hp:t></hp:run></hp:p></hs:sec>',
        'Contents/content.hpf': '<opf:package xmlns:opf="http://www.idpf.org/2007/opf"><opf:metadata><opf:meta name="creator" content="비공개작성자"/></opf:metadata><opf:manifest/></opf:package>',
        'Preview/PrvText.txt': '성명 홍길동 900101-1234567',
        'Preview/PrvImage.png': b'OLD_PRIVATE_PREVIEW',
    }
    parts.update(extra or {})
    return package(path, parts)


def all_text(path):
    with ZipFile(path) as z:
        return '\n'.join(z.read(n).decode('utf-8', errors='replace') for n in z.namelist())
