import hashlib
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZipFile
import pytest
from rename_app.rules import Options
from rename_app.session import Session
from rename_app.documents import DocumentError, MAX_FILE
from tests.helpers import xlsx, hwpx, all_text, package


@pytest.mark.parametrize('factory,suffix', [(xlsx,'.xlsx'), (hwpx,'.hwpx')])
def test_roundtrip_removes_pii_keeps_original(tmp_path, factory, suffix):
    source = factory(tmp_path / ('홍길동' + suffix))
    original = hashlib.sha256(source.read_bytes()).hexdigest()
    session = Session.load([source], Options())
    assert any(c.kind == 'name' and c.original == '홍길동' for c in session.candidates)
    result = session.export(tmp_path / 'out')
    assert len(result.files) == 1
    out = result.files[0]
    text = all_text(out)
    for value in ['홍길동', '900101-1234567', '비공개작성자', '미사용원문', 'OLD_PRIVATE_PREVIEW']:
        assert value not in text
    assert '대상자001' in text
    assert '홍길동' not in out.name
    assert hashlib.sha256(source.read_bytes()).hexdigest() == original
    report = (out.parent / '처리결과.json').read_text('utf8')
    assert '홍길동' not in report and str(source) not in report


def test_opt_out_and_edit_and_amount(tmp_path):
    session = Session.load([xlsx(tmp_path / 'a.xlsx')], Options(amounts=True))
    for c in session.candidates:
        if c.kind == 'name': c.replacement = '검토대상'
        if c.kind == 'rrn': c.selected = False
        if c.kind == 'amount': c.replacement = '0'
    text = all_text(session.export(tmp_path / 'out').files[0])
    assert '검토대상' in text and '900101-1234567' in text
    assert '1234500' not in text
    assert '<f>' not in text and 'sharedStrings.xml' not in text
    assert 'mergeCell' in text


def test_repeat_exports_never_overwrite(tmp_path):
    s = Session.load([xlsx(tmp_path / 'a.xlsx')], Options())
    first = s.export(tmp_path / 'out').files[0]
    second = s.export(tmp_path / 'out').files[0]
    assert first != second and first.exists() and second.exists()


def test_images_need_acknowledgement(tmp_path):
    s = Session.load([xlsx(tmp_path / 'a.xlsx', {'xl/media/image1.png': b'PNG'})], Options())
    with pytest.raises(DocumentError): s.export(tmp_path / 'out')
    assert s.export(tmp_path / 'out', acknowledge_images=True).files


def test_opaque_embedded_data_refused(tmp_path):
    with pytest.raises(DocumentError):
        Session.load([xlsx(tmp_path / 'a.xlsx', {'xl/embeddings/oleObject1.bin': b'private'})], Options())


@pytest.mark.parametrize('part', [
    'xl/connections.xml',
    'xl/queryTables/queryTable1.xml',
    'xl/dde/ddeLink1.xml',
    'xl/ctrlProps/ctrlProp1.xml',
    'xl/macrosheets/sheet1.xml',
    'xl/dialogsheets/sheet1.xml',
])
def test_unhandled_xlsx_package_parts_refused(tmp_path, part):
    with pytest.raises(DocumentError, match='지원하지 않습니다'):
        Session.load([xlsx(tmp_path / 'a.xlsx', {part: b'<x/>'})], Options())


def test_zip_traversal_refused(tmp_path):
    with pytest.raises(DocumentError):
        Session.load([package(tmp_path / 'a.hwpx', {'../x': 'bad'})], Options())


def test_xml_entities_refused(tmp_path):
    with pytest.raises(DocumentError):
        Session.load([xlsx(tmp_path / 'a.xlsx', {'xl/workbook.xml': '<!DOCTYPE x [<!ENTITY a SYSTEM "file:///etc/passwd">]><x>&a;</x>'})], Options())


def test_source_changed_since_preview_refused(tmp_path):
    source = xlsx(tmp_path / 'a.xlsx')
    s = Session.load([source], Options())
    source.write_bytes(b'changed')
    with pytest.raises(DocumentError): s.export(tmp_path / 'out')


def test_oversized_source_refused_before_reading(tmp_path, monkeypatch):
    source = xlsx(tmp_path / 'a.xlsx')
    session = Session.load([source], Options())
    real_stat = Path.stat

    def oversized_stat(self, *args, **kwargs):
        if self == source:
            return SimpleNamespace(st_size=MAX_FILE + 1)
        return real_stat(self, *args, **kwargs)

    def forbidden(*args, **kwargs):
        raise AssertionError('oversized source was read instead of refused')

    monkeypatch.setattr(Path, 'stat', oversized_stat)
    monkeypatch.setattr(Path, 'read_bytes', forbidden)
    monkeypatch.setattr(Path, 'open', forbidden)
    with pytest.raises(DocumentError, match='변경'):
        session.export(tmp_path / 'out')


def test_zero_detection_still_cleans_metadata(tmp_path):
    s = Session.load([xlsx(tmp_path / 'a.xlsx')], Options(names=False, rrn=False, birth=False))
    assert not s.candidates
    text = all_text(s.export(tmp_path / 'out').files[0])
    assert '홍길동' in text and '비공개작성자' not in text


def test_replacement_cannot_retain_original(tmp_path):
    s=Session.load([xlsx(tmp_path/'a.xlsx')],Options())
    for c in s.candidates:
        if c.kind=='name': c.replacement='홍길동'
    with pytest.raises(DocumentError): s.export(tmp_path/'out')
    assert not (tmp_path/'out').exists()


def test_formula_injection_stays_literal_text(tmp_path):
    s=Session.load([xlsx(tmp_path/'a.xlsx')],Options())
    for c in s.candidates:
        if c.kind=='name': c.replacement='=HYPERLINK("https://example.invalid","test")'
    text=all_text(s.export(tmp_path/'out').files[0])
    assert '<f>' not in text and '=HYPERLINK' in text


def test_numeric_rrn_scientific_notation(tmp_path):
    from tests.helpers import X
    sheet=f'<worksheet xmlns="{X}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>주민등록번호</t></is></c></row><row r="2"><c r="A2"><v>9.001011234567E+12</v></c></row></sheetData></worksheet>'
    s=Session.load([xlsx(tmp_path/'a.xlsx',{'xl/worksheets/sheet1.xml':sheet})],Options())
    assert any(c.kind=='rrn' for c in s.candidates)
    out=s.export(tmp_path/'out').files[0]
    assert '9.001011234567E+12' not in all_text(out)


def test_numeric_birth_and_rrn_leading_zero(tmp_path):
    from tests.helpers import X
    sheet=f'<worksheet xmlns="{X}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>생년월일</t></is></c><c r="B1" t="inlineStr"><is><t>주민등록번호</t></is></c></row><row r="2"><c r="A2"><v>10101</v></c><c r="B2"><v>101013123456</v></c></row></sheetData></worksheet>'
    s=Session.load([xlsx(tmp_path/'a.xlsx',{'xl/worksheets/sheet1.xml':sheet})],Options())
    assert {c.kind for c in s.candidates}=={'birth','rrn'}


def test_uninspected_text_parts_blocked(tmp_path):
    with pytest.raises(DocumentError):
        Session.load([xlsx(tmp_path/'a.xlsx',{'xl/private.txt':b'private copy'})],Options())


def test_internal_hyperlink_targets_removed(tmp_path):
    from tests.helpers import X
    sheet=f'<worksheet xmlns="{X}"><sheetData/><hyperlinks><hyperlink ref="A1" location="홍길동!A1"/></hyperlinks></worksheet>'
    s=Session.load([xlsx(tmp_path/'a.xlsx',{'xl/worksheets/sheet1.xml':sheet})],Options())
    assert '홍길동!A1' not in all_text(s.export(tmp_path/'out').files[0])


def test_native_bridge_returns_clear_error_off_windows(tmp_path):
    import sys
    if sys.platform=='win32': pytest.skip('Real Office integration is a manual acceptance test')
    from rename_app.windows import convert_legacy
    with pytest.raises(DocumentError,match='Windows'): convert_legacy(tmp_path/'a.hwp')
