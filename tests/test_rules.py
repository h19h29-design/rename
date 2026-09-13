import pytest
from rename_app.rules import Options, detect, replacement, apply_text, Hit


def kinds(text, hint='', **kwargs):
    return [(h.kind, h.value) for h in detect(text, hint, Options(**kwargs))]


def test_labeled_name_and_repeat():
    assert ('name', '홍길동') in kinds('성명: 홍길동 / 확인')


def test_name_header_and_manual_english():
    assert ('name', '남궁민수') in kinds('남궁민수', '성명')
    assert ('name', 'Alice Kim') in kinds('서명 Alice Kim 확인', manual_names=('Alice Kim',))


def test_rrn_precedes_birthday_overlap():
    found = kinds('900101-1234567 9001011234567', dates=True)
    assert found == [('rrn', '900101-1234567'), ('rrn', '9001011234567')]


def test_masked_rrn_detected():
    assert ('rrn', '900101-1******') in kinds('900101-1******')


def test_invalid_calendar_date_not_selected():
    assert not kinds('2026-02-31 20261301', dates=True)


def test_birth_context_not_general_dates_by_default():
    assert ('birth', '1990. 1. 2.') in kinds('생년월일: 1990. 1. 2.')
    assert not kinds('계약일 2026-09-13')


def test_compact_birth_and_date_option():
    assert ('birth', '900102') in kinds('900102', '생년월일')
    assert ('date', '2026년 9월 13일') in kinds('계약일 2026년 9월 13일', dates=True)


def test_amount_opt_in_ignores_unlabelled_number():
    assert not kinds('1,234,500원')
    assert ('amount', '1,234,500원') in kinds('1,234,500원', amounts=True)
    assert ('amount', '1234500') in kinds('1234500', '금액', amounts=True)
    assert not kinds('1234500', amounts=True)


def test_options_disable_default_categories():
    assert not kinds('성명 홍길동 900101-1234567 생년월일 1990-01-01', names=False, rrn=False, birth=False)


def test_replacements_and_non_cascading_edits():
    assert replacement('name', 1, 'token') == '대상자001'
    assert replacement('rrn', 1, 'mask') == '******-*******'
    assert replacement('name', 1, 'delete') == ''
    hits = [Hit('name', '홍길동', 0, 3), Hit('name', '김민수', 4, 7)]
    assert apply_text('홍길동 김민수', [(hits[0], '김민수'), (hits[1], '대상자002')]) == '김민수 대상자002'


def test_spaced_korean_name_and_manual_variant():
    assert ('name','홍 길 동') in kinds('홍 길 동','성 명')
    assert ('name','홍 길 동') in kinds('홍 길 동 님',manual_names=('홍길동',))


def test_mixed_birth_and_contract_dates():
    assert kinds('생년월일 1990-01-01 / 계약일 2026-09-13') == [('birth','1990-01-01')]


def test_rrn_disabled_does_not_partially_replace_as_date():
    assert not kinds('2001011234567',rrn=False,dates=True)
