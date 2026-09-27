#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Tests for HWP Reader (Windows/한글 프로그램 없이 실행 가능)
"""

import struct

from src.tools.hwp_reader import (
    TAG_CTRL_HEADER, TAG_LIST_HEADER, TAG_PARA_HEADER, TAG_PARA_TEXT, TAG_TABLE,
    _Renderer, _build_tree, _decode_para_text, _parse_records,
)


def _record(tag, level, data=b""):
    size = len(data)
    if size >= 0xFFF:
        return struct.pack("<II", tag | (level << 10) | (0xFFF << 20), size) + data
    return struct.pack("<I", tag | (level << 10) | (size << 20)) + data


def _text(s):
    return s.encode("utf-16le")


def _ctrl_char(code, ctrl_id):
    # 확장 컨트롤: 코드 + 컨트롤 ID(4바이트) + 예약(8바이트) + 코드 = 8 WCHAR
    return struct.pack("<H", code) + ctrl_id[::-1] + b"\x00" * 8 + struct.pack("<H", code)


def _paragraph(level, text_bytes):
    return _record(TAG_PARA_HEADER, level, b"\x00" * 24) + _record(TAG_PARA_TEXT, level + 1, text_bytes)


def _cell(level, row, col, text, rowspan=1, colspan=1):
    header = struct.pack("<HHIHHHH", 1, 0, 0, col, row, colspan, rowspan)
    return _record(TAG_LIST_HEADER, level, header) + _paragraph(level, _text(text + "\r"))


def _render(stream, as_markdown=True):
    renderer = _Renderer(as_markdown)
    return "\n\n".join(renderer.render_paragraphs(_build_tree(_parse_records(stream))))


def test_decode_para_text_handles_control_chars():
    data = _text("A") + _ctrl_char(11, b"tbl ") + _text("B") + struct.pack("<H", 10) + _text("C\r")
    text, n_ctrl = _decode_para_text(data)
    assert text == "A\x00B\nC"
    assert n_ctrl == 1


def test_parse_records_extended_size():
    payload = b"x" * 5000
    records = _parse_records(_record(TAG_PARA_TEXT, 1, payload))
    assert len(records) == 1
    assert records[0].data == payload


def test_render_plain_paragraphs():
    stream = _paragraph(0, _text("첫 문단\r")) + _paragraph(0, _text("둘째 문단\r"))
    assert _render(stream) == "첫 문단\n\n둘째 문단"


def test_render_table_with_merged_cell():
    table = (
        _record(TAG_CTRL_HEADER, 1, b" lbt" + b"\x00" * 40)
        + _record(TAG_TABLE, 2, struct.pack("<IHH", 0, 2, 2) + b"\x00" * 16)
        + _cell(2, 0, 0, "항목", colspan=2)
        + _cell(2, 1, 0, "물")
        + _cell(2, 1, 1, "120,000")
    )
    stream = (
        _paragraph(0, _text("앞") + _ctrl_char(11, b"tbl ") + _text("뒤\r"))
        + table
    )
    # CTRL_HEADER는 문단(PARA_HEADER)의 자식이어야 하므로 PARA_TEXT 뒤에 이어 붙인다
    assert _render(stream) == "앞\n\n| 항목 |  |\n|---|---|\n| 물 | 120,000 |\n\n뒤"
    assert _render(stream, as_markdown=False) == "앞\n\n항목\t\n물\t120,000\n\n뒤"
