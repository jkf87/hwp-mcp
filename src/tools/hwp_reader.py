"""
한글(HWP 5.0) 파일을 한글 프로그램 없이 직접 읽는 모듈
OLE 복합 문서 구조를 파싱하므로 Windows가 아닌 환경(macOS, Linux)에서도 동작합니다.
"""

import struct
import zlib
import logging
from typing import List, Optional, Tuple

import olefile

logger = logging.getLogger("hwp-reader")

# HWP 5.0 레코드 태그 (HWPTAG_BEGIN = 0x10)
TAG_PARA_HEADER = 0x10 + 50
TAG_PARA_TEXT = 0x10 + 51
TAG_CTRL_HEADER = 0x10 + 55
TAG_LIST_HEADER = 0x10 + 56
TAG_TABLE = 0x10 + 61

# 확장/인라인 컨트롤 문자는 8 WCHAR(16바이트)를 차지합니다.
_EXTENDED_CTRL_CHARS = {1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23}
_INLINE_CTRL_CHARS = {4, 5, 6, 7, 8, 9, 19, 20}

_FLAG_COMPRESSED = 0x01
_FLAG_PASSWORD = 0x02
_FLAG_DISTRIBUTION = 0x04


class HwpReadError(Exception):
    """HWP 파일을 읽을 수 없을 때 발생하는 예외"""


class _Record:
    __slots__ = ("tag", "level", "data", "children")

    def __init__(self, tag: int, level: int, data: bytes):
        self.tag = tag
        self.level = level
        self.data = data
        self.children: List["_Record"] = []


def _parse_records(data: bytes) -> List[_Record]:
    records = []
    pos = 0
    while pos + 4 <= len(data):
        header, = struct.unpack_from("<I", data, pos)
        pos += 4
        tag = header & 0x3FF
        level = (header >> 10) & 0x3FF
        size = (header >> 20) & 0xFFF
        if size == 0xFFF:
            size, = struct.unpack_from("<I", data, pos)
            pos += 4
        records.append(_Record(tag, level, data[pos:pos + size]))
        pos += size
    return records


def _build_tree(records: List[_Record]) -> List[_Record]:
    """레코드의 level 값을 이용해 트리를 구성하고 최상위 레코드 목록을 반환합니다."""
    roots: List[_Record] = []
    stack: List[_Record] = []
    for rec in records:
        while stack and stack[-1].level >= rec.level:
            stack.pop()
        (stack[-1].children if stack else roots).append(rec)
        stack.append(rec)
    return roots


def _decode_para_text(data: bytes) -> Tuple[str, int]:
    """
    PARA_TEXT 레코드를 문자열로 변환합니다.
    확장 컨트롤(표, 그림 등) 위치에는 '\\x00'을 남겨 이후 컨트롤 내용을 끼워 넣을 수 있게 합니다.

    Returns:
        (텍스트, 확장 컨트롤 개수)
    """
    out = []
    n_ctrl = 0
    i = 0
    n = len(data) // 2
    while i < n:
        ch, = struct.unpack_from("<H", data, i * 2)
        if ch >= 32:
            out.append(chr(ch))
            i += 1
        elif ch in _EXTENDED_CTRL_CHARS:
            out.append("\x00")
            n_ctrl += 1
            i += 8
        elif ch in _INLINE_CTRL_CHARS:
            if ch == 9:
                out.append("\t")
            i += 8
        else:
            if ch == 10:
                out.append("\n")
            elif ch in (24, 30, 31):  # 하이픈, 묶음 빈칸, 고정폭 빈칸
                out.append("-" if ch == 24 else " ")
            # 13(문단 끝) 등 나머지 문자 컨트롤은 무시
            i += 1
    return "".join(out), n_ctrl


def _escape_cell(text: str) -> str:
    return text.strip().replace("|", "\\|").replace("\n", "<br>")


class _Renderer:
    def __init__(self, as_markdown: bool):
        self.as_markdown = as_markdown
        self.table_count = 0

    def render_paragraphs(self, nodes: List[_Record]) -> List[str]:
        blocks = []
        for node in nodes:
            if node.tag == TAG_PARA_HEADER:
                blocks.extend(self._render_paragraph(node))
        return blocks

    def _render_paragraph(self, para: _Record) -> List[str]:
        text = ""
        for child in para.children:
            if child.tag == TAG_PARA_TEXT:
                text, _ = _decode_para_text(child.data)
                break
        ctrls = [c for c in para.children if c.tag == TAG_CTRL_HEADER]

        # 확장 컨트롤 위치를 기준으로 문단 텍스트와 컨트롤 내용을 순서대로 배치
        blocks: List[str] = []
        parts = text.split("\x00")
        for idx, part in enumerate(parts):
            if part.strip():
                blocks.append(part.rstrip())
            if idx < len(parts) - 1 and idx < len(ctrls):
                blocks.extend(self._render_ctrl(ctrls[idx]))
        # PARA_TEXT에 표시되지 않은 나머지 컨트롤
        for ctrl in ctrls[len(parts) - 1:]:
            blocks.extend(self._render_ctrl(ctrl))
        return blocks

    def _render_ctrl(self, ctrl: _Record) -> List[str]:
        ctrl_id = ctrl.data[:4][::-1].decode("latin-1") if len(ctrl.data) >= 4 else ""
        if ctrl_id == "tbl ":
            return [self._render_table(ctrl)]
        # 머리말/꼬리말/각주/글상자 등 내부 문단이 있는 컨트롤
        return self.render_paragraphs(ctrl.children)

    def _render_table(self, ctrl: _Record) -> str:
        self.table_count += 1
        rows = cols = 0
        cells = []  # (row, col, rowspan, colspan, text)
        current: Optional[list] = None
        for child in ctrl.children:
            if child.tag == TAG_TABLE and len(child.data) >= 8:
                rows, cols = struct.unpack_from("<HH", child.data, 4)
            elif child.tag == TAG_LIST_HEADER and len(child.data) >= 16:
                col, row, colspan, rowspan = struct.unpack_from("<HHHH", child.data, 8)
                current = [row, col, rowspan, colspan, []]
                cells.append(current)
            elif child.tag == TAG_PARA_HEADER and current is not None:
                current[4].extend(self._render_paragraph(child))

        if not cells:
            return ""
        rows = max(rows, max(c[0] + max(c[2], 1) for c in cells))
        cols = max(cols, max(c[1] + max(c[3], 1) for c in cells))
        grid = [["" for _ in range(cols)] for _ in range(rows)]
        for row, col, _, _, blocks in cells:
            if row < rows and col < cols:
                grid[row][col] = "\n".join(blocks)

        if not self.as_markdown:
            return "\n".join("\t".join(c.strip().replace("\n", " ") for c in r) for r in grid)

        lines = ["| " + " | ".join(_escape_cell(c) for c in grid[0]) + " |",
                 "|" + "---|" * cols]
        for r in grid[1:]:
            lines.append("| " + " | ".join(_escape_cell(c) for c in r) + " |")
        return "\n".join(lines)


def _open_ole(file_path: str) -> olefile.OleFileIO:
    if not olefile.isOleFile(file_path):
        raise HwpReadError("HWP 5.0(OLE) 형식이 아닙니다. HWPX 또는 HWP 3.x 파일은 지원하지 않습니다.")
    return olefile.OleFileIO(file_path)


def _read_flags(ole: olefile.OleFileIO) -> int:
    header = ole.openstream("FileHeader").read()
    if not header.startswith(b"HWP Document File"):
        raise HwpReadError("HWP 파일 시그니처가 올바르지 않습니다.")
    flags, = struct.unpack_from("<I", header, 36)
    return flags


def _section_streams(ole: olefile.OleFileIO) -> List[str]:
    sections = [e for e in ole.listdir() if len(e) == 2 and e[0] == "BodyText" and e[1].startswith("Section")]
    sections.sort(key=lambda e: int(e[1][len("Section"):] or 0))
    return ["/".join(e) for e in sections]


def read_hwp(file_path: str, as_markdown: bool = True) -> str:
    """
    HWP 파일의 본문 텍스트를 추출합니다. 표는 Markdown 표(또는 탭 구분 텍스트)로 변환합니다.

    Args:
        file_path: HWP 파일 경로
        as_markdown: True이면 표를 Markdown 표로, False이면 탭으로 구분된 텍스트로 출력

    Returns:
        str: 추출된 텍스트
    """
    ole = _open_ole(file_path)
    try:
        flags = _read_flags(ole)
        if flags & _FLAG_PASSWORD:
            raise HwpReadError("암호가 걸린 문서는 읽을 수 없습니다.")
        if flags & _FLAG_DISTRIBUTION:
            raise HwpReadError("배포용 문서는 읽을 수 없습니다.")

        renderer = _Renderer(as_markdown)
        blocks: List[str] = []
        for stream in _section_streams(ole):
            raw = ole.openstream(stream).read()
            if flags & _FLAG_COMPRESSED:
                raw = zlib.decompress(raw, -15)
            roots = _build_tree(_parse_records(raw))
            blocks.extend(renderer.render_paragraphs(roots))
        return "\n\n".join(b for b in blocks if b.strip())
    finally:
        ole.close()


def read_hwp_preview_text(file_path: str) -> str:
    """HWP 파일에 저장된 미리보기 텍스트(PrvText)를 반환합니다. 암호/배포용 문서에서도 동작할 수 있습니다."""
    ole = _open_ole(file_path)
    try:
        if not ole.exists("PrvText"):
            raise HwpReadError("미리보기 텍스트(PrvText)가 없습니다.")
        return ole.openstream("PrvText").read().decode("utf-16le", errors="replace").rstrip("\x00")
    finally:
        ole.close()
