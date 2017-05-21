import re
import struct
import zlib
from io import BytesIO
from itertools import islice

import olefile

HWPTAG_BEGIN = 0x10
tag_table = {
    HWPTAG_BEGIN: 'HWPTAG_DOCUMENT_PROPERTIES',
    HWPTAG_BEGIN + 1: 'HWPTAG_ID_MAPPINGS',
    HWPTAG_BEGIN + 2: 'HWPTAG_BIN_DATA',
    HWPTAG_BEGIN + 3: 'HWPTAG_FACE_NAME',
    HWPTAG_BEGIN + 4: 'HWPTAG_BORDER_FILL',
    HWPTAG_BEGIN + 5: 'HWPTAG_CHAR_SHAPE',
    HWPTAG_BEGIN + 6: 'HWPTAG_TAB_DEF',
    HWPTAG_BEGIN + 7: 'HWPTAG_NUMBERING',
    HWPTAG_BEGIN + 8: 'HWPTAG_BULLET',
    HWPTAG_BEGIN + 9: 'HWPTAG_PARA_SHAPE',
    HWPTAG_BEGIN + 10: 'HWPTAG_STYLE',
    HWPTAG_BEGIN + 11: 'HWPTAG_DOC_DATA',
    HWPTAG_BEGIN + 12: 'HWPTAG_DISTRIBUTE_DOC_DATA',
    HWPTAG_BEGIN + 13: 'HWPTAG__RESERVED',
    HWPTAG_BEGIN + 14: 'HWPTAG_COMPATIBLE_DOCUMENT',
    HWPTAG_BEGIN + 15: 'HWPTAG_LAYOUT_COMPATIBILITY',
    HWPTAG_BEGIN + 16: 'HWPTAG_TRACKCHANGE',
    HWPTAG_BEGIN + 50: 'HWPTAG_PARA_HEADER',
    HWPTAG_BEGIN + 51: 'HWPTAG_PARA_TEXT',
    HWPTAG_BEGIN + 52: 'HWPTAG_PARA_CHAR_SHAPE',
    HWPTAG_BEGIN + 53: 'HWPTAG_PARA_LINE_SEG',
    HWPTAG_BEGIN + 54: 'HWPTAG_PARA_RANGE_TAG',
    HWPTAG_BEGIN + 55: 'HWPTAG_CTRL_HEADER',
    HWPTAG_BEGIN + 56: 'HWPTAG_LIST_HEADER',
    HWPTAG_BEGIN + 57: 'HWPTAG_PAGE_DEF',
    HWPTAG_BEGIN + 58: 'HWPTAG_FOOTNOTE_SHAPE',
    HWPTAG_BEGIN + 59: 'HWPTAG_PAGE_BORDER_FILL',
    HWPTAG_BEGIN + 60: 'HWPTAG_SHAPE_COMPONENT',
    HWPTAG_BEGIN + 61: 'HWPTAG_TABLE',
    HWPTAG_BEGIN + 62: 'HWPTAG_SHAPE_COMPONENT_LINE',
    HWPTAG_BEGIN + 63: 'HWPTAG_SHAPE_COMPONENT_RECTANGLE',
    HWPTAG_BEGIN + 64: 'HWPTAG_SHAPE_COMPONENT_ELLIPSE',
    HWPTAG_BEGIN + 65: 'HWPTAG_SHAPE_COMPONENT_ARC',
    HWPTAG_BEGIN + 66: 'HWPTAG_SHAPE_COMPONENT_POLYGON',
    HWPTAG_BEGIN + 67: 'HWPTAG_SHAPE_COMPONENT_CURVER',
    HWPTAG_BEGIN + 68: 'HWPTAG_SHAPE_COMPONENT_OLE',
    HWPTAG_BEGIN + 69: 'HWPTAG_SHAPE_COMPONENT_PICTURE',
    HWPTAG_BEGIN + 70: 'HWPTAG_SHAPE_COMPONENT_CONTAINER',
    HWPTAG_BEGIN + 71: 'HWPTAG_CTRL_DATA',
    HWPTAG_BEGIN + 72: 'HWPTAG_EQEDIT',
    HWPTAG_BEGIN + 73: 'HWPTAG_RESERVED',
    HWPTAG_BEGIN + 74: 'HWPTAG_SHAPE_COMPONENT_TEXTART',
    HWPTAG_BEGIN + 75: 'HWPTAG_FORM_OBJECT',
    HWPTAG_BEGIN + 76: 'HWPTAG_MEMO_SHAPE',
    HWPTAG_BEGIN + 77: 'HWPTAG_MEMO_LIST',
    HWPTAG_BEGIN + 76: 'HWPTAG_MEMO_SHAPE',
    HWPTAG_BEGIN + 78: 'HWPTAG_FORBIDDEN_CHAR',
    HWPTAG_BEGIN + 79: 'HWPTAG_CHART_DATA',
    HWPTAG_BEGIN + 80: 'HWPTAG_TRACK_CHANGE',
    HWPTAG_BEGIN + 81: 'HWPTAG_TRACK_CHANGE_AUTHOR',
    HWPTAG_BEGIN + 82: 'HWPTAG_VIDEO_DATA',
    HWPTAG_BEGIN + 99: 'HWPTAG_SHAPE_COMPONENT_UNKNOWN',
}

class char(object):
    size = 1

class inline(object):
    size = 8

class extended(object):
    size = 8

control_char_table = {
    0x00: ('UNUSABLE', char),
    0x01: ('RESERVED0', extended),
    0x02: ('SECTION_OR_COLUMN_DEF', extended),
    0x03: ('FIELD_START', extended),
    0x04: ('FIELD_END', inline),
    0x05: ('RESERVED1', inline),
    0x06: ('RESERVED2', inline),
    0x07: ('RESERVED3', inline),
    0x08: ('TITLE_MARK', inline),
    0x09: ('TAB', inline),
    0x0a: ('LINE_BREAK', char),
    0x0b: ('DRAWING_OR_TABLE', extended),
    0x0c: ('RESERVED4', extended),
    0x0d: ('PARA_BREAK', char),
    0x0e: ('RESERVED5', extended),
    0x0f: ('HIDDEN_EXPLANATION', extended),
    0x10: ('HEADER_OR_FOOTER', extended),
    0x11: ('FOOTNOTE_OR_ENDNOTE', extended),
    0x12: ('AUTO_NUMBERING', extended),
    0x13: ('RESERVED6', inline),
    0x14: ('RESERVED7', inline),
    0x15: ('PAGE_CONTROL', extended),
    0x16: ('BOOKMARK', extended),
    0x17: ('DUTMAL_OR_CHAR_OVERLAP', extended),
    0x18: ('HYPEN', char),
    0x19: ('RESERVED8', char),
    0x1a: ('RESERVED9', char),
    0x1b: ('RESERVED10', char),
    0x1c: ('RESERVED11', char),
    0x1d: ('RESERVED12', char),
    0x1e: ('NONBREAK_SPACE', char),
    0x1f: ('FIXEDWIDTH_SPACE', char),
}

class Record(object):
    def __init__(self, tag_id, payload, parent=None):
        self.parent = parent
        self.children = []

        self.tag_id = tag_id
        self.tag_name = tag_table.get(self.tag_id, '<ROOT>')
        self.payload = payload

    def __repr__(self):
        return '<Record %s>' % self.tag_name

    def get_next_siblings(self, count=None):
        start_idx = self.parent.children.index(self) + 1
        if count is None:
            end_idx = None
        else:
            end_idx = start_idx + count

        return islice(self.parent.children, start_idx, end_idx)

    @staticmethod
    def build_tree_from_stream(stream):
        root = Record(None, None)

        while True:
            header = stream.read(4)
            if not header:
                break

            header = struct.unpack('<I', header)[0]

            tag_id = header & 0x3ff
            level = (header >> 10) & 0x3ff
            size = (header >> 20) & 0xfff

            if size == 0xfff:
                size = struct.unpack('<I', stream.read(4))[0]

            payload = stream.read(size)

            last_record = root
            for _ in range(level):
                last_record = last_record.children[-1]

            last_record.children.append(Record(tag_id, payload, last_record))

        return root

def get_paragraph_text(record):
    regex = re.compile(rb'([\x00-\x1f])\x00')

    length = len(record.payload)
    text = ''

    idx = 0
    while idx < length:
        searched = regex.search(record.payload, idx)
        if searched:
            control_char_pos = searched.start()

            if control_char_pos & 1:
                idx = control_char_pos + 1
            elif control_char_pos > idx:
                text += record.payload[idx:control_char_pos].decode('utf-16')
                idx = control_char_pos
            else:
                control_char = ord(searched.group(1))
                control_char_size = control_char_table[control_char][1].size

                if control_char == 0x0a:
                    text += '\n'

                idx = control_char_pos + control_char_size * 2
        else:
            text += record.payload[idx:].decode('utf-16')
            break

    return text

class Table(object):
    def __init__(self, caption, row_cnt, col_cnt):
        self.caption = caption
        self.row_cnt = row_cnt
        self.col_cnt = col_cnt

        self.rows = [[] for _ in range(row_cnt)]

    def __repr__(self):
        return '<Table %s>' % self.caption

class TableCell(object):
    def __init__(self, lines, row, col, row_span, col_span):
        self.lines = lines
        self.row = row
        self.col = col
        self.row_span = row_span
        self.col_span = col_span

    def __repr__(self):
        return '<TableCell(%d, %d) %s>' % (self.row, self.col, self.lines)

def make_tables(record_tree_root):
    def traverse(record, depth=0):
        # print('  ' * depth + repr(record))
        # if (record.tag_name == 'HWPTAG_PARA_TEXT'
        #     and record.parent.parent.tag_name == '<ROOT>'
        #     and record.payload[0] != 0x0b):
        #     ctx['table_caption'] = get_paragraph_text(record).strip()
        if record.tag_name == 'HWPTAG_TABLE':
            if 'current_table_idx' not in ctx:
                ctx['current_table_idx'] = 0
            else:
                ctx['current_table_idx'] += 1

            row_cnt = struct.unpack('<H', record.payload[4:6])[0]
            col_cnt = struct.unpack('<H', record.payload[6:8])[0]

            ctx['tables'].append(Table(None, row_cnt, col_cnt))
            # ctx['tables'].append(Table(ctx['table_caption'], row_cnt, col_cnt))
        elif (record.tag_name == 'HWPTAG_LIST_HEADER'
              and record.parent.tag_name == 'HWPTAG_CTRL_HEADER'
              and record.parent.payload[:4][::-1] == b'tbl '):
            paragraph_count = struct.unpack('<H', record.payload[:2])[0]
            col = struct.unpack('<H', record.payload[8:10])[0]
            row = struct.unpack('<H', record.payload[10:12])[0]
            col_span = struct.unpack('<H', record.payload[12:14])[0]
            row_span = struct.unpack('<H', record.payload[14:16])[0]

            lines = []
            for sibling in record.get_next_siblings(paragraph_count):
                for child in sibling.children:
                    if child.tag_name == 'HWPTAG_PARA_TEXT':
                        lines.extend(get_paragraph_text(child).strip().splitlines())
                        break

            ctx['tables'][ctx['current_table_idx']].rows[row].append(
                TableCell(lines, row, col, row_span, col_span)
            )

        for child in record.children:
            traverse(child, depth + 1)

    ctx = {'tables': []}
    traverse(record_tree_root)

    return ctx['tables']

class ZlibDecompressStream(object):
    def __init__(self, stream, wbits=15, chunk_size=4096):
        self._stream = stream
        self._decompressor = zlib.decompressobj(wbits)
        self.chunk_size = chunk_size
        self.buffer = b''

    def read(self, size):
        while len(self.buffer) < size and not self._decompressor.eof:
            chunk = self._decompressor.unconsumed_tail
            if not chunk:
                chunk = self._stream.read(self.chunk_size)
                if not chunk:
                    break

            self.buffer += self._decompressor.decompress(chunk, self.chunk_size)

        result = self.buffer[:size]
        self.buffer = self.buffer[size:]

        return result

class HwpFile(object):
    def __init__(self, file):
        self.ole = olefile.OleFileIO(file)

    @property
    def compressed(self):
        if not hasattr(self, '_compressed'):
            with self.ole.openstream('FileHeader') as stream:
                stream.seek(36)
                flag = struct.unpack('<I', stream.read(4))[0]
                self._compressed = bool(flag & 1)

        return self._compressed

    def get_record_tree(self, section_idx):
        with self.ole.openstream('BodyText/Section%d' % section_idx) as stream:
            if self.compressed:
                stream = ZlibDecompressStream(stream, -15)

            record_tree_root = Record.build_tree_from_stream(stream)

        return record_tree_root