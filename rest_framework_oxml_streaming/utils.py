# -*- coding: utf-8 -*-
from io import BytesIO
from xml.etree import ElementTree as ETree

import openpyxl
from openpyxl.writer.excel import save_virtual_workbook

OPENXML_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def create_xlsx_template(first_row):
    """Return an in memory xlsx template with one row, whose cells are taken
    from `first_row`.

    Args:
        first_row: OpenXMLSerializer instance
    Return:
        BytesIO, in memory xlsx file with one row
    """
    wb = openpyxl.Workbook()
    for col, data in enumerate(first_row.values()):
        cell = wb.active.cell(row=1, column=col + 1)
        cell.value = data.pop('value')
        for key, value in data.items():
            setattr(cell, key, value)
        data['value'] = cell.internal_value
    return BytesIO(save_virtual_workbook(wb))


def extract_column_attributes(xml):
    """Given an excel sheet xml search for the first row and extract each
    column attributes.

    .. note:: the cell identifier attribute (`r`) is not returned

    Args:
        xml: str, a valid excel sheet xml
    Return:
        a list of `{attribute: value}`
    """
    tree = ETree.fromstring(xml)
    sheet_data = tree.find('%ssheetData' % OPENXML_NS)
    first_row = sheet_data.find('%srow' % OPENXML_NS)
    column_attributes = [el.attrib for el in first_row.findall('%sc' % OPENXML_NS)]
    for attrib in column_attributes:
        attrib.pop('r')  # remove cell reference
    return _replace_reference_str_by_inline(column_attributes)

def _replace_reference_str_by_inline(column_attributes):
    ret = [d.copy() for d in column_attributes]
    for d in ret:
        if d.get('t') == 's':
            d['t'] = 'inlineStr'
    return ret
