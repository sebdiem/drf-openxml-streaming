# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from functools import wraps

from django.utils.text import capfirst
from django.utils.translation import ugettext_lazy as _

from openpyxl import Workbook
from openpyxl.cell import Cell

from rest_framework import serializers


def openpyxl_decorator(f, openpyxl_args=None, convert=True):
    @wraps(f)
    def wrapper(*args, **kwargs):
        value = f(*args, **kwargs)
        cell = Cell(worksheet=Workbook().active, column=0, row=0, value=value)
        ret = openpyxl_args.copy() if openpyxl_args else {}
        ret.update(value=cell.internal_value if convert else value)
        return ret
    return wrapper


class OpenXMLSerializer(serializers.Serializer):
    """
    kwargs:
        convert: boolean, optional, if False the value is not converted to
        OpenXML internal format but kept as a python datatype. Defaults to
        True.
    """
    def __init__(self, *args, **kwargs):
        convert = kwargs.pop('convert', True)
        super(OpenXMLSerializer, self).__init__(*args, **kwargs)
        for label, field in self.fields.items():
            field.to_representation = openpyxl_decorator(
                field.to_representation,
                openpyxl_args=getattr(field, 'openpyxl_args', None),
                convert=convert,
            )

    def get_column_headers(self):
        return dict((f.field_name, _(f.label)) for f in self.fields.values())


class OpenXMLField(serializers.Field):
    """Base field to serialize data for OpenXML

    Typically fields will be defined as built-in DRF fields but for "handmade"
    field we need a `verbose_name` to automatically get the column header from
    the serializer.
    """

    def __init__(self, *args, **kwargs):
        verbose_name = kwargs.pop('verbose_name')
        super(OpenXMLField, self).__init__(*args, **kwargs)
        self.label = capfirst(verbose_name)


class CurrencyField(OpenXMLField, serializers.ReadOnlyField):
    def __init__(self, currency, *args, **kwargs):
        super(CurrencyField, self).__init__(*args, **kwargs)
        self.openpyxl_args = {'number_format': r'#,##0.00_)"{symb}";\-#,##0.00"{symb}"'.format(symb=currency)}
