from functools import wraps

from django.utils.translation import ugettext_lazy as _

from openpyxl import Workbook
from openpyxl.cell import Cell

from rest_framework import serializers


def openpyxl_decorator(f, openpyxl_args=None):
    @wraps(f)
    def wrapper(*args, **kwargs):
        cell = Cell(worksheet=Workbook().active, column=0, row=0, value=f(*args, **kwargs))
        ret = openpyxl_args.copy() if openpyxl_args else {}
        ret.update(value=cell.internal_value)
        return ret
    return wrapper


class OpenXMLSerializer(serializers.Serializer):
    def __init__(self, *args, **kwargs):
        super(OpenXMLSerializer, self).__init__(*args, **kwargs)
        for label, field in self.fields.items():
            field.to_representation = openpyxl_decorator(field.to_representation, getattr(field, 'openpyxl_args', None))

    def get_column_headers(self):
        return dict((f.field_name, _(f.label)) for f in self.fields.values())


class CurrencyField(serializers.ReadOnlyField):
    def __init__(self, *args, **kwargs):
        super(CurrencyField, self).__init__(*args, **kwargs)
        self.openpyxl_args = {'style': '%%%as'}

