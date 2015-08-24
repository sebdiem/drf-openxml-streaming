from django.utils.text import capfirst

from rest_framework import fields


class OpenXMLField(fields.Field):
    """Base field to serialize data for OpenXML

    Typically fields will be defined as built-in DRF fields but for "handmade"
    field we need a `verbose_name` to automatically get the column header from
    the serializer.
    """

    def __init__(self, *args, **kwargs):
        verbose_name = kwargs.pop('verbose_name')
        super(OpenXMLField, self).__init__(*args, **kwargs)
        self.label = capfirst(verbose_name)

    def to_representation(self, obj):
        raise NotImplementedError()

    def to_internal_value(self, data):
        raise NotImplementedError()
