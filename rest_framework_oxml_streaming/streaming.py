# -*- coding: utf-8 -*-
# Copyright (c) 2011-2015 Polyconseil SAS. All rights reserved.

from __future__ import absolute_import, division, print_function, unicode_literals

import datetime
from itertools import chain
from xml.etree import ElementTree as ETree
from xml.dom.minidom import Text
import zipfile

from django.http import StreamingHttpResponse
from django.utils.translation import ugettext_lazy as _

import openpyxl
from openpyxl.cell import get_column_letter
from openpyxl.utils.datetime import to_excel as datetime_to_excel
from openpyxl.writer.excel import save_virtual_workbook

from rest_framework import metadata
from rest_framework import mixins
from rest_framework import renderers
from rest_framework import serializers
from rest_framework.decorators import list_route

import six
import zipstream

from . import utils

OPEN_XML_MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
OPEN_XML_FORMAT = 'xlsx'


class OpenXMLRenderer(renderers.BaseRenderer):
    """Serialize to OpenXML xlsx."""

    media_type = OPEN_XML_MEDIA_TYPE
    format = OPEN_XML_FORMAT
    charset = None
    render_style = 'binary'

    # the custom keys in renderer_context
    column_headers_key = 'column_headers'
    column_attributes_key = 'column_attributes'

    @staticmethod
    def render_header(renderer_context):
        # the header row is all string, no need for column_attributes:
        header_context = dict(
            (k, v)
            for k, v in six.iteritems(renderer_context)
            if k != OpenXMLRenderer.column_attributes_key
        )
        ws_namespace = (
            '<worksheet '
            'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            '">'
        )
        sheet_data_tag = '<sheetData>'
        column_headers = renderer_context.get(OpenXMLRenderer.column_headers_key)
        header_row = OpenXMLRenderer.render_row(data=column_headers, line=1, renderer_context=header_context)

        return ''.join([ws_namespace, sheet_data_tag, header_row])

    @staticmethod
    def render_cell(data, line, column, renderer_context):
        attr = renderer_context.get(OpenXMLRenderer.column_attributes_key, None)
        column = ETree.Element('c', attr[column] if attr else {})
        column.set('r', "{column}{line}".format(column=get_column_letter(column + 1), line=line))
        if attr is not None and attr[column].get('t') == 'n':
            # column is of type number in excel
            value = ETree.SubElement(column, 'v')
        else:
            # we use the inlineStr excel type to avoid references to another xml file
            value = ETree.SubElement(column, 'is').append(ETree.Element('t'))
        value.text = six.text_type(data)
        return column

    @staticmethod
    def render_row(data, line, renderer_context):
        if data:
            row = ETree.Element('row', r=six.text_type(line))
            for col, cell in enumerate(data):
                row.append(OpenXMLRenderer.render_cell(cell, line, col, renderer_context))
            return ETree.tostring(row)
        else:
            return ''

    @staticmethod
    def render_rows(data, start_line, renderer_context, return_next_line=False):
        ret = ''.join([
            OpenXMLRenderer.render_row(item.values(), start_line + i, renderer_context)
            for i, item in enumerate(data)
        ])
        if return_next_line:
            return ret, start_line + len(data)
        else:
            return ret

    @staticmethod
    def render_footer():
        return '</sheetData></worksheet>'

    @staticmethod
    def render_worksheet(data, renderer_context):
        yield OpenXMLRenderer.render_header(renderer_context).encode('utf-8')
        current_line = 2 if renderer_context.get(OpenXMLRenderer.column_headers_key) else 1
        for rows in data:
            data, current_line = OpenXMLRenderer.render_rows(
                rows, start_line=current_line, renderer_context=renderer_context, return_next_line=True)
            yield data.encode('utf-8')
        yield OpenXMLRenderer.render_footer().encode('utf-8')

    def render(self, data, accepted_media_type=None, renderer_context=None):
        """Generates a xlsx file whose rows are generated from the `data` iterator.

        We use `openpyxl` to generate a xlsx file based on the first row of data and
        then reuse the generated files to build the final xlsx.
        The returned file is a `zipstream.Zipfile` archive which can be streamed over
        Http without loading its content in memory (useful for huge datasets).

        Return: the zipstream.ZipFile archive representing the xlsx file
        """
        renderer_context = renderer_context or {}
        sheet_name = 'xl/worksheets/sheet1.xml'

        print(data)
        first_chunk = next(data)  # extract the first rows of data to determine column types and column orders
        data = chain([first_chunk], data)
        print(data)
        print(first_chunk)
        zip_template = zipfile.ZipFile(utils.create_xlsx_template(first_chunk[0]), mode='r')


        # Copy `zip_template` in a zipstream object, except the sheet data
        stream = zipstream.ZipFile(mode='w', compression=zipstream.ZIP_DEFLATED)
        for file_name in zip_template.namelist():
            if file_name != sheet_name:
                stream.write_iter(
                    arcname=file_name,
                    iterable=iter([zip_template.read(file_name)]),
                    compress_type=zipstream.ZIP_DEFLATED,
                )
            else:
                renderer_context[self.column_attributes_key] = utils.extract_column_attributes(
                    zip_template.read(file_name).decode('utf-8'))

        # Write our data to the stream
        stream.write_iter(
            arcname=sheet_name,
            iterable=self.render_worksheet(data, renderer_context),
            compress_type=zipstream.ZIP_DEFLATED
        )
        return stream


class DynamicMetadata(metadata.SimpleMetadata):
    """This metadata class uses `get_renderers` instead of the class attribute to
    answer correctly to `OPTIONS` requests."""
    def determine_metadata(self, request, view):
        ret = super(DynamicMetadata, self).determine_metadata(request, view)
        ret['renders'] = [renderer.media_type for renderer in view.get_renderers()]
        return ret


class OpenXMLListMixin(mixins.ListModelMixin):
    """A view that can serve a list as a streamed Open XML `.xlsx` document in addition
    to other DRST default renderers.

    It requires an additional class attribute `open_xml_serializer_class` which defines
    the serializer to be used for the Open XML export.
    If this attribute is present, then the `OpenXMLRenderer` is automatically added to
    the view's list of supported renderers.
    """
    metadata_class = DynamicMetadata
    streaming_page_size = 1000

    def __init__(self, *args, **kwargs):
        if (
            not hasattr(self.__class__, 'open_xml_serializer_class') or
            not hasattr(self.open_xml_serializer_class(), 'column_headers')
        ):
            raise AttributeError('OpenXMLListMixin shall define an open_xml_serializer_class attribute')
        super(OpenXMLListMixin, self).__init__(*args, **kwargs)

    def get_serializer_class(self, *args, **kwargs):
        accepted_renderer = getattr(self.request, 'accepted_renderer', None)
        if accepted_renderer and accepted_renderer.media_type == OPEN_XML_MEDIA_TYPE:
            return self.open_xml_serializer_class
        else:
            return super(OpenXMLListMixin, self).get_serializer_class(*args, **kwargs)

    def get_renderers(self):
        ret = [renderer for renderer in super(OpenXMLListMixin, self).get_renderers()]
        if self.suffix and self.suffix.lower() == 'list':
            ret.append(OpenXMLRenderer())
        return ret

    def data_stream(self, request):
        serializer_class = self.open_xml_serializer_class
        queryset = self.filter_queryset(self.get_queryset())
        paginate_by = self.streaming_page_size
        for start in range(0, queryset.count(), paginate_by):
            end = start + paginate_by
            yield serializer_class(queryset[start:end], many=True, context={'request': request}).data

    def list_as_open_xml(self, request):
        column_headers = utils.get_column_headers(self.open_xml_serializer_class)
        response = StreamingHttpResponse(
            OpenXMLRenderer().render(
                data=self.data_stream(request),
                renderer_context={OpenXMLRenderer.column_headers_key: column_headers},
            ),
            content_type=request.accepted_media_type,
        )
        response['Content-Disposition'] = 'attachment; filename={0}.{1}'.format(
            self.get_view_name(),
            OpenXMLRenderer.format,
        )
        return response

    def list(self, request, *args, **kwargs):
        accepted_renderer = getattr(request, 'accepted_renderer', None)
        if accepted_renderer and accepted_renderer.media_type == OPEN_XML_MEDIA_TYPE:
            return self.list_as_open_xml(request)
        else:
            return super(OpenXMLListMixin, self).list(request, *args, **kwargs)

    @list_route(methods=['get'], url_path='download')
    def download_list(self, request):
        return self.list_as_open_xml(request)
