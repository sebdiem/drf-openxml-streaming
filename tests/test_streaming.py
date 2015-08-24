from collections import namedtuple
import datetime

from django.db import models
from django.test import TestCase

from rest_framework import serializers as drf_serializers

from rest_framework_oxml_streaming import serializers
from rest_framework_oxml_streaming import streaming

class TestModel(models.Model):
    field_one = models.TextField()
    field_two = models.IntegerField()
    field_three = models.DateTimeField()

class TestSerializer(serializers.OpenXMLSerializer):
    field_one = drf_serializers.CharField()
    field_two = drf_serializers.IntegerField()
    field_three = drf_serializers.ReadOnlyField()
    field_four = serializers.CurrencyField()

class StreamingTestCase(TestCase):
    def test_streaming(self):
        def test_iterator():
            TestTuple = namedtuple('Test', 'field_one field_two field_three field_four')
            data = [TestTuple('Test string', 2, datetime.datetime.now(), 100)]
            for _i in range(10):
                yield TestSerializer(data, many=True).data
        streaming.OpenXMLRenderer().render(
            data=test_iterator(),
            renderer_context={'column_headers': ['field one', 'field two', 'field three', 'field four']}
        )
