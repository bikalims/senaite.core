# -*- coding: utf-8 -*-

import unittest

from bika.lims.interfaces import IRoutineAnalysis
from senaite.core.browser.worksheets.worksheet.referencesamples import \
    ReferenceSamplesView
from zope.interface import implementer


class Worksheet(object):

    def __init__(self, analyses=None, template=None):
        self.analyses = analyses or []
        self.template = template

    def getAnalyses(self):
        return self.analyses

    def getWorksheetTemplate(self):
        return self.template


class WorksheetTemplate(object):

    def __init__(self, services):
        self.services = services

    def getServices(self):
        return self.services


@implementer(IRoutineAnalysis)
class RoutineAnalysis(object):

    def __init__(self, service):
        self.service = service

    def getAnalysisService(self):
        return self.service


class TestReferenceSamplesView(unittest.TestCase):

    def make_view(self, context):
        view = ReferenceSamplesView.__new__(ReferenceSamplesView)
        view.context = context
        return view

    def test_template_services_are_used_for_qc_only_worksheet(self):
        services = [object(), object()]
        template = WorksheetTemplate(services)
        view = self.make_view(Worksheet(template=template))

        self.assertEqual(services, view.get_assigned_services())

    def test_routine_analysis_services_take_precedence(self):
        routine_service = object()
        template_service = object()
        analysis = RoutineAnalysis(routine_service)
        template = WorksheetTemplate([template_service])
        view = self.make_view(Worksheet([analysis], template))

        self.assertEqual([routine_service], view.get_assigned_services())


def test_suite():
    suite = unittest.TestSuite()
    suite.addTest(unittest.makeSuite(TestReferenceSamplesView))
    return suite
