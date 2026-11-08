import unittest

from bika.lims.browser import workflow
from bika.lims.interfaces import IDuplicateAnalysis
from senaite.core.browser.workflow.worksheet import WorkflowActionUnassignAdapter
from zope.interface import alsoProvides


class Item(object):

    def __init__(self, uid, source=None):
        self.uid = uid
        self.source = source
        if source:
            alsoProvides(self, IDuplicateAnalysis)

    def getAnalysis(self):
        return self.source


class TestWorkflowBulkActions(unittest.TestCase):

    def setUp(self):
        self.original_uid = workflow.api.get_uid
        self.original_action = workflow.WorkflowActionGenericAdapter.do_action
        workflow.api.get_uid = lambda obj: obj.uid
        workflow.WorkflowActionGenericAdapter.do_action = (
            lambda adapter, action, objects: objects)
        self.adapter = object.__new__(WorkflowActionUnassignAdapter)
        self.routine = Item('routine')
        self.duplicate = Item('duplicate', self.routine)
        self.other = Item('other')

    def tearDown(self):
        workflow.api.get_uid = self.original_uid
        workflow.WorkflowActionGenericAdapter.do_action = self.original_action

    def test_source_and_duplicate_selected(self):
        result = self.adapter.do_action(
            'unassign', [self.routine, self.duplicate, self.other])
        self.assertEqual(result, [self.routine, self.other])

    def test_duplicate_before_source(self):
        result = self.adapter.do_action(
            'unassign', [self.duplicate, self.routine])
        self.assertEqual(result, [self.routine])

    def test_duplicate_only_selected(self):
        result = self.adapter.do_action('unassign', [self.duplicate])
        self.assertEqual(result, [self.duplicate])

    def test_unrelated_source_selected(self):
        result = self.adapter.do_action(
            'unassign', [self.other, self.duplicate])
        self.assertEqual(result, [self.other, self.duplicate])


def test_suite():
    return unittest.defaultTestLoader.loadTestsFromTestCase(
        TestWorkflowBulkActions)
