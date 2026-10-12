# -*- coding: utf-8 -*-
import unittest

from senaite.core.browser.form.adapters import data_import
from senaite.core.exportimport.load_setup_data import LoadSetupData


class Request(object):
    form = {}


class TestSetupImportFeedback(unittest.TestCase):

    def setUp(self):
        self.loader = data_import.LoadSetupData
        self.abort = data_import.transaction.abort
        self.log_error = data_import.logger.error
        self.aborted = []
        data_import.transaction.abort = lambda: self.aborted.append(True)
        data_import.logger.error = lambda message: None

    def tearDown(self):
        data_import.LoadSetupData = self.loader
        data_import.transaction.abort = self.abort
        data_import.logger.error = self.log_error

    def test_success_returns_success_message(self):
        data_import.LoadSetupData = lambda context, request: lambda: None
        adapter = data_import.EditForm(None, Request())
        self.assertTrue(adapter.handle_data_import())
        message = adapter.data["messages"][0]
        self.assertEqual(message["title"], "Success")
        self.assertEqual(message["level"], "success")
        self.assertFalse(self.aborted)

    def test_failure_returns_error_and_aborts(self):
        def fail():
            raise ValueError("Invalid setup workbook")
        data_import.LoadSetupData = lambda context, request: fail
        adapter = data_import.EditForm(None, Request())
        self.assertFalse(adapter.handle_data_import())
        self.assertEqual(len(adapter.data["messages"]), 1)
        message = adapter.data["messages"][0]
        self.assertEqual(message["level"], "danger")
        self.assertIn("Invalid setup workbook", message["message"])
        self.assertEqual(self.aborted, [True])

    def test_missing_workbook_raises_instead_of_redirecting(self):
        with self.assertRaises(ValueError):
            LoadSetupData(None, Request())()


def test_suite():
    return unittest.TestLoader().loadTestsFromTestCase(TestSetupImportFeedback)
