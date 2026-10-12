# -*- coding: utf-8 -*-
from bika.lims import api
from plone.app.testing import setRoles
from plone.app.testing import TEST_USER_ID
from senaite.core.catalog import SETUP_CATALOG
from senaite.core.exportimport.setupdata import Sample_Types
from senaite.core.tests.base import BaseTestCase


class TestSampleTypeImportReuse(BaseTestCase):

    def setUp(self):
        super(TestSampleTypeImportReuse, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ["Manager"])
        self.folder = self.portal.setup.sampletypes
        self.catalog = api.get_tool(SETUP_CATALOG)
        self.importer = Sample_Types(self.portal)
        self.importer.context = self.portal

    def test_import_updates_type_created_by_earlier_sheet(self):
        slurry = api.create(self.folder, "SampleType", title=u"Slurry")
        self.importer.get_rows = lambda startrow: iter([{
            "title": u"Slurry", "description": u"Updated from workbook",
            "Hazardous": False, "Prefix": "SLR", "MinimumVolume": "250 g",
            "RetentionPeriod": 30,
        }])
        self.importer.Import()
        self.importer.Import()
        matches = self.catalog(portal_type="SampleType", title=u"Slurry")
        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0].getObject().UID(), slurry.UID())
        self.assertEqual(slurry.Description(), u"Updated from workbook")
        self.assertEqual(slurry.getPrefix(), "SLR")
        self.assertEqual(slurry.getMinimumVolume(), "250 g")
        self.assertEqual(slurry.getRetentionPeriod()["days"], 30)

    def test_word_match_is_not_reused_as_exact_type(self):
        water = api.create(self.folder, "SampleType", title=u"Solution (water)")
        result = self.importer.get_or_create_sample_type(
            self.folder, self.catalog, u"Solution", u"Exact type")
        self.assertNotEqual(result.UID(), water.UID())
        self.assertEqual(result.Title(), u"Solution")

    def test_real_duplicate_records_are_reported(self):
        api.create(self.folder, "SampleType", title=u"Slurry")
        api.create(self.folder, "SampleType", title=u"Slurry")
        with self.assertRaisesRegexp(ValueError, "multiple records.*Slurry"):
            self.importer.get_or_create_sample_type(
                self.folder, self.catalog, u"Slurry", u"")


def test_suite():
    from unittest import TestLoader
    return TestLoader().loadTestsFromTestCase(TestSampleTypeImportReuse)
