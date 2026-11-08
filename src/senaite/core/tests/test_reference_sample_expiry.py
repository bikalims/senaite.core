from DateTime import DateTime
from bika.lims import api
from bika.lims.browser.referencesample import ReferenceSamplesView
from plone.app.testing import TEST_USER_ID
from plone.app.testing import setRoles
from senaite.core.browser.controlpanel.suppliers.referencesamples import \
    SupplierReferenceSamplesView
from senaite.core.tests.base import BaseTestCase


class TestReferenceSampleExpiry(BaseTestCase):

    def setUp(self):
        super(TestReferenceSampleExpiry, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ["Manager"])
        self.supplier = api.create(
            self.portal.setup.suppliers, "Supplier", Name="Supplier")
        self.past = api.create(
            self.supplier, "ReferenceSample", title="Past",
            ExpiryDate=DateTime() - 1)
        self.future = api.create(
            self.supplier, "ReferenceSample", title="Future",
            ExpiryDate=DateTime() + 1)
        self.undated = api.create(
            self.supplier, "ReferenceSample", title="Undated")

    def listing(self, state="default", supplier=None):
        self.request.form["list_review_state"] = state
        cls = SupplierReferenceSamplesView if supplier else ReferenceSamplesView
        context = supplier or self.portal.bika_setup.referencesamples
        view = cls(context, self.request)
        view.workflow = api.get_tool("portal_workflow")
        view.portal = self.portal
        view.mtool = api.get_tool("portal_membership")
        view.member = api.get_current_user()
        return view.get_folderitems()

    def test_current_excludes_overdue_and_keeps_future_and_undated(self):
        items = self.listing()
        uids = [item["uid"] for item in items]
        self.assertNotIn(api.get_uid(self.past), uids)
        self.assertIn(api.get_uid(self.future), uids)
        self.assertIn(api.get_uid(self.undated), uids)
        self.assertEqual(api.get_workflow_status_of(self.past), "expired")

    def test_opening_expired_tab_first_expires_and_displays_sample(self):
        items = self.listing("expired")
        self.assertEqual(len(items), 1)
        item = items[0]
        self.assertEqual(item["uid"], api.get_uid(self.past))
        self.assertEqual(item["review_state"], "expired")
        self.assertEqual(item["state_title"], "Expired")
        self.assertIn("exclamation", item["after"]["ExpiryDate"])
        self.assertIn('title="Expired"', item["after"]["ExpiryDate"])

    def test_all_tab_shows_fresh_state_and_no_future_expiry_icon(self):
        items = dict((item["uid"], item) for item in self.listing("all"))
        self.assertEqual(items[api.get_uid(self.past)]["state_title"], "Expired")
        self.assertNotIn("ExpiryDate", items[api.get_uid(self.future)]["after"])

    def test_disposed_sample_is_not_transitioned_again(self):
        api.do_transition_for(self.past, "expire")
        api.do_transition_for(self.past, "dispose")
        items = self.listing("disposed")
        self.assertEqual(items[0]["review_state"], "disposed")

    def test_supplier_listing_does_not_expire_other_suppliers_samples(self):
        other = api.create(
            self.portal.setup.suppliers, "Supplier", Name="Other")
        sample = api.create(other, "ReferenceSample", title="Other lot",
                            ExpiryDate=DateTime() - 1)
        items = self.listing("expired", supplier=self.supplier)
        self.assertEqual([item["uid"] for item in items],
                         [api.get_uid(self.past)])
        self.assertEqual(api.get_workflow_status_of(sample), "current")
