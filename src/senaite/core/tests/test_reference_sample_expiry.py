from DateTime import DateTime
from bika.lims import api
from bika.lims.browser.referencesample import ReferenceSamplesView
from plone.app.testing import TEST_USER_ID
from plone.app.testing import setRoles
from senaite.core.api.referencesample import expire_samples
from senaite.core.browser.controlpanel.suppliers.referencesamples import \
    SupplierReferenceSamplesView
from senaite.core.browser.viewlets.referencesamples import EXPIRY_CHECKED
from senaite.core.browser.viewlets.referencesamples import \
    ReferenceSampleExpiryViewlet
from senaite.core.tests.base import BaseTestCase


class TestReferenceSampleExpiry(BaseTestCase):

    def setUp(self):
        super(TestReferenceSampleExpiry, self).setUp()
        self.request.other.pop(EXPIRY_CHECKED, None)
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
        expire_samples()
        items = self.listing()
        uids = [item["uid"] for item in items]
        self.assertNotIn(api.get_uid(self.past), uids)
        self.assertIn(api.get_uid(self.future), uids)
        self.assertIn(api.get_uid(self.undated), uids)
        self.assertEqual(api.get_workflow_status_of(self.past), "expired")

    def test_expired_tab_displays_transitioned_sample(self):
        expire_samples()
        items = self.listing("expired")
        self.assertEqual(len(items), 1)
        item = items[0]
        self.assertEqual(item["uid"], api.get_uid(self.past))
        self.assertEqual(item["review_state"], "expired")
        self.assertEqual(item["state_title"], "Expired")
        self.assertIn("exclamation", item["after"]["ExpiryDate"])
        self.assertIn('title="Expired"', item["after"]["ExpiryDate"])
        self.assertIn("text-danger", item["after"]["ExpiryDate"])

    def test_all_tab_shows_fresh_state_and_no_future_expiry_icon(self):
        expire_samples()
        items = dict((item["uid"], item) for item in self.listing("all"))
        self.assertEqual(items[api.get_uid(self.past)]["state_title"], "Expired")
        self.assertNotIn("ExpiryDate", items[api.get_uid(self.future)]["after"])

    def test_disposed_sample_is_not_transitioned_again(self):
        api.do_transition_for(self.past, "expire")
        api.do_transition_for(self.past, "dispose")
        items = self.listing("disposed")
        self.assertEqual(items[0]["review_state"], "disposed")

    def test_expiration_covers_all_suppliers_and_supplier_list_stays_scoped(self):
        other = api.create(
            self.portal.setup.suppliers, "Supplier", Name="Other")
        sample = api.create(other, "ReferenceSample", title="Other lot",
                            ExpiryDate=DateTime() - 1)
        self.assertEqual(expire_samples(), 2)
        items = self.listing("expired", supplier=self.supplier)
        self.assertEqual([item["uid"] for item in items],
                         [api.get_uid(self.past)])
        self.assertEqual(api.get_workflow_status_of(sample), "expired")

    def test_listing_does_not_change_workflow_state(self):
        items = self.listing("all")
        item = next(item for item in items
                    if item["uid"] == api.get_uid(self.past))
        self.assertEqual(item["review_state"], "current")
        self.assertIn("text-danger", item["after"]["ExpiryDate"])
        self.assertEqual(api.get_workflow_status_of(self.past), "current")
        self.assertEqual(self.listing("expired", supplier=self.supplier), [])
        self.assertEqual(api.get_workflow_status_of(self.past), "current")

    def test_expiration_is_idempotent(self):
        self.assertEqual(expire_samples(), 1)
        self.assertEqual(expire_samples(), 0)
        self.assertEqual(api.get_workflow_status_of(self.future), "current")
        self.assertEqual(api.get_workflow_status_of(self.undated), "current")

    def test_expiration_leaves_disposed_and_inactive_samples_unchanged(self):
        api.do_transition_for(self.past, "expire")
        api.do_transition_for(self.past, "dispose")
        self.future.setExpiryDate(DateTime() - 1)
        self.future.reindexObject()
        api.do_transition_for(self.future, "deactivate")
        self.assertEqual(expire_samples(), 0)
        self.assertEqual(api.get_workflow_status_of(self.past), "disposed")
        self.assertEqual(api.get_workflow_status_of(self.future), "inactive")

    def test_viewlet_expires_samples_before_list_is_loaded(self):
        viewlet = ReferenceSampleExpiryViewlet(
            self.portal, self.request, None)
        viewlet.update()
        self.assertEqual(viewlet.render(), "")
        self.assertEqual(api.get_workflow_status_of(self.past), "expired")
        items = self.listing("expired", supplier=self.supplier)
        self.assertEqual([item["uid"] for item in items],
                         [api.get_uid(self.past)])

    def test_viewlet_does_not_transition_without_manage_reference_permission(self):
        setRoles(self.portal, TEST_USER_ID, ["Member"])
        viewlet = ReferenceSampleExpiryViewlet(
            self.portal, self.request, None)
        viewlet.update()
        self.assertEqual(api.get_workflow_status_of(self.past), "current")
        self.assertFalse(self.request.get(EXPIRY_CHECKED))

    def test_viewlets_check_expiration_once_per_request(self):
        viewlet = ReferenceSampleExpiryViewlet(
            self.portal, self.request, None)
        viewlet.update()
        self.future.setExpiryDate(DateTime() - 1)
        self.future.reindexObject()
        other = ReferenceSampleExpiryViewlet(
            self.supplier, self.request, None)
        other.update()
        self.assertEqual(api.get_workflow_status_of(self.future), "current")
        # A subsequent page request checks again.
        self.request.other.pop(EXPIRY_CHECKED, None)
        other.update()
        self.assertEqual(api.get_workflow_status_of(self.future), "expired")
