# -*- coding: utf-8 -*-
"""Reference sample expiration, independent of the listing context.
"""

from bika.lims.api.security import check_permission
from plone.app.layout.viewlets import ViewletBase
from senaite.core.api.referencesample import expire_samples
from senaite.core.permissions import ManageReference

EXPIRY_CHECKED = "senaite.core.reference_sample_expiry_checked"


class ReferenceSampleExpiryViewlet(ViewletBase):
    """Transition overdue samples when an authorized user loads a page.
    """

    def update(self):
        super(ReferenceSampleExpiryViewlet, self).update()
        if not check_permission(ManageReference, self.context):
            return
        # CRM's alert viewlet can inherit this hook too. Check only once when
        # both viewlets are present in the same page request.
        if self.request.other.get(EXPIRY_CHECKED):
            return
        expire_samples()
        self.request.other[EXPIRY_CHECKED] = True

    def render(self):
        return ""
