# -*- coding: utf-8 -*-
"""Reference sample maintenance.
"""

from DateTime import DateTime
from bika.lims import api


def expire_samples():
    """Expire overdue current reference samples across all suppliers.

    The caller manages the transaction.
    """
    now = DateTime()
    query = {
        "portal_type": "ReferenceSample",
        "review_state": "current",
        "getExpiryDate": {"query": now, "range": "max"},
    }
    workflow = api.get_tool("portal_workflow")
    count = 0
    for brain in api.search(query, "senaite_catalog"):
        obj = api.get_object(brain)
        expiry = obj.getExpiryDate()
        if not expiry or expiry > now:
            continue
        if api.get_workflow_status_of(obj) != "current":
            continue
        workflow.doActionFor(obj, "expire")
        count += 1
    return count
