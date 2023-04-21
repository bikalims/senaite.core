# -*- coding: utf-8 -*-

from bika.lims import api
from bika.lims.interfaces import IClient
from plone.indexer import indexer
from senaite.core.interfaces.catalog import IClientCatalog


@indexer(IClient, IClientCatalog)
def client_searchable_text(instance):
    """Extract search tokens for ZC text index
    """

    tokens = [
        str(instance.getClientID()),
        str(instance.getName()),
        str(instance.getPhone()),
        str(instance.getFax()),
        str(instance.getEmailAddress()),
        str(instance.getTaxNumber()),
    ]

    # extend address lines
    tokens.extend(instance.getPrintAddress())

    # remove duplicates and filter out emtpies
    tokens = filter(None, set(tokens))

    # return a single unicode string with all the concatenated tokens
    return u" ".join(map(api.safe_unicode, tokens))
