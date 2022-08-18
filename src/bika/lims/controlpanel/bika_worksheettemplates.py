# -*- coding: utf-8 -*-
#
# This file is part of SENAITE.CORE.
#
# SENAITE.CORE is free software: you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free Software
# Foundation, version 2.
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details.
#
# You should have received a copy of the GNU General Public License along with
# this program; if not, write to the Free Software Foundation, Inc., 51
# Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
#
# Copyright 2018-2021 by it's authors.
# Some rights reserved, see README and LICENSE.

import collections

from bika.lims import api
from bika.lims import bikaMessageFactory as _
from bika.lims.browser.bika_listing import BikaListingView
from bika.lims.config import PROJECTNAME
from bika.lims.interfaces import IWorksheetTemplates
from bika.lims.utils import get_link
from plone.app.folder.folder import ATFolder
from plone.app.folder.folder import ATFolderSchema
from Products.Archetypes import atapi
from Products.ATContentTypes.content import schemata
from senaite.core.interfaces import IHideActionsMenu
from senaite.core.permissions import AddWorksheetTemplate
from zope.interface.declarations import implements


class WorksheetTemplatesView(BikaListingView):
    """Listing View for Worksheet Templates
    """

    def __init__(self, context, request):
        super(WorksheetTemplatesView, self).__init__(context, request)

        self.catalog = "senaite_catalog_setup"

        self.contentFilter = {
            "portal_type": "WorksheetTemplate",
            "sort_on": "sortable_title",
            "sort_order": "ascending",
        }

        self.context_actions = {
            _("Add"):
            {
                "url": "createObject?type_name=WorksheetTemplate",
                "permission": AddWorksheetTemplate,
                "icon": "++resource++bika.lims.images/add.png"
            }
        }

        self.title = self.context.translate(_("Worksheet Templates"))
        self.icon = "{}/{}".format(
            self.portal_url,
            "/++resource++bika.lims.images/worksheettemplate_big.png"
        )

        self.show_select_row = False
        self.show_select_column = True

        self.columns = collections.OrderedDict((
            ("Title", {
                "title": _("Title"),
                "index": "sortable_title",
            }),
            ("Description", {
                "title": _("Description"),
                "index": "description",
                "toggle": True,
            }),
            ("Method", {
                "title": _("Method"),
                "toggle": True}),
            ("Instrument", {
                "title": _("Instrument"),
                "index": "instrument_title",
                "toggle": True,}),
            ("NumberOfPositions", {
                "title": _("Number of Positions"),
                "toggle": True}),
            ("Blanks", {
                "title": _("Blanks"),
                "toggle": True}),
            ("Controls", {
                "title": _("Controls"),
                "toggle": True}),
            ("NumberOfDuplicates", {
                "title": _("Number of Duplicates"),
                "toggle": True}),
        ))

        self.review_states = [
            {
                "id": "default",
                "title": _("Active"),
                "contentFilter": {"is_active": True},
                "columns": self.columns.keys(),
            }, {
                "id": "inactive",
                "title": _("Inactive"),
                "contentFilter": {'is_active': False},
                "columns": self.columns.keys()
            }, {
                "id": "all",
                "title": _("All"),
                "contentFilter": {},
                "columns": self.columns.keys(),
            },
        ]

    def folderitem(self, obj, item, index):
        """Service triggered each time an item is iterated in folderitems.
        The use of this service prevents the extra-loops in child objects.
        :obj: the instance of the class to be foldered
        :item: dict containing the properties of the object to be used by
            the template
        :index: current index of the item
        """
        obj = api.get_object(obj)
        item["Description"] = obj.Description()
        item["replace"]["Title"] = get_link(item["url"], item["Title"])

        instrument = obj.getInstrument()
        if instrument:
            instrument_url = api.get_url(instrument)
            instrument_title = api.get_title(instrument)
            item["Instrument"] = instrument_title
            item["replace"]["Instrument"] = get_link(
                instrument_url, value=instrument_title)

        # Method
        method_uid = obj.getMethodUID()
        if method_uid:
            method = api.get_object_by_uid(method_uid)
            method_url = api.get_url(method)
            method_title = api.get_title(method)
            item["Method"] = method_title
            item["replace"]["Method"] = get_link(
                method_url, value=method_title)
        
        # NumberOfPostions
        item["NumberOfPositions"] = len(obj.getLayout())

        Raw_layout = obj.getLayout()
        num = 0
        blanks = []
        controls = []
        controlUrls = []
        blankUrls = []
        for position in Raw_layout:
            if position.get('type') == "d":
                num = num + 1
            if position.get('type') == "b":
                blank_id = position.get('blank_ref')
                blank_obj = api.get_object_by_uid(blank_id)
                blank_title = blank_obj.Title()
                blank_url = api.get_url(blank_obj)
                if blank_title not in blanks:
                    blanks.append(blank_title)
                    blankUrls.append(get_link(blank_url,blank_title))
            if position.get('type') == "c":
                control_id = position.get('control_ref')
                control_obj = api.get_object_by_uid(control_id)
                control_title = control_obj.Title()
                control_url = api.get_url(control_obj)
                if control_title not in controls:
                    controls.append(control_title)
                    controlUrls.append(get_link(control_url,control_title))

        # Blanks
        if blanks:
            item["replace"]["Blanks"] = blankUrls

        # Controls
        if controls:
            item["replace"]["Controls"] = controlUrls

        # NummberOfDuplicates
        if num:
            item["NumberOfDuplicates"] = num
        return item


schema = ATFolderSchema.copy()


class WorksheetTemplates(ATFolder):
    implements(IWorksheetTemplates, IHideActionsMenu)
    displayContentsTab = False
    schema = schema


schemata.finalizeATCTSchema(schema, folderish=True, moveDiscussion=False)
atapi.registerType(WorksheetTemplates, PROJECTNAME)
