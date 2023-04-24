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
import itertools

from AccessControl import ClassSecurityInfo
from bika.lims import api
from bika.lims import bikaMessageFactory as _
from bika.lims.api.security import check_permission
from bika.lims.browser.bika_listing import BikaListingView
from senaite.core.permissions import FieldEditProfiles
from bika.lims.utils import format_supsub
from bika.lims.utils import get_image
from bika.lims.utils import get_link
from plone.memoize import view
from Products.Archetypes.Registry import registerWidget
from Products.Archetypes.Widget import TypesWidget
from zope.i18n.locales import locales
from Products.CMFCore.utils import getToolByName


class AnalysisProfileAnalysesView(BikaListingView):
    """Listing table to display Analyses Services for Analysis Profiles
    """

    def __init__(self, context, request):
        super(AnalysisProfileAnalysesView, self).__init__(context, request)

        self.an_cats = None
        self.an_cats_order = None
        self.catalog = "senaite_catalog_setup"
        self.contentFilter = {
            "portal_type": "AnalysisService",
            "sort_on": "sortable_title",
            "sort_order": "ascending",
            "is_active": True,
        }
        self.context_actions = {}

        self.show_column_toggles = False
        self.show_select_column = True
        self.show_select_all_checkbox = False
        self.pagesize = 999999
        self.allow_edit = True
        self.show_search = True
        self.omit_form = True
        self.fetch_transitions_on_select = False

        # Categories
        if self.show_categories_enabled():
            self.categories = []
            self.show_categories = True
            self.expand_all_categories = False

        self.columns = collections.OrderedDict((
            ("Title", {
                "title": _("Service"),
                "index": "sortable_title",
                "sortable": False}),
            ("Keyword", {
                "title": _("Keyword"),
                "sortable": False}),
            ("Methods", {
                "title": _("Methods"),
                "sortable": False}),
            ("Unit", {
                "title": _("Unit"),
                "sortable": False}),
            ("Price", {
                "title": _("Price"),
                "sortable": False,
            }),
            ("Hidden", {
                "title": _("Hidden"),
                "sortable": False}),
        ))

        columns = ["Title", "Keyword", "Methods", "Unit", "Price", "Hidden"]
        if not self.show_prices():
            columns.remove("Price")

        self.review_states = [
            {
                "id": "default",
                "title": _("All"),
                "contentFilter": {"is_active": True},
                "transitions": [{"id": "disallow-all-possible-transitions"}],
                "columns": columns,
            },
        ]

    def update(self):
        """Update hook
        """
        super(AnalysisProfileAnalysesView, self).update()
        self.allow_edit = self.is_edit_allowed()
        self.configuration = self.get_configuration()

    def get_settings(self):
        """Returns a mapping of UID -> setting
        """
        settings = self.context.getAnalysisServicesSettings()
        mapping = dict(map(lambda s: (s.get("uid"), s), settings))
        return mapping

    def get_configuration(self):
        """Returns a mapping of UID -> configuration
        """
        mapping = {}
        settings = self.get_settings()
        for service in self.context.getService():
            uid = api.get_uid(service)
            setting = settings.get(uid, {})
            config = {
                "hidden": setting.get("hidden", False),
            }
            mapping[uid] = config
        return mapping

    @view.memoize
    def show_categories_enabled(self):
        """Check in the setup if categories are enabled
        """
        return self.context.bika_setup.getCategoriseAnalysisServices()

    @view.memoize
    def show_prices(self):
        """Checks if prices should be shown or not
        """
        setup = api.get_setup()
        return setup.getShowPrices()

    @view.memoize
    def get_currency_symbol(self):
        """Get the currency Symbol
        """
        locale = locales.getLocale("en")
        setup = api.get_setup()
        currency = setup.getCurrency()
        return locale.numbers.currencies[currency].symbol

    @view.memoize
    def get_decimal_mark(self):
        """Returns the decimal mark
        """
        setup = api.get_setup()
        return setup.getDecimalMark()

    @view.memoize
    def format_price(self, price):
        """Formats the price with the set decimal mark and correct currency
        """
        return u"{} {}{}{:02d}".format(
            self.get_currency_symbol(),
            price[0],
            self.get_decimal_mark(),
            price[1],
        )

    @view.memoize
    def is_edit_allowed(self):
        """Check if edit is allowed
        """
        return check_permission(FieldEditProfiles, self.context)

    @view.memoize
    def get_editable_columns(self):
        """Return editable fields
        """
        columns = []
        if self.is_edit_allowed():
            columns = ["Hidden"]
        return columns

    def folderitems(self):
        """Sort by Categories
        """
        bsc = getToolByName(self.context, "senaite_catalog_setup")
        self.an_cats = bsc(
            portal_type="AnalysisCategory",
            sort_on="sortable_title")
        self.an_cats_order = dict([
            (b.Title, "{:04}".format(a))
            for a, b in enumerate(self.an_cats)])
        items = super(AnalysisProfileAnalysesView, self).folderitems()
        if self.show_categories_enabled():
            self.categories = map(lambda x: x[0],
                                    sorted(self.categories, key=lambda x: x[1]))
        else:
            self.categories.sort()
        return items

    def folderitem(self, obj, item, index):
        """Service triggered each time an item is iterated in folderitems.

        The use of this service prevents the extra-loops in child objects.

        :obj: the instance of the class to be foldered
        :item: dict containing the properties of the object to be used by
            the template
        :index: current index of the item
        """
        # ensure we have an object and not a brain
        obj = api.get_object(obj)
        uid = api.get_uid(obj)
        url = api.get_url(obj)
        title = api.get_title(obj)
        cat = obj.getCategoryTitle()
        cat_order = self.an_cats_order.get(cat)

        # get the category
        if self.show_categories_enabled():
            category = obj.getCategoryTitle()
            if (category,cat_order) not in self.categories:
                self.categories.append((category,cat_order))
            item["category"] = category

        config = self.configuration.get(uid, {})
        hidden = config.get("hidden", False)

        item["replace"]["Title"] = get_link(url, value=title)
        item["Price"] = self.format_price(obj.Price)
        item["allow_edit"] = self.get_editable_columns()
        item["selected"] = False
        item["Hidden"] = hidden
        item["selected"] = uid in self.configuration
        item["Keyword"] = obj.getKeyword()

        # Add methods
        methods = obj.getMethods()
        if methods:
            links = map(
                lambda m: get_link(
                    m.absolute_url(), value=m.Title(), css_class="link"),
                methods)
            item["replace"]["Methods"] = ", ".join(links)
        else:
            item["methods"] = ""

        # Unit
        unit = obj.getUnit()
        item["Unit"] = unit and format_supsub(unit) or ""

        # Icons
        after_icons = ""
        if obj.getAccredited():
            after_icons += get_image(
                "accredited.png", title=_("Accredited"))
        if obj.getAttachmentRequired():
            after_icons += get_image(
                "attach_reqd.png", title=_("Attachment required"))
        if after_icons:
            item["after"]["Title"] = after_icons

        return item


class AnalysisProfileAnalysesWidget(TypesWidget):
    """Analysis Profile Analyses Widget
    """
    _properties = TypesWidget._properties.copy()
    _properties.update({
        "macro": "bika_widgets/analysisprofileanalyseswidget",
        "helper_js": ("bika_widgets/analysisprofileanalyseswidget.js",),
        "helper_css": ("bika_widgets/analysisprofileanalyseswidget.css",),
    })

    security = ClassSecurityInfo()

    security.declarePublic("process_form")

    def process_form(self, instance, field, form, empty_marker=None,
                     emptyReturnsMarker=False):
        """Return UIDs of the selected services for the AnalysisProfile reference field
        """

        # selected services
        service_uids = form.get("uids", [])
        # hidden services
        hidden_services = form.get("Hidden", {})

        # get the service objects
        services = map(api.get_object_by_uid, service_uids)
        # get dependencies
        dependencies = map(lambda s: s.getServiceDependencies(), services)
        dependencies = list(itertools.chain.from_iterable(dependencies))
        # Merge dependencies and services
        services = set(services + dependencies)

        as_settings = []
        for service in services:
            service_uid = api.get_uid(service)
            hidden = hidden_services.get(service_uid, "") == "on"
            as_settings.append({"uid": service_uid, "hidden": hidden})

        # set the analysis services settings
        instance.setAnalysisServicesSettings(as_settings)

        return map(api.get_uid, services), {}

    security.declarePublic("Analyses")

    def Analyses(self, field, allow_edit=False):
        """Render Analyses Listing Table
        """
        instance = getattr(self, "instance", field.aq_parent)
        table = api.get_view("table_analysis_profile_analyses",
                             context=instance,
                             request=self.REQUEST)

        # Call listing hooks
        table.update()
        table.before_render()

        if allow_edit is False:
            return table.contents_table_view()
        return table.ajax_contents_table()


registerWidget(AnalysisProfileAnalysesWidget,
               title="Analysis Profile Analyses selector",
               description=("Analysis Profile Analyses selector"),)
