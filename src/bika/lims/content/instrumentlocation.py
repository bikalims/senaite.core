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

from zope.interface import implements

from AccessControl import ClassSecurityInfo

from Products.Archetypes.public import BaseContent
from Products.Archetypes import atapi

from bika.lims.content.bikaschema import BikaSchema
from bika.lims.config import PROJECTNAME
from bika.lims.interfaces import IInstrumentLocation, IDeactivable

schema = BikaSchema.copy()

schema['description'].schemata = 'default'
schema['description'].widget.visible = True


class InstrumentLocation(BaseContent):
    """A physical place, where an Instrument is located
    """
    implements(IInstrumentLocation, IDeactivable)
    security = ClassSecurityInfo()
    displayContentsTab = False
    schema = schema

    _at_rename_after_creation = True

    def _renameAfterCreation(self, check_auto_id=False):
        from senaite.core.idserver import renameAfterCreation
        renameAfterCreation(self)

atapi.registerType(InstrumentLocation, PROJECTNAME)
