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


from Products.Archetypes.event import ObjectEditedEvent
from zope.event import notify

from bika.lims import api
from bika.lims.catalog import SETUP_CATALOG
from senaite.core import logger
from senaite.core.config import PROJECTNAME as product
from senaite.core.upgrade import upgradestep
from senaite.core.upgrade.utils import UpgradeUtils

version = "2.2.0"  # Remember version number in metadata.xml and setup.py
profile = "profile-{0}:default".format(product)


@upgradestep(product, version)
def upgrade(tool):
    portal = tool.aq_inner.aq_parent
    setup = portal.portal_setup  # noqa
    ut = UpgradeUtils(portal)
    ver_from = ut.getInstalledVersion(product)

    if ut.isOlderVersion(product, version):
        logger.info(
            "Skipping upgrade of {0}: {1} > {2}".format(product, ver_from, version)
        )
        return True

    logger.info("Upgrading {0}: {1} -> {2}".format(product, ver_from, version))

    # -------- ADD YOUR STUFF BELOW --------
    setup.runImportStepFromProfile(profile, "viewlets")
    migrate_calculations_of_analysisservice(portal)

    logger.info("{0} upgraded to version {1}".format(product, version))
    return True


def migrate_calculations_of_analysisservice(portal):
    logger.info("Migrate AnalysisService `Calculation` field ...")
    query = {"portal_type": "AnalysisService"}
    for brain in api.search(query, SETUP_CATALOG):
        obj = api.get_object(brain)
        methods = obj.getMethods()
        defaultcalculation = None
        for m, meth in enumerate(methods):
            if meth.getCalculations():
                for c, calc in enumerate(meth.getCalculations()):
                    if calc.Title() == "Correction":
                        defaultcalculation = calc
                        break
                if defaultcalculation:
                    break

        if defaultcalculation:
            logger.info(
                "{0} calculation is {1}".format(obj.title, obj.getCalculation())
            )
            logger.info("Method calculation is {0}".format(defaultcalculation.Title()))
            obj.setCalculation(defaultcalculation)
            notify(ObjectEditedEvent(obj))
            logger.info("Migrate AnalysisService `Calculation` field ...")

    logger.info("Migrate AnalysisService `Calculation` field ... [DONE]")
