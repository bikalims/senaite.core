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
# Copyright 2018-2023 by it's authors.
# Some rights reserved, see README and LICENSE.

import transaction
from Acquisition import aq_base
from bika.lims import api
from senaite.core import logger
from senaite.core.api.catalog import add_index
from senaite.core.api.catalog import del_column
from senaite.core.api.catalog import del_index
from senaite.core.api.catalog import reindex_index
from senaite.core.catalog import CLIENT_CATALOG
from senaite.core.catalog import REPORT_CATALOG
from senaite.core.catalog import SAMPLE_CATALOG
from senaite.core.config import PROJECTNAME as product
from senaite.core.permissions import ManageBika
from senaite.core.registry import get_registry_record
from senaite.core.setuphandlers import add_dexterity_items
from senaite.core.setuphandlers import setup_catalog_mappings
from senaite.core.setuphandlers import setup_core_catalogs
from senaite.core.upgrade import upgradestep
from senaite.core.upgrade.utils import UpgradeUtils
from senaite.core.upgrade.utils import uncatalog_brain

version = "2.5.0"  # Remember version number in metadata.xml and setup.py
profile = "profile-{0}:default".format(product)

CONTENT_ACTIONS = [
    # portal_type, action
    ("Client", {
        "id": "manage_access",
        "name": "Manage Access",
        "action": "string:${object_url}/@@sharing",
        # NOTE: We use this permission to hide the action from logged in client
        # contacts
        "permission": ManageBika,
        # "permission": "Sharing page: Delegate roles",
        "category": "object",
        "visible": True,
        "icon_expr": "",
        "link_target": "",
        "condition": "",
        "insert_after": "edit",
    }),
]


@upgradestep(product, version)
def upgrade(tool):
    portal = tool.aq_inner.aq_parent
    ut = UpgradeUtils(portal)
    ver_from = ut.getInstalledVersion(product)

    if ut.isOlderVersion(product, version):
        logger.info("Skipping upgrade of {0}: {1} > {2}".format(
            product, ver_from, version))
        return True

    logger.info("Upgrading {0}: {1} -> {2}".format(product, ver_from, version))

    # -------- ADD YOUR STUFF BELOW --------

    logger.info("{0} upgraded to version {1}".format(product, version))
    return True


def rebuild_sample_zctext_index_and_lexicon(tool):
    """Recreate sample listing_searchable_text ZCText index and Lexicon
    """
    # remove the existing index
    index = "listing_searchable_text"
    del_index(SAMPLE_CATALOG, index)
    # remove the Lexicon
    catalog = api.get_tool(SAMPLE_CATALOG)
    if "Lexicon" in catalog.objectIds():
        catalog.manage_delObjects("Lexicon")
    # recreate the index + lexicon
    add_index(SAMPLE_CATALOG, index, "ZCTextIndex")
    # reindex
    reindex_index(SAMPLE_CATALOG, index)


@upgradestep(product, version)
def setup_labels(tool):
    """Setup labels for SENAITE
    """
    logger.info("Setup Labels")
    portal = api.get_portal()

    tool.runImportStepFromProfile(profile, "typeinfo")
    tool.runImportStepFromProfile(profile, "workflow")
    tool.runImportStepFromProfile(profile, "plone.app.registry")
    setup_core_catalogs(portal)

    items = [
        ("labels",
         "Labels",
         "Labels")
    ]
    setup = api.get_senaite_setup()
    add_dexterity_items(setup, items)


def setup_client_catalog(tool):
    """Setup client catalog
    """
    logger.info("Setup Client Catalog ...")
    portal = api.get_portal()

    # setup and rebuild client_catalog
    setup_catalog_mappings(portal)
    setup_core_catalogs(portal)
    client_catalog = api.get_tool(CLIENT_CATALOG)
    client_catalog.clearFindAndRebuild()

    # portal_catalog cleanup
    uncatalog_type("Client", catalog="portal_catalog")

    logger.info("Setup Client Catalog [DONE]")


def uncatalog_type(portal_type, catalog="portal_catalog", **kw):
    """Uncatalog all entries of the given type from the catalog
    """
    query = {"portal_type": portal_type}
    query.update(kw)
    brains = api.search(query, catalog=catalog)
    for brain in brains:
        uncatalog_brain(brain)


def setup_catalogs(tool):
    """Setup all core catalogs and ensure all indexes are present
    """
    logger.info("Setup Catalogs ...")
    portal = api.get_portal()

    setup_catalog_mappings(portal)
    setup_core_catalogs(portal)

    logger.info("Setup Catalogs [DONE]")


def update_report_catalog(tool):
    """Update indexes in report catalog and add new metadata columns
    """
    logger.info("Update report catalog ...")
    portal = api.get_portal()

    # ensure new indexes are created
    setup_catalog_mappings(portal)
    setup_core_catalogs(portal)

    # remove columns
    del_column(REPORT_CATALOG, "getClientTitlegetClientURL")
    del_column(REPORT_CATALOG, "getDatePrinted")

    logger.info("Update report catalog [DONE]")


def import_registry(tool):
    """Import registry step from profiles
    """
    portal = tool.aq_inner.aq_parent
    setup = portal.portal_setup
    setup.runImportStepFromProfile(profile, "plone.app.registry")


def create_client_groups(tool):
    """Create for all Clients an explicit Group
    """
    logger.info("Create client groups ...")
    clients = api.search({"portal_type": "Client"}, CLIENT_CATALOG)
    total = len(clients)
    for num, client in enumerate(clients):
        obj = api.get_object(client)
        logger.info("Processing client %s/%s: %s"
                    % (num+1, total, obj.getName()))

        # recreate the group
        obj.remove_group()

        # skip group creation
        if not get_registry_record("auto_create_client_group", True):
            logger.info("Auto group creation is disabled in registry. "
                        "Skipping group creation ...")
            continue

        group = obj.create_group()
        # add all linked client contacts to the group
        for contact in obj.getContacts():
            user = contact.getUser()
            if not user:
                continue
            logger.info("Adding user '%s' to the client group '%s'"
                        % (user.getId(), group.getId()))
            obj.add_user_to_group(user)

    logger.info("Create client groups [DONE]")


def reindex_client_security(tool):
    """Reindex client object security to grant the owner role for the client
       group to all contents
    """
    logger.info("Reindex client security ...")

    clients = api.search({"portal_type": "Client"}, CLIENT_CATALOG)
    total = len(clients)
    for num, client in enumerate(clients):
        obj = api.get_object(client)
        logger.info("Processing client %s/%s: %s"
                    % (num+1, total, obj.getName()))

        if not obj.get_group():
            logger.info("No client group exists for client %s. "
                        "Skipping reindexing ..." % obj.getName())
            continue

        _recursive_reindex_object_security(obj)

        logger.info("Commiting client %s/%s" % (num+1, total))
        transaction.commit()
        logger.info("Commit done")

        # Flush the object from memory
        obj._p_deactivate()

    logger.info("Reindex client security [DONE]")


def _recursive_reindex_object_security(obj):
    """Recursively reindex object security for the given object
    """
    if hasattr(aq_base(obj), "objectValues"):
        for child_obj in obj.objectValues():
            _recursive_reindex_object_security(child_obj)
    obj.reindexObject(idxs=["allowedRolesAndUsers"])
    obj._p_deactivate()


def add_content_actions(tool):
    logger.info("Add cotent actions ...")
    portal_types = api.get_tool("portal_types")
    data = {'AA Test': 'AAT',
            'ARC - Animal Production: Milk Recording': 'ARC',
            'Action Classics': 'ACS',
            'Adam Farming (Pty) Ltd': 'ADF',
            'Afgri Animal Feeds': 'AAF',
            'Afraceuticals (Pty) Ltd': 'ACT',
            'African Cures': 'ANC',
            'Afrigetics Botanicals': 'AGB',
            'Agri Hygiene': 'AHE',
            'Akacia Medical (Pty) Ltd t/a CliniSut': 'AMC',
            'Aluminium Foil Converters': 'AFC',
            'Amanzi-4-All': 'AMA',
            'Andermatt PHP': 'APH',
            'Andrade International': 'AEI',
            'Andrew Butt': 'ABU',
            'Assign Trading cc': 'ATC',
            'Avis Milling': 'AVM',
            'BASF Agricultural Specialities (Pty) Ltd': 'BAS',
            'Balcairn (Pty) Ltd': 'BCL',
            'Bandini Group': 'BDG',
            'Bio-Science Laboratories': 'BIS',
            'Biomimetics (Pty) Ltd': 'BML',
            'Blomeyer Farming': 'BLF',
            'Bonle Dairy': 'BDY',
            'Bonsma Farming Trust': 'BFT',
            'Bosveld Hides': 'BVH',
            'Brandkraal Farm': 'BKF',
            'Brenell Desserts': 'BED',
            'Brenn-O-Kem': 'BOK',
            'Britos Food International': 'BFI',
            'Bryden Farming': 'BRN',
            'Burnview Farm (Pty) Ltd': 'BVF',
            'CPAC': 'CPA',
            'Cake Board Suppliers cc': 'CBS',
            'Caldecott Farming - Nutcombe': 'CFN',
            'Cape Nuts': 'CPN',
            'Carara Agro Processing': 'CAP',
            'Cardiblox': 'CDX',
            'Carisbrooke Valley Citrus': 'CVC',
            'Chartwell Farms': 'CWF',
            'City of Tshwane - Municipal Health Services': 'COT',
            'Clint Swartz': 'CSZ',
            'Clive Prince': 'CPE',
            'Clover': 'CVR',
            'Commercial & Agric Trading cc.': 'CAT',
            'Corporate Cleaning': 'CCS',
            'Corporate Services': 'CSD',
            'Creighton Cheese': 'CNC',
            'Creighton Dairies': 'CND',
            'Culverwell Trading - Rosetta Farm': 'CTR',
            'Dairy 52': 'DTF',
            'Dairy Group (Pty) Ltd': 'DGL',
            'Dairy Standards Agency': 'DSA',
            'Dairy Tech': 'DTH',
            'Dargle Valley Meats': 'DVM',
            'De Heus': 'DHS',
            'Dee Import and Export (Pty) Ltd': 'DIE',
            'Defence Farm Dairy (Pty) Ltd': 'DFD',
            'Dendolor Investments (PTY) Ltd (Broadside)': 'DDB',
            'Desmanda': 'DMA',
            'Douglasdale Dairy': 'DDY',
            'Dr Rick Mapham': 'DRM',
            'Dr Tod Collins Livestock Consultancy': 'DTC',
            'Drummond Tor': 'DMT',
            'Dynamed Pharmaceuticals': 'DDP',
            'E.G. Veterinary Services \xe2\x80\x93 AJ Joubert': 'EGV',
            'EcoWize': 'EZE',
            'Edkins Farming cc': 'EFC',
            'Eggbert': 'EGG',
            'Elliott Farm': 'ELF',
            'Elmwood Farms': 'ELM',
            'Essenwood Micro Dairy': 'EMD',
            'Essity (BSN Medical (Pty) Ltd)': 'EBM',
            'Estcourt Veterinary Clinic': 'EVC',
            'Etlin International (Pty) Ltd': 'ETL',
            'Exceptional Find Investments': 'EFI',
            'Fairfield Dairy (Pty) Ltd': 'FDL',
            'Far End Dairy (Pty) Ltd': 'FED',
            "Farmer's Agricare": 'FAC',
            'Farmgate Dairy': 'FDA',
            'Farningholm Farms (Pty) Ltd': 'FHF',
            'Federated Meats': 'FDM',
            'Fleures Honey (Pty) Ltd': 'FHL',
            'FreeMe Wildlife Rehabilitation KZN': 'FWR',
            'Frisch Consulting (Pty) Ltd': 'FCG',
            'Froozels': 'FZL',
            'Frost Farming': 'FFG',
            'Fry Group Foods (Pty) Ltd': 'FRY',
            'Future Farmers Invest (Pty) Ltd': 'FFI',
            'Fyvie Estates Richmond Partnership': 'FER',
            'G.A. Carr T/A Netherby Farm': 'GCN',
            'G3 Global SA t/a Jersey Cow Co.': 'GGC',
            'GG & QG Elliott Farm': 'GQE',
            'GHL - Method Validations': 'GMV',
            'GI Science Solutions': 'GIS',
            'Gace Farming': 'GFA',
            'Garland Farming': 'GLF',
            'Geochem (Pty) Ltd': 'GCL',
            'Gilcraft cc': 'GCC',
            'Glendy Farming': 'GDF',
            'Global Solutions': 'GLS',
            'Grant Irons': 'GIS',
            'Green Farming Trust': 'GFT',
            'Green Farms Nut Company (Pty) Ltd - White River': 'GFN',
            'Green Farms Nut Company Levubu (Pty) Ltd': 'GFL',
            'Harmony Labs': 'HML',
            'Hen-Li Consulting': 'HLC',
            'Hendrik Smit': 'HDS',
            'Highlands Investments': 'HIS',
            'Highveld Farm': 'HVF',
            'Hilton Foods': 'HFS',
            'Hlogoma Farming Trust': 'HFT',
            'Hulley Bros (Pty) Ltd': 'HBL',
            'Hume International': 'HEI',
            'Impact Distributors (Pty) Ltd T/A Bandini Cheese': 'IBC',
            'Inchgarth Dairies': 'IDS',
            'Ingeli Group - Sarsgrove/Grassy Park': 'IGP',
            'Instru-serv': 'ITS',
            'JA Theron': 'JAT',
            'Jabula Lamb': 'JBL',
            'Jerry Gengan': 'JGN',
            'John Stimson': 'JSN',
            'Just Pies': 'PIE',
            'Kara Nichhas': 'KAR',
            'Kelpack': 'KPK',
            'Kevin Barnsley': 'KBY',
            'Khanyisa Projects': 'KPS',
            'Kroy Distributions': 'KDS',
            'LMVP Products': 'LMV',
            'Lauviv (Pty) Ltd t/a Indezi River Creamery': 'LIC',
            "Lionel's Veterinary Supplies (Pty) Ltd": 'DVL',
            'Liselo Labs': 'LLS',
            'Lloyd Kirk': 'LDK',
            'Luggage Protector': 'LPR',
            'Lyn Robert Blomeyer': 'LRB',
            'Lynn Blomeyer': 'LBR',
            'M.T. Hodgson': 'MTH',
            'MNR Agri (Pty) Ltd': 'MNR',
            'MOP Foods cc (USE TAGONIST)': 'MOP',
            'Macallum (Pty) Ltd': 'MCM',
            'Mackenzie Farms': 'MKF',
            'Macston cc': 'MCS',
            'Madumbi Sustainable Agriculture': 'MSA',
            'Mangwane Investments cc': 'MIC',
            'Mantos Foods (Pty) Ltd': 'MFL',
            'Mark Hodgson': 'MAH',
            'Mark Willment': 'MWT',
            'Mary-Ann Murphy': 'MAM',
            'Meadow Farm': 'MFM',
            'Melda Dairies': 'MDS',
            'Merlog Foods (Pty) Ltd': 'MFS',
            'Merrivale Poultry Farm': 'MPF',
            'Michigan Equipment (Pty) Ltd': 'MEL',
            'Middledale Farming (Pty) Ltd': 'MEF',
            'Midlands Eggs (Pty) Ltd': 'MES',
            'Midlands Foods Co': 'MFC',
            'Mighty Meats (Pty) Ltd': 'MMS',
            "Miguel's Bakery": 'MBY',
            'Milchstef Dairy Tech': 'MDT',
            'Mistbelt AfroChemicals': 'MAC',
            'Modern Value Meats (Pty) Ltd': 'MVM',
            'Moller Farming': 'MFG',
            'Mostly Milk': 'MYM',
            'Mountain Valley': 'MVY',
            'Mountainview Dairy': 'MVD',
            'Mukesh': 'MKH',
            'Mycelia Projects Pty Ltd': 'MPL',
            'National Institute for Communicable Diseases': 'NCID',
            'Natural & Organic Formulations': 'NOF',
            'Ndiza Poultry Rearers (Pty) Ltd': 'NPR',
            'Nedbank Limited': 'NBK',
            'Nestle (South Africa) (Pty) Ltd': 'NSL',
            'Niekerksfontein': 'NKF',
            'Ntsika Tech': 'NST',
            'Nutrochem': 'NCH',
            'Oaksprings Dairy': 'OSD',
            'Orange Grove Dairies': 'OGD',
            "Oscar's Meats": 'OSC',
            'P.N. Kean': 'PNK',
            'PSG Meats cc t/a Modern Butchery': 'PSG',
            "Page's Chocolate Crunchies": 'PCC',
            'Peels Holdings (Pty) Ltd': 'PHL',
            'Peels Honey': 'PSH',
            'Pepper Agri Holdings': 'PAH',
            'Petlin': 'PET',
            'Praecautio (Pty) Ltd': 'PRC',
            'Pride Milling Company (Pty) Ltd': 'PMC',
            'R. A. Oldfield': 'RAO',
            'Reddy Bio-Clean cc': 'RBC',
            'Ribs N Meat cc': 'RAM',
            'Rich Products Corporation Africa': 'RPC',
            'Rodger Spencer': 'RSR',
            'Salt Springs Farming (Pty) Ltd': 'SSF',
            'Sani Agri (Pty) Ltd \xe2\x80\x93 Glen Gowrie Farm': 'SAL',
            'Sapuma Eggs cc': 'SAE',
            'Savour Solutions (Pty) Ltd': 'SSL',
            'Scotston Farm': 'SNF',
            'Seara Africa': 'SEA',
            'Senga Farming Pty Ltd': 'SFL',
            'Shield Health Care': 'SHC',
            'Silchem Solutions': 'SSN',
            'Siphokit': 'SKT',
            'Smee Dubazane': 'SDZ',
            'South Atlantic Meat Imp & Exp (Pty) Ltd T/A Transtrade International': 'SAM',
            'Southern Oil (Pty) Ltd': 'SOL',
            'Southlands Burnside': 'SBE',
            'Southlands Rosedale': 'SRE',
            'Spring Meadow Dairy Farm (Pty) Ltd': 'SMD',
            'Springvale Farm Trust': 'SVT',
            'Squirrels Nuts (Pty) Ltd': 'SNL',
            'Stapylton-Smith Farming cc': 'SSF',
            'SteriTech': 'STH',
            'Stockton Trading Trust': 'STT',
            'Stratford Farming (Pty) Ltd': 'SFL',
            'Struan Farm cc': 'SFC',
            'Summerhill Farm - G Ngcobo': 'SMF',
            'TGS Property Investments cc T/A Drakensberg Abattoir': 'TGS',
            'Tagonist Trading cc T/A TFT Foods': 'TFT',
            'Talbot Laboratories (Pty) Ltd': 'TBL',
            'Teichmann Eggs': 'TEG',
            'The Anolyte Company (Pty) Ltd': 'TAC',
            'The Food Laboratory': 'TFL',
            'The Goble Family Trust': 'GFT',
            'The Natal Pepper Company (Pty) Ltd': 'NPC',
            'Thinking Strings Media': 'TSM',
            'Torr Farming': 'TFG',
            'Transtrade': 'TTE',
            'UKZN Inqubate': 'UKZN',
            'Underberg Dairy (Pty) Ltd': 'UND',
            'Underberg Veterinary Surgery cc': 'UVS',
            'Underbush Valley Farms': 'UVF',
            'Unity Food Products': 'UFP',
            'Veterinary House Hospital CC \xe2\x80\x93 TD Marwick & Sons \xe2\x80\x93 Little Harmony Farm': 'VHH',
            'Vileshen / Reshmika': 'VNR',
            'Vitam International': 'VIT',
            'Waterfall Dairy': 'WFD',
            'Westfalia Fruit Estates (Pty) Ltd t/a Everdon Estates': 'WFE',
            'Westfalia Fruit Products': 'WFP',
            'Wezco': 'WEZ',
            'Willowton Oil': 'WWO',
            'Woodlands Dairy (Pty) Ltd': 'WDL',
            'Worrall Farming (Pty) Ltd': 'WFL',
            'Yusuf Jugmohan': 'YFJ',
            'Zelphy 2673 (Pty) Ltd t/a St Louis': 'ZPL',
            'Zenco farming': 'ZEN',
            'Ziklag Farming & Agricultural Ent. cc': 'ZFA',
            'Zoetis South Africa': 'ZSA',
            'Zylem Pty Ltd': 'ZPL',
            'uMoya Management Services': 'UMS'}
    clients = api.search({"portal_type": "Client"}, CLIENT_CATALOG)
    for client in clients:
        obj = client.getObject()
        obj.setClientID(str(data[client.Title]))
        # obj.reindexObject()
    client_catalog = api.get_tool(CLIENT_CATALOG)
    client_catalog.clearFindAndRebuild()
    # for record in CONTENT_ACTIONS:
    #     portal_type, action = record
    #     type_info = portal_types.getTypeInfo(portal_type)
    #     action_id = action.get("id")
    #     # remove any previous added actions with the same ID
    #     _remove_action(type_info, action_id)
    #     # pop out the position info
    #     insert_after = action.pop("insert_after", None)
    #     # add the action
    #     type_info.addAction(**action)
    #     # sort the action to the right position
    #     actions = type_info._cloneActions()
    #     action_ids = map(lambda a: a.id, actions)
    #     if insert_after in action_ids:
    #         ref_index = action_ids.index(insert_after)
    #         index = action_ids.index(action_id)
    #         action = actions.pop(index)
    #         actions.insert(ref_index + 1, action)
    #         type_info._actions = tuple(actions)

    #     logger.info("Added action id '%s' to '%s'",
    #                action_id, portal_type)
    logger.info("Add content actions [DONE]")


def _remove_action(type_info, action_id):
    """Removes the action id from the type passed in
    """
    actions = map(lambda action: action.id, type_info._actions)
    if action_id not in actions:
        return True
    index = actions.index(action_id)
    type_info.deleteActions([index])
    return _remove_action(type_info, action_id)
