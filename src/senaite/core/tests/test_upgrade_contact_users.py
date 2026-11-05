# -*- coding: utf-8 -*-

import unittest

from senaite.core.upgrade import v02_07_000 as upgrade


class ContactStub(object):

    def __init__(self, username="user", matches=None):
        self.username = username
        self.matches = matches
        self.linked = []

    def getUsername(self):
        return self.username

    def getContactByUsername(self, username):
        return self.matches

    def setUser(self, username):
        self.linked.append(username)
        return True


class TestUpgradeContactUsers(unittest.TestCase):

    def test_duplicate_username_is_not_linked(self):
        contact = ContactStub(matches=[ContactStub(), ContactStub()])
        get_path = upgrade.api.get_path
        upgrade.api.get_path = lambda obj: "/portal/contact"
        try:
            self.assertFalse(upgrade.link_contact_user(contact))
        finally:
            upgrade.api.get_path = get_path
        self.assertEqual(contact.linked, [])

    def test_unique_username_is_linked(self):
        contact = ContactStub()
        contact.matches = contact
        self.assertTrue(upgrade.link_contact_user(contact))
        self.assertEqual(contact.linked, ["user"])

    def test_no_username_is_skipped(self):
        contact = ContactStub(username="")
        self.assertFalse(upgrade.link_contact_user(contact))
        self.assertEqual(contact.linked, [])

    def test_unrelated_link_error_is_not_suppressed(self):
        class BrokenContact(ContactStub):
            def setUser(self, username):
                raise ValueError("Cannot link user")

        self.assertRaises(ValueError, upgrade.link_contact_user, BrokenContact())


def test_suite():
    return unittest.defaultTestLoader.loadTestsFromTestCase(
        TestUpgradeContactUsers)
