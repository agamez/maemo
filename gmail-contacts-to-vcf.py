#!/usr/bin/env python
# -*- coding: utf-8 -*-
u"""
GMail contacts to VCF
---------------------

Exports Gmail contacts to a vcard (VCF) file. This can be done through the
Gmail web interface, but this script is more complete (more fields are
exported) and can be run automatically and periodically, for example as a
backup system.


.. :Authors:
       Aurélien Bompard <aurelien@bompard.org> <http://aurelien.bompard.org>

.. :License:
       GNU GPL v3 or later

"""

from __future__ import with_statement # compat python 2.5

import os
import sys
import getpass
import base64
from urlparse import urlparse
from cStringIO import StringIO
from optparse import OptionParser

import atom
import gdata.contacts
import gdata.contacts.service
import gdata.contacts.client
import vobject



class Contacts(object):

    def __init__(self, email, password, picsdir=None):
        """
        Takes an email and password corresponding to a gmail account to
        connect to the Contacts feed.

        :param email: The e-mail address of the account to use.
        :param password: The password corresponding to the account specified by
            the email parameter.
        """
        self.gd_client = gdata.contacts.client.ContactsClient()
        self.gd_client.source = os.path.basename(sys.argv[0])
        self.gd_client.ClientLogin(email, password, self.gd_client.source)
        self.groups = {}
        self.maingroup = None
        self.picsdir = picsdir
        if self.picsdir and not os.path.exists(self.picsdir):
            os.makedirs(self.picsdir)


    def dump(self, filename):
        self.list_groups()
        query = gdata.contacts.client.ContactsQuery()
        query.max_results = 999
        feed = self.gd_client.GetContacts(q=query)

        with open(filename, "w") as vcf_file:
            for i, entry in enumerate(feed.entry):
                if entry.title.text is None:
                    continue # likely to be a collected address
                all_group_ids = [ g.href for g in entry.group_membership_info ]
                if self.maingroup not in all_group_ids:
                    # Don't store this contact, it's a collected address
                    continue
                print i+1, entry.title.text.encode("utf8")
                #print entry
                contact = self._make_contact(entry)
                if contact is None:
                    continue
                vcf_file.write(contact.serialize())
                vcf_file.write("\r\n")


    def _make_contact(self, entry):
        """Builds a VCard entry from a Google Atom entry and returns it"""

        # Name
        contact = vobject.vCard()
        contact.add("n")
        contact.n.value = vobject.vcard.Name()
        if entry.name.given_name:
            contact.n.value.given = entry.name.given_name.text
        if entry.name.family_name:
            contact.n.value.family = entry.name.family_name.text
        if entry.name.name_prefix:
            contact.n.value.prefix = entry.name.name_prefix.text
        if entry.name.additional_name:
            contact.n.value.additional = entry.name.additional_name.text
        contact.add("fn")
        contact.fn.value = entry.name.full_name.text
        contact.add("name").value = entry.title.text

        # Email addresses
        for email in entry.email:
            c_email = contact.add("email")
            c_email.value = email.address
            if email.primary and email.primary == 'true':
                c_email.type_param = "PREF"

        # Note
        if entry.content:
            contact.add("note").value = entry.content.text

        # Groups
        groups = []
        for group in entry.group_membership_info:
            if group.href not in self.groups:
                continue
            groups.append(self.groups[group.href])
        if groups:
            contact.add("categories").value = groups

        # Modification time
        contact.add("rev")
        contact.rev.value = entry.updated.text

        # Phone
        for phone in entry.phone_number:
            phone_type = urlparse(phone.rel).fragment
            if phone_type == "mobile":
                phone_type = "cell"
            elif phone_type == "work_fax":
                phone_type = "fax"
            tel = contact.add("tel")
            tel.value = phone.text
            tel.type_param = phone_type.upper()

        # Organization
        if entry.organization:
            contact.add("org").value = [entry.organization.name.text]

        # Birthday
        if entry.birthday:
            contact.add("bday").value = entry.birthday.when

        # Address
        for address in entry.structured_postal_address:
            adr = contact.add("adr")
            adr.value = vobject.vcard.Address()
            if address.street:
                adr.value.street = address.street.text,
            if address.city:
                adr.value.city = address.city.text,
            if address.region:
                adr.value.region = address.region.text,
            if address.neighborhood:
                adr.value.code = address.neighborhood.text,
            if address.postcode:
                adr.value.code = address.postcode.text,
            if address.country:
                adr.value.country = address.country.text,
            if address.po_box:
                adr.value.box = address.po_box.text,
            adr_type = urlparse(address.rel).fragment
            adr.type_param = adr_type.upper()

        # Photo
        for link in entry.link:
            if link.rel != "http://schemas.google.com/contacts/2008/rel#photo":
                continue
            if "{http://schemas.google.com/g/2005}etag" not in link._other_attributes:
                continue
            hosted_image_binary = self.gd_client.GetPhoto(entry)
            if hosted_image_binary:
                contact.add("photo")
                contact.photo.value = hosted_image_binary
                contact.photo.encoding_param = "b"
                contact.photo.type_param = "image/jpeg"
            if self.picsdir:
                with open(os.path.join(self.picsdir,
                          "%s.jpg" % entry.title.text), "w") as img:
                    img.write(hosted_image_binary)

        # IM
        im_addrs = []
        for im in entry.im:
            proto = urlparse(im.protocol).fragment
            im_addrs.append( (proto, im.address) )
        if im_addrs:
            c_im = contact.add("x-kaddressbook-x-imaddress")
            c_im.value = " ".join("(%s)%s" % addr for addr in im_addrs)

        # Website
        for website in entry.website:
            contact.add("url").value = website.href

        # Display extended properties.
        for extended_property in entry.extended_property:
            if extended_property.value:
                value = extended_property.value
            else:
                value = extended_property.GetXmlBlob()
            print '    Extended Property - %s: %s' % (extended_property.name, value)

        return contact


    def list_groups(self):
        """
        Lists all Google contact groups and stores them in self.groups.
        The main "My Contacts" group is stored in self.maingroup to filter contacts.
        """
        query = gdata.service.Query(feed='/m8/feeds/groups/default/full')
        query.max_results = 2
        feed = self.gd_client.GetGroups()
        for entry in feed.entry:
            if entry.system_group is not None:
                if entry.system_group.id == "Contacts":
                    self.maingroup = entry.id.text
                continue
            self.groups[entry.id.text] = entry.title.text



def parse_opts():
    usage = "%prog [--user email_address] [--password password] [--filename vcard_file]"
    parser = OptionParser(usage)
    parser.add_option("-u", "--user", help="full email address")
    parser.add_option("-p", "--password")
    parser.add_option("-f", "--filename", help="VCard file to write to")
    parser.add_option("--pics", help="dump contact pictures in this folder")
    opts, args = parser.parse_args()
    while not opts.user:
        opts.user = raw_input("Please enter your username: ")
    while not opts.password:
        print "Please enter your password: ",
        opts.password = getpass.getpass()
        if not opts.password:
            print "Password cannot be blank."
    while not opts.filename:
        opts.filename = raw_input("Please enter the VCard file: ")
    return opts


def main():
    opts = parse_opts()

    try:
        contacts = Contacts(opts.user, opts.password, opts.pics)
    except gdata.service.BadAuthentication:
        print 'Invalid user credentials given.'
        return

    contacts.dump(opts.filename)



if __name__ == '__main__':
    main()
