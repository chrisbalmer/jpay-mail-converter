#!/usr/bin/env python3
"""Go through the mail XML files from a JPAY tablet and conver the mail to
easy to read text files sorted into folders by the datetime and sender of
each email."""

import xml.etree.ElementTree as elementtree
from datetime import datetime
import os
import shutil

OUTPUT_DIR = './output/'
ATTACHMENTS_PATH = './Mail/Attachments/'
MAILBOXES = [
    {
        'name': 'Inbox',
        'filename': './Mail/Inbox.xml',
        'name_field': 'From'
    },
    {
        'name': 'Drafts',
        'filename': './Mail/Drafts.xml',
        'name_field': 'To'
    },
    {
        'name': 'Sent',
        'filename': './Mail/Sent.xml',
        'name_field': 'To'
    }
]

TYPE_FOLDERS = {
    'Image': 'Img'
}

BODY_FILENAME = 'body.txt'


def create_letter_directory(name, timestamp, name_value):
    """Creates the directory for the specified letter"""

    # Create the name for the letter folder.
    folder_name = timestamp.isoformat() + '-' + name_value
    folder_name = folder_name.replace(' ', '-')
    folder_name = folder_name.replace(':', '-')

    folder_path = os.path.join(OUTPUT_DIR, name, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def save_body(folder_path, body):
    """Saves the body of the letter to a file in the folder"""

    file_name = os.path.join(folder_path, BODY_FILENAME)
    with open(file_name, "w") as body_file:
        body_file.write(body)


def parse_attachments(attachments, dest_folder_path, letter_id):
    """Go through the attachments for the letter and copy them
    to the letter's directory"""

    for attachment in attachments:
        extension = ''
        name = ''
        sub_folder = ''
        for child in attachment:
            if child.tag == 'Name':
                extension = child.text.split('.')[-1]
            if child.tag == 'Id':
                name = child.text
            if child.tag == 'Type':
                sub_folder = TYPE_FOLDERS[child.text]

        filename = name + '.' + extension
        file_path = os.path.join(ATTACHMENTS_PATH,
                                 letter_id,
                                 sub_folder,
                                 filename)

        if os.path.isfile(file_path):
            shutil.copy2(file_path, dest_folder_path)
        else:
            error_path = os.path.join(dest_folder_path,
                                      'attachment_error.txt')
            with open(error_path, "w") as error_file:
                error_file.write(file_path + ' not found.')


def parse_letters(name, name_field, root):
    """Go through the letters for the mailbox and write the body of each letter
    to a file in a folder named for the letter's date and sender. Then process
    attachments for the letter and save them to the same directory."""

    for letter in root:
        timestamp = ''
        name_value = ''
        attachments = []
        body = ''
        letter_id = None

        for child in letter:
            if child.tag == 'DateTime':
                timestamp = datetime.strptime(child.text, '%m/%d/%Y %I:%M %p')
            elif child.tag == name_field:
                name_value = child.text
            elif child.tag == 'Attachment':
                attachments += [child]
            elif child.tag == 'Body':
                body = child.text
            elif child.tag == 'ID':
                letter_id = child.text

        if (timestamp and name_value):
            folder_path = create_letter_directory(name, timestamp, name_value)

        if (folder_path and body):
            save_body(folder_path, body)

        if attachments:
            parse_attachments(attachments, folder_path, letter_id)


def parse_mailbox(name, filename, name_field):
    """Grab the XML for the mailbox and go through the letters in the XML
    root element."""

    tree = elementtree.parse(filename)
    root = tree.getroot()
    parse_letters(name, name_field, root)


for mailbox in MAILBOXES:
    parse_mailbox(mailbox['name'],
                  mailbox['filename'],
                  mailbox['name_field'])
