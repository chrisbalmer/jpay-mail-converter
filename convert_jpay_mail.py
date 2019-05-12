#!/usr/bin/env python3

import xml.etree.ElementTree as elementtree
from datetime import datetime
import os
import shutil

output_dir = './output/'
attachments_path = './Mail/Attachments/'
mailboxes = [
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

type_folders = {
    'Image': 'Img'
}

def create_letter_directory(name, timestamp, name_value):
    folder_name = timestamp.isoformat() + '-' + name_value
    folder_name = folder_name.replace(' ', '-')
    folder_name = folder_name.replace(':', '-')
    folder_path = os.path.join(output_dir, name, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

def save_body(folder_path, body):
    file_name = os.path.join(folder_path, 'body.txt')
    with open(file_name, "w") as body_file:
        body_file.write(body)

def parse_attachments(attachments, dest_folder_path, letter_id):
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
                sub_folder = type_folders[child.text]

        filename = name + '.' + extension
        file_path = os.path.join(attachments_path,
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
    for letter in root:
        timestamp = ''
        mailfrom = ''
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

        if (attachments):
            parse_attachments(attachments, folder_path, letter_id)

def parse_mailbox(name, filename, name_field):
    tree = elementtree.parse(filename)
    root = tree.getroot()
    parse_letters(name, name_field, root)






for mailbox in mailboxes:
    parse_mailbox(mailbox['name'],
                    mailbox['filename'],
                    mailbox['name_field'])
