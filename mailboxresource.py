#!/usr/bin/env python
# -*- coding:utf-8 -*-

from __future__ import print_function

import imaplib, email
import re
import os
import hashlib
from message import Message
import datetime
import urllib


class MailboxClient:
    """Operations on a mailbox"""

    def __init__(self, host, port, username, password, remote_folder, ssl):
        """
        Initializes a MailboxClient object.

        Parameters:
            host (str): The hostname of the IMAP server.
            port (int): The port number of the IMAP server.
            username (str): The login id for the IMAP server.
            password (str): The password for the IMAP server.
            remote_folder (str): The remote folder to select on the IMAP server.
            ssl (bool): Whether to use SSL for the IMAP connection.

        Raises:
            Exception: If the remote folder could not be selected on the IMAP server.

        Returns:
            None
        """
        if not ssl:
            self.mailbox = imaplib.IMAP4(host, port)
        else:
            self.mailbox = imaplib.IMAP4_SSL(host, port)
        self.mailbox.login(username, password)
        status, folders = self.mailbox.list()
        print(f"Status: {status}, folders: {folders}")
        # typ, data = self.mailbox.select(remote_folder, readonly=True)
        typ, data = self.mailbox.select("INBOX", readonly=True)
        if typ != "OK":
            # Handle case where Exchange/Outlook uses '.' path separator when
            # reporting subfolders. Adjust to use '/' on remote.
            adjust_remote_folder = re.sub(r"\.", "/", remote_folder)
            typ, data = self.mailbox.select(adjust_remote_folder, readonly=True)
            if typ != "OK":
                print("MailboxClient: Could not select remote folder '%s'" % remote_folder)

    def copy_emails(self, days, local_folder, wkhtmltopdf):
        """
        Copies emails from the mailbox to a local folder.

        Parameters:
            days (int): The number of days back to get emails from. If None, gets all emails.
            local_folder (str): The full path to the folder where emails should be saved.
            wkhtmltopdf (str): The location of the wkhtmltopdf binary.

        Returns:
            tuple: A tuple containing the number of emails saved and the number of emails that already exist locally.
        """

        n_saved = 0
        n_exists = 0

        self.local_folder = local_folder
        self.wkhtmltopdf = wkhtmltopdf
        criterion = "ALL"

        if days:
            date = (datetime.date.today() - datetime.timedelta(days)).strftime("%d-%b-%Y")
            criterion = "(SENTSINCE {date})".format(date=date)

        typ, data = self.mailbox.search(None, criterion)
        if data and data[0]:
            for num in data[0].split():
                typ, data = self.mailbox.fetch(num, "(BODY.PEEK[])")
                if self.saveEmail(data):
                    n_saved += 1
                else:
                    n_exists += 1

        return (n_saved, n_exists)

    def cleanup(self):
        """
        Closes the IMAP mailbox connection and logs out of the server.

        This method is used to clean up resources after the mailbox operations are complete.

        Parameters:
            None

        Returns:
            None
        """
        self.mailbox.close()
        self.mailbox.logout()

    def getEmailFolder(self, msg, data):
        """
        Returns the path to a folder where an email should be saved.

        The folder name is derived from the email's Message-Id if it exists and is less than 255 characters.
        Otherwise, a SHA-224 hash of the email data is used.

        The folder path is constructed by joining the local_folder, the year the email was sent (if available), and the folder name.

        Parameters:
            msg (dict): The email message.
            data (bytes): The email data.

        Returns:
            str: The path to the folder where the email should be saved.
        """
        # 255is the max filename length on all systems
        if msg["Message-Id"] and len(msg["Message-Id"]) < 255:
            foldername = re.sub(r"[^a-zA-Z0-9_\-\.() ]+", "", msg["Message-Id"])
        else:
            foldername = hashlib.sha224(data).hexdigest()

        year = "None"
        if msg["Date"]:
            match = re.search(r"\d{1,2}\s\w{3}\s(\d{4})", msg["Date"])
            if match:
                year = match.group(1)

        return os.path.join(self.local_folder, year, foldername)

    def saveEmail(self, data):
        """
        Saves an email to a local directory.

        Parameters:
            data (list): A list containing the email data.

        Returns:
            bool: True if the email was saved successfully, False if the directory already exists.
        """
        for response_part in data:
            if isinstance(response_part, tuple):
                msg = ""
                # Handle Python version differences:
                # Python 2 imaplib returns bytearray, Python 3 imaplib
                # returns str.
                if isinstance(response_part[1], str):
                    msg = email.message_from_string(response_part[1])
                else:
                    try:
                        msg = email.message_from_string(response_part[1].decode("utf-8"))
                    except:
                        print("couldn't decode message with utf-8 - trying 'ISO-8859-1'")
                        msg = email.message_from_string(response_part[1].decode("ISO-8859-1"))

                directory = '.' + str(self.getEmailFolder(msg, data[0][1]))
                print(directory)
                if os.path.exists(directory):
                    return False

                os.makedirs(directory)

                try:
                    message = Message(directory, msg)
                    message.createRawFile(data[0][1])
                    message.createMetaFile()
                    message.extractAttachments()

                    if self.wkhtmltopdf:
                        message.createPdfFile(self.wkhtmltopdf)

                except Exception as e:
                    # ex: Unsupported charset on decode
                    print(directory)
                    if hasattr(e, "strerror"):
                        print("MailboxClient.saveEmail() failed:", e.strerror)
                    else:
                        print("MailboxClient.saveEmail() failed")
                        print(e)

        return True


def save_emails(account, options):
    mailbox = MailboxClient(
        account["host"],
        account["port"],
        account["username"],
        account["password"],
        account["remote_folder"],
        account["ssl"],
    )
    stats = mailbox.copy_emails(options["days"], options["local_folder"], options["wkhtmltopdf"])
    mailbox.cleanup()
    if stats[0] == 0 and stats[1] == 0:
        print("Folder {} is empty".format(account["remote_folder"]))
    else:
        print("{} emails created, {} emails already exists".format(stats[0], stats[1]))


def get_folder_fist(account):
    if not account["ssl"]:
        mailbox = imaplib.IMAP4(account["host"], account["port"])
    else:
        mailbox = imaplib.IMAP4_SSL(account["host"], account["port"])
    mailbox.login(account["username"], account["password"])
    folder_list = mailbox.list()[1]
    mailbox.logout()
    return folder_list


# DSN:
# defaults to INBOX, path represents a single folder:
#  imap://username:password@imap.gmail.com:993/
#  imap://username:password@imap.gmail.com:993/INBOX
#
# get all folders
#  imap://username:password@imap.gmail.com:993/__ALL__
#
# singe folder with ssl, both are the same:
#  imaps://username:password@imap.gmail.com:993/INBOX
#  imap://username:password@imap.gmail.com:993/INBOX?ssl=true
#
# folder as provided as path or as query param "remote_folder" with comma separated list
#  imap://username:password@imap.gmail.com:993/INBOX.Drafts
#  imap://username:password@imap.gmail.com:993/?remote_folder=INBOX.Drafts
#
# combined list of folders with path and ?remote_folder
#  imap://username:password@imap.gmail.com:993/INBOX.Drafts?remote_folder=INBOX.Sent
#
# with multiple remote_folder:
#  imap://username:password@imap.gmail.com:993/?remote_folder=INBOX.Drafts
#  imap://username:password@imap.gmail.com:993/?remote_folder=INBOX.Drafts,INBOX.Sent
#
# setting other parameters
#  imap://username:password@imap.gmail.com:993/?name=Account1
def get_account(dsn, name=None):
    account = {
        "name": "account",
        "host": None,
        "port": 993,
        "username": None,
        "password": None,
        "remote_folder": "INBOX",  # String (might contain a comma separated list of folders)
        "ssl": False,
    }

    parsed_url = urllib.parse.urlparse(dsn)

    if parsed_url.scheme.lower() not in ["imap", "imaps"]:
        raise ValueError('Scheme must be "imap" or "imaps"')

    account["ssl"] = parsed_url.scheme.lower() == "imaps"

    if parsed_url.hostname:
        account["host"] = parsed_url.hostname

    if parsed_url.port:
        account["port"] = parsed_url.port
    if parsed_url.username:
        account["username"] = urllib.parse.unquote(parsed_url.username)
    if parsed_url.password:
        account["password"] = urllib.parse.unquote(parsed_url.password)

    # prefill account name, if none was provided (by config.cfg) in case of calling it from commandline. can be overwritten by the query param 'name'
    if name:
        account["name"] = name

    else:
        if account["username"]:
            account["name"] = account["username"]

        if account["host"]:
            account["name"] += "@" + account["host"]

    if parsed_url.path != "":
        account["remote_folder"] = parsed_url.path.lstrip("/").rstrip("/")

    if parsed_url.query != "":
        query_params = urllib.parse.parse_qs(parsed_url.query)

        # merge query params into account
        for key, value in query_params.items():

            if key == "remote_folder":
                if account["remote_folder"] is not None:
                    account["remote_folder"] += "," + value[0]
                else:
                    account["remote_folder"] = value[0]

            elif key == "ssl":
                account["ssl"] = value[0].lower() == "true"

            # merge all others params, to be able to overwrite username, password, ... and future account options
            else:
                account[key] = value[0] if len(value) == 1 else value

    return account
