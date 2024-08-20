"""The 'mails.py' module creates and sends of emails directly via SMTP server.
It also uses the exchangelib library to connect to the Exchange server via
Exchange Web Services (EWS) in order to retrieve messages and save message
attachment under a specified account.

Version history:
----------------
1.0.20220112 - Initial version.
1.1.20220614 - Added retrieving message objects from the Exchange server accounts and
               saving message attachments to a local file.
"""
from datetime import datetime
from email.mime.multipart import MIMEMultipart
import os
from os.path import exists, isfile, join

import exchangelib as xlib
from exchangelib.queryset import QuerySet
from exchangelib import (
    Account, Build, Configuration, Identity,
    Message, Version, EWSDateTime, OAuth2Credentials
)

# custom message classes
class SmtpMessage(MIMEMultipart):
    """A wrapper for MIMEMultipart
    objects that are sent using
    SMTP server.
    """

# custom warnings
class UndeliveredWarning(Warning):
    """Raised when message is not
    delivered to all recipients.
    """

# custom exceptions
class FolderNotFoundError(Exception):
    """Raised when a directory is
    requested but doesn't exist.
    """

class AttachmentSavingError(Exception):
    """Raised when any exception is
    caught durng writing of attachment
    data to a file.
    """

class AttachmentNotFoundError(Exception):
    """Raised when a file attachment
    is requested but doesn't exist.
    """

class ParamNotFoundError(Exception):
    """Raised when a parameter required
    for creating credentials is not
    found in the source file.
    """

class InvalidSmtpHostError(Exception):
    """Raised when an invalid host name
    is used for SMTP connection.
    """


def _get_credentials(acc_name: str) -> OAuth2Credentials:
    """Returns credentails for a given account.

    Parameters:
    -----------
    acc_name:
        Name of the account for which
        the credentails will be obtained.

    Returns:
    --------
    The Credentials object.
    """

    cred_dir = join(os.environ["APPDATA"], "bia")
    cred_path = join(cred_dir, f"{acc_name.lower()}.token.email.dat")

    if not isfile(cred_path):
        raise FileNotFoundError(f"Credentials file not found: {cred_path}")

    with open(cred_path, 'r', encoding = "utf-8") as stream:
        lines = stream.readlines()

    params = dict(
        client_id = None,
        client_secret = None,
        tenant_id = None,
        identity = Identity(primary_smtp_address = acc_name)
    )

    for line in lines:

        if ":" not in line:
            continue

        tokens = line.split(":")
        param_name = tokens[0].strip()
        param_value = tokens[1].strip()

        if param_name == "Client ID":
            key = "client_id"
        elif param_name == "Client Secret":
            key = "client_secret"
        elif param_name == "Tenant ID":
            key = "tenant_id"

        params[key] = param_value

    # verify loaded parameters
    if params["client_id"] is None:
        raise ValueError("Parameter 'client_id' not found!")

    if params["client_secret"] is None:
        raise ValueError("Parameter 'client_secret' not found!")

    if params["tenant_id"] is None:
        raise ValueError("Parameter 'tenant_id' not found!")

    # params OK, create credentials
    creds = OAuth2Credentials(
        params["client_id"],
        params["client_secret"],
        params["tenant_id"],
        params["identity"]
    )

    return creds

def get_account(mailbox: str, name: str, x_server: str) -> Account:
    """Returns an account for a shared mailbox.

    Parameters:
    -----------
    mailbox:
        Name of the shared mailbox.

    name:
        Name of the account for which
        the credentails will be obtained.

    x_server:
        Name of the MS Exchange server.

    Returns:
    --------
    An exchangelib.Account object.
    """

    build = Build(major_version = 15, minor_version = 20)
    creds = _get_credentials(name)

    cfg = Configuration(creds,
        server = x_server,
        auth_type = xlib.OAUTH2,
        version = Version(build)
    )

    cfg = Configuration(server = x_server, credentials = creds)

    acc = Account(
        mailbox,
        config = cfg,
        autodiscover = False,
        access_type = xlib.IMPERSONATION
    )

    return acc

def fetch_messages(
        acc: Account, email_from: str, from_date: datetime = None,
        search_subfolders: bool = False) -> QuerySet:
    """Collects message objects from a given mailbox.

    Parameters:
    -----------
    acc:
        Account object containing reference to a mailbox.

    cust_dir:
        Customer-specific folder contained in the account's inbox.

    Returns:
    --------
    An exchangelib.QuerySet object that represents
    a collection of emails from a mailbox folder.
    """

    if from_date is not None:
        timezone = acc.default_timezone
        start = EWSDateTime.from_datetime(from_date).astimezone(timezone)
        end = EWSDateTime.from_datetime(datetime.now()).astimezone(timezone)

    if search_subfolders:
        folders = acc.inbox.walk()
    else:
        folders = acc.inbox.all()

    emails = folders.filter(
        datetime_received__range = (start, end),
        sender = email_from
    ).only(
        "subject", "sender", "attachments", "message_id",
    )

    return list(emails)

def download_attachments(
        msg: Message, dst_folder: str,  ext: str = None,
        overwrite: bool = False) -> list:
    """Saves message attachment(s) to file(s).

    Parameters:
    -----------
    msg:
        An exchangelib.Message object containing attachment(s).

    dst_folder:
        Path to the folder where attachments will be stored.

    ext:
        If None is used (default), then all attachments will be downloaded,
        regardless of their file type. If a file extension (e.g. '.pdf') is used,
        then only attachmnets of that particular file type will be downloaded.
        The parameter is case insensitive.

    Returns:
    --------
    A list of file paths to the stored attachments.
    """

    if not exists(dst_folder):
        raise FolderNotFoundError(f"Folder does not exist: {dst_folder}")

    file_paths = []

    for att in msg.attachments:

        file_path = join(dst_folder, att.name)

        if not (ext is None or file_path.lower().endswith(ext)):
            continue

        if isfile(file_path):

            if not overwrite:
                print("WARNING: File already exists. Attachment won't be saved.")
                continue

            print("WARNING: File already exists. Data will be ovewritten.")

        with open(file_path, 'wb') as stream:
            stream.write(att.content)

        if not isfile(file_path):
            raise FileNotFoundError(f"Error writing attachment data to file: {file_path}")

        file_paths.append(file_path)

    return file_paths
