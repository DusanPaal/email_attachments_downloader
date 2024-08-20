# pylint: disable = W0703

"""
Downloads message attachments.
"""

from datetime import datetime
import sys
from engine import mails

if __name__ == "__main__":

    print("Connecting to account...")
    account = mails.get_account(
        mailbox = "EMEA-GSS-AR@ledvance.com",
        name = "lbs.robot@ledvance.com",
        x_server = "outlook.office365.com"
    )

    print("Collecting messages...")
    messages = mails.fetch_messages(
        acc = account,
        email_from = "elab_apm_de@obi.de",
        from_date = datetime(2023, 7, 23),
        search_subfolders = False
    )

    n_total = len(messages)

    if n_total == 0:
        print("ERROR: No message found using the search criteria!")
        sys.exit(0)

    for i, msg in enumerate(messages, start = 1):

        print(f"Processig message {i} of {n_total} ...", end = " ")

        try:
            pdf_path = mails.download_attachments(
                msg, ext = ".pdf",
                dst_folder = r"C:\bia\AttDownloader\temp\pdf"
            )
        except Exception as exc:
            print(str(exc))
            continue

        if len(pdf_path) != 0:
            print(f"Attachment(s) saved to: {'; '.join(pdf_path)}")

    sys.exit(0)
