#%%
from win32com.client import Dispatch
from bs4 import BeautifulSoup


class OutlookConnection:
    def __init__(self):
        self.outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(6)
        self.messages = self.inbox.items
        self.new_RFI = {}
        self.new_submittal = {}

    def check_email(self):
        for message in self.messages:
            sender = str(message.SenderEmailAddress)
            received = str(message.ReceivedTime.strftime("%m/%d/%Y"))
            subject = str(message.Subject)

            if sender == "CAToolsSystem@nbbj.com":
                if "RFI" in subject:
                    self.parse_rfi_emails(message, received)
                else:
                    self.parse_submittal_emails(message, received)

    def parse_rfi_emails(self, message, received):
        # Convert email to HTML for parsing
        message_html = BeautifulSoup(message.HTMLBody, "lxml")

        # An RFI may be a revision (e.g. .1 added to end) and we don't want to remove that information
        rfi_number = float(
            message_html.find("td", text=("RFI No:")).find_next().get_text(strip=True)
        )

        if rfi_number.is_integer():
            rfi_number = str(int(rfi_number))
        else:
            rfi_number = str(round(rfi_number, 1))

        # Get RFI subject (removing the RFI-#### tag) and link to the RFI attachment
        rfi_subject = str(
            message_html.find("td", text=("Subject:")).find_next().get_text(strip=True)
        )
        rfi_subject = rfi_subject.split(f"{rfi_number} ")[-1]
        rfi_file_url = message_html.find("div", text=("RFI Attachments:")).find_next(
            "a"
        )["href"]

        # Save information to dictionary
        self.new_RFI[rfi_number] = {
            "Description": rfi_subject,
            "dateReceived": received,
            "URL": rfi_file_url,
        }

    def parse_submittal_emails(self, message, received):
        # Convert email to HTML for parsing
        message_html = BeautifulSoup(message.HTMLBody, "lxml")

        # Find and save information about the submittal
        submittal_number = (
            message_html.find("td", text=("NBBJ Sub No:"))
            .find_next()
            .get_text(strip=True)
        )
        submittal_subject = str(
            message_html.find("td", text=("Specific Item:"))
            .find_next()
            .get_text(strip=True)
        )
        submittal_spec = str(
            message_html.find("td", text=("Spec. Number:"))
            .find_next()
            .get_text(strip=True)
        )
        submittal_type = str(
            message_html.find("td", text=("Spec. Description:"))
            .find_next()
            .get_text(strip=True)
        )
        submittal_url = message_html.find(
            "div", text=("Submittal Attachments:")
        ).find_next("a")["href"]

        #  Save information to dictionary
        self.new_submittal[submittal_number] = {
            "Description": submittal_subject,
            "SpecNumber": submittal_spec,
            "SpecType": submittal_type,
            "dateReceived": received,
            "URL": submittal_url,
        }


# %%
