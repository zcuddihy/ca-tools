#%%
import xlwings as xw
from datetime import datetime
import utils
import os
import re


class Submittal_Log:
    def __init__(self, new_submittals: dict, project_obj: object):
        self.excel_file = project_obj.excel_file
        self.main_file_path = project_obj.main_file_path
        self.wb = xw.Book(f"{self.main_file_path}\{self.excel_file}")
        self.new_submittals = new_submittals

    def get_lastRow(self, sheetName: str) -> int:

        lastRow = (
            self.wb.sheets[sheetName]
            .range("A" + str(self.wb.sheets[sheetName].cells.last_cell.row))
            .end("up")
            .row
            + 1
        )

        return lastRow

    def save(self):

        sheet_name = "Submittal Log"
        lastRow = self.get_lastRow(sheet_name)
        current_submittals = self.wb.sheets[sheet_name].range(f"A6:A{lastRow-1}").value

        for submittal in self.new_submittals:
            if submittal in current_submittals:
                continue
            else:
                rfi_description = self.new_submittals[RFI]["Description"]
                rfi_url = self.new_submittals[RFI]["URL"]

                # Create a new folder for the RFI and save the PDF to that location
                file_path, file_name = self.get_RFI_file_path(RFI, rfi_description)
                pdf_file = utils.get_pdf(rfi_url)
                utils.save_to_folder(file_name, pdf_file, file_path)

                # Update the excel log
                self.wb.sheets[sheet_name].range(lastRow, 1).add_hyperlink(
                    file_path, text_to_display=RFI
                )
                self.wb.sheets[sheet_name].range(lastRow, 2).value = self.new_RFI[RFI][
                    "dateReceived"
                ]
                self.wb.sheets[sheet_name].range(lastRow, 6).value = self.new_RFI[RFI][
                    "Description"
                ]

                lastRow += 1

    def get_RFI_file_path(self, rfi_number: str, rfi_subject: str) -> str:
        folder_name = f"{rfi_number} - {rfi_subject}"
        file_path = f"{self.main_file_path}\RFI\{folder_name}"
        return file_path


#%%

wb = xw.Book("I:\OHMC-Exp\Construction Admin\OHMC NET - CA Log.xlsx")
current_submittals = wb.sheets["Submittal Log"].range("A6:C10").value
current_submittals
# %%
