"""
This class will take the raw 'Agency - Exclusions Report' from ServicePoint's ART and process it into a DQ report for
tracking inconsistent data entry by TPI staff members.  The rules this class checks against with the missing_data_check
method are not particularly flexible at this time and heavy modifications will need to be made to make this script work
for ServicePoint using agencies.
"""

import numpy as np
import pandas as pd

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

class incidentReport:
    def __init__(self):
        self.file = askopenfilename(title="Open the Agency - Exclusion Report")
        self.raw_data = pd.read_excel(self.file, sheet_name="Exclusions")

    def missing_data_check(self, data_frame):
        """
        The raw 'Agency - Exclusion Report' will be processed by this method using multiple numpy.select calls to make
        sure that each of the fields matches the current best practices for the TPI agency.

        :param data_frame: This should be a pandas data frame created from the 'Agency - Exclusion Report' ART report
        :return: a data frame showing incidents with errors will be returned
        """

        missing_df = data_frame
        check_columns = {
            "Infraction Provider": "Provider Error",
            "Infraction Banned End Date": "End Date Error",
            "Infraction Staff Person": "Staff Name Error",
            "Infraction Type": "Incident Error",
            "Infraction Banned Code": "Incident Code Error",
            "Infraction Banned Sites": "Sites Excluded From Error",
            "Infraction Notes": "Notes Error"
        }
        incident_types = [
            "Non-compliance with program",
            "Violent Behavior",
            "Police Called",
            "Alcohol",
            "Drugs"
        ]
        incident_codes = [
            "Bar - Other",
            "TPI_Exclusion - Agency (requires reinstatement)"
        ]

        for column in check_columns.keys():
            if column == "Infraction Provider":
                conditions = [(missing_df[column] == "Transition Projects (TPI) - Agency - SP(19)")]
                choices = ["Incorrect Provider"]
                missing_df[check_columns[column]] = np.select(conditions, choices, default=None)
            elif column == "Infraction Banned End Date":
                conditions = [
                    (
                        missing_df[column].notna() &
                        (missing_df["Infraction Banned Code"] == "TPI_Exclusion - Agency (Requires Reinstatement)")
                    )
                ]
                choices = ["Banned Date Should Be Blank"]
                missing_df[check_columns[column]] = np.select(conditions, choices, default=None)
            elif column == "Infraction Staff Person":
                conditions = [(missing_df[column].isna())]
                choices = ["No Staff Name Entered"]
                missing_df[check_columns[column]] = np.select(conditions, choices, default=None)
            elif column == "Infraction Type":
                conditions = [missing_df[column].isna(), ~(missing_df[column].isin(incident_types))]
                choices = ["No Incident Selected", "Non-TPI Incident Selected"]
                missing_df[check_columns[column]] = np.select(conditions, choices, default=None)
            elif column == "Infraction Banned Code":
                conditions = [(missing_df[column].isna()), ~(missing_df[column].isin(incident_codes))]
                choices = ["No Incident Code Selected", "Non-TPI Incident Code Selected"]
                missing_df[check_columns[column]] = np.select(conditions, choices, default=None)
            elif column == "Infraction Banned Sites":
                conditions = [(missing_df[column].isna())]
                choices = ["No Sites Excluded From Entry"]
                missing_df[check_columns[column]] = np.select(conditions, choices, default=None)
            elif column == "Infraction Notes":
                conditions = [
                    missing_df[column].isna(),
                    (
                        missing_df[column].str.contains("uno") |
                        missing_df[column].str.contains("UNO")
                    )
                ]
                choices = ["No Notes Entered", "Use of department specific shorthand"]
                missing_df[check_columns[column]] = np.select(conditions, choices, default=None)
            else:
                pass

        missing_df =  missing_df[[
            "Client Uid",
            "Infraction User Creating",
            "Infraction User Updating",
            "Infraction Banned Start Date",
            "Provider Error",
            "End Date Error",
            "Staff Name Error",
            "Incident Error",
            "Incident Code Error",
            "Sites Excluded From Error",
            "Notes Error"
        ]]
        missing_df["Infraction Banned Start Date"] = missing_df["Infraction Banned Start Date"].dt.date
        return missing_df

    def process(self):
        """
        This method will call the missing_data_check method then create a excel spreadsheet with moth an Errors sheet
        and a Raw Data sheet.  These will then be saved using an asksaveasfilename function call.

        :return: True will be returned if the method completes correctly.
        """
        raw = self.raw_data.copy()[[
            "Client Uid",
            "Infraction User Creating",
            "Infraction User Updating",
            "Infraction Provider",
            "Infraction Date Added",
            "Infraction Banned Start Date",
            "Infraction Banned End Date",
            "Infraction Staff Person",
            "Infraction Type",
            "Infraction Banned Code",
            "Infraction Banned Sites",
            "Infraction Notes"
        ]]
        errors = self.missing_data_check(self.raw_data.copy())
        writer = pd.ExcelWriter(asksaveasfilename(title="Save As"), engine="xlsxwriter")
        errors.to_excel(writer, sheet_name="Errors", index=False)
        raw.to_excel(writer, sheet_name="Raw Data", index=False)
        writer.save()
        return True

if __name__ == "__main__":
    a = incidentReport()
    a.process()