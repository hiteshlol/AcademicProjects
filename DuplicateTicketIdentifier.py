from datetime import date
import pandas as pd
from fuzzywuzzy import fuzz

"""Module Description: Create a module/model that compares open tickets to the historical set of closed tickets and 
current tickets in progress and returns a list of tickets that are most likely duplicates"""


class Ticket:
    def __init__(self, description, status='Open'):
        """
               Represents a ticket with a description, date entered, and status.

               Parameters:
               - description (str): The description of the ticket.
               - status (str): The status of the ticket (default is 'Open').
        """
        self.description = description
        self.date_entered = date.today()
        self.status = status

    def close_ticket(self):
        self.status = 'Closed'


class TicketSystem:
    """
    A system to check the simalirity between a newly submitted ticket and an entire list of closed and current
    tickets recorded and provide information about all duplicates present
    """

    def __init__(self):
        self.system = []

    def is_duplicate(self, description):
        """
        Checks if a ticket with a specified similarity in description already exists in the system.

        Parameters:
        - description (str): The description of the ticket to check.

        Returns:
        - bool: True if a duplicate is found, False otherwise.
        """

        # Check if a ticket with a 70% match in description already exists in the Excel file
        df = pd.DataFrame(self.system)
        if not df.empty:
            ratio_threshold = 70
            duplicates = df[df['Description'].apply(lambda desc: fuzz.ratio(description, desc) >= ratio_threshold)]
            num_duplicates = len(duplicates)
            if num_duplicates > 0:
                print(f'{num_duplicates} duplicate(s) found for the same issue.')
                return True

        return False

    def create_ticket(self, description):
        """
        Creates a new ticket and adds it to the system if it is not a duplicate.

        Parameters:
        - description (str): The description of the new ticket.
        """
        if not self.is_duplicate(description):
            tick_num = Ticket(description)
            tick_num.ticket_number = len(self.system) + 1
            self.system.append(tick_num)

    def export_to_excel(self, file_path='ticket_system.xlsx'):
        """
        Exports ticket information to an Excel file and handles duplicates.

        Parameters:
        - file_path (str): The file path for the Excel file (default is 'ticket_system.xlsx').
        """
        try:
            existing_df = pd.read_excel(file_path)
        except FileNotFoundError:
            existing_df = pd.DataFrame()

        # Determine the next available ticket number based on existing data in Excel
        next_ticket_number = existing_df['Ticket Number'].max() + 1 if not existing_df.empty else 1000
        data = {
            'Ticket Number': [],
            'Description': [],
            'Date Entered': [],
            'Status': []
        }
        duplicate_rows = []
        for ticket in self.system:
            # Check for duplicates based on the existing Excel data
            if not existing_df.empty and any(
                    fuzz.ratio(ticket.description, desc) >= 70 for desc in existing_df['Description'].values):
                print(f'Duplicate ticket found for the same issue')
                # Append the duplicate row to duplicate_rows
                duplicate_rows.append(existing_df[existing_df['Description'] == ticket.description])
                continue
            else:
                print('No duplicates exist')
            # Add ticket information to the data dictionary
            data['Ticket Number'].append(next_ticket_number)
            data['Description'].append(ticket.description)
            data['Date Entered'].append(ticket.date_entered)
            data['Status'].append(ticket.status)

            next_ticket_number += 1

        # Create a DataFrame from the new data
        new_df = pd.DataFrame(data)

        # Concatenate the existing DataFrame and the new data, excluding the index
        df = pd.concat([existing_df, new_df], ignore_index=True, axis=0)

        # Write to Excel file without the index and unnamed columns
        df.to_excel(file_path, index=False)

        # Print the entire row of duplicate tickets
        if duplicate_rows:
            count = 0
            print("\nDuplicate Ticket Entries:")
            for dup_df in duplicate_rows:
                for index, row in dup_df.iterrows():
                    count += 1
                    print(row[['Ticket Number', 'Description', 'Date Entered', 'Status']])
            print(f'Total {count} duplicate(s) found.')


ticket_system = TicketSystem()
description = input('What is the issue?:')
ticket_system.is_duplicate(description)
ticket_system.create_ticket(description)
ticket_system.export_to_excel()

