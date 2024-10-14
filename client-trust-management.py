# ----------------
# App Improvements
# ----------------

# 1 Separate button for daily deposits to store list
# 5 when you make add_deposits_to_store remember to check the previous thursday's store books
# 5 add_ins_outs needs to account for people phasing up or down on tuesdays
#7 generate_deposits_sheet can be updated to make the totals cells dynamically update as deposits are added
#8 add_withdrawals should check that the patient has the money first
#9 Make a NO REFILL check for loading money onto account

# -------
# Imports
# -------

import os
import sys
import openpyxl
import pyodbc
from PyQt5.QtWidgets import (QApplication, QMainWindow,
                             QPushButton, QVBoxLayout,
                             QHBoxLayout, QTextEdit,
                             QLabel, QWidget, QGroupBox)
from PyQt5.QtCore import Qt
import pandas as pd
import xlwings as xw
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment
import shutil

# --------------
# ALL FILE PATHS
# --------------

# Access Database:
database_path = r'I:\Client Trust\Client Trust.accdb'

# Automated-Deposits Excel sheet:
auto_deposits_path = r'C:\Users\Dcude\OneDrive - Principles Inc\Desktop\2024 Deposits\Automated-Deposits-Sheet.xlsx'

# Automated-Withdrawals Excel sheet:
auto_withdrawals_path = r'C:\Users\Dcude\OneDrive - Principles Inc\Desktop\2024 Withdrawals\Automated-Withdrawals-Sheet.xlsx'

# Automated-InsOuts Excel sheet:
auto_ins_outs_path = r'C:\Users\Dcude\OneDrive - Principles Inc\Desktop\Ins N Outs\Automated-InsOuts.xlsx'

# Store List Folder:
store_list_folder_path = r'C:\Users\Dcude\OneDrive - Principles Inc\Desktop\Store List 2024'

# Store List Linked To Access Excel sheet:
linked_to_access_path = r'C:\Users\Dcude\OneDrive - Principles Inc\Desktop\Store List 2024\Store List Linked To Access.xlsx'

# Deposits Folder path:
deposits_folder_path = r'C:\Users\Dcude\OneDrive - Principles Inc\Desktop\2024 Deposits'

# Withdrawals Folder path:
withdrawals_folder_path = r'C:\Users\Dcude\OneDrive - Principles Inc\Desktop\2024 Withdrawals'

# Database Connection String
connection_string = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + database_path + ';'
)

# -----------------
# Main Window Class
# -----------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set up the main window
        self.setWindowTitle("Client Trust Management")
        self.setGeometry(200, 200, 600, 400)

        # Main layout
        main_layout = QVBoxLayout()

        # Horizontal layout to hold the groups
        button_layout = QHBoxLayout()

        # 1. Organization and Deposits Group
        org_deposits_group = QGroupBox("Organization and Deposits")
        org_deposits_group.setAlignment(Qt.AlignCenter)
        org_deposits_layout = QVBoxLayout()

        self.ins_and_outs_button = QPushButton("Add Ins && Outs to Access")
        self.ins_and_outs_button.clicked.connect(self.add_ins_outs)
        org_deposits_layout.addWidget(self.ins_and_outs_button)

        self.generate_deposit_sheet_button = QPushButton("Generate New Deposits Sheet")
        self.generate_deposit_sheet_button.clicked.connect(self.generate_deposits_sheet)
        org_deposits_layout.addWidget(self.generate_deposit_sheet_button)

        org_deposits_group.setLayout(org_deposits_layout)
        button_layout.addWidget(org_deposits_group)

        # 2. Store List Group
        store_list_group = QGroupBox("Store List")
        store_list_group.setAlignment(Qt.AlignCenter)
        store_list_layout = QVBoxLayout()

        self.store_list_button = QPushButton("Generate Today's Store List")
        self.store_list_button.clicked.connect(self.generate_store_list)
        store_list_layout.addWidget(self.store_list_button)

        self.daily_store_deposits_button = QPushButton("Add Daily Deposits to Store List")
        self.daily_store_deposits_button.clicked.connect(self.deposits_to_store)
        store_list_layout.addWidget(self.daily_store_deposits_button)

        self.replenish_books_thurs_button = QPushButton("Store Balances to $100")
        self.replenish_books_thurs_button.clicked.connect(self.replenish_store_balances_thurs)
        store_list_layout.addWidget(self.replenish_books_thurs_button)

        store_list_group.setLayout(store_list_layout)
        button_layout.addWidget(store_list_group)

        # 3. Balancing the Client Trust Database Group
        balance_group = QGroupBox("Balancing Client Trust")
        balance_group.setAlignment(Qt.AlignCenter)
        balance_layout = QVBoxLayout()

        self.generate_withdrawal_sheet_button = QPushButton("Generate New Withdrawals Sheet")
        self.generate_withdrawal_sheet_button.clicked.connect(self.generate_withdrawals_sheet)
        balance_layout.addWidget(self.generate_withdrawal_sheet_button)

        self.deposit_button = QPushButton("Add Deposits to Access")
        self.deposit_button.clicked.connect(self.add_deposits)
        balance_layout.addWidget(self.deposit_button)

        self.withdrawal_button = QPushButton("Add Withdrawals to Access")
        self.withdrawal_button.clicked.connect(self.add_withdrawals)
        balance_layout.addWidget(self.withdrawal_button)

        self.discharge_button = QPushButton("Discharge $0.00 Balances")
        self.discharge_button.clicked.connect(self.discharge_patients)
        balance_layout.addWidget(self.discharge_button)

        balance_group.setLayout(balance_layout)
        button_layout.addWidget(balance_group)

        # Add the horizontal layout to the main layout
        main_layout.addLayout(button_layout)

        # Text box for displaying results
        self.result_box = QTextEdit()
        self.result_box.setReadOnly(True)
        main_layout.addWidget(self.result_box)

        # Container widget
        container = QWidget()
        container.setLayout(main_layout)

        # Set the central widget to the container
        self.setCentralWidget(container)

    def discharge_patients(self):
        # Default connection
        connection = None
        try:
            # Establish the connection
            connection = pyodbc.connect(connection_string)

            # Create a cursor object using the connection
            cursor = connection.cursor()
            query = '''
                SELECT FirstName, LastName, Phase, [Sum Of DepositAmount], [Sum Of WithdrawalAmount]
                FROM Balance
                '''
            cursor.execute(query)
            rows = cursor.fetchall()

            # Clear previous results
            self.result_box.clear()

            # Keep count of discharged patients
            patient_discharge_count = 0

            # Iterate over the rows and print them
            for row in rows:
                # Get name
                first_name = row[0]
                last_name = row[1]

                # Get phase
                phase = int(row[2])

                # Get deposits and withdrawals
                deposit = float(row[3])
                withdrawal = float(row[4])

                # Calculate balance
                balance = deposit - withdrawal

                # Mark Discharged if patient is phase 4 and balance = 0.00
                if phase == 4 and balance == 0.0000:
                    self.result_box.append(f"Discharging: {last_name}, {first_name}, "
                                           f"Phase: {phase}, Balance: {balance}")
                    update_query = '''
                        UPDATE Clients
                        SET Discharged = True
                        WHERE FirstName = ? AND LastName = ?
                    '''
                    cursor.execute(update_query, (first_name, last_name))
                    connection.commit()
                    patient_discharge_count = patient_discharge_count + 1

            # If no patients were discharge, print message for the user
            if patient_discharge_count == 0:
                self.result_box.append(f"No patients met the discharge criteria.")

        # Print error if found
        except pyodbc.Error as e:
            self.result_box.setText(f"Error: {e}")
        except Exception as e:
            self.result_box.setText(f"Unexpected error: {e}")
        finally:
            if connection:
                connection.close()

    def add_deposits(self):
        # Path to Excel file
        excel_path = auto_deposits_path

        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel(excel_path)
        connection = None

        try:
            # Establish the connection to Access Database
            connection = pyodbc.connect(connection_string)

            # Create a cursor object using the Access connection
            cursor = connection.cursor()

            # Clear previous results
            self.result_box.clear()

            # Iterate over each row in the Excel Dataframe
            for index, row in df.iterrows():
                excel_first_name = row['FirstName']
                excel_last_name = row['LastName']
                excel_type = row['Type']
                excel_amount = row['Amount']

                # Query the Access database for the client
                query = """
                                SELECT ClientID, FirstName, LastName, Phase
                                FROM Clients
                                WHERE FirstName = ? AND LastName = ?
                            """

                cursor.execute(query, (excel_first_name, excel_last_name))
                client = cursor.fetchone()

                # Check if the client exists
                if not client:
                    # If no client is found, print a message
                    self.result_box.append(
                        f"{excel_first_name} {excel_last_name} not found in Access. Deposit was not added.")
                    self.result_box.append("")
                    continue

                client_id, first_name, last_name, phase = client

                # Check if the client's phase is 4
                if phase == 4:
                    # If phase is 4, print a message and skip the deposit
                    self.result_box.append(
                        f"{excel_first_name} {excel_last_name} found in Access, but Phase: 4. Deposit was not added.")
                    self.result_box.append("")
                    continue

                # Query the Access database
                query = """
                    SELECT ClientID, FirstName, LastName
                    FROM Clients
                    WHERE FirstName = ? AND LastName = ?
                """

                cursor.execute(query, (excel_first_name, excel_last_name))
                client = cursor.fetchone()

                if client:
                    client_id = client[0]

                    # Get today's date for the transaction
                    today_date = datetime.today().strftime('%m/%d/%Y')

                    # Check if a matching transaction already exists
                    check_query = """
                    SELECT TransactionID
                    FROM Transactions
                    WHERE TransactionDate = ? AND TransactionDescription = ? 
                    AND DepositAmount = ? AND ClientID = ?
                    """
                    cursor.execute(check_query, (today_date, excel_type, excel_amount, client_id))
                    existing_transaction = cursor.fetchone()

                    if existing_transaction:
                        # If transaction already exists, print message
                        self.result_box.append(
                            f"Transaction already exists for {excel_first_name} {excel_last_name} "
                            f"with the amount {excel_amount} on {today_date}.")
                        self.result_box.append("")
                    else:
                        # Get the highest current transaction number from the Transactions table
                        max_transaction_query = "SELECT MAX(TransactionID) FROM Transactions"
                        cursor.execute(max_transaction_query)
                        max_transaction = cursor.fetchone()[0]  # Fetch the current maximum

                        # Create the new transaction number by adding 1
                        new_transaction_number = max_transaction + 1

                        # Get today's date for the transaction
                        today_date = datetime.today().strftime('%m/%d/%Y')

                        # Insert the new transaction entry
                        insert_query = """
                                        INSERT INTO Transactions (TransactionID, TransactionDate, 
                                        TransactionDescription, DepositAmount, WithdrawalAmount, ClientID)
                                        VALUES (?, ?, ?, ?, ?, ?)
                                    """

                        deposit_amount = excel_amount
                        withdrawal_amount = 0.00

                        cursor.execute(insert_query, (new_transaction_number, today_date,
                                                      excel_type, deposit_amount, withdrawal_amount, client_id))

                        self.result_box.append(
                            f"Deposited: {deposit_amount} to {excel_first_name} "
                            f"{excel_last_name}'s account on {today_date}.")
                        self.result_box.append("")

                        # Commit the transaction
                        connection.commit()

        except FileNotFoundError:
            self.result_box.setText("Deposits Excel file not found.")

        except pyodbc.Error as e:
            self.result_box.setText(f"Error: {e}")

        except Exception as e:
            self.result_box.setText(f"Unexpected error: {e}")

        finally:
            if connection:
                connection.close()

    def add_withdrawals(self):
        # Path to Excel file
        excel_path = auto_withdrawals_path
        connection = None
        try:
            # Load the Excel file into a pandas DataFrame
            df = pd.read_excel(excel_path)

            # Establish the connection to Access Database
            connection = pyodbc.connect(connection_string)

            # Create a cursor object using the Access connection
            cursor = connection.cursor()

            # Clear previous results
            self.result_box.clear()

            # Iterate over each row in the Excel Dataframe
            for index, row in df.iterrows():
                excel_first_name = row['FirstName']
                excel_last_name = row['LastName']
                excel_type = row['Type']
                excel_amount = row['Amount']

                # Query the Access database for the client
                query = """
                                SELECT ClientID, FirstName, LastName, Phase
                                FROM Clients
                                WHERE FirstName = ? AND LastName = ?
                            """

                cursor.execute(query, (excel_first_name, excel_last_name))
                client = cursor.fetchone()

                # Check if the client exists
                if not client:
                    # If no client is found, print a message
                    self.result_box.append(
                        f"{excel_first_name} {excel_last_name} not found in Access. Withdrawal was not added.")
                    self.result_box.append("")
                    continue

                client_id, first_name, last_name, phase = client

                # Check if the client's phase is 4
                if phase == 4:
                    # If phase is 4, print a message and skip the deposit
                    self.result_box.append(
                        f"{excel_first_name} {excel_last_name} found in Access, but Phase: 4. Withdrawal was not added.")
                    self.result_box.append("")
                    continue

                # Get today's date for the transaction
                today_date = datetime.today().strftime('%m/%d/%Y')

                # Check if a matching transaction already exists
                check_query = """
                            SELECT TransactionID
                            FROM Transactions
                            WHERE TransactionDate = ? AND TransactionDescription = ? 
                            AND WithdrawalAmount = ? AND ClientID = ?
                            """
                cursor.execute(check_query, (today_date, excel_type, excel_amount, client_id))
                existing_transaction = cursor.fetchone()

                if existing_transaction:
                    # If transaction already exists, print message
                    self.result_box.append(
                        f"Transaction already exists for {excel_first_name} {excel_last_name} "
                        f"with the amount {excel_amount} on {today_date}.")
                    self.result_box.append("")
                else:
                    # Get the highest current transaction number from the Transactions table
                    max_transaction_query = "SELECT MAX(TransactionID) FROM Transactions"
                    cursor.execute(max_transaction_query)
                    max_transaction = cursor.fetchone()[0]  # Fetch the current maximum

                    # Create the new transaction number by adding 1
                    new_transaction_number = max_transaction + 1

                    # Insert the new transaction entry
                    insert_query = """
                                                INSERT INTO Transactions (TransactionID, TransactionDate, 
                                                TransactionDescription, DepositAmount, WithdrawalAmount, ClientID)
                                                VALUES (?, ?, ?, ?, ?, ?)
                                            """

                    deposit_amount = 0.00
                    withdrawal_amount = excel_amount

                    cursor.execute(insert_query, (new_transaction_number, today_date,
                                                  excel_type, deposit_amount, withdrawal_amount, client_id))

                    self.result_box.append(
                        f"Withdrew: {withdrawal_amount} from {excel_first_name} "
                        f"{excel_last_name}'s account on {today_date}.")
                    self.result_box.append("")

                    # Commit the transaction
                    connection.commit()

        except FileNotFoundError:
            self.result_box.setText("Withdrawal Excel file not found.")

        except pyodbc.Error as e:
            self.result_box.setText(f"Error: {e}")

        except Exception as e:
            self.result_box.setText(f"Unexpected error: {e}")
        finally:
            if connection:
                connection.close()

    def add_ins_outs(self):
        # Default connection
        connection = None

        try:
            # Ins and Outs file path
            file_path = auto_ins_outs_path

            # Read into a dataframe
            df = pd.read_excel(file_path)

            # Establish connection to Access Database
            connection = pyodbc.connect(connection_string)
            cursor = connection.cursor()

            # Step 1: Get the max ClientID from the Clients table
            max_id_query = "SELECT MAX(ClientID) FROM Clients"
            cursor.execute(max_id_query)
            max_client_id = cursor.fetchone()[0]

            self.result_box.clear()

            # Step 2: Insert new clients from admissions_df into the Clients table
            for index, row in df.iterrows():
                entry_type = row['Type']
                first_name = row['FirstName']
                last_name = row['LastName']

                if entry_type == 'A':
                    # Check if the client already exists in the database
                    client_check_query = """
                            SELECT ClientID FROM Clients
                            WHERE FirstName = ? AND LastName = ?
                        """

                    cursor.execute(client_check_query, first_name, last_name)
                    existing_client = cursor.fetchone()

                    if existing_client is not None:
                        self.result_box.append(f"{first_name} {last_name} is already in Access Database.")
                        self.result_box.append(
                            "Check with records and manually enter the patient into Access Database.")
                        self.result_box.append("")
                    else:
                        client_id = max_client_id + 1  # Incrementing for each new client
                        max_client_id = client_id
                        first_name = row['FirstName']
                        last_name = row['LastName']
                        phase = 1
                        discharged = False
                        comments = None  # No comments as per your request
                        contract = row['Contract']

                        # SQL query to insert the new client
                        insert_query = """
                            INSERT INTO Clients (ClientID, FirstName, LastName, Phase, Discharged, Comments, Contract)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                            """

                        # Execute the query for each client
                        cursor.execute(insert_query, client_id, first_name, last_name, phase, discharged, comments,
                                       contract)

                        # Get the max Transaction ID from the Transactions table
                        max_transaction_id_query = "SELECT MAX(TransactionID) FROM Transactions"
                        cursor.execute(max_transaction_id_query)
                        max_transaction_id = cursor.fetchone()[0]

                        if max_transaction_id is None:
                            max_transaction_id = 1
                        else:
                            max_transaction_id += 1

                        # Transaction details
                        transaction_id = max_transaction_id
                        transaction_date = datetime.now().strftime('%m/%d/%Y')  # Today's date
                        transaction_description = 'Beginning Balance'
                        deposit_amount = 0.00
                        withdrawal_amount = 0.00

                        # SQL query to insert a new transaction
                        transaction_query = """
                                INSERT INTO Transactions (TransactionID, TransactionDate, TransactionDescription, DepositAmount, WithdrawalAmount, ClientID)
                                VALUES (?, ?, ?, ?, ?, ?)
                                """

                        # Execute the query to add the new transaction
                        cursor.execute(transaction_query, transaction_id, transaction_date, transaction_description,
                                       deposit_amount,
                                       withdrawal_amount, client_id)

                        self.result_box.append(f"{first_name} {last_name} added to Access Database.")
                        self.result_box.append("")
                elif entry_type == 'D':
                    # Step 2: Insert Discharges from admissions_df into the Clients table on Access Database
                    reason_for_discharge = row['ReasonForDischarge']

                    # Find the Access ClientID by matching first and last names
                    client_id_query = """
                        SELECT ClientID, Phase FROM Clients
                        WHERE FirstName = ? AND LastName = ?
                    """

                    cursor.execute(client_id_query, first_name, last_name)
                    client_data = cursor.fetchone()

                    if client_data is not None:
                        client_id, current_phase = client_data

                        # Check if the patient is already phase 4
                        if int(current_phase) == 4:
                            self.result_box.append(f"{first_name} {last_name} has already been discharged.")
                            self.result_box.append("")
                        else:
                            # Get the max TransactionID from the Transactions table
                            max_transaction_id_query = """
                                    SELECT MAX(TransactionID) FROM Transactions
                                    """
                            cursor.execute(max_transaction_id_query)
                            max_transaction_id = cursor.fetchone()[0]

                            max_transaction_id = max_transaction_id + 1

                            # Transaction details for discharge
                            transaction_id = max_transaction_id
                            transaction_date = datetime.now().strftime("%m/%d/%Y")
                            transaction_description = reason_for_discharge
                            deposit_amount = 0.00
                            withdrawal_amount = 0.00

                            # SQL query to insert a new discharge transaction
                            transaction_query = """
                                        INSERT INTO Transactions (TransactionID, TransactionDate, TransactionDescription, 
                                        DepositAmount, WithdrawalAmount, ClientID)
                                        VALUES (?, ?, ?, ?, ?, ?)
                                    """

                            cursor.execute(transaction_query, transaction_id, transaction_date,
                                           transaction_description, deposit_amount, withdrawal_amount, client_id)

                            update_phase_query = """
                                        UPDATE Clients
                                        SET Phase = 4
                                        WHERE ClientID = ?
                                    """
                            cursor.execute(update_phase_query, client_id)

                            self.result_box.append(f"{first_name} {last_name} discharged from Access Database.")
                            self.result_box.append("")
                    else:
                        self.result_box.append(
                            f"Client {first_name} {last_name} not found in Access Database. No discharge entry added.")
                        self.result_box.append("")
                else:
                    self.result_box.append(f"Invalid entry type for {first_name} {last_name}")

            # Commit the transaction
            connection.commit()

        except FileNotFoundError:
            self.result_box.setText("Ins n Outs Excel file not found.")
        except pyodbc.Error as e:
            self.result_box.setText(f"Error: {e}")
        except Exception as e:
            self.result_box.setText(f"Unexpected error: {e}")
        finally:
            if connection:
                connection.close()

    def replenish_store_balances_thurs(self):
        # Default connection
        connection = None

        try:
            connection = pyodbc.connect(connection_string)
            cursor = connection.cursor()

            # Get today's date in the format MM-DD-YY
            today = datetime.now().strftime("%m-%d-%y")

            # Construct the expected file name
            expected_file_name = f"Store List_{today}.xlsx"

            # Build the full file path
            folder_path = store_list_folder_path
            excel_file_path = os.path.join(folder_path, expected_file_name)

            self.result_box.clear()

            try:
                # Use xlwings to open the workbook and recalculate all formulas
                app = xw.App(visible=False)  # Run Excel in the background
                workbook = xw.Book(excel_file_path)
                worksheet = workbook.sheets[0]

                # Recalculate the workbook to ensure all formulas are updated
                workbook.api.RefreshAll()  # Refresh all data connections and formulas
                worksheet.api.Calculate()  # Recalculate worksheet formulas

                # Iterate over the rows, starting at row 2
                for row in range(2, worksheet.range('A1').current_region.last_cell.row + 1):
                    last_name = worksheet.range(f'A{row}').value
                    first_name = worksheet.range(f'B{row}').value
                    store_balance = worksheet.range(f'G{row}').value  # Get the value from formula in Column G

                    # Stop the loop if last_name is None
                    if last_name is None:
                        break
                    # Query the Access database for the matching LastName and FirstName
                    query = """
                    SELECT [Sum of DepositAmount], [Sum of WithdrawalAmount]
                    FROM Balance
                    WHERE [LastName] = ? AND [FirstName] = ?
                    """

                    cursor.execute(query, (last_name, first_name))
                    result = cursor.fetchone()

                    if result:
                        deposit_sum = result[0] if result[0] is not None else 0
                        withdrawal_sum = result[1] if result[1] is not None else 0

                        balance = float(deposit_sum - withdrawal_sum)
                        current_add_col = worksheet.range(f'F{row}').value \
                            if worksheet.range(f'F{row}').value is not None else 0

                        store_balance = float(store_balance)

                        # Only proceed if balance > 0 to avoid negative values being added
                        if balance > 0:
                            # Check if adding balance exceeds 100
                            if (balance + store_balance) > 100:
                                add_amount = 100 - store_balance - current_add_col
                                worksheet.range(f'F{row}').value = current_add_col + add_amount
                                self.result_box.append(
                                    f"Added {add_amount} to {first_name} {last_name}'s store balance.")
                            else:
                                worksheet.range(f'F{row}').value = current_add_col + balance
                                self.result_box.append(f"Added {balance} to {first_name} {last_name}'s store balance.")

                # Save the workbook to store recalculated values (optional)
                workbook.save()
                workbook.close()
                app.quit()

            except Exception as e:
                self.result_box.setText(f"Error loading Excel file: {e}")
                return

        except FileNotFoundError:
            self.result_box.setText("Store List Excel file not found.")
        except pyodbc.Error as e:
            self.result_box.setText(f"Error: {e}")
        except Exception as e:
            self.result_box.setText(f"Unexpected error: {e}")
        finally:
            if connection:
                connection.close()

    def generate_store_list(self):
        # Default connection
        connection = None
        try:
            # Create a connection to the Access database
            connection = pyodbc.connect(connection_string)
            cursor = connection.cursor()

            # Query to fetch clients with Phase = 1
            cursor.execute("SELECT LastName, FirstName FROM Clients WHERE Phase = '1'"
                           "ORDER BY LastName ASC, FirstName ASC")

            # Fetch all results
            clients = cursor.fetchall()

            # Set the folder path to save the new Store List workbook
            folder_path = store_list_folder_path

            # Get today's data and format it as MM-DD-YY
            today = datetime.today().strftime('%m-%d-%y')

            # Create the file name with today's date
            file_name = f'Store List_{today}.xlsx'

            # Full path to save the workbook
            file_path = os.path.join(folder_path, file_name)

            # Create a new blank Store List and select the active sheet
            sl = openpyxl.Workbook()
            ws = sl.active

            # Add headers
            ws['A1'] = 'LastName'
            ws['B1'] = 'FirstName'
            ws['C1'] = 'Store List Balance'
            ws['D1'] = 'Spent At Store'
            ws['E1'] = 'Quarters'
            ws['F1'] = 'Add to List'
            ws['G1'] = 'Balance'

            # Load the linked Excel file (Store List_Linked_To_Access)
            linked_file_path = linked_to_access_path
            linked_wb = openpyxl.load_workbook(linked_file_path, data_only=True)
            linked_ws = linked_wb.active

            # Iterate over clients and populate new file
            for idx, client in enumerate(clients, start=2):
                last_name = client[0]
                first_name = client[1]

                # Add names to the new sheet
                ws[f'A{idx}'] = last_name  # LastName in column A
                ws[f'B{idx}'] = first_name  # FirstName in column B

                # Search for the matching name in the linked file
                for row in linked_ws.iter_rows(min_row=2, values_only=True):  # Start from row 2, skip header
                    linked_last_name = row[0]
                    linked_first_name = row[1]
                    linked_balance = row[6]  # Column G (balance) is the 7th column, index 6

                    # Check if the names match
                    if last_name.lower() == linked_last_name.lower() and first_name.lower() == linked_first_name.lower():
                        ws[f'C{idx}'] = linked_balance  # Add the balance to the new file (column C)
                        break  # Stop searching once the match is found

                # Add formula to the G column (Balance) for the current row
                ws[f'G{idx}'] = f'=C{idx}-D{idx}-E{idx}+F{idx}'

            # After all clients are added, find the row after the last client
            last_row = len(clients) + 2  # Row after the last client
            first_client_row = 2  # First client always starts at row 2

            # Add the SUM formulas from the first to the last client
            ws[f'C{last_row}'] = f"=SUM(C{first_client_row}:C{last_row - 1})"
            ws[f'D{last_row}'] = f"=SUM(D{first_client_row}:D{last_row - 1})"
            ws[f'E{last_row}'] = f"=SUM(E{first_client_row}:E{last_row - 1})"
            ws[f'F{last_row}'] = f"=SUM(F{first_client_row}:F{last_row - 1})"
            ws[f'G{last_row}'] = f"=SUM(G{first_client_row}:G{last_row - 1})"

            # Create the Left list of clients
            left_title_row = len(clients) + 4

            # Merge columns A and B in the left_title_row
            ws.merge_cells(f'A{left_title_row}:B{left_title_row}')

            # Add the text "Left" in the merged cell
            ws[f'A{left_title_row}'] = "Left"
            ws[f'C{left_title_row}'] = "Store List Balance"
            ws[f'D{left_title_row}'] = "Spent At Store"
            ws[f'E{left_title_row}'] = "Quarters"
            ws[f'F{left_title_row}'] = "Balance"

            # Apply yellow background to the merged cell
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            ws[f'A{left_title_row}'].fill = yellow_fill

            # Center the text in the merged cell
            center_alignment = Alignment(horizontal="center", vertical="center")
            ws[f'A{left_title_row}'].alignment = center_alignment

            # Create a set of full names (LastName, FirstName) for easy comparison
            client_names = {(client[0].lower(), client[1].lower()) for client in clients}

            # Start filling the "Left" list, one row below the "Left" title row
            left_list_row = left_title_row + 1

            # Iterate through the linked Excel file and find names not in the clients list
            for row in linked_ws.iter_rows(min_row=2, values_only=True):
                linked_last_name = row[0]
                linked_first_name = row[1]
                linked_balance = row[6]

                # Only proceed if both last name and first name are not None
                if linked_last_name and linked_first_name:
                    # Compare full name from linked list with client names
                    if (linked_last_name.lower(), linked_first_name.lower()) not in client_names:
                        # Add unmatched name to the "Left" list
                        ws[f'A{left_list_row}'] = linked_last_name
                        ws[f'B{left_list_row}'] = linked_first_name
                        # Populate column C with the balance (Column G in linked file)
                        ws[f'C{left_list_row}'] = linked_balance
                        ws[f'F{left_list_row}'] = f'=C{left_list_row}-D{left_list_row}-E{left_list_row}'
                        left_list_row += 1  # Move to the next row for the next unmatched name

            # After all clients are added, find the row after the last client
            last_left_row = left_list_row
            first_left_client_row = left_title_row + 1

            # After all clients are added, check if any clients were added to the "Left" list
            if left_list_row > left_title_row + 1:
                # Add the SUM formulas from the first to the last client
                ws[f'C{last_left_row}'] = f"=SUM(C{first_left_client_row}:C{last_left_row - 1})"
                ws[f'D{last_left_row}'] = f"=SUM(D{first_left_client_row}:D{last_left_row - 1})"
                ws[f'E{last_left_row}'] = f"=SUM(E{first_left_client_row}:E{last_left_row - 1})"
                ws[f'F{last_left_row}'] = f"=SUM(F{first_left_client_row}:F{last_left_row - 1})"
            else:
                # If no clients were added, set the cells to 0.00
                ws[f'C{left_list_row}'] = 0.00
                ws[f'D{left_list_row}'] = 0.00
                ws[f'E{left_list_row}'] = 0.00
                ws[f'F{left_list_row}'] = 0.00

            # Title the final Calculation row
            final_calc_row = last_left_row + 2

            ws[f'C{final_calc_row}'] = "Total"
            ws[f'D{final_calc_row}'] = "Spent"
            ws[f'E{final_calc_row}'] = "Quarters"
            ws[f'F{final_calc_row}'] = "Left"
            ws[f'G{final_calc_row}'] = "Added"
            ws[f'H{final_calc_row}'] = "Balance"

            # Highlight the final calculation row
            ws[f'C{final_calc_row}'].fill = yellow_fill
            ws[f'D{final_calc_row}'].fill = yellow_fill
            ws[f'E{final_calc_row}'].fill = yellow_fill
            ws[f'F{final_calc_row}'].fill = yellow_fill
            ws[f'G{final_calc_row}'].fill = yellow_fill
            ws[f'H{final_calc_row}'].fill = yellow_fill

            # Enter the formulas
            ws[f'C{final_calc_row + 1}'] = f'=C{last_row}+C{last_left_row}'
            ws[f'D{final_calc_row + 1}'] = f'=D{last_row}+D{last_left_row}'
            ws[f'E{final_calc_row + 1}'] = f'=E{last_row}+E{last_left_row}'
            ws[f'F{final_calc_row + 1}'] = f'=F{last_left_row}'
            ws[f'G{final_calc_row + 1}'] = f'=F{last_row}'
            ws[f'H{final_calc_row + 1}'] = (f'=C{final_calc_row + 1}-D{final_calc_row + 1}'
                                            f'-E{final_calc_row + 1}-F{final_calc_row + 1}+G{final_calc_row + 1}')

            # Save and open the new workbook
            sl.save(file_path)
            os.startfile(file_path)

            # Clear previous results and print confirmation
            self.result_box.clear()
            self.result_box.append(f"{file_name} created.")

        except FileNotFoundError:
            self.result_box.setText("Linked file not found.")
        except pyodbc.Error as e:
            self.result_box.setText(f"Error: {e}")
        except Exception as e:
            self.result_box.setText(f"Unexpected error: {e}")
        finally:
            if connection:
                connection.close()

    def generate_deposits_sheet(self):
        try:
            # Path to the folder containing the files
            folder_path = deposits_folder_path

            # Get the list of files in the folder
            files = os.listdir(folder_path)

            # Filter only the files with the correct prefix
            deposit_files = [f for f in files if f.startswith("Deposits for Client Trust")]

            if not deposit_files:
                self.result_box.setText("No deposit files found.")
                return

            # Sort files by the date in the format MM-DD-YY (removing .xlsx extension)
            deposit_files.sort(key=lambda x: datetime.strptime(x.split()[-1].replace('.xlsx', ''),
                                                               '%m-%d-%y'), reverse=True)
            # Get the latest file
            latest_file = deposit_files[0]

            # Create the full path to the source file
            source_file = os.path.join(folder_path, latest_file)

            # Get today's date
            today_date = datetime.now().strftime('%m-%d-%y')

            # Create the full path to the destination file with today's date
            destination_file = os.path.join(folder_path, f"Deposits for Client Trust {today_date}.xlsx")

            # Copy the file to the destination
            shutil.copy(source_file, destination_file)

            # Now modify the new copy to clear specified cells
            # Load the workbook
            workbook = openpyxl.load_workbook(destination_file)

            # Iterate over both Sheet1 and Sheet2
            for sheet_name in ['Sheet 1', 'Sheet 2']:
                sheet = workbook[sheet_name]

                # Edit cell B5 to contain the current date
                sheet['B5'].value = f"Date: {today_date}"

                # Clear columns B, D, F, H in rows 8 through 33
                for row in range(8, 29):
                    for col in ['B', 'D', 'F', 'H']:
                        sheet[f'{col}{row}'].value = None  # Clear the content

            # Save the modified workbook
            workbook.save(destination_file)

            # Update the result box to show success
            self.result_box.setText(f"Successfully created file: {destination_file}")

            # Open the newly created file
            os.startfile(destination_file)

        except FileNotFoundError:
            self.result_box.setText("Deposit file not found.")
        except pyodbc.Error as e:
            self.result_box.setText(f"Database error: {e}")
        except Exception as e:
            self.result_box.setText(f"Unexpected error: {e}")

    def generate_withdrawals_sheet(self):
        try:
            # Path to the folder containing the files
            folder_path = withdrawals_folder_path

            # Get the list of files in the folder
            files = os.listdir(folder_path)

            # Filter only the files with the correct prefix
            withdrawal_files = [f for f in files if f.startswith("Withdrawals for Client Trust")]

            if not withdrawal_files:
                self.result_box.setText("No withdrawal files found.")
                return

            # Sort files by the date in the format MM-DD-YY (removing .xlsx extension)
            withdrawal_files.sort(key=lambda x: datetime.strptime(x.split()[-1].replace('.xlsx', ''),
                                                                  '%m-%d-%y'), reverse=True)
            # Get the latest file
            latest_file = withdrawal_files[0]

            # Create the full path to the source file
            source_file = os.path.join(folder_path, latest_file)

            # Get today's date
            today_date = datetime.now().strftime('%m-%d-%y')

            # Create the full path to the destination file with today's date
            destination_file = os.path.join(folder_path, f"Withdrawals for Client Trust {today_date}.xlsx")

            # Copy the file to the destination
            shutil.copy(source_file, destination_file)

            # Now modify the new copy to clear specified cells
            # Load the workbook
            workbook = openpyxl.load_workbook(destination_file)

            # Iterate over both Sheet1 and Sheet2
            for sheet_name in ['Sheet 1', 'Sheet 2']:
                sheet = workbook[sheet_name]

                # Edit cell B5 to contain the current date
                sheet['B5'].value = f"Date: {today_date}"

                # Clear columns B, D, F, H in rows 8 through 33
                for row in range(8, 30):
                    for col in ['B', 'D', 'F', 'H']:
                        sheet[f'{col}{row}'].value = None  # Clear the content

            # Save the modified workbook
            workbook.save(destination_file)

            # Update the result box to show success
            self.result_box.setText(f"Successfully created file: {destination_file}")

            # Open the newly created file
            os.startfile(destination_file)

        except FileNotFoundError:
            self.result_box.setText("Withdrawal file not found.")
        except pyodbc.Error as e:
            self.result_box.setText(f"Database error: {e}")
        except Exception as e:
            self.result_box.setText(f"Unexpected error: {e}")

    def deposits_to_store(self):
        try:
            # Load the Auto-Deposits Excel file into a DataFrame
            auto_deposits_df = pd.read_excel(auto_deposits_path, engine='openpyxl')

            # Load today's Store List Excel file into a Dataframe
            folder_path = store_list_folder_path

            # Load Today's Store List Excel file into a DataFrame
            today = datetime.now()
            today_date = today.strftime("%m-%d-%y")
            today_store_file_name = f"Store List_{today_date}.xlsx"
            today_store_file_path = os.path.join(folder_path, today_store_file_name)
            today_store_df = pd.read_excel(today_store_file_path, engine='openpyxl')

            # ---- Now handle the previous Thursday ----

            # Get last thursday's date
            # If today is Thursday (weekday == 3), subtract 7 days to get the last Thursday
            if today.weekday() == 3:
                last_thursday = today - timedelta(days=7)
            else:
                # Otherwise, find how many days ago the last Thursday was (Thursday is weekday 3)
                days_ago = (today.weekday() - 3) % 7
                last_thursday = today - timedelta(days=days_ago)

            # Get last Thursday's Store List
            last_thursday_date = last_thursday.strftime("%m-%d-%y")
            last_thursday_file_name = f"Store List_{last_thursday_date}.xlsx"

            # Create the full file path for last Thursday's store list
            last_thursday_store_list_path = os.path.join(folder_path, last_thursday_file_name)

            # Load last Thursday's store list into a DataFrame (Third DataFrame)
            thursday_df = pd.read_excel(last_thursday_store_list_path, engine='openpyxl')

            # Filter out rows with NaN values in 'First Name' or 'Last Name', and exclude rows with 'Last Name' as 'Left'
            thursday_df = thursday_df.dropna(
                subset=['FirstName', 'LastName'])  # Only drop NaNs in 'FirstName' or 'LastName'
            thursday_df = thursday_df[thursday_df['LastName'] != 'Left']  # Exclude rows with 'LastName' == 'Left'

            # Get Today's Store List as a workbook you can edit
            today_store_wb = openpyxl.load_workbook(today_store_file_path)
            today_store_ws = today_store_wb.active

            # Iterate over the common names and check for a match in thursday_df
            for index, row in auto_deposits_df.iterrows():
                match = thursday_df[
                    (thursday_df['FirstName'] == row['FirstName']) & (thursday_df['LastName'] == row['LastName'])]
                if not match.empty:
                    last_name = match['LastName'].values[0]
                    first_name = match['FirstName'].values[0]
                    today_deposit = auto_deposits_df.loc[(auto_deposits_df['LastName'] == last_name)
                                                         & (auto_deposits_df['FirstName'] == first_name), 'Amount'].values[0]
                    thurs_add = thursday_df.loc[(thursday_df['LastName'] == last_name)
                                                         & (thursday_df['FirstName'] == first_name), 'Add to List'].values[0]
                    # Check if thurs_add is NaN, and if so, set it to 0.00
                    if pd.isna(thurs_add):
                        thurs_add = 0.00

                    today_bal = today_store_df.loc[(today_store_df['LastName'] == last_name)
                                                         & (today_store_df['FirstName'] == first_name), 'Balance'].values[0]
#### STILL NOT WORKING ##########
                    if (today_deposit + thurs_add) > 100:
                        add_val = 100 - thurs_add
                        if (add_val + today_bal) > 100:
                            add_val = 100 - today_bal
                            # Iterate through rows of the worksheet to find the match in columns A (FirstName) and B (LastName)
                            for ws_row in today_store_ws.iter_rows(min_row=2, max_col=6,
                                                                   values_only=False):  # Assuming column F is column 6
                                ws_first_name = ws_row[0].value  # FirstName in column A
                                ws_last_name = ws_row[1].value  # LastName in column B

                                # If FirstName and LastName match, modify column F
                                if ws_first_name == first_name and ws_last_name == last_name:
                                    today_store_ws[f"F{ws_row[0].row}"].value = add_val  # Modify column F at this row
                                    print(f"Updated {first_name} {last_name}'s cell in column F with value: {add_val}")
                                else:
                                    print(f"No match for {last_name} {first_name} found in today's Store List")
                            # Save the workbook to apply the changes
                            try:
                                today_store_wb.save(today_store_file_path)
                                print("Workbook saved successfully.")
                            except Exception as e:
                                print(f"Error saving workbook: {e}")
                else:
                    print(f"{row['FirstName']} {row['LastName']}: No match found")
##### STILL NOT WORKING #######

        except FileNotFoundError:
            self.result_box.setText("Deposit or Store List file not found.")
        except pyodbc.Error as e:
            self.result_box.setText(f"Database error: {e}")
        except Exception as e:
            self.result_box.setText(f"Unexpected error: {e}")

# ------------------
# Main Program Logic
# ------------------

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())