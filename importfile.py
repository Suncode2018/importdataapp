import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from dotenv import load_dotenv
import os
from sqlalchemy import create_engine, text
import pandas as pd
import io
import xml.etree.ElementTree as ET
import warnings
import threading  # Import threading module

# Ignore warnings
warnings.simplefilter("ignore")

# Load environment variables from .env file
load_dotenv()

# Retrieve database values from environment variables
server = os.getenv("DB_SERVER")
database = os.getenv("DB_DATABASE")
username = os.getenv("DB_USERNAME")
password = os.getenv("DB_PASSWORD")
driver = os.getenv("DB_DRIVER")
dbwh = os.getenv("DB_WH")  # Database warehouse identifier

# Create the database connection engine
def create_db_connection():
    try:
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect=DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}")
        return engine
    except Exception as e:
        messagebox.showerror("Database Connection Error", f"Failed to connect to database: {str(e)}")
        return None

# Function to fetch values for a combobox from the database
def fetch_combobox_values():
    engine = create_db_connection()
    if engine:
        try:
            with engine.connect() as connection:
                sql_query = text("EXEC [dbo].[spFileJob] 'showFile'")
                result = connection.execute(sql_query).fetchall()
                combobox_values = [row[0] for row in result]
                return combobox_values
        except Exception as e:
            messagebox.showerror("SQL Query Error", f"Failed to execute query: {str(e)}")
            return []
    else:
        return []

# Function to check if a file with the same name already exists in the database
def is_duplicate_file(file_name):
    engine = create_db_connection()
    if engine:
        try:
            with engine.connect() as connection:
                sql_query = text("SELECT COUNT(*) FROM [dbo].[tblFile] WHERE NameFile = :name")
                result = connection.execute(sql_query, {"name": file_name}).scalar()
                return result > 0  # Returns True if a duplicate exists
        except Exception as e:
            messagebox.showerror("SQL Query Error", f"Failed to check for duplicate file: {str(e)}")
            return False
    else:
        return False

# Function to validate the file type against the expected type for the selected job
def validate_file_type(name_job, file_path):
    engine = create_db_connection()
    if engine:
        try:
            with engine.connect() as connection:
                sql_query = text("SELECT xFile FROM [dbo].[tblFileJob] WHERE xNameJob = :namejob")
                result = connection.execute(sql_query, {"namejob": name_job}).fetchone()
                if result:
                    expected_file_type = result[0]  # Expected file type from the database
                    actual_file_type = os.path.splitext(file_path)[1].lower()  # Actual file type from the file path
                    if actual_file_type == f".{expected_file_type.lower()}":
                        return True
                    else:
                        messagebox.showwarning("File Type Mismatch", f"Expected file type: {expected_file_type}, but selected file type: {actual_file_type}")
                        return False
                else:
                    messagebox.showwarning("Invalid Job", f"No file type found for job: {name_job}")
                    return False
        except Exception as e:
            messagebox.showerror("Validation Error", f"Failed to validate file type: {str(e)}")
            return False
    else:
        return False

# Function to insert file data into the database
def insert_file_to_db(file_path, name_job):
    try:
        # Step 1: Validate the file type
        if not validate_file_type(name_job, file_path):
            return  # Stop if validation fails

        # Step 2: Extract file name
        file_name = os.path.basename(file_path)

        # Step 3: Check for duplicate file names
        if is_duplicate_file(file_name):
            messagebox.showwarning("Duplicate File", f"A file with the name '{file_name}' already exists in the database.")
            return

        # Step 4: Read the file as binary data
        with open(file_path, "rb") as file:
            binary_data = file.read()

        # Step 5: Extract file type
        file_type = os.path.splitext(file_name)[1]  # Get file extension

        # Step 6: Create the SQL query to insert data into tblFile
        sql_query = text("""
            INSERT INTO [dbo].[tblFile] (NameJob, NameFile, TypeFile, binDataFile)
            VALUES (:namejob, :name, :type, :data)
        """)

        # Step 7: Execute the query
        engine = create_db_connection()
        if engine:
            with engine.connect() as connection:
                connection.execute(sql_query, {"namejob": name_job, "name": file_name, "type": file_type, "data": binary_data})
                connection.commit()
                messagebox.showinfo("Success", "File inserted into the database successfully!")
        else:
            messagebox.showerror("Error", "Failed to connect to the database.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to insert file into the database: {str(e)}")

# Function to fetch data from tblFile and return both column names and data
def fetch_file_data():
    engine = create_db_connection()
    if engine:
        try:
            with engine.connect() as connection:
                sql_query = text("SELECT NameJob, NameFile, TypeFile FROM [dbo].[tblFile]")
                result = connection.execute(sql_query)
                columns = result.keys()
                cleaned_columns = [col.replace("RMKeyView(['", "").replace("'])", "") for col in columns]
                data = [tuple(row) for row in result.fetchall()]
                if len(data) == 0:
                    return [], []
                return cleaned_columns, data
        except Exception as e:
            messagebox.showerror("SQL Query Error", f"Failed to fetch data: {str(e)}")
            return [], []
    else:
        return [], []

# Function to open the Import File form
def open_import_file_form():
    import_file_window = tk.Toplevel()
    import_file_window.title(f"Import File - {dbwh}")

    # Set size to 70% of monitor
    screen_width = import_file_window.winfo_screenwidth()
    screen_height = import_file_window.winfo_screenheight()
    width = int(screen_width * 0.7)
    height = int(screen_height * 0.7)
    import_file_window.geometry(f"{width}x{height}")

    import_file_window.resizable(True, True)

    # Center the Import File window
    center_window(import_file_window, width, height)

    # Set a background color for the entire window
    import_file_window.configure(bg="#f0f0f0")

    # Create a frame for the form (fills the entire window)
    form_frame = tk.Frame(import_file_window, bg="#ffffff", padx=20, pady=20)
    form_frame.pack(expand=True, fill="both", padx=1, pady=1)

    # Combined Frame for File Selection, Combobox, and Text File
    combined_frame = tk.Frame(form_frame, bg="#ffffff")
    combined_frame.pack(fill="x", pady=10)

    # Combobox Label and Combobox
    combobox_label = tk.Label(combined_frame, text="Select Option:", font=("Helvetica", 14), bg="#ffffff")
    combobox_label.pack(side="left", padx=5)

    combobox_values = fetch_combobox_values()  # Fetch values from the database
    combobox = ttk.Combobox(combined_frame, values=combobox_values, font=("Helvetica", 12), width=30)
    combobox.pack(side="left", padx=5)

    # New Row for TextBox, Browse Button, and Import Button
    new_row_frame = tk.Frame(form_frame, bg="#ffffff")
    new_row_frame.pack(fill="x", pady=10)

    # TextBox Label and Entry (TextBox)
    new_textbox_label = tk.Label(new_row_frame, text="TextBox:", font=("Helvetica", 14), bg="#ffffff")
    new_textbox_label.pack(side="left", padx=5)

    new_textbox_entry = tk.Entry(new_row_frame, font=("Helvetica", 12), width=50)
    new_textbox_entry.pack(side="left", padx=5)

    # "Browse" Button for New Row
    def new_browse_file():
        file_path = filedialog.askopenfilename(title="Select a File", filetypes=(("All files", "*.*"), ("Text files", "*.txt"), ("CSV files", "*.csv")))
        if file_path:
            new_textbox_entry.delete(0, tk.END)  # Clear the current entry
            new_textbox_entry.insert(0, file_path)  # Insert the selected file path

    new_browse_button = tk.Button(new_row_frame, text="Browse", font=("Helvetica", 12, "bold"), fg="white", bg="#4CAF50", width=12, height=1, command=new_browse_file, relief="flat", bd=0)
    new_browse_button.pack(side="left", padx=5)

    # "Import" Button in the same row as "Browse"
    def import_action():
        file_path = new_textbox_entry.get().strip()
        name_job = combobox.get().strip()  # Get the selected value from the combobox
        if file_path and name_job:  # Ensure both file path and combobox value are provided
            # Use threading to avoid freezing the GUI
            def on_import_success():
                load_file_data()  # Refresh the table data after successful import
                # messagebox.showinfo("Success", "File imported and data refreshed successfully!")

            def import_thread():
                try:
                    # Step 1: Execute the stored procedure to drop temporary tables
                    engine = create_db_connection()
                    if engine:
                        with engine.connect() as connection:
                            # Call the stored procedure to drop temporary tables
                            sql_query = text("EXEC [dbo].[spDropTableTemp]")
                            connection.execute(sql_query)
                            connection.commit()
                            # print("Temporary tables dropped successfully.")

                    # Step 2: Insert the file into the database
                    insert_file_to_db(file_path, name_job)

                    # Step 3: Schedule the callback on the main thread
                    import_file_window.after(0, on_import_success)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to execute stored procedure or import file: {str(e)}")

            # Start the import process in a separate thread
            threading.Thread(target=import_thread).start()
        else:
            messagebox.showwarning("Input Error", "Please select a file and a job option first.")

    import_button = tk.Button(new_row_frame, text="Import", font=("Helvetica", 12, "bold"), fg="white", bg="#4CAF50", width=12, height=1, command=import_action, relief="flat", bd=0)
    import_button.pack(side="left", padx=5)

    # Add a Table (Treeview) with dynamic columns
    table_frame = tk.Frame(form_frame, bg="#ffffff")
    table_frame.pack(fill="both", expand=True, pady=(10, 20))

    # Add vertical scrollbar
    vertical_scrollbar = ttk.Scrollbar(table_frame, orient="vertical")
    vertical_scrollbar.pack(side="right", fill="y")

    # Add horizontal scrollbar
    horizontal_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal")
    horizontal_scrollbar.pack(side="bottom", fill="x")

    # Create the Treeview widget
    table = ttk.Treeview(table_frame, show="headings", height=10, yscrollcommand=vertical_scrollbar.set, xscrollcommand=horizontal_scrollbar.set)
    table.pack(fill="both", expand=True)

    # Configure scrollbars
    vertical_scrollbar.config(command=table.yview)
    horizontal_scrollbar.config(command=table.xview)

    # Define alternating row colors
    table.tag_configure("evenrow", background="#f0f0f0")  # Light gray for even rows
    table.tag_configure("oddrow", background="#ffffff")   # White for odd rows
    table.tag_configure("placeholder", background="#ffcccc")  # Light red for placeholder rows

    # Function to load data from tblFile into the table
    def load_file_data():
        try:
            # Fetch column names and data from the database
            columns, file_data = fetch_file_data()

            # Clear existing data in the table
            for row in table.get_children():
                table.delete(row)

            # Clear existing columns in the table
            table["columns"] = columns
            for col in table["columns"]:
                table.heading(col, text=col, anchor="center")  # Center-align headers
                table.column(col, width=100, stretch=True, anchor="w")  # Left-align data

            # If no data is found, display a placeholder message
            if len(file_data) == 0:
                table.insert("", "end", values=["No data found"] * len(columns), tags=("placeholder",))
            else:
                # Insert data into the table
                for i, row in enumerate(file_data):
                    tag = "evenrow" if i % 2 == 0 else "oddrow"  # Alternating row colors
                    table.insert("", "end", values=row, tags=(tag,))

                # Dynamically adjust column widths based on content
                for col in columns:
                    max_width = max([len(str(table.set(item, col))) for item in table.get_children()] + [len(col)]) * 10
                    table.column(col, width=max_width)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file data: {str(e)}")

    # Schedule load_file_data to run after the form is shown
    import_file_window.after(0, load_file_data)

    # Exit Button Section
    exit_button_frame = tk.Frame(form_frame, bg="#ffffff")
    exit_button_frame.pack(fill="x", pady=20)

    # "Reset" Button
    def reset_action():
        try:
            confirm = messagebox.askyesno("Confirm Reset", "Are you sure you want to reset the table? This will delete all data in tblFile.")
            if confirm:
                engine = create_db_connection()
                if engine:
                    with engine.connect() as connection:
                        sql_query = text("TRUNCATE TABLE [dbo].[tblFile]")
                        connection.execute(sql_query)
                        connection.commit()
                        messagebox.showinfo("Reset Successful", "The table tblFile has been truncated.")
                        load_file_data()  # Refresh the table to show it's empty
                else:
                    messagebox.showerror("Error", "Failed to connect to the database.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to reset the table: {str(e)}")

    reset_button = tk.Button(exit_button_frame, text="Reset", font=("Helvetica", 12, "bold"), fg="white", bg="#FF9800", width=12, height=1, command=reset_action, relief="flat", bd=0)
    reset_button.pack(side="left", padx=5)

    # Add a Progress Bar
    progress_frame = tk.Frame(form_frame, bg="#ffffff")
    progress_frame.pack(fill="x", pady=10)

    progress_label = tk.Label(progress_frame, text="Progress:", font=("Helvetica", 14), bg="#ffffff")
    progress_label.pack(side="left", padx=5)

    progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=400, mode="determinate")
    progress_bar.pack(side="left", padx=5)

    # "Gen File" Button
    def gen_file_action():
        try:
            # Fetch data from the tblFile table
            columns, file_data = fetch_file_data()

            # Check if the table is empty
            if len(file_data) == 0:
                messagebox.showwarning("No Data", "The table tblFile is empty. No data to generate.")
                return  # Stop further processing if the table is empty

            # Reset progress bar
            progress_bar["value"] = 0
            progress_bar["maximum"] = len(file_data)

            # Use threading to avoid freezing the GUI
            threading.Thread(target=process_files, args=(file_data,)).start()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate file: {str(e)}")

    def process_files(file_data):
        # Loop through each row in the file_data
        for index, row in enumerate(file_data):
            # Extract data from the row
            name_job, name_file, type_file = row

            # Query the database to fetch TypeFile and binDataFile for the current NameJob
            engine = create_db_connection()
            if engine:
                with engine.connect() as connection:
                    sql_query = text("SELECT TypeFile, binDataFile FROM [dbo].[tblFile] WHERE NameFile = :namefile")
                    result = connection.execute(sql_query, {"namefile": name_file}).fetchone()

                    if result:
                        db_type_file, bin_data_file = result  # Fetch TypeFile and binDataFile from the database

                        # Replace "." with an empty string in db_type_file
                        db_type_file = db_type_file.replace(".", "")

                        # Check the file type and perform actions
                        if db_type_file.lower() == "xml":
                            # Action for XML files
                            try:
                                xml_data = io.BytesIO(bin_data_file)
                                ns = {"doc": "urn:schemas-microsoft-com:office:spreadsheet"}
                                tree = ET.parse(xml_data)
                                root = tree.getroot()
                                data = []
                                for i, node in enumerate(root.findall('.//doc:Row', ns)):
                                    row_data = {}
                                    cells = node.findall('doc:Cell', ns)
                                    for j, cell in enumerate(cells):
                                        data_node = cell.find('doc:Data', ns)
                                        row_data[f'{j + 1}'] = data_node.text if data_node is not None else None
                                    row_data[f'{len(cells) + 1}'] = name_file
                                    data.append(row_data)
                                dfx = pd.DataFrame(data)
                                if not dfx.empty:
                                    table_name = f'tblImport{dbwh}_{index}_xml'.replace('-', '').lower()
                                    dfx.to_sql(table_name, con=engine, if_exists='replace', index=False)
                            except ET.ParseError as e:
                                messagebox.showerror("Error", f"Error parsing XML file '{name_file}': {e}")
                            except Exception as e:
                                messagebox.showerror("Error", f"Error inserting into the database for file '{name_file}': {e}")

                        elif db_type_file.lower() == "xlsx":
                            # Action for XLSX files
                            try:
                                xlsx_data = io.BytesIO(bin_data_file)
                                excel_file = pd.ExcelFile(xlsx_data, engine='openpyxl')
                                all_sheets = excel_file.sheet_names
                                for sheet_index in range(len(all_sheets)):
                                    df_sheet = excel_file.parse(sheet_index, header=None) #=None  header=0
                                    df_sheet['name_file'] = name_file
                                    df_sheet.columns = [f'{i}' for i in range(len(df_sheet.columns))]
                                    table_name = f'tblImportExcel{dbwh}_{sheet_index}_xlsx'.replace("-", "").lower()
                                    try:
                                        with engine.begin():
                                            df_sheet.to_sql(table_name, con=engine, if_exists='replace', index=False)
                                    except Exception as e:
                                        messagebox.showerror("Error", f"Failed to save data from sheet index {sheet_index} to database: {e}")
                            except Exception as e:
                                messagebox.showerror("Error", f"Failed to process XLSX file: {str(e)}")

                        elif db_type_file.lower() == "csv":
                            # Action for CSV files
                            try:
                                csv_data = io.BytesIO(bin_data_file)
                                sample = csv_data.read(1024)
                                csv_data.seek(0)

                                # Check the delimiter
                                if b"|" in sample:
                                    delimiter = "|"
                                    table_prefix = "Pipe"
                                elif b"," in sample:
                                    delimiter = ","
                                    table_prefix = "Comma"
                                else:
                                    messagebox.showwarning("Delimiter Error", f"Unsupported delimiter in CSV file: {name_file}")
                                    continue  # Skip this file

                                # Read the CSV file using pandas
                                df_csv = pd.read_csv(csv_data, delimiter=delimiter, encoding="iso8859_11", header=None)

                                # Add the 'name_file' column to the DataFrame
                                df_csv['name_file'] = name_file

                                # Rename columns dynamically as [0], [1], [2], ..., [N]
                                df_csv.columns = [f'{i}' for i in range(len(df_csv.columns))]

                                # Clean up table name (replace spaces with underscores and convert to lowercase)
                                table_name = f'tblImport{table_prefix}{dbwh}_{index}_csv'.replace("-", "").lower()

                                try:
                                    with engine.begin():
                                        df_csv.to_sql(table_name, con=engine, if_exists='replace', index=False)
                                        # messagebox.showinfo("Success", f"CSV data for {name_file} imported into table '{table_name}' successfully!")
                                except Exception as e:
                                    messagebox.showerror("Error", f"Failed to save CSV data to database: {e}")
                            except Exception as e:
                                messagebox.showerror("Error", f"Failed to process CSV file: {str(e)}")

                        else:
                            # Action for other file types
                            messagebox.showinfo("File Type", f"Unsupported file type: {db_type_file} for job: {name_job}")
                    else:
                        messagebox.showwarning("No Data", f"No data found for file: {name_file}")
            else:
                messagebox.showerror("Error", "Failed to connect to the database.")

            # Update progress bar
            progress_bar["value"] = index + 1
            import_file_window.update_idletasks()  # Refresh the GUI

        messagebox.showinfo("Success", "All files processed successfully!")

    gen_file_button = tk.Button(exit_button_frame, text="Gen File", font=("Helvetica", 12, "bold"), fg="white", bg="#9C27B0", width=12, height=1, command=gen_file_action, relief="flat", bd=0)
    gen_file_button.pack(side="left", padx=5)

    # "Exit" Button
    exit_button = tk.Button(exit_button_frame, text="Exit", font=("Helvetica", 12, "bold"), fg="white", bg="#F44336", width=12, height=1, command=import_file_window.destroy, relief="flat", bd=0)
    exit_button.pack(side="left", padx=5)

# Function to center the window on the screen
def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

# Example usage (for testing)
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Set the application icon
    try:
        root.iconbitmap("app_icon.ico")  # Replace with the path to your .ico file
    except Exception as e:
        print(f"Failed to load icon: {e}")

    open_import_file_form()  # Open the Import File form
    root.mainloop()