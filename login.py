import tkinter as tk
from tkinter import messagebox
from dotenv import load_dotenv
import os
from sqlalchemy import create_engine, text
import mainmenu  # Import the mainmenu module
import threading  # Import the threading module

# Load environment variables from .env file
load_dotenv()

# Retrieve database values from environment variables
server = os.getenv("DB_SERVER")
database = os.getenv("DB_DATABASE")
username = os.getenv("DB_USERNAME")
password = os.getenv("DB_PASSWORD")
driver = os.getenv("DB_DRIVER")
dbwh = os.getenv("DB_WH")  # Database warehouse identifier

# Function to handle the login action
def login_action():
    user = entry_username.get()
    passw = entry_password.get()

    def perform_login():
        try:
            # Create the connection engine
            engine = create_engine(
                f"mssql+pyodbc:///?odbc_connect=DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}"
            )

            # Define the SQL query to call the stored procedure
            sql_query = text(f"EXEC [dbo].[spUsers] 'myLogin','{user}', '{passw}'")

            # Execute the stored procedure
            with engine.connect() as conn:
                user_data = conn.execute(sql_query).fetchone()

            if not user_data:
                messagebox.showerror("Login Failed", "Invalid Username or Password")
            else:
                root.withdraw()  # Hide the login window
                mainmenu.open_main_menu()  # Open the main menu from mainmenu.py

        except Exception as e:
            messagebox.showerror("Database Connection Failed", f"Error: {str(e)}")

    # Run the login operation in a separate thread
    threading.Thread(target=perform_login).start()

# Function to center the window on the screen
def center_window(window, width, height):
    # Get screen dimensions
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # Calculate x and y coordinates to center the window
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)

    # Set the window's position
    window.geometry(f"{width}x{height}+{x}+{y}")

# Function to switch focus to the password field
def focus_password(event):
    if entry_username.get().strip() != "":  # Check if username is not empty
        entry_password.focus_set()  # Focus on the password field

# Function to trigger login action on Enter key in password field
def login_on_enter(event):
    login_action()  # Call the login_action function

# Function to set the icon for a window
def set_icon(window):
    try:
        window.iconbitmap("app_icon.ico")  # Replace with the path to your .ico file
    except Exception as e:
        print(f"Failed to load icon: {e}")

# Create the main window
root = tk.Tk()
root.title(f"Login - {dbwh}")  # Dynamic title with database warehouse identifier
root.geometry("400x300")  # Initial window size
root.resizable(False, False)

# Set the application icon for the main window
set_icon(root)

# Center the window on the screen
center_window(root, 400, 300)

# Set a background color
root.configure(bg="#f0f0f0")

# Create a frame for the form (dynamic size)
form_frame = tk.Frame(root, bg="#ffffff", padx=20, pady=20)
form_frame.pack(expand=True, fill="both", padx=1, pady=1)  # Dynamic size with padding

# Add a title label
title_label = tk.Label(
    form_frame,
    text="Login",
    font=("Helvetica", 24, "bold"),
    fg="#4CAF50",  # Green color for the title
    bg="#ffffff"
)
title_label.pack(pady=(0, 20))

# Username field
username_frame = tk.Frame(form_frame, bg="#ffffff")
username_frame.pack(fill="x", pady=(0, 10))

label_username = tk.Label(
    username_frame,
    text="Username:",
    font=("Helvetica", 12),
    bg="#ffffff"
)
label_username.pack(side="left", padx=(0, 10))

entry_username = tk.Entry(
    username_frame,
    font=("Helvetica", 12),
    width=25,
    bd=1,
    relief="solid"
)
entry_username.pack(side="right", expand=True, fill="x")

# Password field
password_frame = tk.Frame(form_frame, bg="#ffffff")
password_frame.pack(fill="x", pady=(0, 10))

label_password = tk.Label(
    password_frame,
    text="Password:",
    font=("Helvetica", 12),
    bg="#ffffff"
)
label_password.pack(side="left", padx=(0, 10))

entry_password = tk.Entry(
    password_frame,
    show="*",
    font=("Helvetica", 12),
    width=25,
    bd=1,
    relief="solid"
)
entry_password.pack(side="right", expand=True, fill="x")

# Login Button
login_button = tk.Button(
    form_frame,
    text="Login",
    font=("Helvetica", 12, "bold"),
    fg="white",
    bg="#4CAF50",
    width=15,
    height=1,
    command=login_action,  # Use the new login_action function
    relief="flat",  # Remove the default button border
    bd=0  # Remove border
)
login_button.pack(pady=(20, 0))

# Set default cursor focus on the username field
entry_username.focus_set()

# Bind the Enter key to switch focus to the password field
entry_username.bind("<Return>", focus_password)

# Bind the Enter key in the password field to trigger login
entry_password.bind("<Return>", login_on_enter)

# Run the application
root.mainloop()