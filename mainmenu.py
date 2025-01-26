import tkinter as tk
from tkinter import messagebox
from dotenv import load_dotenv
import os
from importfile import open_import_file_form  # Import the function from importfile.py
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

# Function to handle the 'Import File' action
def import_file():
    # Run the file import operation in a separate thread
    threading.Thread(target=open_import_file_form).start()

# Function to open the main menu
def open_main_menu():
    main_menu = tk.Toplevel()
    main_menu.title(f"Main Menu - {dbwh}")  # Dynamic title with database warehouse identifier

    # Set size to 80% of monitor
    screen_width = main_menu.winfo_screenwidth()
    screen_height = main_menu.winfo_screenheight()
    width = int(screen_width * 0.75)
    height = int(screen_height * 0.75)
    main_menu.geometry(f"{width}x{height}")

    main_menu.resizable(True, True)

    # Center the main menu window
    center_window(main_menu, width, height)

    # Set a background color for the entire window
    main_menu.configure(bg="#f0f0f0")

    # Create a menubar
    menubar = tk.Menu(main_menu)

    # Create the 'File' menu
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)

    # Add 'Import File' option in the 'File' menu
    file_menu.add_command(label="Import File", command=import_file)

    # Add the menubar to the main menu window
    main_menu.config(menu=menubar)

    # Create a frame for the menu (fills the entire form)
    menu_frame = tk.Frame(main_menu, bg="#ffffff", padx=20, pady=20)
    menu_frame.pack(expand=True, fill="both", padx=1, pady=1)  # Expand to fill the entire window

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

# Example usage (for testing)
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Set the application icon
    try:
        root.iconbitmap("app_icon.ico")  # Replace with the path to your .ico file
    except Exception as e:
        print(f"Failed to load icon: {e}")

    open_main_menu()  # Open the main menu
    root.mainloop()