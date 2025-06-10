"""
InventorAutomationApplication is a class for handling automation of Inventor.

Created on Monday 9th June 2025.
@author: Harry New

"""

from tkinter import Tk, Text, END
from tkinter.filedialog import askopenfilename
import win32com.client
import logging.config
import time
from datetime import datetime
import json
import os

from .FileAccessEventsHandler import FileAccessEventsHandler
from .MainWindow import MainWindow

# - - - - - - - - - - - - - - - - - - - - -

global logger
logger = logging.getLogger()

# - - - - - - - - - - - - - - - - - - - - -


class InventorAutomationApplication:
    def __init__(self):
        """
        Class for managing automation of Inventor.

        Attributes:
        app : win32com.client.Dispatch
            Inventor Application
        root :
            Window for tkinter UI.
        """
        self.app = None  # Inventor Application
        self.root = None  # Tkinter window.

        # Connect to Inventor application.
        ret = self.connect_to_inventor()
        if not ret:
            logger.error("Failed to connect to Inventor application.")
        else:
            logger.info("Successfully connected to Inventor application.")

        # Set file access handler.
        ret = self.add_file_access_handler(FileAccessEventsHandler)
        if not ret:
            logger.error("Failed to add file access handler.")
        else:
            logger.info("Successfully added file access handler.")

        # Create tkinter user interface.
        self.create_user_interface()

    def run(self):
        """
        Running Inventor Automation Application.
        """
        # Run user interface.
        self.root.mainloop()

    # - - - - - - - - - - - - - - - -
    # Methods for managing Inventor application connection

    def connect_to_inventor(self) -> bool:
        """
        Connect to an existing Inventor application instance.

        Returns:
            bool: True if connection is successful, False otherwise.
        """
        try:
            self.app = win32com.client.Dispatch("Inventor.Application")
            self.app.Visible = True

            return True
        except Exception as e:
            logger.error(f"Error connecting to Inventor: {e}")
            return False

    def disconnect(self):
        """
        Disconnect from the Inventor application.
        """
        try:
            if self.app:
                self.app.Quit()
                time.sleep(1)
                logger.info("Disconnected from Inventor application.")
                self.app = None
        except Exception as e:
            logger.error(f"Error disconnecting from Inventor: {e}")

    def add_file_access_handler(self, file_access_handler) -> bool:
        """
        Add handle for file access events.

        Args:
            file_access_handler: Handler for file access events.

        Returns:
            bool: True if added successfully, False otherwise.
        """
        try:
            # Set file access events.
            file_access_events = self.app.FileAccessEvents
            win32com.client.WithEvents(file_access_events, file_access_handler)
            return True
        except Exception as e:
            logger.error(f"Error adding file access handler: {e}")

    # - - - - - - - - - - - - - - - -
    # Methods for managing user interface.

    def create_user_interface(self):
        """
        Create tkinter user interface.
        """
        # Create root window.
        self.root = Tk()
        self.root.title("Inventor Automation Application")
        self.root.geometry("1200x540")

        # Get command mapping.
        commands = {"select_file": self.select_file, "export_parts_list": None}

        # Add main window.
        main_window = MainWindow(self.root, commands)
        main_window.pack()

    # - - - - - - - - - - - - - - - -
    # Methods for using Inventor.

    def select_file(self, filename: str = None, text_var: Text = None) -> bool:
        """
        Method for selecting file in Inventor.

        Args:
            filename (str): Filename, optional
            text_var (Text): Text variable to update, optional

        Returns:
            bool: Successful or not.
        """
        # Create dialogue for selecting file.
        if not filename:
            filename = askopenfilename(title="Select a file")

        # Open file.
        try:
            # Open document.
            self.doc = self.app.Documents.Open(filename, False)
            logger.info(f"Selected document: {filename}")
            if self.doc.DocumentType == 12291:
                self.assembly_doc = win32com.client.CastTo(self.doc, "AssemblyDocument")
            else:
                self.assembly_doc = win32com.client.CastTo(self.doc, "PartDocument")

            # Update text variable.
            if text_var:
                text_var.delete("1.0", END)
                text_var.insert(END, filename.split("/")[-1])
            return True, filename
        except Exception as e:
            logger.error(f"Error selecting document '{filename}': {e}")
            return False, None


# - - - - - - - - - - - - - - - - - - - - -


if __name__ == "__main__":
    # Set up logging to output to console.
    with open("config/logging_config.json") as f:
        config_dict = json.load(f)
        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

        # Creating log path.
        global log_path
        log_path = "logs/" + current_time

        # Create directory for current time.
        os.mkdir(log_path)
        # Set the filename for the log file.
        config_dict["handlers"]["file"][
            "filename"
        ] = f"logs/{current_time}/{current_time}.log"
        logging.config.dictConfig(config_dict)

    # Create an instance of Inventor Automation Application.
    application = InventorAutomationApplication()
    application.run()
