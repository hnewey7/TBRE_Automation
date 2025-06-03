import win32com.client
from win32com.client import gencache
import logging.config
import os
from datetime import datetime
import json
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import time

# - - - - - - - - - - - - - - - - - - - - -

global logger
logger = logging.getLogger()

# - - - - - - - - - - - - - - - - - - - - -


class InventorManager:
    def __init__(self):
        """
        A class for managing connection to Inventor application.

        Attributes
        ----------
        app : win32com.client.Dispatch
            Inventor Application
        """
        self.app = None  # Inventor Application

        # Connect to Inventor application.
        ret = self.connect_to_inventor()
        if not ret:
            logger.error("Failed to connect to Inventor application.")
        else:
            logger.info("Successfully connected to Inventor application.")

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
            CastTo = gencache.EnsureModule(
                "{D98A091D-3A0F-4C3E-B36E-61F62068D488}", 0, 1, 0
            )
            self.app = CastTo.Application(self.app)
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

    # - - - - - - - - - - - - - - - -
    # Methods for selecting documents.

    def select_document(self, filename: str = None) -> bool:
        """
        Select a document by prompting user to file.

        Args:
            filename (str): Filename, optional

        Returns:
            bool: True if the document is found and selected, False otherwise.
        """
        if not filename:
            # Open a file dialog to select an Inventor document.
            Tk().withdraw()
            filename = askopenfilename()

        if not filename:
            logger.warning("No file selected.")
            return False
        else:
            try:
                self.app.Documents.Open(filename)
                logger.info(f"Selected document: {filename}")
                return True
            except Exception as e:
                logger.error(f"Error selecting document '{filename}': {e}")
                return False

    # - - - - - - - - - - - - - - - -
    # Methods for getting document properties.

    def get_property(self, property_set: str, property: str) -> str:
        """
        Get property of document.

        Args:
            property_set (str): Set property in.
            property (str): Name of property.

        Returns:
            str: Property value.
        """
        doc = self.app.ActiveDocument

        # Get property set.
        try:
            property_set = doc.PropertySets.Item(property_set)
        except Exception as e:
            logger.error(f"Error getting property set: {e}")
            return None

        # Get property value.
        try:
            property_value = property_set.Item(property).Value
            logger.info("Successfully got property value.")
            return property_value
        except Exception as e:
            logger.error(f"Error getting property: {e}")
            return None


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
        ] = f"logs/{current_time}/{current_time}_RoverRPi.log"
        logging.config.dictConfig(config_dict)

    # Create an instance of InventorManager to test the connection.
    manager = InventorManager()
    manager.select_document(
        r"C:\Users\hnewe\Documents\tbre_automation_tools\example_files\INVENTOR\DOG TOOTH GEARBOX.iam"
    )
    part_number = manager.get_property("Design Tracking Properties", "Part Number")
    description = manager.get_property("Design Tracking Properties", "Description")
    print(part_number)
    print(description)
