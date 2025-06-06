"""
InventorManager is a class that handles communcation with Inventor API.

Created on Friday 30th May 2025.
@author: Harry New

"""

import win32com.client
import logging.config
import os
from datetime import datetime
import json
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import time

from src.Part import Part

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
                self.doc = self.app.Documents.Open(filename, False)
                logger.info(f"Selected document: {filename}")
                if self.doc.DocumentType == 12291:
                    self.assembly_doc = win32com.client.CastTo(
                        self.doc, "AssemblyDocument"
                    )
                else:
                    self.assembly_doc = win32com.client.CastTo(self.doc, "PartDocument")
                return True
            except Exception as e:
                logger.error(f"Error selecting document '{filename}': {e}")
                return False

    def get_parts_list(self, document: str = None) -> dict:
        """
        Get parts list for assembly document.

        Args:
            document (str): Document, optional

        Returns:
            dict: Parts list with number of occurrences.
        """
        # Opening document.
        if not document:
            self.select_document()
        else:
            self.select_document(document)

        # Get occurrences.
        occurrences = self.assembly_doc.ComponentDefinition.Occurrences
        all_parts = []
        self.get_part_occurrences(occurrences, all_parts)

        return all_parts

    def get_part_occurrences(self, occurrences, parts_list: list):
        """
        Get part occurrences.

        Args:
            occurrences: Occurrences
            parts_list (list): List containing all parts.
        """
        for occ in occurrences:
            doc_type = occ.DefinitionDocumentType
            if doc_type == 12290:  # Part
                self.select_document(occ.Definition.Document.FullDocumentName)
                part = Part(self.assembly_doc)
                parts_list.append(part)
            elif doc_type == 12291:  # Assembly
                self.get_part_occurrences(occ.SubOccurrences, parts_list)

    def get_part_details(self, part: Part):
        """
        Get all details for an individual part.

        Args:
            part (Part): Part object.
        """
        self.select_document(part.filename)
        part.part_number = self.get_property(
            "Design Tracking Properties", "Part Number"
        )
        part.part_name = self.get_property("Design Tracking Properties", "Description")
        part.mass = self.get_property("Design Tracking Properties", "Mass")
        part.mass = self.get_property("Design Tracking Properties", "Mass")

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
        # Get property set.
        try:
            property_set = self.assembly_doc.PropertySets.Item(property_set)
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
    parts_list = manager.get_parts_list()
