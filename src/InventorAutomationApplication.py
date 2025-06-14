"""
InventorAutomationApplication is a class for handling automation of Inventor.

Created on Monday 9th June 2025.
@author: Harry New

"""

from tkinter import Tk, Text, END, messagebox, Frame, Toplevel
from tkinter.filedialog import askopenfilename
import win32com.client
import logging.config
import time
from datetime import datetime
import json
import os

from .MainWindow import MainWindow
from .Part import Part
from .ProgressBarWindow import ProgressBarWindow

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
        self.root.resizable(False, False)

        # Get command mapping.
        commands = {
            "select_file": self.select_file,
            "export_parts_list": self.export_parts_list,
        }

        # Add main window.
        self.main_window = MainWindow(self.root, commands)
        self.main_window.pack()

    def display_progress_bar(self, name: str):
        """
        Display progress bar.

        Args:
            name (str): Progress bar name/
        """
        # Create new window.
        self.subwindow = Toplevel(self.root, takefocus=True)
        self.subwindow.title(name)
        self.subwindow.geometry("300x150")
        self.subwindow.resizable(False, False)

        # Create progress bar.
        self.progress_bar = ProgressBarWindow(self.subwindow, name)
        self.progress_bar.pack()

        # Update subwindow.
        self.subwindow.update()

    # - - - - - - - - - - - - - - - -
    # Methods for using Inventor.

    def select_file(
        self,
        filename: str = None,
        text_var: Text = None,
        suppress_message: bool = False,
    ) -> bool:
        """
        Method for selecting file in Inventor.

        Args:
            filename (str): Filename, optional
            text_var (Text): Text variable to update, optional

        Returns:
            bool: Successful or not.
        """
        # Create dialog for selecting file.
        if not filename:
            filename = askopenfilename(title="Select a file")

        # Info for resolve file dialog.
        if not suppress_message:
            messagebox.showinfo(
                "Resolve File Dialog",
                "Check Inventor for resolve file dialog and resolve if neccessary.",
            )

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
                text_var.configure(state="normal")
                text_var.delete("1.0", END)
                text_var.insert(END, filename.split("/")[-1])
                text_var.configure(state="disabled")
            return True, filename
        except Exception as e:
            logger.error(f"Error selecting document '{filename}': {e}")
            return False, None

    def export_parts_list(self, file_text: Text, checkbox_frame: Frame) -> bool:
        """
        Export parts list function.

        Args:
            file_text (Text): Text variable for open file.
            checkbox_frame (Frame): Checkbox frame containing options.

        Return:
            bool: Successful or not.
        """
        # Check valid file open.
        if file_text.get("1.0", "1.4") == "None":
            messagebox.showerror(
                "Invalid Assembly File",
                "Please select a valid assembly file before exporting the parts list.",
            )
            return False

        # Display progress bar.
        self.display_progress_bar("Exporting parts list...")

        # Get parts list of current assembly file.
        occurrences = self.assembly_doc.ComponentDefinition.Occurrences
        all_parts = []
        self.get_part_occurrences(occurrences, all_parts, self.progress_bar)

        # Close progress bar.
        self.subwindow.destroy()

        # Store parts list for later use.
        self.recent_parts_list = all_parts

        # Generate HTML content.
        self.recent_html_preview = self.create_html_parts_list(self.recent_parts_list)

        # Display HTML content.
        self.main_window.right_side_frame.update_html_preview(self.recent_html_preview)

    def get_part_occurrences(
        self, occurrences, parts_list: list, progress_bar: ProgressBarWindow = None
    ):
        """
        Get part occurrences.

        Args:
            occurrences: Occurrences
            parts_list (list): List containing all parts.
        """
        # Set maximum for progress bar.
        if progress_bar:
            max_length = len(occurrences)
            self.progress_bar.set_length(max_length)

        for occ in occurrences:
            # Add to progress bar.
            if progress_bar:
                progress_bar.add_to_progress_bar()

            # Handle occurrence.
            doc_type = occ.DefinitionDocumentType
            if doc_type == 12290:
                # Update current task.
                if progress_bar:
                    progress_bar.update_task(
                        f"Getting part details ({occ.Definition.Document.DisplayName})..."
                    )
                # Get part.
                part = Part(occ)
                parts_list.append(part)
            elif doc_type == 12291:
                # Update current task.
                if progress_bar:
                    progress_bar.update_task(
                        f"Getting occurrences ({occ.Definition.Document.DisplayName})..."
                    )
                # Get occurrencess.
                self.get_part_occurrences(occ.SubOccurrences, parts_list)

    def create_html_parts_list(self, parts_list: list):
        """
        Create HTML parts list.

        Args:
            parts_list (list): Parts list.

        Return:
            str: HTML content.
        """
        # Create table.
        table_content = "<table><tr><th>P/N</th><th>Part Name</th><th>Mass (kg)</th><th>X-axis (mm)</th><th>X-axis mass</th><th>Y-axis (mm)</th><th>Y-axis mass</th><th>Z-axis (mm)</th><th>Z-axis mass</th></tr>"
        for part in parts_list:
            table_content += f"<tr><td>{part.part_number}</td><td>{part.part_name}</td><td>{part.mass}</td><td>{part.x_axis}</td><td>{part.x_axis_mass}</td><td>{part.y_axis}</td><td>{part.y_axis_mass}</td><td>{part.z_axis}</td><td>{part.z_axis_mass}</td></tr>"
        table_content += "</table>"

        # HTML content.
        html_template = f"""<html>
            <head>
            <title>PARTS LIST</title>
            </head>
            <body>
            {table_content}
            </body>
            </html>
        """
        return html_template


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
