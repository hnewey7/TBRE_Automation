"""
InventorAutomationApplication is a class for handling automation of Inventor.

Created on Monday 9th June 2025.
@author: Harry New

"""

from tkinter import Tk, Text, END, messagebox, Toplevel
from tkinter.filedialog import askopenfilename, asksaveasfile
import win32com.client
import logging.config
import time
from datetime import datetime
import json
import os
import pandas as pd

from .MainWindow import MainWindow, CheckButtonFrame
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
        self.recent_html_preview = None  # HTML preview.
        self.assembly_doc = None  # Assembly doc.

        # Connect to Inventor application.
        ret = self.connect_to_inventor()
        if not ret:
            logger.error("Failed to connect to Inventor application.")
        else:
            logger.info("Successfully connected to Inventor application.")

        # Create tkinter user interface.
        self.create_user_interface()

        # Set flag for updating file name.
        self.update_file_name_flag = True

    def run(self):
        """
        Running Inventor Automation Application.
        """
        # Start update file loop.
        self.update_file_name()
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
            "save_parts_list": self.save_parts_list,
        }

        # Get options.
        self.options_config = {}
        with open("config/option_config.json") as f:
            self.options_config = json.load(f)

        # Add main window.
        self.main_window = MainWindow(self.root, commands, self.options_config)
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

    def get_option_variables(self, checkbutton_frame: CheckButtonFrame) -> list:
        """
        Getting option variables in CheckButtonFrame.

        Args:
            checkbutton_frame (CheckButtonFrame): Checkbutton frame.

        Returns:
            list: Selected options.
        """
        selected_variables = []
        for variable in checkbutton_frame.option_variables.keys():
            if checkbutton_frame.option_variables[variable].get() == 1:
                selected_variables.append(variable)
        return selected_variables

    def update_file_name(self):
        """
        Updating file name.
        """
        # Flag for stopping.
        if not self.update_file_name_flag:
            return

        # Check for active document.
        if not self.check_active_document():
            document_name = "None"
        else:
            # Get active document name.
            document_name = self.app.ActiveDocument.FullFileName

            # Select document.
            if (
                self.assembly_doc is None
                or document_name != self.assembly_doc.FullFileName
            ):
                self.select_file(document_name, suppress_message=True)

        # Update file name widget.
        text_var = self.main_window.left_side_frame.file_label
        text_var.configure(state="normal")
        text_var.delete("1.0", END)
        text_var.insert(END, document_name.split("\\")[-1])
        text_var.configure(state="disabled")

        # Repeat every second.
        self.root.after(1000, self.update_file_name)

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
        # Disable update file name.
        self.update_file_name_flag = False

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

            # Re-enable update file name flag.
            self.update_file_name_flag = True

            return True, filename
        except Exception as e:
            logger.error(f"Error selecting document '{filename}': {e}")
            # Re-enable update file name flag.
            self.update_file_name_flag = True
            self.update_file_name()

            return False, None

    def export_parts_list(
        self, file_text: Text, checkbox_frame: CheckButtonFrame
    ) -> bool:
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

        # Get selected options.
        selected_options = self.get_option_variables(checkbox_frame)

        # Check valid options.
        if len(selected_options) == 0:
            messagebox.showerror(
                "Invalid Parts List Options",
                "Please select options to export the parts list.",
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

        # Create pandas df.
        self.dataframe = self.create_dataframe(self.recent_parts_list, selected_options)

        # Copy to clipboard.
        self.dataframe.to_clipboard(index=False, excel=True)

        # Generate HTML content.
        self.recent_html_preview = self.create_html_parts_list(
            self.recent_parts_list, selected_options
        )

        # Display HTML content.
        self.main_window.right_side_frame.update_html_preview(self.recent_html_preview)

        # Information box about clipboard.
        messagebox.showinfo(
            "Copied To Clipboard",
            "Parts list is copied to clipboard so can be directly pasted into excel.",
        )

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
            # Check valid occurence.
            if not self.check_valid_occurrence_definition(occ):
                continue

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

    def create_html_parts_list(self, parts_list: list, selected_options: list):
        """
        Create HTML parts list.

        Args:
            parts_list (list): Parts list.
            selected_options (list): Options to export.

        Return:
            str: HTML content.
        """
        # Get full info about options.
        full_option_info = []
        for selected in selected_options:
            for i, option in enumerate(self.options_config["options"]):
                if option["option_name"] == selected:
                    full_option_info.append(self.options_config["options"][i])

        # Create table heading.
        table_content = "<table><tr>"
        for option in full_option_info:
            for display_name in option["display_name"]:
                table_content += f"<th>{display_name}</th>"
        table_content += "</tr>"

        # Add table content.
        for part in parts_list:
            table_content += "<tr>"
            for option in full_option_info:
                for attribute in option["attribute_name"]:
                    value = eval(f"part.{attribute}")
                    table_content += f"<td>{value}</td>"
            table_content += "</tr>"
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

    def save_parts_list(self):
        """
        Save parts list to computer.
        """
        # Check if parts list to save.
        if self.recent_html_preview is None:
            messagebox.showerror(
                "Invalid Parts List",
                "Please export a parts list to save.",
            )
            return False

        # Get save location.
        filetype = [("HTML file", "*.html")]
        f = asksaveasfile("w", filetypes=filetype, defaultextension=filetype)
        if f is None:
            return
        f.write(self.recent_html_preview)
        f.close()

        # Display save location.
        messagebox.showinfo(
            "Saved Parts List",
            f"Parts list saved at {os.path.abspath(f.name)}",
        )

    def create_dataframe(
        self, parts_list: list, selected_options: list
    ) -> pd.DataFrame:
        """
        Create pandas dataframe from parts list.

        Args:
            parts_list (list): List of parts.
            selected_options (list): List of selected options.
        """
        # Get full info about options.
        full_option_info = []
        for selected in selected_options:
            for i, option in enumerate(self.options_config["options"]):
                if option["option_name"] == selected:
                    full_option_info.append(self.options_config["options"][i])

        # Get columns.
        columns = []
        for option in full_option_info:
            for display_name in option["display_name"]:
                columns.append(display_name)

        # Add each part to dataframe.
        rows = []
        for part in parts_list:
            new_row = {}
            for option in full_option_info:
                for i, display_name in enumerate(option["display_name"]):
                    new_row[display_name] = eval(f'part.{option["attribute_name"][i]}')
            rows.append(new_row)

        # Create dataframe.
        dataframe = pd.DataFrame(rows, columns=columns)
        return dataframe

    def check_active_document(self) -> bool:
        """
        Check if document is active in Inventor.

        Returns:
            bool: Active document or not.
        """
        if self.app.ActiveDocument is None:
            return False
        else:
            return True

    def check_valid_occurrence_definition(self, occurence) -> bool:
        """
        Check valid definition for occurrence.

        Args:
            occurence: Occurrence.

        Returns:
            bool: Valid occurrence.
        """
        try:
            type(occurence.Definition)
            return True
        except Exception as e:
            logger.error(f"Invalid occurrence definition: {e}")
            return False


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
