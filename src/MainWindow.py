"""
Main Window of Inventor Automation Application.

Created on Monday 9th June 2025.
@author: Harry New

"""

from tkinter import Frame, Label, Button, Text, WORD, END, IntVar, Checkbutton
import tkinter.font as tkFont
from tkhtmlview import HTMLLabel
import re
import logging.config

# - - - - - - - - - - - - - - - - - - - - -

global logger
logger = logging.getLogger()

# - - - - - - - - - - - - - - - - - - - - -


class MainWindow(Frame):
    def __init__(self, window, commands: dict, options_config: dict):
        """
        Main window of UI.

        Args:
            window: Top level window.
            commands (dict): Dictionary of commands.
            options_config (dict): Dictionary of options.
        """
        # Create frame.
        Frame.__init__(self, window)

        # Create side frames.
        self.left_side_frame = LeftSideFrame(self, commands, options_config)
        self.left_side_frame.grid(row=0, column=0, padx=20, sticky="n", pady=20)
        self.right_side_frame = RightSideFrame(self)
        self.right_side_frame.grid(row=0, column=1, padx=20, pady=20)


# - - - - - - - - - - - - - - - - - - - - -


class LeftSideFrame(Frame):
    def __init__(self, window, commands: dict, options_config: dict):
        """
        Left side frame for holding buttons.

        Args:
            window: Parent window.
            commands (dict): Dictionary of commands.
            options_config (dict): Dictionary of options.
        """
        # Create frame.
        Frame.__init__(self, window)

        # Custom font.
        normal_font = tkFont.Font(family="Ubuntu", size=10)
        title_font = tkFont.Font(family="Ubuntu", size=20)

        # Add title label.
        title_label = Label(
            self, text="Inventor Automation Application", font=title_font
        )
        title_label.grid(row=0, column=0, pady=10, columnspan=3)

        # Add selected file label.
        selected_file_label = Label(self, text="Selected File:", font=normal_font)
        selected_file_label.grid(row=1, column=0, padx=5)
        self.file_label = Text(self, wrap=WORD, height=1, width=30)
        self.file_label.grid(row=1, column=1, padx=5)
        self.file_label.insert(END, "None")
        self.file_label.configure(state="disabled")

        # Add select file button.
        select_file_button = Button(
            self,
            text="Select File",
            command=lambda: commands["select_file"](text_var=self.file_label),
        )
        select_file_button.grid(row=1, column=2, padx=5, pady=10)

        # Create check button frame.
        options = [option["option_name"] for option in options_config["options"]]
        checkbutton_frame = CheckButtonFrame(self, options, normal_font, 3)

        # Add check button frame.
        checkbutton_frame.grid(row=2, column=0, columnspan=2, pady=10)

        # Add export parts list.
        export_parts_list_button = Button(
            self,
            text="Export Parts List",
            command=lambda: commands["export_parts_list"](
                file_text=self.file_label, checkbox_frame=checkbutton_frame
            ),
        )
        export_parts_list_button.grid(row=2, column=2, pady=10)

        # Save parts list.
        save_part_list_button = Button(
            self, text="Save Parts List", command=commands["save_parts_list"], width=40
        )
        save_part_list_button.grid(row=3, column=0, columnspan=3, pady=10)


# - - - - - - - - - - - - - - - - - - - - -


class RightSideFrame(Frame):
    def __init__(self, window):
        """
        Right side frame for HTML preview.

        Args:
            window: Parent window.
        """
        # Create frame.
        Frame.__init__(self, window)

        # Custom font.
        normal_font = tkFont.Font(family="Ubuntu", size=10)

        # Create title label.
        title_label = Label(self, text="HTML Preview:", font=normal_font)
        title_label.grid(row=0, column=0, pady=10)

        # Create HTML preview.
        self.html_preview = HTMLLabel(
            self, html="", background="#FFFFFF", width=80, height=28
        )
        self.html_preview.grid(row=1, column=0, padx=20)

    def round_numbers_in_html(self, html, decimals=3):
        """
        Round numbers in html file.

        Args:
            html: HTML file.
            decimals (int, optional): Number of decimals to round. Defaults to 3.
        """

        def round_match(match):
            num_str = match.group()
            try:
                num = float(num_str)
                return f"{round(num, decimals):.{decimals}f}"
            except ValueError:
                return num_str  # Leave unmodified if somehow not parseable

        # Regex to match numbers (including decimals, optional sign, and scientific notation)
        number_regex = r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?"

        return re.sub(number_regex, round_match, html)

    def update_html_preview(self, html_content: str):
        """
        Update HTML preview.

        Args:
            html_content (str): HTML cotent to update.
        """
        # Round numbers for better formatting.
        rounded_count = self.round_numbers_in_html(html_content, 2)

        # Additional formatting.
        formatted_content = f'<div style="font-size: 8px; zoom: 0.5; transform: scale(0.8); transform-origin: top left;">{rounded_count}</div>'

        # Update content.
        self.html_preview.set_html(formatted_content)


# - - - - - - - - - - - - - - - - - - - - -


class CheckButtonFrame(Frame):
    def __init__(self, window, options, font, column_span):
        """
        Check button frame.

        Args:
            window: Window to place check button frame.
            options: Options for checkbuttons.
            font: Font for check buttons.
            colucolumn_spanmns: Number of columns.
        """
        # Create frame.
        Frame.__init__(self, window)

        # Store font.
        self.font = font

        # Add title label.
        title_label = Label(self, text="Properties to export:", font=font)
        title_label.grid(row=0, column=0, columnspan=column_span)

        # Create options.
        self.option_variables = {}
        self.option_buttons = {}

        # Counters.
        self.max_column = column_span
        self.row_counter = 1
        self.column_counter = 0

        for option in options:
            self.create_option(option)

    def create_option(self, name: str):
        """
        Create option for checkbutton list.

        Args:
            name (str): Name of option
            variable: Variable for storing value.
        """
        # Create variable.
        var = IntVar()
        self.option_variables[name] = var

        # Create check button.
        checkbutton = Checkbutton(
            self, text=name, variable=var, onvalue=1, offvalue=0, font=self.font
        )
        self.option_buttons[name] = checkbutton

        # Add to grid.
        checkbutton.grid(row=self.row_counter, column=self.column_counter)

        # Update counters.
        self.column_counter += 1
        if self.column_counter >= self.max_column:
            self.column_counter = 0
            self.row_counter += 1
