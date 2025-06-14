"""
Progress Bar Window.

Created on Friday 13th June 2025.
@author: Harry New

"""

from tkinter import Frame, Label, DoubleVar, StringVar
from tkinter import ttk
import tkinter.font as tkFont
import logging.config

# - - - - - - - - - - - - - - - - - - - - -

global logger
logger = logging.getLogger()

# - - - - - - - - - - - - - - - - - - - - -


class ProgressBarWindow(Frame):
    def __init__(self, window, name):
        # Create frame.
        Frame.__init__(self, window)

        self.progress = 0

        # Create fonts.
        normal_font = tkFont.Font(family="Ubuntu", size=10)
        small_font = tkFont.Font(family="Ubuntu", size=8)

        # Label.
        title_label = Label(self, text=name, font=normal_font)
        title_label.grid(row=0, column=0, pady=(10, 0))

        # Progress bar.
        self.progress_bar = ttk.Progressbar(self, length=180)
        self.progress_bar.grid(row=1, column=0, pady=10)

        # Create task label.
        self.task_var = StringVar(self)
        task_label = Label(
            self, textvariable=self.task_var, font=small_font, wraplength=250
        )
        task_label.grid(row=2, column=0)

        # Create percentage Label.
        self.percentage_var = DoubleVar(self)
        self.percentage_string_var = StringVar(self)
        percentage_label = Label(
            self, textvariable=self.percentage_string_var, font=normal_font
        )
        percentage_label.grid(row=3, column=0)

    def reset_progress(self):
        """
        Reset progress bar.
        """
        self.progress = 0

    def set_length(self, max_length: int):
        """
        Set max length of progress bar.

        Args:
            max_length (int): Max length of progress bar.
        """
        self.maximum = max_length
        self.update()

    def add_to_progress_bar(self):
        """
        Add to progress bar.
        """
        self.progress += 1
        self.percentage_var.set(round(self.progress * 100 / self.maximum, 2))
        self.percentage_string_var.set(f"{self.percentage_var.get()}%")
        self.progress_bar.step(100 / self.maximum)
        self.update()

    def update_task(self, task: str):
        """
        Update task.

        Args:
            task (str): Current task.
        """
        self.task_var.set(task)
        self.update()
