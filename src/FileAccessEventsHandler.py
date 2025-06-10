"""
File Access Events Handler is a handler for when an Inventor file requires resolution.

Created on Monday 9th June 2025.
@author: Harry New

"""

import os
import logging.config
from tkinter import messagebox

# - - - - - - - - - - - - - - - - - - - - -

global logger
logger = logging.getLogger()

# - - - - - - - - - - - - - - - - - - - - -


class FileAccessEventsHandler:
    def OnFileResolution(self, file_name, *args):
        logger.info(f"[Intercepted] Trying to resolve: {file_name}")
        if not os.path.exists(file_name):
            logger.error("Unable to resolve file link.")
            messagebox.showerror(
                "Resolve Link Error",
                f"Could not open file due to unknown part reference: {file_name} \n\n Close all error messages and then open Inventor to resolve.",
            )
        return file_name
