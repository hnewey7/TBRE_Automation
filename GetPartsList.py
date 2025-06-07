"""
Script to get the parts list of an assembly document.

Created on Saturday 7th June 2025.
@author: Harry New

"""

import logging.config
import json
import os
from datetime import datetime

from src.InventorManager import InventorManager

# - - - - - - - - - - - - - - - - - - - - -

global logger
logger = logging.getLogger()

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

    # Start InventorManager.
    manager = InventorManager()

    # Get parts list.
    parts_list = manager.get_parts_list()

    # Export parts list.
    manager.export_parts_list(parts_list)
