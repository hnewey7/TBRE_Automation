import win32com.client
from win32com.client import gencache
import logging.config
import os
from datetime import datetime
import json

# - - - - - - - - - - - - - - - - - - - - -

global logger
logger = logging.getLogger()

# - - - - - - - - - - - - - - - - - - - - -


class InventorManager:
    def __init__(self):
        # Connect to Inventor application.
        ret = self.connect_to_inventor()
        if not ret:
            logger.error("Failed to connect to Inventor application.")
        else:
            logger.info("Successfully connected to Inventor application.")

    def connect_to_inventor(self):
        """
        Connect to an existing Inventor application instance.

        Returns:
            bool: True if connection is successful, False otherwise.
        """
        try:
            self.app = win32com.client.Dispatch("Inventor.Application")
            self.app.Visible = True
            CastTo = gencache.EnsureModule("{D98A091D-3A0F-4C3E-B36E-61F62068D488}", 0, 1, 0)
            self.app = CastTo.Application(self.app)
            return True
        except Exception as e:
            logger.error(f"Error connecting to Inventor: {e}")
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
        config_dict["handlers"]["file"]["filename"] = f"logs/{current_time}/{current_time}_RoverRPi.log"
        logging.config.dictConfig(config_dict)

    # Create an instance of InventorManager to test the connection.
    manager = InventorManager()
    logger.info(
        f"Active Document: {manager.app.ActiveDocument.Name if manager.app.ActiveDocument else 'No active document'}"
    )
