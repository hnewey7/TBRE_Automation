"""
Part is a class representing a PartDocument in Inventor

Created on Thursday 5th Jue 2025.
@author: Harry New

"""

import logging.config

# - - - - - - - - - - - - - - - - - - - - -

global logger
logger = logging.getLogger()

# - - - - - - - - - - - - - - - - - - - - -


class Part:
    """
    A class representing each part.

    Attributes
    ----------
    filename: str
        Filename for part.
    part_number: str
        Part number.
    part_name: str
        Part name.
    mass: double
        Mass.
    x_axis: double
        Length along x-axis
    x_axis_mass: double
        Centre of mass along x-axis.
    y_axis: double
        Length along y-axis
    y_axis_mass: double
        Centre of mass along y-axis.
    z_axis: double
        Length along z-axis
    z_axis_mass: double
        Centre of mass along z-axis.
    """

    def __init__(self, assembly_doc):
        """
        Initialise.

        Args:
            assembly_doc: Assembly document for part.
        """
        prop_set = assembly_doc.ComponentDefinition.Document.PropertySets.Item(
            "Design Tracking Properties"
        )
        comp_def = assembly_doc.ComponentDefinition.Document.ComponentDefinition

        # Initial values.
        self.filename = assembly_doc.FullFileName
        self.part_number = prop_set.Item("Part Number").Value
        self.part_name = prop_set.Item("Description").Value

        try:
            self.mass = comp_def.MassProperties.Mass
            logger.info("Successfully got mass of part.")
        except Exception as e:
            self.mass = None
            logger.error(f"Unable to get mass of part, {e}")

        try:
            self.x_axis = comp_def.MassProperties.CenterOfMass.X * 10
            self.y_axis = comp_def.MassProperties.CenterOfMass.Y * 10
            self.z_axis = comp_def.MassProperties.CenterOfMass.Z * 10
            logger.info("Successfully got centre of mass of part.")

            self.x_axis_mass = self.x_axis * self.mass
            self.y_axis_mass = self.y_axis * self.mass
            self.z_axis_mass = self.z_axis * self.mass

        except Exception as e:
            self.x_axis = None
            self.y_axis = None
            self.z_axis = None
            self.x_axis_mass = None
            self.y_axis_mass = None
            self.z_axis_mass = None
            logger.error(f"Unable to get centre of mass of part, {e}")
