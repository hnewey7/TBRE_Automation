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

    def __init__(self, occurrence=None, assembly_doc=None):
        """
        Initialise.

        Args:
            occurrence: Part occurrence, optional.
            assembly_doc: Assembly document, optional.
        """
        if occurrence:
            prop_set = occurrence.Definition.Document.PropertySets.Item(
                "Design Tracking Properties"
            )
        elif assembly_doc:
            prop_set = assembly_doc.ComponentDefinition.Document.PropertySets.Item(
                "Design Tracking Properties"
            )
        else:
            logger.error("No occurrence or assembly document provided.")
            raise ValueError

        # Initial values.
        self.part_number = prop_set.Item("Part Number").Value
        self.part_name = prop_set.Item("Description").Value

        # Set reference document.
        ref_doc = occurrence if occurrence else assembly_doc.ComponentDefinition

        try:
            self.mass = ref_doc.MassProperties.Mass
            logger.info("Successfully got mass of part.")
        except Exception as e:
            self.mass = None
            logger.error(f"Unable to get mass of part, {e}")

        try:
            self.x_axis = ref_doc.MassProperties.CenterOfMass.X * 10
            self.y_axis = ref_doc.MassProperties.CenterOfMass.Y * 10
            self.z_axis = ref_doc.MassProperties.CenterOfMass.Z * 10
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
