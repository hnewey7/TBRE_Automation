import pytest

from src.InventorManager import InventorManager
from src.Part import Part

from example_files.INVENTOR.example_test_data import example_assembly_data

# - - - - - - - - - - - - - - - - -


def test_connection():
    """
    Test connection to Inventor application.
    """
    manager = InventorManager()
    assert manager.app is not None


def test_select_document():
    """
    Test selecting document method.
    """
    manager = InventorManager()
    assert manager.app is not None

    ret = manager.select_document(
        r"C:\Users\hnewe\Documents\tbre_automation_tools\example_files\INVENTOR\DOG TOOTH GEARBOX.iam"
    )
    assert ret


property_test_data = [
    (
        r"C:\Users\hnewe\Documents\tbre_automation_tools\example_files\INVENTOR\DOG TOOTH GEARBOX.iam",
        "Design Tracking Properties",
        "Part Number",
        "DOG TOOTH GEARBOX",
    ),
    (
        r"C:\Users\hnewe\Documents\tbre_automation_tools\example_files\INVENTOR\DOG TOOTH GEARBOX.iam",
        "Design Tracking Properties",
        "Description",
        "",
    ),
]


@pytest.mark.parametrize(
    "document,property_set,property_item,value", property_test_data
)
def test_get_property(document, property_set, property_item, value):
    """
    Test getting properties from document.
    """
    manager = InventorManager()
    assert manager.app is not None

    # Selecting document.
    ret = manager.select_document(document)
    assert ret

    # Getting property.
    property_value = manager.get_property(property_set, property_item)
    assert property_value == value


assembly_test_data = [
    (
        r"C:\Users\hnewe\Documents\tbre_automation_tools\example_files\INVENTOR\DOG TOOTH GEARBOX.iam",
        example_assembly_data,
    )
]


@pytest.mark.parametrize("document,assembly_data", assembly_test_data)
def test_get_parts_list(document, assembly_data):
    """
    Test get parts list for a given assembly document.

    Args:
        document (str): File path to assembly document.
        assembly_data (dict): Dictionary containing all part files in assembly.
    """
    manager = InventorManager()
    assert manager.app is not None

    # Selecting document.
    ret = manager.select_document(document)
    assert ret

    # Get parts list.
    parts_list = manager.get_parts_list(document)

    # Check each part.
    for i, part_name in enumerate(assembly_data[document]):
        # Get example part separately.
        manager.select_document(part_name)
        example_part = Part(manager.assembly_doc)

        # Check each property.
        assert example_part.filename == parts_list[i].filename
        assert example_part.part_number == parts_list[i].part_number
        assert example_part.part_name == parts_list[i].part_name
        assert example_part.mass == parts_list[i].mass
        assert example_part.x_axis == parts_list[i].x_axis
        assert example_part.x_axis_mass == parts_list[i].x_axis_mass
        assert example_part.y_axis == parts_list[i].y_axis
        assert example_part.y_axis_mass == parts_list[i].y_axis_mass
        assert example_part.z_axis == parts_list[i].z_axis
        assert example_part.z_axis_mass == parts_list[i].z_axis_mass
