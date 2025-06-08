import pytest

from src.InventorManager import InventorManager

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
