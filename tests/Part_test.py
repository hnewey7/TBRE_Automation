import pytest

from src.InventorManager import InventorManager
from src.Part import Part

from example_files.INVENTOR.example_test_data import example_part_data

# - - - - - - - - - - - - - - - - -

test_part_data = [(key, example_part_data[key]) for key in example_part_data.keys()]


@pytest.mark.parametrize("document,data", test_part_data)
def test_part_init(document, data):
    """
    Test correct initialisation of part object.

    Args:
        document (str): Document file name.
        data (str): Correct data for file.
    """
    # Start InventorManager.
    inv_manager = InventorManager()
    assert inv_manager.app is not None

    # Select valid part.
    ret = inv_manager.select_document(document)
    assert ret

    # Get part.
    part = Part(assembly_doc=inv_manager.assembly_doc)

    # Check part values.
    assert part.part_number == data["part_number"]
    assert part.part_name == data["part_name"]
    assert round(part.mass, 3) == round(data["mass"], 3)

    if (
        data["x_axis"] is not None
        or data["y_axis"] is not None
        or data["z_axis"] is not None
    ):
        assert round(part.x_axis, 3) == round(data["x_axis"], 3)
        assert round(part.y_axis, 3) == round(data["y_axis"], 3)
        assert round(part.z_axis, 3) == round(data["z_axis"], 3)
        assert round(part.x_axis_mass, 1) == round(data["mass"] * data["x_axis"], 1)
        assert round(part.y_axis_mass, 1) == round(data["mass"] * data["y_axis"], 1)
        assert round(part.z_axis_mass, 1) == round(data["mass"] * data["z_axis"], 1)
    else:
        assert part.x_axis is None
        assert part.y_axis is None
        assert part.z_axis is None
        assert part.x_axis_mass is None
        assert part.y_axis_mass is None
        assert part.z_axis_mass is None
