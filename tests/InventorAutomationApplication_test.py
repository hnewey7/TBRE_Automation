import pytest

from src.InventorAutomationApplication import InventorAutomationApplication

# - - - - - - - - - - - - - - - - -


def test_init():
    """
    Test initialisation of InventorAutomationApplication.
    """
    app = InventorAutomationApplication()
    assert app.app is not None


filenames = [
    r"C:\Users\hnewe\Documents\tbre_automation_tools\example_files\INVENTOR\DOG TOOTH GEARBOX.iam",
    r"C:\Users\hnewe\Documents\tbre_automation_tools\example_files\INVENTOR\27T WIDTH 15.iam",
    r"C:\Users\hnewe\Documents\tbre_automation_tools\example_files\INVENTOR\38T WIDTH 15 MOD 2.iam",
]


@pytest.mark.parametrize("filename", filenames)
def test_select_file(filename: str):
    """
    Test select file method.

    Args:
        filename (str): Filename
    """
    app = InventorAutomationApplication()
    assert app.app is not None

    ret, ret_filename = app.select_file(filename)
    assert ret
    assert ret_filename == filename
    assert app.assembly_doc is not None
