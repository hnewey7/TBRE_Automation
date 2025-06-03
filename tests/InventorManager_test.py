from src.InventorManager import InventorManager

# - - - - - - - - - - - - - - - - -


def test_connection():
    """
    Test connection to Inventor application.
    """
    manager = InventorManager()
    assert manager.app is not None


def test_disconnect():
    """
    Test disconnecting from Inventor application.
    """
    manager = InventorManager()
    manager.disconnect()
    assert manager.app is None


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
