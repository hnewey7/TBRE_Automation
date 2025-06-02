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