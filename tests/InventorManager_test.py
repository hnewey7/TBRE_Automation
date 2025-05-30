from src.InventorManager import InventorManager

# - - - - - - - - - - - - - - - - -


def test_connection():
    """
    Test connection to Inventor application.
    """
    manager = InventorManager()
    assert manager.app is not None
