# errors.py
## custom error classes

class NoAccessError(Exception):
    """
    raised whenever account doesn't have access
    """
    pass

class SheetError(Exception):
	"""
	raised whenever an illegal sheet operation is attempted
	"""
	pass