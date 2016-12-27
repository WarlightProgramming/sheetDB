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

class DataError(Exception):
    """
    raise whenever an invariant is violated
    """
    pass