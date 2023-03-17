import time


def wait_for_attribute(obj, attr_name, *args):
    """
    Waits for a CDispatch attribute to be retrieved.
    """
    while True:
        try:
            # Try to get the attribute value
            attr_value = getattr(obj, attr_name)
            return attr_value
        except:
            # If an exception occurs, wait for a short time and try again
            time.sleep(0.1)

def wait_for_method_return(obj, method_name, *args):
    """
    Waits for a CDispatch method to return a value.
    """
    while True:
        try:
            # Try to call the method and get the return value
            method = getattr(obj, method_name)
            return_value = method(*args)
            return return_value
        except:
            # If an exception occurs, wait for a short time and try again
            time.sleep(0.1)
