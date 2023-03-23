import time

def wait_for_attribute(obj, attr_name, *args):
    """
    Waits for a CDispatch attribute to be retrieved.
    """
    counter = 0
    while True:
        try:
            # Try to get the attribute value
            attr_value = getattr(obj, attr_name)
            return attr_value
        except:
            # If an exception occurs, wait for a short time and try again
            time.sleep(0.1)
            counter += 1
            if counter > 5:
                break
            

def set_attribute(obj, attr_name, arg):
    """
    Waits for a CDispatch attribute to be retrieved.
    """
    counter = 0
    while True:
        try:
            # Try to get the attribute value
            setattr(obj, attr_name, arg)
            return 0
        except:
            # If an exception occurs, wait for a short time and try again
            time.sleep(0.1)
            counter += 1
            if counter > 5:
                break

def wait_for_method_return(obj, method_name, *args):
    """
    Waits for a CDispatch method to return a value.
    """
    counter = 0
    while True:
        try:
            # Try to call the method and get the return value
            method = getattr(obj, method_name)
            return_value = method(*args)
            return return_value
        except:
            # If an exception occurs, wait for a short time and try again
            time.sleep(0.3)
            counter += 1
            if counter > 5:
                break
