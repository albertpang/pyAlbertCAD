import time
import win32com.client
import pythoncom

def APoint(x, y, z = 0):
     return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


def aDouble(*seq):
    """ Returns :class:`array.array` of doubles ('d' code) for passing to AutoCAD
    For 3D points use :class:`APoint` instead.
    """
    return _sequence_to_comtypes('d', *seq)


def aInt(*seq):
    """ Returns :class:`array.array` of ints ('l' code) for passing to AutoCAD
    """
    return _sequence_to_comtypes('l', *seq)


def aShort(*seq):
    """ Returns :class:`array.array` of shorts ('h' code) for passing to AutoCAD
    """
    return _sequence_to_comtypes('h', *seq)


def _sequence_to_comtypes(typecode='d', *sequence):
    if len(sequence) == 1:
        return array.array(typecode, sequence[0])
    return array.array(typecode, sequence)