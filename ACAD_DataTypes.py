import win32com.client
import pythoncom

def APoint(x, y = 0, z = 0):
     return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))
