from vb2py.vbfunctions import *
from vb2py.vbdebug import *



def my_write_file3(data1, data2, data3, FileName):
    _ret = None
    n = UBound(data1())
    VBFiles.openFile(1, FileName, 'w')
    for I in vbForRange(1, n):
        if GetInputState:
            DoEvents()
        VBFiles.writeText(1, Format(data1(I), '00.00'), ',', "\t", Format(data2(I), '0000.00'), ',', "\t", Format(data3(I), '0000.00'), '\n')
    VBFiles.closeFile(1)
    return _ret

def my_write_file2(data1, data2, FileName):
    _ret = None
    n = UBound(data1())
    VBFiles.openFile(1, FileName, 'w')
    for I in vbForRange(1, n):
        VBFiles.writeText(1, Format(data1(I), '00.00'), ',', "\t", Format(data2(I), '0000.00'), '\n')
    VBFiles.closeFile(1)
    return _ret

def my_read_file(data1, data2, data3, FileName):
    _ret = None
    VBFiles.openFile(1, FileName, 'r')
    while EOF(1) == False:
        #If GetInputState Then DoEvents
        n = n + 1
        data1 = vbObjectInitialize((n,), Variant, data1)
        data2 = vbObjectInitialize((n,), Variant, data2)
        data3 = vbObjectInitialize((n,), Variant, data3)
        data1[n], data2[n], data3[n] = VBFiles.getInput(1, 3)
    VBFiles.closeFile(1)
    return _ret

# VB2PY (UntranslatedCode) Attribute VB_Name = "Module2"
