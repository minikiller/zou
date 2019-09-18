from vb2py.vbfunctions import *
from vb2py.vbdebug import *

"""============= DelayNum为延时的毫秒数"""

SecondsInDay = 86400

def DelayTime(DelayNum):
    Ctr1 = Variant()

    Ctr2 = Variant()

    Freq = Currency()

    Count = Double()
    if QueryPerformanceFrequency(Freq):
        QueryPerformanceCounter(Ctr1)
        while 1:
            QueryPerformanceCounter(Ctr2)
            if GetInputState:
                DoEvents()
            if not (( Ctr2 - Ctr1 )  / Freq * 1000 < DelayNum):
                break
    else:
        MsgBox('不支持高精度计数器!')

# VB2PY (UntranslatedCode) Attribute VB_Name = "Module12"
# VB2PY (UntranslatedCode) Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
# VB2PY (UntranslatedCode) Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
# VB2PY (UntranslatedCode) Public Declare Function GetInputState Lib "user32" () As Long
