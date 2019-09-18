from vb2py.vbfunctions import *
from vb2py.vbdebug import *

"""Option Explicit
*模块********************************************************
FFT0 数组下标以0开始                                                   经检验好用
AR() 数据实部         AI() 数据虚部
N 数据点数，为2的整数次幂
NI 变换方向 1为正变换，-1为反变换
***************************************************************
"""

fftIn = 256
Pi = 3.14159265358979

def FFT0(AR, AI, n, NI):
    _ret = None
    I = Long()

    J = Long()

    k = Long()

    L = Long()

    M = Long()

    IP = Long()

    LE = Long()

    L1 = Long()

    N1 = Long()

    N2 = Long()

    SN = Double()

    Tr = Double()

    Ti = Double()

    Wr = Double()

    Wi = Double()

    UR = Double()

    UI = Double()

    US = Double()
    M = NTOM(n)
    N2 = n / 2
    N1 = n - 1
    SN = NI
    J = 1
    for I in vbForRange(1, N1):
        if I < J:
            Tr = AR(J - 1)
            AR[J - 1] = AR(I - 1)
            AR[I - 1] = Tr
            Ti = AI(J - 1)
            AI[J - 1] = AI(I - 1)
            AI[I - 1] = Ti
        k = N2
        while ( k < J ):
            J = J - k
            k = k / 2
        J = J + k
    for L in vbForRange(1, M):
        LE = 2 ** L
        L1 = LE / 2
        UR = 1
        UI = 0
        Wr = Cos(Pi / L1)
        Wi = SN * Sin(Pi / L1)
        for J in vbForRange(1, L1):
            for I in vbForRange(J, n, LE):
                IP = I + L1
                Tr = AR(IP - 1) * UR + AI(IP - 1) * UI
                Ti = AI(IP - 1) * UR - AR(IP - 1) * UI
                AR[IP - 1] = AR(I - 1) - Tr
                AI[IP - 1] = AI(I - 1) - Ti
                AR[I - 1] = AR(I - 1) + Tr
                AI[I - 1] = AI(I - 1) + Ti
            US = UR
            UR = US * Wr - UI * Wi
            UI = UI * Wr + US * Wi
    if SN == -1:
        for I in vbForRange(1, n):
            AR[I - 1] = AR(I - 1) / n
            AI[I - 1] = AI(I - 1) / n
    return _ret

def FFT1(AR, AI, n, NI):
    _ret = None
    I = Long()

    J = Long()

    k = Long()

    L = Long()

    M = Long()

    IP = Integer()

    LE = Integer()

    L1 = Integer()

    N1 = Integer()

    N2 = Integer()

    SN = Double()

    Tr = Double()

    Ti = Double()

    Wr = Double()

    Wi = Double()

    UR = Double()

    UI = Double()

    US = Double()
    M = NTOM(n)
    N2 = n / 2
    N1 = n - 1
    SN = NI
    J = 1
    for I in vbForRange(1, N1):
        if I < J:
            Tr = AR(J)
            AR[J] = AR(I)
            AR[I] = Tr
            Ti = AI(J)
            AI[J] = AI(I)
            AI[I] = Ti
        k = N2
        while ( k < J ):
            J = J - k
            k = k / 2
        J = J + k
    for L in vbForRange(1, M):
        LE = 2 ** L
        L1 = LE / 2
        UR = 1
        UI = 0
        Wr = Cos(Pi / L1)
        Wi = SN * Sin(Pi / L1)
        for J in vbForRange(1, L1):
            for I in vbForRange(J, n, LE):
                IP = I + L1
                Tr = AR(IP) * UR + AI(IP) * UI
                Ti = AI(IP) * UR - AR(IP) * UI
                AR[IP] = AR(I) - Tr
                AI[IP] = AI(I) - Ti
                AR[I] = AR(I) + Tr
                AI[I] = AI(I) + Ti
            US = UR
            UR = US * Wr - UI * Wi
            UI = UI * Wr + US * Wi
    if SN == -1:
        for I in vbForRange(1, n):
            AR[I] = AR(I) / n
            AI[I] = AI(I) / n
    return _ret

def NTOM(n):
    _ret = None
    ND = Single()
    ND = n
    _ret = 0
    while ( ND > 1 ):
        ND = ND / 2
        _ret = NTOM() + 1
    return _ret

# VB2PY (UntranslatedCode) Attribute VB_Name = "Module6"
