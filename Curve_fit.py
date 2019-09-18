from vb2py.vbfunctions import *
from vb2py.vbdebug import *


x = vbObjectInitialize(objtype=Double)
y = vbObjectInitialize(objtype=Double)
a = vbObjectInitialize((20, 20,), Double)
m = Integer()
b = vbObjectInitialize(objtype=Double)
n = Integer()
I = Integer()
J = Integer()
xishu = vbObjectInitialize(objtype=Double)
Xmin = Double()
Xmax = Double()
Ymin = Double()
Ymax = Double()
Xo = Double()
Yo = Double()

def My_curvefit(x, y, TimeOfFit, xishu):
    global a, b, m, n
    Xh = Integer()
    m = TimeOfFit + 1
    n = UBound(x())
    Erase(b)
    Erase(xishu)
    Erase(a)
    b = vbObjectInitialize((m,), Variant)
    xishu = vbObjectInitialize(((1, m),), Variant)
    #形成方程组的各元素
    a[1, 1] = n
    for I in vbForRange(1, n):
        b[1] = b(1) + y(I)
    for J in vbForRange(2, m):
        for I in vbForRange(1, n):
            a[1, J] = a(1, J) + x(I) **  ( J - 1 )
    for I in vbForRange(2, m):
        for J in vbForRange(1, m):
            for Xh in vbForRange(1, n):
                a[I, J] = a(I, J) + x(Xh) **  ( I + J - 2 )
                if J == 1:
                    b[I] = b(I) + x(Xh) **  ( I - 1 )  * y(Xh)
    Call(My_fit(a, b, xishu))
    #For I = 1 To M
    #  Debug.Print xishu(I)
    #Next I
    #Dim Str As String: Str = "y="
    #For I = 1 To M    '写方程
    # If I < M Then
    #    Str = Str & xishu(I) & "*I^" & I - 1 & "+"
    #Else
    #   Str = Str & xishu(I) & "*I^" & I - 1
    #End If
    #Next I
    #Debug.Print Str

def My_fit(a, b, x):
    global n
    _ret = None
    TempA = Double()

    L = Integer()

    k = Integer()

    kk = Integer()

    Ii = Integer()

    ChuShu = Double()

    Sum = Double()
    n = UBound(b)
    for I in vbForRange(1, n):
        L = 0
        kk = 0
        for J in vbForRange(I, n):
            # If GetInputState Then DoEvents
            if a(J, I) == 0:
                L = L + 1
        for J in vbForRange(I, n - L):
            if a(J, I) == 0:
                kk = kk + 1
                for k in vbForRange(I, n):
                    # If GetInputState Then DoEvents
                    TempA = a(J, k)
                    a[J, k] = a(n - kk + 1, k)
                    a[n - kk + 1, k] = TempA
                TempA = b(J)
                b[J] = b(n - kk + 1)
                b[n - kk + 1] = TempA
        for Ii in vbForRange(I, n - L):
            ChuShu = a(Ii, I)
            for J in vbForRange(I, n):
                # If GetInputState Then DoEvents
                a[Ii, J] = a(Ii, J) / ChuShu
            b[Ii] = b(Ii) / ChuShu
        for Ii in vbForRange(I + 1, n - L):
            for J in vbForRange(I, n):
                #If GetInputState Then DoEvents
                a[Ii, J] = a(Ii, J) - a(I, J)
            b[Ii] = b(Ii) - b(I)
    for I in vbForRange(1, n):
        for J in vbForRange(1, I - 1):
            #If GetInputState Then DoEvents
            a[I, J] = 0
    x[n] = b(n) / a(n, n)
    for I in vbForRange(n - 1, 1, -1):
        Sum = 0
        for J in vbForRange(I + 1, n):
            # If GetInputState Then DoEvents
            Sum = Sum + a(I, J) * x(J)
        x[I] = ( b(I) - Sum )  / a(I, I)
    return _ret

# VB2PY (UntranslatedCode) Attribute VB_Name = "Module8"
# VB2PY (UntranslatedCode) Option Explicit
