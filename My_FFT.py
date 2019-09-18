from vb2py.vbfunctions import *
from vb2py.vbdebug import *

"""Option Explicit
*模块********************************************************
FFT0 数组下标以0开始;输出复数
PR() 数据实部         PI() 数据虚部
N 数据点数，为2的整数次幂
K 长度的幂数
SIGN 变换方向 0为正变换，1为逆变换
IL  目前未知功能！！！！均取大于1的整数即可
***************************************************************
此函数找出N的最小的二次幂的幂值
此函数找出N的二次幂的幂值
"""


def My_KFFt(PR, Pi, n, k, FR, FI, SIGN, il):
    _ret = None
    p = Variant()

    q = Variant()

    s = Variant()

    VR = Variant()

    VI = Variant()

    PODDR = Variant()

    PODDI = Double()

    m = Variant()

    NV = Variant()

    L0 = Variant()

    I = Variant()

    J = Variant()

    it = Variant()

    ISS = Integer()
    for it in vbForRange(0, n - 1):
        m = it
        ISS = 0
        for I in vbForRange(0, k - 1):
            J = m // 2
            ISS = 2 * ISS +  ( m - 2 * J )
            m = J
        FR[it + 1] = PR(ISS + 1)
        FI[it + 1] = Pi(ISS + 1)
    PR[1] = 1
    Pi[1] = 0
    PR[2] = Cos(6.283185306 / n)
    Pi[2] = - Sin(6.283185306 / n)
    if  ( SIGN > 0 ) :
        Pi[2] = - Pi(2)
    for I in vbForRange(3, n):
        p = PR(I - 1) * PR(2)
        q = Pi(I - 1) * Pi(2)
        s = ( PR(I - 1) + Pi(I - 1) )  *  ( PR(2) + Pi(2) )
        PR[I] = p - q
        Pi[I] = s - p - q
    for it in vbForRange(0, n - 2, 2):
        VR = FR(it + 1)
        VI = FI(it + 1)
        FR[it + 1] = VR + FR(it + 2)
        FI[it + 1] = VI + FI(it + 2)
        FR[it + 2] = VR - FR(it + 2)
        FI[it + 2] = VI - FI(it + 2)
    m = n / 2
    NV = 2
    for L0 in vbForRange(k - 2, 0, -1):
        m = m / 2
        NV = 2 * NV
        for it in vbForRange(0, ( m - 1 )  * NV, NV):
            for J in vbForRange(0, NV / 2 - 1):
                p = PR(m * J + 1) * FR(it + J + 1 + NV / 2)
                q = Pi(m * J + 1) * FI(it + J + 1 + NV / 2)
                s = PR(m * J + 1) + Pi(m * J + 1)
                s = s *  ( FR(it + J + 1 + NV / 2) + FI(it + J + 1 + NV / 2) )
                PODDR = p - q
                PODDI = s - p - q
                FR[it + J + 1 + NV / 2] = FR(it + J + 1) - PODDR
                FI[it + J + 1 + NV / 2] = FI(it + J + 1) - PODDI
                FR[it + J + 1] = FR(it + J + 1) + PODDR
                FI[it + J + 1] = FI(it + J + 1) + PODDI
    if  ( SIGN > 0 ) :
        for I in vbForRange(1, n):
            FR[I] = FR(I) / n
            FI[I] = FI(I) / n
    if  ( il > 0 ) :
        for I in vbForRange(1, n):
            PR[I] = Sqr(FR(I) * FR(I) + FI(I) * FI(I))
            if FR(I) == 0:
                Pi[I] = 90
            else:
                Pi[I] = Atn(FI(I) / FR(I)) * 360 / 6.283185306
    return _ret

def NextPow2(n):
    _ret = None
    for I in vbForRange(1, n):
        if 2 ** I < n and 2 **  ( I + 1 )  > n:
            _ret = I + 1
            return _ret
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

def My_Array_VarToDbl(myarray):
    _ret = None
    # 把想要修改的数组传入 myarray(),
    # 因为是引用调用, 所以对 myarray(), 的修改会保留下来.
    # 而对 byval s 的修改不会保留下来.
    # 你应该了解一下 byref 引用调用 和 byval 传值调用 的区别.
    #学过C/C++语言,这个应该比较好理解.
    pass
    return _ret

def ds():
    _ret = None
    _ret = Array(1, 2, 3, 4, 5)
    return _ret

# VB2PY (UntranslatedCode) Attribute VB_Name = "Module4"
