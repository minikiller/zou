from vb2py.vbfunctions import *
from vb2py.vbdebug import *

""" 函数 My_Statistic(sign_in() As Double, mean As Double, Var As Double, Std As Double, c() As Double) 对输入序列sign_in() As Double进行统计运算
 输入参数：Sign_in——小波细节系数序列；
          mean——系数序列均值；
          var——系数序列方差；
          Std——系数序列标准差（均方差）。
 输出参数：c——输出的重构序列；
''''''''''''''''''''''''''''''
 函数 My_Feature_Vector(sign_in() As Double, mean As Double, Var As Double, Std As Double, c() As Double) 对输入序列sign_in() As Double进行统计运算
 输入参数：Sign_in——信号序列；
          mean——系数序列均值；
          var——系数序列方差；
          Std——系数序列标准差（均方差）
          xiedu——系数序列的升降趋势
          skew——系数序列基于均值的左右偏
          curt——系数序列振动程度。
 输出参数：c——输出的重构序列；
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 函数 My_swap((a As Variant, b As Variant) 对输入参数a,b互换位置运算
 输入参数：a b——为输如；

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 函数 Public Function My_Median(a() As Double, Median As Double) 对输入参数a(),先进行排序，后求中位数
 输入参数：a ——为输如数组；

 函数 My_qsort(a() As Double, b() As Integer) 对输入参数a(),b()进行排序，注：以a()为排序主体，b()随a()动且a,b等长
 输入参数：a b——为输如数组；

 函数 My_qsort(a() As Double, b() As Integer) 对输入参数a(),b()进行排序，注：以a()为排序主体，b()随a()动且a,b等长
 输入参数：a b——为输如数组；

*模块********************************************************
My_Wrev 数组下标以1开始
data_in() 数据向量         Wrev_out() 倒转后的输出向量
***************************************************************
*模块********************************************************
My_QMF 数组下标以1开始
data_in() 数据向量         Wmirr_out() 镜像后的输出向量
***************************************************************
*模块********************************************************
My_Cmult_Conj 复数数组乘法
data_ai,data_ar,data_bi,data_br 输入数据向量         data_oi(),data_or()复数数组乘法后的输出向量
***************************************************************
*模块********************************************************
My_Cmult 复数数组乘法
data_ai,data_ar,data_bi,data_br 输入数据向量         data_oi(),data_or()复数数组乘法后的输出向量
***************************************************************
*模块********************************************************
My_Cdive 复数数组除法
data_ai,data_ar,data_bi,data_br 输入数据向量         data_oi(),data_or()复数数组除法后的输出向量
***************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  函数名：My_Trapzd
  功能：  用梯形求积法求积
  参数：  data_y_in   - Double型变量，波强数据数组
          data_x_in   - Double型变量，波长数据数组
          a     - Long型变量，数组下限
          b     - Long型变量，数组上限，要求 b>a
          result   - Double型变量，返回的结果
  返回值：Double型，积分值
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 函数 My_ave(a() As Double) 对输入参数a()进行平均
 输入参数：a——为输如数组；

 函数 My_ave(a() As Double) 对输入参数a()取中值;偶数取中间二数平均,奇数取中值
 输入参数：a——为输如数组；

 Sort an array of longs.
 Sort an array of longs.
 Sort an array of singles.
 Sort an array of doubles.
 Sort an array of strings.
 Sort an array of variants.
"""


def My_Statistic(sign_in, Mean, var, std):
    _ret = None
    n = Long()
    n = UBound(sign_in())
    if n <= 1:
        #  Debug.Print " 'N must be at least 2' "
        return _ret
    #''''
    s = 0
    for i in vbForRange(1, n):
        #s = s + Abs(sign_in(I))
        s = s + sign_in(i)
    if s == 0:
        return _ret
    Mean = s / n
    var = 0
    std = 0
    for i in vbForRange(1, n):
        s = sign_in(i) - Mean
        p = s * s
        var = var + p
    var = var /  ( n - 1 )
    std = Sqr(var)
    #'''''
    return _ret

def My_Feature_Vector(x, Mean, avedev, std0, std1, xiedu, skew, curt):
    _ret = None
    n = Long()
    n = UBound(x())
    if n <= 1:
        #  Debug.Print " 'N must be at least 2' "
        return _ret
    #''''
    for i in vbForRange(1, 50):
        #dif = dif + (x(i + 1) - x(i)) ^ 3
        dif = dif + x(n - i) - x(i + 1)
    max = x(1)
    min = x(1)
    for J in vbForRange(1, n):
        if max < x(J):
            max = x(J)
        if min > x(J):
            min = ( J )
    #如果极差为0说明数列中所有的数据都相等
    if max - min == 0:
        Mean = x(1)
        range = 0
        avedev = 0
        var = 0
        std1 = 0
        xiedu = 0
        skew = 0
        curt = 0
        return _ret
    range = max - min
    #''''平均值
    Mean = 0
    for J in vbForRange(1, n):
        Mean = Mean + x(J)
    if Mean == 0:
        return _ret
    Mean = Mean / n
    #''''''''
    adev = 0
    sdev = 0
    var = 0
    xiedu = 0
    skew = 0
    curt = 0
    #''''''''
    xiedu = dif /  ( n / 4 )
    for J in vbForRange(1, n):
        s = s + Abs(x(J) - Mean)
        ss = ss +  ( x(J) - Mean )  ** 2
        sss = sss +  ( x(J) - Mean )  ** 3
        ssss = ssss +  ( x(J) - Mean )  ** 4
    avedev = s / n
    var1 = ss /  ( n - 1 )
    var0 = ss / n
    std0 = Sqr(var0)
    std1 = Sqr(var1)
    RRR = ss / Mean
    if var0 <> 0:
        #xiedu = dif / (n - 1) / ((ss / n) ^ 0.5) ^ 2
        skew = sss / n /  ( ( ss / n )  ** 0.5 )  ** 3
        curt = ssss / n /  ( std0 ** 4 )
    else:
        # Debug.Print "no skew or kurtosis when zero variance"
        pass
    return _ret

def My_swap(a, B):
    _ret = None
    t = a
    a = B
    B = t
    return _ret

def My_Median(a, median):
    _ret = None
    Temp = vbObjectInitialize(objtype=Double)

    n = Long()

    i = Long()
    n = UBound(a())
    Temp = vbObjectInitialize((n,), Variant)
    Call(CopyMemory(Temp(1), a(1), n * 8))
    #tttt = Timer()
    #Call My_qsort1(a, Temp)
    Call(SortDoubleArray(Temp))
    #Debug.Print Timer() - tttt
    if n % 2 == 1:
        median = Abs(Temp(( n + 1 )  / 2))
    else:
        median = Abs(Temp(n / 2) + Temp(n / 2 + 1)) / 2
    return _ret

def My_qsort(a, B):
    _ret = None
    M = Long()

    n = Long()

    i = Long()

    J = Long()
    M = UBound(a())
    n = LBound(a())
    for i in vbForRange(1, M):
        for J in vbForRange(i + 1, M):
            if a(i) < a(J):
                Call(My_swap(a(i), a(J)))
                Call(My_swap(B(i), B(J)))
    return _ret

def My_qsort1(a, B):
    _ret = None
    M = Long()

    n = Long()

    i = Long()

    J = Long()
    M = UBound(a())
    Call(CopyMemory(B(1), a(1), M * 8))
    for i in vbForRange(1, M):
        for J in vbForRange(i + 1, M):
            if Abs(B(i)) < Abs(B(J)):
                Call(My_swap(B(i), B(J)))
                #Call My_swap(b(i), b(j))
                #B(I) = a(I)
                #B(J) = a(J)
    return _ret

def My_Wrev(data_in, Wrev_out):
    _ret = None
    length_data_in = UBound(data_in())
    Wrev_out = vbObjectInitialize((length_data_in,), Double)
    for i in vbForRange(1, length_data_in):
        Wrev_out[length_data_in - i + 1] = data_in(i)
    return _ret

def My_QMF(data_in, QMF_out):
    _ret = None
    length_data_in = UBound(data_in())
    QMF_out = vbObjectInitialize((length_data_in,), Double)
    for i in vbForRange(1, length_data_in):
        if length_data_in % 2 == 0:
            if i % 2 == 1:
                sign_out = - 1
                QMF_out[length_data_in - i + 1] = sign_out * data_in(i)
            else:
                sign_out = 1
                QMF_out[length_data_in - i + 1] = sign_out * data_in(i)
        else:
            if i % 2 == 1:
                sign_out = 1
                QMF_out[length_data_in - i + 1] = sign_out * data_in(i)
            else:
                sign_out = - 1
                QMF_out[length_data_in - i + 1] = sign_out * data_in(i)
    return _ret

def My_Cmult_Conj(data_ar, data_ai, data_br, data_bi, data_or, data_oi):
    _ret = None
    length_data_in = UBound(data_ai())
    #ReDim data_oi(1 To length_data_in) As Double
    #ReDim data_or(1 To length_data_in) As Double
    for i in vbForRange(1, length_data_in):
        data_or[i] = data_ar(i) * data_br(i) - data_ai(i) *  ( data_bi(i) )
        data_oi[i] = data_ar(i) *  ( data_bi(i) )  + data_ai(i) * data_br(i)
    return _ret

def My_Cmult(data_ar, data_ai, data_br, data_bi, data_or, data_oi):
    _ret = None
    length_data_in = UBound(data_ai())
    #ReDim data_oi(1 To length_data_in) As Double
    #ReDim data_or(1 To length_data_in) As Double
    for i in vbForRange(1, length_data_in):
        data_or[i] = data_ar(i) * data_br(i) - data_ai(i) * data_bi(i)
        data_oi[i] = data_ar(i) * data_bi(i) + data_ai(i) * data_br(i)
    return _ret

def My_Cdive(data_ar, data_ai, data_br, data_bi, data_or, data_oi):
    _ret = None
    e = Double()

    F = Double()
    length_data_in = UBound(data_ai())
    #ReDim data_oi(1 To length_data_in) As Double
    #ReDim data_or(1 To length_data_in) As Double
    #For I = 1 To length_data_in
    # data_or(I) = (data_ar(I) * data_br(I) + data_ai(I) * data_bi(I)) / (data_bi(I) ^ 2 + data_br(I) ^ 2)
    # data_oi(I) = (data_ai(I) * data_br(I) - data_bi(I) * data_ar(I)) / (data_bi(I) ^ 2 + data_br(I) ^ 2)
    #Next I
    for i in vbForRange(1, length_data_in):
        if Abs(data_br(i)) == 0 and Abs(data_bi(i)) == 0:
            MsgBox(( '除数为零，退出计算！！！' ))
            return _ret
        elif Abs(data_br(i)) >= Abs(data_bi(i)):
            e = data_bi(i) / data_br(i)
            F = data_br(i) + e * data_bi(i)
            data_or[i] = ( data_ar(i) + data_ai(i) * e )  / F
            data_oi[i] = ( data_ai(i) - data_ar(i) * e )  / F
        else:
            e = data_br(i) / data_bi(i)
            F = data_bi(i) + e * data_br(i)
            data_or[i] = ( data_ai(i) + data_ar(i) * e )  / F
            data_oi[i] = ( data_ar(i) + data_ai(i) * e )  / F
    return _ret

def My_Trapzd(data_x_in, data_y_in, a, B, result):
    _ret = None
    n = Integer()

    K = Integer()

    fa = Double()

    fb = Double()

    H = Double()

    t1 = Double()

    p = Double()

    s = Double()

    x = Double()

    t = Double()
    if a > B:
        return _ret
    elif a == B:
        result = 0
    else:
        n = B - a + 1
    # 积分区间端点的函数值
    fa = data_x_in(a)
    fb = data_x_in(B)
    H = ( fb - fa )  /  ( n - 1 )
    Temp = 0
    for i in vbForRange(1, n - 1):
        Temp = Temp + H / 2 *  ( data_y_in(a + i - 1) + data_y_in(a + i) )
        # Debug.Print i; temp
    # 返回满足精度的积分值
    result = Temp
    return _ret

def My_ave(a):
    _ret = None
    B = vbObjectInitialize(objtype=Double)

    M = Integer()
    M = UBound(a())
    B = vbObjectInitialize((M,), Double)
    for i in vbForRange(1, M):
        B[i] = a(i)
        Temp = Temp + B(i)
    _ret = Temp / M
    return _ret

def My_mida(a):
    _ret = None
    C = vbObjectInitialize(objtype=Double)

    M = Integer()
    #Dim b() As Double
    M = UBound(a())
    B = vbObjectInitialize((M,), Double)
    C = vbObjectInitialize((M,), Double)
    for i in vbForRange(1, M):
        B[i] = a(i)
    for i in vbForRange(1, M):
        for J in vbForRange(i + 1, M):
            if B(i) > B(J):
                Call(My_swap(B(i), B(J)))
    if M % 2 == 0:
        _ret = ( B(M / 2) + B(M / 2 + 1) )  / 2
    else:
        _ret = ( B(M // 2 - 1) + B(M // 2) + B(M // 2 + 1) + B(M // 2 + 2) + B(M // 2 + 3) )  / 5
    return _ret

def QuicksortInt(list, min, max):
    med_value = Long()

    hi = Long()

    lo = Long()

    i = Long()
    # If the list has no more than CutOff elements,
    # finish it off with SelectionSort.
    if max <= min:
        return
    # Pick the dividing value.
    i = Int(( max - min + 1 )  * Rnd() + min)
    med_value = list(i)
    # Swap it to the front.
    list[i] = list(min)
    lo = min
    hi = max
    while 1:
        # Look down from hi for a value < med_value.
        while list(hi) >= med_value:
            hi = hi - 1
            if hi <= lo:
                break
        if hi <= lo:
            list[lo] = med_value
            break
        # Swap the lo and hi values.
        list[lo] = list(hi)
        # Look up from lo for a value >= med_value.
        lo = lo + 1
        while list(lo) < med_value:
            lo = lo + 1
            if lo >= hi:
                break
        if lo >= hi:
            lo = hi
            list[hi] = med_value
            break
        # Swap the lo and hi values.
        list[hi] = list(lo)
    # Sort the two sublists.
    QuicksortInt(list(), min, lo - 1)
    QuicksortInt(list(), lo + 1, max)

def QuicksortSingle(list, min, max):
    med_value = Single()

    hi = Long()

    lo = Long()

    i = Long()
    # If the list has no more than CutOff elements,
    # finish it off with SelectionSort.
    if max <= min:
        return
    # Pick the dividing value.
    i = Int(( max - min + 1 )  * Rnd() + min)
    med_value = list(i)
    # Swap it to the front.
    list[i] = list(min)
    lo = min
    hi = max
    while 1:
        # Look down from hi for a value < med_value.
        while list(hi) >= med_value:
            hi = hi - 1
            if hi <= lo:
                break
        if hi <= lo:
            list[lo] = med_value
            break
        # Swap the lo and hi values.
        list[lo] = list(hi)
        # Look up from lo for a value >= med_value.
        lo = lo + 1
        while list(lo) < med_value:
            lo = lo + 1
            if lo >= hi:
                break
        if lo >= hi:
            lo = hi
            list[hi] = med_value
            break
        # Swap the lo and hi values.
        list[hi] = list(lo)
    # Sort the two sublists.
    QuicksortSingle(list(), min, lo - 1)
    QuicksortSingle(list(), lo + 1, max)

def QuicksortDouble(list, min, max):
    med_value = Double()

    hi = Long()

    lo = Long()

    i = Long()
    # If the list has no more than CutOff elements,
    # finish it off with SelectionSort.
    if max <= min:
        return
    # Pick the dividing value.
    i = Int(( max - min + 1 )  * Rnd() + min)
    med_value = Abs(list(i))
    # Swap it to the front.
    list[i] = Abs(list(min))
    lo = min
    hi = max
    while 1:
        DoEvents()
        # Look down from hi for a value < med_value.
        while Abs(list(hi)) >= med_value:
            DoEvents()
            hi = hi - 1
            if hi <= lo:
                break
        if hi <= lo:
            list[lo] = med_value
            break
        # Swap the lo and hi values.
        list[lo] = Abs(list(hi))
        # Look up from lo for a value >= med_value.
        lo = lo + 1
        while Abs(list(lo)) < med_value:
            DoEvents()
            lo = lo + 1
            if lo >= hi:
                break
        if lo >= hi:
            lo = hi
            list[hi] = med_value
            break
        # Swap the lo and hi values.
        list[hi] = Abs(list(lo))
    # Sort the two sublists.
    QuicksortDouble(list(), min, lo - 1)
    QuicksortDouble(list(), lo + 1, max)

def QuicksortString(list, min, max):
    med_value = String()

    hi = Long()

    lo = Long()

    i = Long()
    # If the list has no more than CutOff elements,
    # finish it off with SelectionSort.
    if max <= min:
        return
    # Pick the dividing value.
    i = Int(( max - min + 1 )  * Rnd() + min)
    med_value = list(i)
    # Swap it to the front.
    list[i] = list(min)
    lo = min
    hi = max
    while 1:
        # Look down from hi for a value < med_value.
        while list(hi) >= med_value:
            hi = hi - 1
            if hi <= lo:
                break
        if hi <= lo:
            list[lo] = med_value
            break
        # Swap the lo and hi values.
        list[lo] = list(hi)
        # Look up from lo for a value >= med_value.
        lo = lo + 1
        while list(lo) < med_value:
            lo = lo + 1
            if lo >= hi:
                break
        if lo >= hi:
            lo = hi
            list[hi] = med_value
            break
        # Swap the lo and hi values.
        list[hi] = list(lo)
    # Sort the two sublists.
    QuicksortString(list(), min, lo - 1)
    QuicksortString(list(), lo + 1, max)

def QuicksortVariant(list, min, max):
    med_value = Variant()

    hi = Long()

    lo = Long()

    i = Long()
    # If the list has no more than CutOff elements,
    # finish it off with SelectionSort.
    if max <= min:
        return
    # Pick the dividing value.
    i = Int(( max - min + 1 )  * Rnd() + min)
    med_value = list(i)
    # Swap it to the front.
    list[i] = list(min)
    lo = min
    hi = max
    while 1:
        # Look down from hi for a value < med_value.
        while list(hi) >= med_value:
            hi = hi - 1
            if hi <= lo:
                break
        if hi <= lo:
            list[lo] = med_value
            break
        # Swap the lo and hi values.
        list[lo] = list(hi)
        # Look up from lo for a value >= med_value.
        lo = lo + 1
        while list(lo) < med_value:
            lo = lo + 1
            if lo >= hi:
                break
        if lo >= hi:
            lo = hi
            list[hi] = med_value
            break
        # Swap the lo and hi values.
        list[hi] = list(lo)
    # Sort the two sublists.
    QuicksortVariant(list(), min, lo - 1)
    QuicksortVariant(list(), lo + 1, max)

def QuicksortLong(list, min, max):
    med_value = Long()

    hi = Long()

    lo = Long()

    i = Long()
    # If the list has no more than CutOff elements,
    # finish it off with SelectionSort.
    if max <= min:
        return
    # Pick the dividing value.
    i = Int(( max - min + 1 )  * Rnd() + min)
    med_value = list(i)
    # Swap it to the front.
    list[i] = list(min)
    lo = min
    hi = max
    while 1:
        # Look down from hi for a value < med_value.
        while list(hi) >= med_value:
            hi = hi - 1
            if hi <= lo:
                break
        if hi <= lo:
            list[lo] = med_value
            break
        # Swap the lo and hi values.
        list[lo] = list(hi)
        # Look up from lo for a value >= med_value.
        lo = lo + 1
        while list(lo) < med_value:
            lo = lo + 1
            if lo >= hi:
                break
        if lo >= hi:
            lo = hi
            list[hi] = med_value
            break
        # Swap the lo and hi values.
        list[hi] = list(lo)
    # Sort the two sublists.
    QuicksortLong(list(), min, lo - 1)
    QuicksortLong(list(), lo + 1, max)

def SortIntArray(list):
    QuicksortInt(list, LBound(list), UBound(list))

def SortLongArray(list):
    QuicksortLong(list, LBound(list), UBound(list))

def SortSingleArray(list):
    QuicksortSingle(list, LBound(list), UBound(list))

def SortDoubleArray(list):
    QuicksortDouble(list, LBound(list), UBound(list))

def SortStringArray(list):
    QuicksortString(list, LBound(list), UBound(list))

def SortVariantArray(list):
    QuicksortVariant(list, LBound(list), UBound(list))

# VB2PY (UntranslatedCode) Attribute VB_Name = "Module9"
