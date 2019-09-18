from vb2py.vbfunctions import *
from vb2py.vbdebug import *

"""''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  功能：  求解线性方程组
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LEGauss
  功能：  使用全选主元高斯消去法求解线性方程组
  参数    n     - Integer型变量，线性方程组的阶数
         dblA   - Double型 n x n 二维数组，线性方程组的系数矩阵
         dblB   - Double型长度为 n 的一维数组，线性方程组的常数向量，返回方程组的解向量
  返回值：Boolean型，求解成功为True，无解或求解失败为False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LEGaussJordan
  功能：  使用全选主元高斯－约当消去法求解线性方程组
  参数    n     - Integer型变量，线性方程组的阶数
          m    - Integer型变量，线性方程组的个数，即右端常数矩阵列向量的个数
         dblA   - Double型 n x n 二维数组，线性方程组的系数矩阵
         dblB   - Double型n x m二维数组，线性方程组的常数矩阵，返回方程组的解矩阵
  返回值：Boolean型，求解成功为True，无解或求解失败为False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LECpxGauss
  功能：  使用全选主元高斯消去法求解复系数线性代数方程组
  参数    n     - Integer型变量，线性代数方程组的阶数
         dblAR  - Double型 n x n 二维数组，线性代数方程组的系数矩阵的实部
         dblAI   - Double型 n x n 二维数组，线性代数方程组的系数矩阵的虚部
         dblBR  - Double型长度为 n 的一维数组，线性代数方程组的常数向量的实部，返回方程组的解向量的实部
         dblBI   - Double型长度为 n 的一维数组，线性代数方程组的常数向量的虚部，返回方程组的解向量的虚部
  返回值：Boolean型，求解成功为True，无解或求解失败为False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LECpxGaussJordan
  功能：  使用全选主元高斯－约当消去法求解复系数线性代数方程组
  参数    n    - Integer型变量，线性代数方程组的阶数
          m    - Integer型变量，方程组右端复常数向量的个数
         dblAR - Double型 n x n 二维数组，线性代数方程组的系数矩阵的实部
         dblAI - Double型 n x n 二维数组，线性代数方程组的系数矩阵的虚部
         dblBR - Double型长度为 n X m 的二维数组，存放方程组右端的m组常数向量的实部，
                 返回时存放m组解向量的实部
         dblBI - Double型长度为 n x m 的二维数组，存放方程组右端的m组常数向量的虚部，
                 返回时存放m组解向量的虚部
  返回值：Boolean型，求解成功为True，无解或求解失败为False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LETrid
  功能：  使用追赶法求解三对角线线性代数方程组
  参数    n   - Integer型变量，线性代数方程组的阶数
          m   - Integer型变量，n阶三对角线矩阵三对角线上元素的个数，即数组b的长度。
                 它的值应为m = 3n -2。函数应对此值进行检验。
         dblB － Double型一维数组，长度为m。以行为主存放三对角线矩阵中三对角线上的元素，即b中依次存放下列元素：
                  a11，a12，a21，a22，a23，a32，a33，a34，…，an,n-1，an,n
         dblD － Double型一维数组，长度为n。作为传入参数，存放方程组右端的常数向量。
                  函数返回时，此数组中存放着方程组的解向量。
  返回值： Integer型。小于0，m的值不正确；为0，求解失败，无解；大于0，成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LEBand
  功能：  一般带型方程组的求解
  参数    n   - Integer型变量，线性代数方程组的阶数
          m  - Integer型变量，为方程组右端的常数向量的个数
          l   - Integer型变量，为系数矩阵的半带宽。
          il   - Integer型变量，为系数矩阵的带宽。
         dblB - Double型n x il二维数组，存放带型矩阵A中带区内的元素
         dblD - Double型n x m二维数组，作为传入参数，存放方程组右端的m组常数向量。
                函数返回时，其中存放着m组解向量。
  返回值： Integer型。小于0，参数中半带宽l与带宽il的关系不对；为0，系数矩阵A奇异，无解；大于0成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LEDjn
  功能：  用分解法求解对称方程组
  参数    n   - Integer型变量，线性代数方程组的阶数
          m  - Integer型变量，为方程组右端的常数向量的个数
         dblA - Double型n x n二维数组，存放系数矩阵
         dblC - Double型n x m二维数组，作为传入参数，存放方程组右端的m组常数向量。
                函数返回时，其中存放着m组解向量。
  返回值： Boolean型。False，失败无解；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LECholesky
  功能：  用乔里斯基分解法求解正定方程组
  参数    n   - Integer型变量，线性代数方程组的阶数
          m  - Integer型变量，为方程组右端的常数向量的个数
         dblA - Double型n x n二维数组，存放系数矩阵（应为对称正定矩阵）；返回时，其上三角部分存放分解后的U矩阵
         dblD - Double型n x m二维数组，作为传入参数，存放方程组右端的m组常数向量。
                函数返回时，其中存放着m组解向量。
  返回值： Boolean型。False，失败无解；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LEGgje
  功能：  用全选主元高斯－约当消去法求解稀疏方程组
  参数    n    - Integer型变量，线性代数方程组的阶数
         dblA  - Double型n x n二维数组，存放系数矩阵（应为稀疏矩阵）
         dblB  - Double一维数组，长度为n，存放方程组右端的常数向量；返回时存放方程组的解
  返回值： Boolean型。False，失败无解；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LETlvs
  功能：  用列文逊递推算法求解对称托伯利兹方程组
  参数    n    - Integer型变量，线性代数方程组的阶数
         dblT  - Double型一维数组，长度为n ，存放对称T型矩阵中的元素
         dblB  - Double一维数组，长度为n，存放方程组右端的常数向量
         dblX  - Double一维数组，长度为n，返回存放方程组的解
  返回值： Boolean型。False，失败无解；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LESeidel
  功能：  用高斯-赛德尔迭代求解系数矩阵主对角线占绝对优势线性方程组
  参数    n    - Integer型变量，线性代数方程组的阶数
         dblA  - Double型n x n二维数组，存放系数矩阵
         dblB  - Double一维数组，长度为n，存放方程组右端的常数向量
         dblX  - Double型一维数组，长度为n。返回方程组的解。
         eps  - Double型变量。给定的精度要求。
  返回值： Boolean型。False，失败无解；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LEGrad
  功能：  用共轭梯度法是求解n阶对称正定方程组
  参数    n    - Integer型变量，线性代数方程组的阶数
         dblA  - Double型n x n二维数组，存放对称正定系数矩阵
         dblB  - Double一维数组，长度为n，存放方程组右端的常数向量
         dblX  - Double型一维数组，长度为n。返回方程组的解。
         eps  - Double型变量。给定的精度要求。
  返回值： Boolean型。False，失败无解；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：MatrixModule.bas
  函数名：MatrixToString
  功能：  将矩阵转换为显示字符串
  参数：  m   - Integer型变量，矩阵的行数
          n   - Integer型变量，矩阵的列数
          mtxA  - Double型m x n二维数组，存放相加的左边矩阵
          sFormat - 显示矩阵各元素的格式控制字符串
  返回值：String型，显示矩阵的字符串
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：MatrixModule.bas
  函数名：MMul
  功能：  计算矩阵的乘法
  参数：  m   - Integer型变量，相乘的左边矩阵的行数
          n   - Integer型变量，相乘的左边矩阵的列数和右边矩阵的行数
          l   -  Integer型变量，相乘的右边矩阵的列数
          mtxA  - Double型m x n二维数组，存放相乘的左边矩阵
          mtxB  - Double型n x l二维数组，存放相乘的右边矩阵
          mtxC  - Double型m x l二维数组，返回矩阵乘积矩阵
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LEMqr
  功能：  用豪斯荷尔德变换法求解线性最小二乘问题方程组
  参数：   m    - Integer型变量。系数矩阵的行数， m>=n
           n    - Integer型变量。系数矩阵的列数，n<=m
          dblA  - Double型二维数组，体积维n x n。存放系数矩阵；返回时，存放分解式中的R矩阵.
          dblB  - Double型一维数组，长度为m。存放方程组右端常数向量；返回时，前n个元素存放方程组的最小二乘解。
          dblQ  - Double型二维数组，体积为m x m。返回时，存放分解式中的Q矩阵
  返回值： Boolean型。False，失败无解；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：MatrixModule.bas
  函数名：MMqr
  功能：  用豪斯荷尔德变换法对矩阵进行QR分解
  参数：   m    - Integer型变量。矩阵的行数， m>=n
           n    - Integer型变量。矩阵的列数，n<=m
          dblA  - Double型二维数组，体积维n x n。存放待分解矩阵；返回时，存放分解式中的R矩阵.
          dblQ  - Double型二维数组，体积为m x m。返回时，存放分解式中的Q矩阵
  返回值： Boolean型。False，失败；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LEMiv
  功能：  用广义逆法求解线性最小二乘问题方程组
  参数：   m    - Integer型变量。系数矩阵的行数， m>=n
           n    - Integer型变量。系数矩阵的列数，n<=m
          dblA  - Double型二维数组，体积维n x n。存放超定方程组系数矩阵；
                  返回时，其对角线存放矩阵的奇异值，其余元素为0。
          dblB  - Double型一维数组，长度为m。存放超定方程组右端常数向量
          dblX  - Double型一维数组，长度为n。返回时，存放超定方程组的最小二乘解。
          dblAP - Double型二维数组，体积维n x m。返回时，存放超定方程组系数矩阵A的广义逆A+。
          dblU  - Double型二维数组，体积维m x m。返回时，存放超定方程组系数矩阵A的奇异值分解式中的
                  左奇异向量U。
          dblV  - Double型二维数组，体积维n x n。返回时，存放超定方程组系数矩阵A的奇异值分解式中的
                  右奇异向量VT。
           ka  - Integer型变量。ka=max(m,n)+1
          eps  - Double型变量。奇异值分解函数中的控制精度参数。
  返回值： Boolean型。False，失败无解；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：MatrixModule.bas
  函数名：MUav
  功能：  用豪斯荷尔德变换及变形QR算法对矩阵进行奇异值分解
  参数：   m    - Integer型变量。系数矩阵的行数， m>=n
           n    - Integer型变量。系数矩阵的列数，n<=m
          dblA  - Double型二维数组，体积维m x n。存放待分解矩阵；
                  返回时，其对角线存放矩阵的奇异值(以非递增次序排列)，其余元素为0。
          dblU  - Double型二维数组，体积维m x m。返回时，存放奇异值分解式中的左奇异向量U。
          dblV  - Double型二维数组，体积维n x n。返回时，存放奇异值分解式中的右奇异向量VT。
           ka  - Integer型变量。ka=max(m,n)+1
          eps  - Double型变量。奇异值分解函数中的控制精度参数。
  返回值： Boolean型。False，失败无解；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：MatrixModule.bas
  函数名：Cal1
  功能：  内部过程，供MUav函数调用
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：MatrixModule.bas
  函数名：Cal2
  功能：  内部过程，供MUav函数调用
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：MatrixModule.bas
  函数名：MInv
  功能：  求矩阵的广义逆
  参数：   m    - Integer型变量。系数矩阵的行数， m>=n
           n    - Integer型变量。系数矩阵的列数，n<=m
          dblA  - Double型二维数组，体积维m x n。存放待分解矩阵；
                  返回时，其对角线存放矩阵的奇异值(以非递增次序排列)，其余元素为0。
          dblAP - Double型二维数组，体积维n x m。返回时存放矩阵的广义逆。
          dblU  - Double型二维数组，体积维m x m。返回时，存放奇异值分解式中的左奇异向量U。
          dblV  - Double型二维数组，体积维n x n。返回时，存放奇异值分解式中的右奇异向量VT。
           ka  - Integer型变量。ka=max(m,n)+1
          eps  - Double型变量。奇异值分解函数中的控制精度参数。
  返回值： Boolean型。False，失败无解；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  模块名：LEModule.bas
  函数名：LEMorbid
  功能：  求解病态方程组
  参数：   n    - Integer型变量，方程组的阶数。
          dblA  - Double型二维数组，体积维n x n。存放病态方程组系数矩阵。
          dblB  - Double型一维数组，长度为n，存放方程组右端常数向量。
          dblX  - Double型一维数组，长度为n。返回时，存放方程组的解向量。
          eps  - Double型变量。控制精度参数。
  返回值： Boolean型。False，失败无解；True, 成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
"""


def LEGauss(n, dblA, dblB):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()

    nIs = Integer()

    d = Double()

    t = Double()
    # 局部变量
    nJs = vbObjectInitialize((n,), Integer)
    # 开始求解
    for k in vbForRange(1, n - 1):
        d = 0
        # 归一
        for I in vbForRange(k, n):
            for J in vbForRange(k, n):
                if GetInputState:
                    DoEvents()
                t = Abs(dblA(I, J))
                if t > d:
                    d = t
                    nJs[k] = J
                    nIs = I
        # 无解，返回
        if d + 1 == 1:
            _ret = False
            return _ret
        # 消元
        if nJs(k) <> k:
            for I in vbForRange(1, n):
                if GetInputState:
                    DoEvents()
                t = dblA(I, k)
                dblA[I, k] = dblA(I, nJs(k))
                dblA[I, nJs(k)] = t
        if nIs <> k:
            for J in vbForRange(k, n):
                if GetInputState:
                    DoEvents()
                t = dblA(k, J)
                dblA[k, J] = dblA(nIs, J)
                dblA[nIs, J] = t
            t = dblB(k)
            dblB[k] = dblB(nIs)
            dblB[nIs] = t
        d = dblA(k, k)
        for J in vbForRange(k + 1, n):
            dblA[k, J] = dblA(k, J) / d
        dblB[k] = dblB(k) / d
        for I in vbForRange(k + 1, n):
            for J in vbForRange(k + 1, n):
                if GetInputState:
                    DoEvents()
                dblA[I, J] = dblA(I, J) - dblA(I, k) * dblA(k, J)
            dblB[I] = dblB(I) - dblA(I, k) * dblB(k)
    d = dblA(n, n)
    # 无解，返回
    if Abs(d) + 1 == 1:
        _ret = False
        return _ret
    # 回代
    dblB[n] = dblB(n) / d
    for I in vbForRange(n - 1, 1, -1):
        t = 0
        for J in vbForRange(I + 1, n):
            if GetInputState:
                DoEvents()
            t = t + dblA(I, J) * dblB(J)
        dblB[I] = dblB(I) - t
    # 调整解的次序
    nJs[n] = n
    for k in vbForRange(n, 1, -1):
        if nJs(k) <> k:
            if GetInputState:
                DoEvents()
            t = dblB(k)
            dblB[k] = dblB(nJs(k))
            dblB[nJs(k)] = t
    # 求解成功
    _ret = True
    return _ret

def LEGaussJordan(n, m, dblA, dblB):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()

    nIs = Integer()

    d = Double()

    q = Double()
    # 局部变量
    nJs = vbObjectInitialize((n,), Integer)
    # 开始求解
    for k in vbForRange(1, n):
        q = 0
        # 归一
        for I in vbForRange(k, n):
            for J in vbForRange(k, n):
                if Abs(dblA(I, J)) > q:
                    q = Abs(dblA(I, J))
                    nJs[k] = J
                    nIs = I
        # 无解，返回
        if q + 1 == 1:
            _ret = False
            return _ret
        # 消元
        # A->
        for J in vbForRange(k, n):
            d = dblA(k, J)
            dblA[k, J] = dblA(nIs, J)
            dblA[nIs, J] = d
        # B->
        for J in vbForRange(1, m):
            d = dblB(k, J)
            dblB[k, J] = dblB(nIs, J)
            dblB[nIs, J] = d
        #A->
        for I in vbForRange(1, n):
            d = dblA(I, k)
            dblA[I, k] = dblA(I, nJs(k))
            dblA[I, nJs(k)] = d
        for J in vbForRange(k + 1, n):
            dblA[k, J] = dblA(k, J) / dblA(k, k)
        for J in vbForRange(1, m):
            dblB[k, J] = dblB(k, J) / dblA(k, k)
        # 回代
        for J in vbForRange(k + 1, n):
            for I in vbForRange(1, n):
                if I <> k:
                    dblA[I, J] = dblA(I, J) - dblA(I, k) * dblA(k, J)
        for J in vbForRange(1, m):
            for I in vbForRange(1, n):
                if I <> k:
                    dblB[I, J] = dblB(I, J) - dblA(I, k) * dblB(k, J)
    # 调整解的次序
    for k in vbForRange(n, 1, -1):
        for J in vbForRange(1, m):
            d = dblB(k, J)
            dblB[k, J] = dblB(nJs(k), J)
            dblB[nJs(k), J] = d
    # 求解成功
    _ret = True
    return _ret

def LECpxGauss(n, dblAR, dblAI, dblBR, dblBI):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()

    nIs = Integer()

    d = Double()

    p = Double()

    q = Double()

    s = Double()
    # 局部变量
    nJs = vbObjectInitialize((n,), Integer)
    # 开始求解
    for k in vbForRange(1, n - 1):
        d = 0
        # 归一
        for I in vbForRange(k, n):
            for J in vbForRange(k, n):
                p = dblAR(I, J) * dblAR(I, J) + dblAI(I, J) * dblAI(I, J)
                if p > d:
                    d = p
                    nJs[k] = J
                    nIs = I
        # 无解，返回
        if d + 1 == 1:
            _ret = False
            return _ret
        # 消元
        for J in vbForRange(k, n):
            p = dblAR(k, J)
            dblAR[k, J] = dblAR(nIs, J)
            dblAR[nIs, J] = p
            p = dblAI(k, J)
            dblAI[k, J] = dblAI(nIs, J)
            dblAI[nIs, J] = p
        p = dblBR(k)
        dblBR[k] = dblBR(nIs)
        dblBR[nIs] = p
        p = dblBI(k)
        dblBI[k] = dblBI(nIs)
        dblBI[nIs] = p
        for I in vbForRange(1, n):
            p = dblAR(I, k)
            dblAR[I, k] = dblAR(I, nJs(k))
            dblAR[I, nJs(k)] = p
            p = dblAI(I, k)
            dblAI[I, k] = dblAI(I, nJs(k))
            dblAI[I, nJs(k)] = p
        # 复数运算
        for J in vbForRange(k + 1, n):
            p = dblAR(k, J) * dblAR(k, k)
            q = - dblAI(k, J) * dblAI(k, k)
            s = ( dblAR(k, k) - dblAI(k, k) )  *  ( dblAR(k, J) + dblAI(k, J) )
            dblAR[k, J] = ( p - q )  / d
            dblAI[k, J] = ( s - p - q )  / d
        p = dblBR(k) * dblAR(k, k)
        q = - dblBI(k) * dblAI(k, k)
        s = ( dblAR(k, k) - dblAI(k, k) )  *  ( dblBR(k) + dblBI(k) )
        dblBR[k] = ( p - q )  / d
        dblBI[k] = ( s - p - q )  / d
        for I in vbForRange(k + 1, n):
            for J in vbForRange(k + 1, n):
                p = dblAR(I, k) * dblAR(k, J)
                q = dblAI(I, k) * dblAI(k, J)
                s = ( dblAR(I, k) + dblAI(I, k) )  *  ( dblAR(k, J) + dblAI(k, J) )
                dblAR[I, J] = dblAR(I, J) - p + q
                dblAI[I, J] = dblAI(I, J) - s + p + q
            p = dblAR(I, k) * dblBR(k)
            q = dblAI(I, k) * dblBI(k)
            s = ( dblAR(I, k) + dblAI(I, k) )  *  ( dblBR(k) + dblBI(k) )
            dblBR[I] = dblBR(I) - p + q
            dblBI[I] = dblBI(I) - s + p + q
    d = dblAR(n, n) * dblAR(n, n) + dblAI(n, n) * dblAI(n, n)
    # 无解，返回
    if d + 1 == 1:
        _ret = False
        return _ret
    p = dblAR(n, n) * dblBR(n)
    q = - dblAI(n, n) * dblBI(n)
    s = ( dblAR(n, n) - dblAI(n, n) )  *  ( dblBR(n) + dblBI(n) )
    dblBR[n] = ( p - q )  / d
    dblBI[n] = ( s - p - q )  / d
    # 回代
    for I in vbForRange(n - 1, 1, -1):
        for J in vbForRange(I + 1, n):
            p = dblAR(I, J) * dblBR(J)
            q = dblAI(I, J) * dblBI(J)
            s = ( dblAR(I, J) + dblAI(I, J) )  *  ( dblBR(J) + dblBI(J) )
            dblBR[I] = dblBR(I) - p + q
            dblBI[I] = dblBI(I) - s + p + q
    # 调整解的次序
    nJs[n] = n
    for k in vbForRange(n, 1, -1):
        p = dblBR(k)
        dblBR[k] = dblBR(nJs(k))
        dblBR[nJs(k)] = p
        p = dblBI(k)
        dblBI[k] = dblBI(nJs(k))
        dblBI[nJs(k)] = p
    # 求解成功
    _ret = True
    return _ret

def LECpxGaussJordan(n, m, dblAR, dblAI, dblBR, dblBI):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()

    nIs = Integer()

    d = Double()

    p = Double()

    q = Double()

    s = Double()
    # 局部变量
    nJs = vbObjectInitialize((n,), Integer)
    # 开始求解
    for k in vbForRange(1, n):
        d = 0
        # 归一
        for I in vbForRange(k, n):
            for J in vbForRange(k, n):
                p = dblAR(I, J) * dblAR(I, J) + dblAI(I, J) * dblAI(I, J)
                if p > d:
                    d = p
                    nJs[k] = J
                    nIs = I
        # 无解，返回
        if d + 1 == 1:
            _ret = False
            return _ret
        # 消元
        if nIs <> k:
            # A->
            for J in vbForRange(k, n):
                p = dblAR(k, J)
                dblAR[k, J] = dblAR(nIs, J)
                dblAR[nIs, J] = p
                p = dblAI(k, J)
                dblAI[k, J] = dblAI(nIs, J)
                dblAI[nIs, J] = p
            # B ->
            for J in vbForRange(1, m):
                p = dblBR(k, J)
                dblBR[k, J] = dblBR(nIs, J)
                dblBR[nIs, J] = p
                p = dblBI(k, J)
                dblBI[k, J] = dblBI(nIs, J)
                dblBI[nIs, J] = p
        if nJs(k) <> k:
            # A->
            for I in vbForRange(1, n):
                p = dblAR(I, k)
                dblAR[I, k] = dblAR(I, nJs(k))
                dblAR[I, nJs(k)] = p
                p = dblAI(I, k)
                dblAI[I, k] = dblAI(I, nJs(k))
                dblAI[I, nJs(k)] = p
        # 复数运算
        for J in vbForRange(k + 1, n):
            p = dblAR(k, J) * dblAR(k, k)
            q = - dblAI(k, J) * dblAI(k, k)
            s = ( dblAR(k, k) - dblAI(k, k) )  *  ( dblAR(k, J) + dblAI(k, J) )
            dblAR[k, J] = ( p - q )  / d
            dblAI[k, J] = ( s - p - q )  / d
        for J in vbForRange(1, m):
            p = dblBR(k, J) * dblAR(k, k)
            q = - dblBI(k, J) * dblAI(k, k)
            s = ( dblAR(k, k) - dblAI(k, k) )  *  ( dblBR(k, J) + dblBI(k, J) )
            dblBR[k, J] = ( p - q )  / d
            dblBI[k, J] = ( s - p - q )  / d
        for I in vbForRange(1, n):
            if I <> k:
                for J in vbForRange(k + 1, n):
                    p = dblAR(I, k) * dblAR(k, J)
                    q = dblAI(I, k) * dblAI(k, J)
                    s = ( dblAR(I, k) + dblAI(I, k) )  *  ( dblAR(k, J) + dblAI(k, J) )
                    dblAR[I, J] = dblAR(I, J) - p + q
                    dblAI[I, J] = dblAI(I, J) - s + p + q
                for J in vbForRange(1, m):
                    p = dblAR(I, k) * dblBR(k, J)
                    q = dblAI(I, k) * dblBI(k, J)
                    s = ( dblAR(I, k) + dblAI(I, k) )  *  ( dblBR(k, J) + dblBI(k, J) )
                    dblBR[I, J] = dblBR(I, J) - p + q
                    dblBI[I, J] = dblBI(I, J) - s + p + q
    # 调整解的次序
    for k in vbForRange(n, 1, -1):
        if nJs(k) <> k:
            for J in vbForRange(1, m):
                p = dblBR(k, J)
                dblBR[k, J] = dblBR(nJs(k), J)
                dblBR[nJs(k), J] = p
                p = dblBI(k, J)
                dblBI[k, J] = dblBI(nJs(k), J)
                dblBI[nJs(k), J] = p
    # 求解成功
    _ret = True
    return _ret

def LETrid(n, m, dblB, dblD):
    _ret = None
    k = Integer()

    J = Integer()

    s = Double()
    # 局部变量
    # 参数校验
    if  ( m <>  ( 3 * n - 2 ) ) :
        _ret = - 1
        return _ret
    # 求解
    for k in vbForRange(1, n - 1):
        J = 3 *  ( k - 1 )  + 1
        s = dblB(J)
        # 无解，返回
        if  ( Abs(s) + 1 == 1 ) :
            _ret = 0
            return _ret
        dblB[J + 1] = dblB(J + 1) / s
        dblD[k] = dblD(k) / s
        dblB[J + 3] = dblB(J + 3) - dblB(J + 2) * dblB(J + 1)
        dblD[k + 1] = dblD(k + 1) - dblB(J + 2) * dblD(k)
    s = dblB(3 * n - 2)
    # 无解，返回
    if  ( Abs(s) + 1 == 1 ) :
        _ret = 0
        return _ret
    dblD[n] = dblD(n) / s
    for k in vbForRange(n - 1, 1, -1):
        dblD[k] = dblD(k) - dblB(3 *  ( k - 1 )  + 2) * dblD(k + 1)
    # 求解成功
    _ret = 1
    return _ret

def LEBand(n, m, L, il, dblB, dblD):
    _ret = None
    ls = Integer()

    k = Integer()

    I = Integer()

    J = Integer()

    nIs = Integer()

    p = Double()

    t = Double()
    # 局部变量
    # 参数校验
    if  ( il <>  ( 2 * L + 1 ) ) :
        _ret = - 1
        return _ret
    # 求解
    ls = L
    for k in vbForRange(1, n - 1):
        p = 0
        for I in vbForRange(k, ls):
            t = Abs(dblB(I, 1))
            if t > p:
                p = t
                nIs = I
        # 无解，返回
        if  ( p + 1 == 1 ) :
            _ret = 0
            return _ret
        for J in vbForRange(1, m):
            t = dblD(k, J)
            dblD[k, J] = dblD(nIs, J)
            dblD[nIs, J] = t
        for J in vbForRange(1, il):
            t = dblB(k, J)
            dblB[k, J] = dblB(nIs, J)
            dblB[nIs, J] = t
        for J in vbForRange(1, m):
            dblD[k, J] = dblD(k, J) / dblB(k, 1)
        for J in vbForRange(2, il):
            dblB[k, J] = dblB(k, J) / dblB(k, 1)
        for I in vbForRange(k + 1, ls + 1):
            t = dblB(I, 1)
            for J in vbForRange(1, m):
                dblD[I, J] = dblD(I, J) - t * dblD(k, J)
            for J in vbForRange(2, il):
                dblB[I, J - 1] = dblB(I, J) - t * dblB(k, J)
            dblB[I, il] = 0
        if  ( ls <> n - 1 ) :
            ls = ls + 1
    p = dblB(n, 1)
    # 无解，返回
    if  ( Abs(p) + 1 == 1 ) :
        _ret = 0
        return _ret
    for J in vbForRange(1, m):
        dblD[n, J] = dblD(n, J) / p
    ls = 1
    for I in vbForRange(n - 1, 1, -1):
        for k in vbForRange(1, m):
            for J in vbForRange(2, ls + 1):
                dblD[I, k] = dblD(I, k) - dblB(I, J) * dblD(I + J - 1, k)
        if  ( ls <>  ( il - 1 ) ) :
            ls = ls + 1
    # 求解成功
    _ret = 1
    return _ret

def LEDjn(n, m, dblA, dblC):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()

    k1 = Integer()

    k2 = Integer()

    k3 = Integer()
    # 局部变量
    # 无解，返回
    if  ( Abs(dblA(1, 1)) + 1 == 1 ) :
        _ret = False
        return _ret
    for I in vbForRange(2, n):
        dblA[I, 1] = dblA(I, 1) / dblA(1, 1)
    for I in vbForRange(2, n - 1):
        for J in vbForRange(2, I):
            dblA[I, I] = dblA(I, I) - dblA(I, J - 1) * dblA(I, J - 1) * dblA(J - 1, J - 1)
        # 无解，返回
        if  ( Abs(dblA(I, I)) + 1 == 1 ) :
            _ret = False
            return _ret
        for k in vbForRange(I + 1, n):
            for J in vbForRange(2, I):
                dblA[k, I] = dblA(k, I) - dblA(k, J - 1) * dblA(I, J - 1) * dblA(J - 1, J - 1)
            dblA[k, I] = dblA(k, I) / dblA(I, I)
    for J in vbForRange(2, n):
        dblA[n, n] = dblA(n, n) - dblA(n, J - 1) * dblA(n, J - 1) * dblA(J - 1, J - 1)
    # 无解，返回
    if  ( Abs(dblA(n, n)) + 1 == 1 ) :
        _ret = False
        return _ret
    for J in vbForRange(1, m):
        for I in vbForRange(2, n):
            for k in vbForRange(2, I):
                dblC[I, J] = dblC(I, J) - dblA(I, k - 1) * dblC(k - 1, J)
    for I in vbForRange(2, n):
        for J in vbForRange(I, n):
            dblA[I - 1, J] = dblA(I - 1, I - 1) * dblA(J, I - 1)
    for J in vbForRange(1, m):
        dblC[n, J] = dblC(n, J) / dblA(n, n)
        for k in vbForRange(2, n):
            k1 = n - k + 2
            for k2 in vbForRange(k1, n):
                k3 = n - k + 1
                dblC[k3, J] = dblC(k3, J) - dblA(k3, k2) * dblC(k2, J)
            dblC[k3, J] = dblC(k3, J) / dblA(k3, k3)
    # 求解成功
    _ret = True
    return _ret

def LECholesky(n, m, dblA, dblD):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()
    # 局部变量
    # 矩阵非正定，求解失败
    if  ( ( dblA(1, 1) + 1 == 1 )  or  ( dblA(1, 1) < 0 ) ) :
        _ret = False
        return _ret
    dblA[1, 1] = Sqr(dblA(1, 1))
    for J in vbForRange(2, n):
        dblA[1, J] = dblA(1, J) / dblA(1, 1)
    for I in vbForRange(2, n):
        for J in vbForRange(2, I):
            dblA[I, I] = dblA(I, I) - dblA(J - 1, I) * dblA(J - 1, I)
        # 求解失败
        if  ( ( dblA(I, I) + 1 == 1 )  or  ( dblA(I, I) < 0 ) ) :
            _ret = False
            return _ret
        dblA[I, I] = Sqr(dblA(I, I))
        if  ( I <> n ) :
            for J in vbForRange(I + 1, n):
                for k in vbForRange(2, I):
                    dblA[I, J] = dblA(I, J) - dblA(k - 1, I) * dblA(k - 1, J)
                dblA[I, J] = dblA(I, J) / dblA(I, I)
    for J in vbForRange(1, m):
        dblD[1, J] = dblD(1, J) / dblA(1, 1)
        for I in vbForRange(2, n):
            for k in vbForRange(2, I):
                dblD[I, J] = dblD(I, J) - dblA(k - 1, I) * dblD(k - 1, J)
            dblD[I, J] = dblD(I, J) / dblA(I, I)
    for J in vbForRange(1, m):
        dblD[n, J] = dblD(n, J) / dblA(n, n)
        for k in vbForRange(n, 2, -1):
            for I in vbForRange(k, n):
                dblD[k - 1, J] = dblD(k - 1, J) - dblA(k - 1, I) * dblD(I, J)
            dblD[k - 1, J] = dblD(k - 1, J) / dblA(k - 1, k - 1)
    # 求解成功
    _ret = True
    return _ret

def LEGgje(n, dblA, dblB):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()

    nIs = Integer()

    d = Double()

    q = Double()
    # 局部变量
    nJs = vbObjectInitialize((n,), Integer)
    # 开始求解
    for k in vbForRange(1, n):
        q = 0
        # 归一
        for I in vbForRange(k, n):
            for J in vbForRange(k, n):
                if Abs(dblA(I, J)) > q:
                    q = Abs(dblA(I, J))
                    nJs[k] = J
                    nIs = I
        # 无解，返回
        if q + 1 == 1:
            _ret = False
            return _ret
        # 消元
        # A->
        for J in vbForRange(k, n):
            d = dblA(k, J)
            dblA[k, J] = dblA(nIs, J)
            dblA[nIs, J] = d
        # B->
        d = dblB(k)
        dblB[k] = dblB(nIs)
        dblB[nIs] = d
        #A->
        for I in vbForRange(1, n):
            d = dblA(I, k)
            dblA[I, k] = dblA(I, nJs(k))
            dblA[I, nJs(k)] = d
        for J in vbForRange(k + 1, n):
            dblA[k, J] = dblA(k, J) / dblA(k, k)
        dblB[k] = dblB(k) / dblA(k, k)
        # 回代
        for J in vbForRange(k + 1, n):
            for I in vbForRange(1, n):
                if I <> k:
                    dblA[I, J] = dblA(I, J) - dblA(I, k) * dblA(k, J)
        for I in vbForRange(1, n):
            if I <> k:
                dblB[I] = dblB(I) - dblA(I, k) * dblB(k)
    # 调整解的次序
    for k in vbForRange(n, 1, -1):
        d = dblB(k)
        dblB[k] = dblB(nJs(k))
        dblB[nJs(k)] = d
    # 求解成功
    _ret = True
    return _ret

def LETlvs(n, dblT, dblB, dblX):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()

    a = Double()

    beta = Double()

    q = Double()

    C = Double()

    h = Double()
    # 局部变量
    y = vbObjectInitialize((n,), Double)
    s = vbObjectInitialize((n,), Double)
    a = dblT(1)
    if  ( Abs(a) + 1 == 1 ) :
        _ret = False
        return _ret
    y[1] = 1
    dblX[1] = dblB(1) / a
    for k in vbForRange(1, n - 1):
        beta = 0
        q = 0
        for J in vbForRange(1, k):
            beta = beta + y(J) * dblT(J + 1)
            q = q + dblX(J) * dblT(k - J + 2)
        if  ( Abs(a) + 1 == 1 ) :
            _ret = False
            return _ret
        C = - beta / a
        s[1] = C * y(k)
        y[k + 1] = y(k)
        if  ( k <> 1 ) :
            for I in vbForRange(2, k):
                s[I] = y(I - 1) + C * y(k - I + 1)
        a = a + C * beta
        if  ( Abs(a) + 1 == 1 ) :
            _ret = False
            return _ret
        h = ( dblB(k + 1) - q )  / a
        for I in vbForRange(1, k):
            dblX[I] = dblX(I) + h * s(I)
            y[I] = s(I)
        dblX[k + 1] = h * y(k + 1)
    _ret = True
    return _ret

def LESeidel(n, dblA, dblB, dblX, eps):
    _ret = None
    I = Integer()

    J = Integer()

    p = Double()

    q = Double()

    s = Double()

    t = Double()
    # 局部变量
    # 校验系数矩阵是否主对角线占绝对优势
    for I in vbForRange(1, n):
        p = 0
        dblX[I] = 0
        for J in vbForRange(1, n):
            if  ( I <> J ) :
                p = p + Abs(dblA(I, J))
        if  ( p >= Abs(dblA(I, I)) ) :
            _ret = False
            return _ret
    # 迭代求解
    p = eps + 1
    while ( p >= eps ):
        p = 0
        for I in vbForRange(1, n):
            t = dblX(I)
            s = 0
            for J in vbForRange(1, n):
                if  ( J <> I ) :
                    s = s + dblA(I, J) * dblX(J)
            dblX[I] = ( dblB(I) - s )  / dblA(I, I)
            q = Abs(dblX(I) - t) /  ( 1 + Abs(dblX(I)) )
            if  ( q > p ) :
                p = q
    # 求解成功
    _ret = True
    return _ret

def LEGrad(n, dblA, dblB, dblX, eps):
    I = Integer()

    k = Integer()

    alpha = Double()

    beta = Double()

    d = Double()

    e = Double()
    # 局部变量
    p = vbObjectInitialize((n, 1,), Double)
    R = vbObjectInitialize((n,), Double)
    s = vbObjectInitialize((n, 1,), Double)
    q = vbObjectInitialize((n, 1,), Double)
    x = vbObjectInitialize((n, 1,), Double)
    # 初始化
    for I in vbForRange(1, n):
        x[I, 1] = 0
        p[I, 1] = dblB(I)
        R[I] = dblB(I)
    # 循环求解
    I = 1
    while ( I <= n ):
        # 矩阵乘法
        Call(MMul(n, n, 1, dblA, p, s))
        d = 0
        e = 0
        for k in vbForRange(1, n):
            d = d + p(k, 1) * dblB(k)
            e = e + p(k, 1) * s(k, 1)
        alpha = d / e
        for k in vbForRange(1, n):
            x[k, 1] = x(k, 1) + alpha * p(k, 1)
        # 矩阵乘法
        Call(MMul(n, n, 1, dblA, x, q))
        d = 0
        for k in vbForRange(1, n):
            R[k] = dblB(k) - q(k, 1)
            d = d + R(k) * s(k, 1)
        beta = d / e
        d = 0
        for k in vbForRange(1, n):
            d = d + R(k) * R(k)
        d = Sqr(d)
        # 求解结束，返回
        if  ( d < eps ) :
            for k in vbForRange(1, n):
                dblX[k] = x(k, 1)
            return
        for k in vbForRange(1, n):
            p[k, 1] = R(k) - beta * p(k, 1)
        I = I + 1
    # 求解结束，返回
    for k in vbForRange(1, n):
        dblX[k] = x(k, 1)

def MatrixToString(m, n, mtxA, sFormat):
    _ret = None
    I = Integer()

    J = Integer()

    s = String()
    s = ''
    for I in vbForRange(1, m):
        for J in vbForRange(1, n):
            s = s + Format(mtxA(I, J), sFormat) + '  '
        s = s + Chr(13)
    _ret = s
    return _ret

def MMul(m, n, L, mtxA, mtxB, mtxC):
    I = Integer()

    J = Integer()

    k = Integer()
    for I in vbForRange(1, m):
        for J in vbForRange(1, L):
            mtxC[I, J] = 0
            for k in vbForRange(1, n):
                mtxC[I, J] = mtxC(I, J) + mtxA(I, k) * mtxB(k, J)

def LEMqr(m, n, dblA, dblB, dblQ):
    _ret = None
    I = Integer()

    J = Integer()

    d = Double()
    # 局部变量
    C = vbObjectInitialize((n,), Double)
    # QR分解失败，返回
    if  ( not MMqr(m, n, dblA, dblQ) ) :
        _ret = False
        return _ret
    for I in vbForRange(1, n):
        d = 0
        for J in vbForRange(1, m):
            d = d + dblQ(J, I) * dblB(J)
        C[I] = d
    dblB[n] = C(n) / dblA(n, n)
    for I in vbForRange(n - 1, 1, -1):
        d = 0
        for J in vbForRange(I + 1, n):
            d = d + dblA(I, J) * dblB(J)
        dblB[I] = ( C(I) - d )  / dblA(I, I)
    # 求解结束，返回
    _ret = True
    return _ret

def MMqr(m, n, dblA, dblQ):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()

    nn = Integer()

    jj = Integer()

    u = Double()

    alpha = Double()

    w = Double()

    t = Double()
    if  ( m < n ) :
        _ret = False
        return _ret
    for I in vbForRange(1, m):
        for J in vbForRange(1, m):
            dblQ[I, J] = 0
            if  ( I == J ) :
                dblQ[I, J] = 1
    nn = n
    if  ( m == n ) :
        nn = m - 1
    for k in vbForRange(1, nn):
        u = 0
        for I in vbForRange(k, m):
            w = Abs(dblA(I, k))
            if  ( w > u ) :
                u = w
        alpha = 0
        for I in vbForRange(k, m):
            t = dblA(I, k) / u
            alpha = alpha + t * t
        if  ( dblA(k, k) > 0 ) :
            u = - u
        alpha = u * Sqr(alpha)
        if  ( Abs(alpha) + 1 == 1 ) :
            _ret = False
            return _ret
        u = Sqr(2 * alpha *  ( alpha - dblA(k, k) ))
        if  ( ( u + 1 )  <> 1 ) :
            dblA[k, k] = ( dblA(k, k) - alpha )  / u
            for I in vbForRange(k + 1, m):
                dblA[I, k] = dblA(I, k) / u
            for J in vbForRange(1, m):
                t = 0
                for jj in vbForRange(k, m):
                    t = t + dblA(jj, k) * dblQ(jj, J)
                for I in vbForRange(k, m):
                    dblQ[I, J] = dblQ(I, J) - 2 * t * dblA(I, k)
            for J in vbForRange(k + 1, n):
                t = 0
                for jj in vbForRange(k, m):
                    t = t + dblA(jj, k) * dblA(jj, J)
                for I in vbForRange(k, m):
                    dblA[I, J] = dblA(I, J) - 2 * t * dblA(I, k)
            dblA[k, k] = alpha
            for I in vbForRange(k + 1, m):
                dblA[I, k] = 0
    for I in vbForRange(1, m - 1):
        for J in vbForRange(I + 1, m):
            t = dblQ(I, J)
            dblQ[I, J] = dblQ(J, I)
            dblQ[J, I] = t
    _ret = True
    return _ret

def LEMiv(m, n, dblA, dblB, dblX, dblAP, dblU, dblV, ka, eps):
    _ret = None
    I = Integer()

    J = Integer()
    # 局部变量
    if  ( not MInv(m, n, dblA, dblAP, dblU, dblV, ka, eps) ) :
        _ret = False
        return _ret
    for I in vbForRange(1, n):
        dblX[I] = 0
        for J in vbForRange(1, m):
            if GetInputState:
                DoEvents()
            dblX[I] = dblX(I) + dblAP(I, J) * dblB(J)
    _ret = True
    return _ret

def MUav(m, n, dblA, dblU, dblV, ka, eps):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()

    L = Integer()

    it = Integer()

    ll = Integer()

    kk = Integer()

    mm = Integer()

    nn = Integer()

    m1 = Integer()

    ks = Integer()

    d = Double()

    dd = Double()

    t = Double()

    sm = Double()

    sm1 = Double()

    em1 = Double()

    sk = Double()

    ek = Double()

    b = Double()

    C = Double()

    shh = Double()

    fg = vbObjectInitialize((2,), Double)

    cs = vbObjectInitialize((2,), Double)
    # 局部变量
    s = vbObjectInitialize((ka,), Double)
    e = vbObjectInitialize((ka,), Double)
    w = vbObjectInitialize((ka,), Double)
    it = 60
    k = n
    if  ( m - 1 < n ) :
        k = m - 1
    L = m
    if  ( n - 2 < m ) :
        L = n - 2
    if  ( L < 0 ) :
        L = 0
    ll = k
    if  ( L > k ) :
        ll = L
    if  ( ll >= 1 ) :
        for kk in vbForRange(1, ll):
            if  ( kk <= k ) :
                d = 0
                for I in vbForRange(kk, m):
                    if GetInputState:
                        DoEvents()
                    d = d + dblA(I, kk) * dblA(I, kk)
                s[kk] = Sqr(d)
                if s(kk) <> 0:
                    if  ( dblA(kk, kk) <> 0 ) :
                        s[kk] = Abs(s(kk))
                        if  ( dblA(kk, kk) < 0 ) :
                            s[kk] = - s(kk)
                    for I in vbForRange(kk, m):
                        if GetInputState:
                            DoEvents()
                        dblA[I, kk] = dblA(I, kk) / s(kk)
                    dblA[kk, kk] = 1 + dblA(kk, kk)
                s[kk] = - s(kk)
            if  ( n >= kk + 1 ) :
                for J in vbForRange(kk + 1, n):
                    if  ( ( kk <= k )  and  ( s(kk) <> 0 ) ) :
                        d = 0
                        for I in vbForRange(kk, m):
                            if GetInputState:
                                DoEvents()
                            d = d + dblA(I, kk) * dblA(I, J)
                        d = - d / dblA(kk, kk)
                        for I in vbForRange(kk, m):
                            if GetInputState:
                                DoEvents()
                            dblA[I, J] = dblA(I, J) + d * dblA(I, kk)
                    e[J] = dblA(kk, J)
            if  ( kk <= k ) :
                for I in vbForRange(kk, m):
                    if GetInputState:
                        DoEvents()
                    dblU[I, kk] = dblA(I, kk)
            if  ( kk <= L ) :
                d = 0
                for I in vbForRange(kk + 1, n):
                    if GetInputState:
                        DoEvents()
                    d = d + e(I) * e(I)
                e[kk] = Sqr(d)
                if  ( e(kk) <> 0 ) :
                    if  ( e(kk + 1) <> 0 ) :
                        e[kk] = Abs(e(kk))
                        if  ( e(kk + 1) < 0 ) :
                            e[kk] = - e(kk)
                    for I in vbForRange(kk + 1, n):
                        if GetInputState:
                            DoEvents()
                        e[I] = e(I) / e(kk)
                    e[kk + 1] = 1 + e(kk + 1)
                e[kk] = - e(kk)
                if  ( ( kk + 1 <= m )  and  ( e(kk) <> 0 ) ) :
                    for I in vbForRange(kk + 1, m):
                        w[I] = 0
                    for J in vbForRange(kk + 1, n):
                        for I in vbForRange(kk + 1, m):
                            if GetInputState:
                                DoEvents()
                            w[I] = w(I) + e(J) * dblA(I, J)
                    for J in vbForRange(kk + 1, n):
                        for I in vbForRange(kk + 1, m):
                            if GetInputState:
                                DoEvents()
                            dblA[I, J] = dblA(I, J) - w(I) * e(J) / e(kk + 1)
                for I in vbForRange(kk + 1, n):
                    dblV[I, kk] = e(I)
    mm = n
    if  ( m + 1 < n ) :
        mm = m + 1
    if  ( k < n ) :
        s[k + 1] = dblA(k + 1, k + 1)
    if  ( m < mm ) :
        s[mm] = 0
    if  ( L + 1 < mm ) :
        e[L + 1] = dblA(L + 1, mm)
    e[mm] = 0
    nn = m
    if  ( m > n ) :
        nn = n
    if  ( nn >= k + 1 ) :
        for J in vbForRange(k + 1, nn):
            for I in vbForRange(1, m):
                if GetInputState:
                    DoEvents()
                dblU[I, J] = 0
            dblU[J, J] = 1
    if  ( k >= 1 ) :
        for ll in vbForRange(1, k):
            kk = k - ll + 1
            if  ( s(kk) <> 0 ) :
                if  ( nn >= kk + 1 ) :
                    for J in vbForRange(kk + 1, nn):
                        d = 0
                        for I in vbForRange(kk, m):
                            if GetInputState:
                                DoEvents()
                            d = d + dblU(I, kk) * dblU(I, J) / dblU(kk, kk)
                        d = - d
                        for I in vbForRange(kk, m):
                            dblU[I, J] = dblU(I, J) + d * dblU(I, kk)
                for I in vbForRange(kk, m):
                    dblU[I, kk] = - dblU(I, kk)
                dblU[kk, kk] = 1 + dblU(kk, kk)
                if  ( kk - 1 >= 1 ) :
                    for I in vbForRange(1, kk - 1):
                        dblU[I, kk] = 0
            else:
                for I in vbForRange(1, m):
                    dblU[I, kk] = 0
                dblU[kk, kk] = 1
    for ll in vbForRange(1, n):
        kk = n - ll + 1
        if  ( ( kk <= L )  and  ( e(kk) <> 0 ) ) :
            for J in vbForRange(kk + 1, n):
                d = 0
                for I in vbForRange(kk + 1, n):
                    d = d + dblV(I, kk) * dblV(I, J) / dblV(kk + 1, kk)
                d = - d
                for I in vbForRange(kk + 1, n):
                    dblV[I, J] = dblV(I, J) + d * dblV(I, kk)
        for I in vbForRange(1, n):
            dblV[I, kk] = 0
        dblV[kk, kk] = 1
    for I in vbForRange(1, m):
        for J in vbForRange(1, n):
            dblA[I, J] = 0
    m1 = mm
    it = 60
    while ( 1 ):
        if  ( mm == 0 ) :
            Call(Cal1(dblA, e, s, dblV, m, n))
            _ret = True
            return _ret
        if  ( it == 0 ) :
            Call(Cal1(dblA, e, s, dblV, m, n))
            _ret = False
            return _ret
        kk = mm - 1
        while ( ( kk <> 0 )  and  ( Abs(e(kk)) <> 0 ) ):
            d = Abs(s(kk)) + Abs(s(kk + 1))
            dd = Abs(e(kk))
            if  ( dd > eps * d ) :
                kk = kk - 1
            else:
                e[kk] = 0
        if  ( kk == mm - 1 ) :
            kk = kk + 1
            if  ( s(kk) < 0 ) :
                s[kk] = - s(kk)
                for I in vbForRange(1, n):
                    dblV[I, kk] = - dblV(I, kk)
            while ( ( kk <> m1 )  and  ( s(kk) < s(kk + 1) ) ):
                d = s(kk)
                s[kk] = s(kk + 1)
                s[kk + 1] = d
                if  ( kk < n ) :
                    for I in vbForRange(1, n):
                        d = dblV(I, kk)
                        dblV[I, kk] = dblV(I, kk + 1)
                        dblV[I, kk + 1] = d
                if  ( kk < m ) :
                    for I in vbForRange(1, m):
                        d = dblU(I, kk)
                        dblU[I, kk] = dblU(I, kk + 1)
                        dblU[I, kk + 1] = d
                kk = kk + 1
            it = 60
            mm = mm - 1
        else:
            ks = mm
            while ( ( ks > kk )  and  ( Abs(s(ks)) <> 0 ) ):
                d = 0
                if  ( ks <> mm ) :
                    d = d + Abs(e(ks))
                if  ( ks <> kk + 1 ) :
                    d = d + Abs(e(ks - 1))
                dd = Abs(s(ks))
                if  ( dd > eps * d ) :
                    ks = ks - 1
                else:
                    s[ks] = 0
            if  ( ks == kk ) :
                kk = kk + 1
                d = Abs(s(mm))
                t = Abs(s(mm - 1))
                if  ( t > d ) :
                    d = t
                t = Abs(e(mm - 1))
                if  ( t > d ) :
                    d = t
                t = Abs(s(kk))
                if  ( t > d ) :
                    d = t
                t = Abs(e(kk))
                if  ( t > d ) :
                    d = t
                sm = s(mm) / d
                sm1 = s(mm - 1) / d
                em1 = e(mm - 1) / d
                sk = s(kk) / d
                ek = e(kk) / d
                b = ( ( sm1 + sm )  *  ( sm1 - sm )  + em1 * em1 )  / 2
                C = sm * em1
                C = C * C
                shh = 0
                if  ( ( b <> 0 )  or  ( C <> 0 ) ) :
                    shh = Sqr(b * b + C)
                    if  ( b < 0 ) :
                        shh = - shh
                    shh = C /  ( b + shh )
                fg[1] = ( sk + sm )  *  ( sk - sm )  - shh
                fg[2] = sk * ek
                for I in vbForRange(kk, mm - 1):
                    Call(Cal2(fg, cs))
                    if  ( I <> kk ) :
                        e[I - 1] = fg(1)
                    fg[1] = cs(1) * s(I) + cs(2) * e(I)
                    e[I] = cs(1) * e(I) - cs(2) * s(I)
                    fg[2] = cs(2) * s(I + 1)
                    s[I + 1] = cs(1) * s(I + 1)
                    if  ( ( cs(1) <> 1 )  or  ( cs(2) <> 0 ) ) :
                        for J in vbForRange(1, n):
                            d = cs(1) * dblV(J, I) + cs(2) * dblV(J, I + 1)
                            dblV[J, I + 1] = - cs(2) * dblV(J, I) + cs(1) * dblV(J, I + 1)
                            dblV[J, I] = d
                    Call(Cal2(fg, cs))
                    s[I] = fg(1)
                    fg[1] = cs(1) * e(I) + cs(2) * s(I + 1)
                    s[I + 1] = - cs(2) * e(I) + cs(1) * s(I + 1)
                    fg[2] = cs(2) * e(I + 1)
                    e[I + 1] = cs(1) * e(I + 1)
                    if  ( I < m ) :
                        if  ( ( cs(1) <> 1 )  or  ( cs(2) <> 0 ) ) :
                            for J in vbForRange(1, m):
                                d = cs(1) * dblU(J, I) + cs(2) * dblU(J, I + 1)
                                dblU[J, I + 1] = - cs(2) * dblU(J, I) + cs(1) * dblU(J, I + 1)
                                dblU[J, I] = d
                e[mm - 1] = fg(1)
                it = it - 1
            else:
                if  ( ks == mm ) :
                    kk = kk + 1
                    fg[2] = e(mm - 1)
                    e[mm - 1] = 0
                    for ll in vbForRange(kk, mm - 1):
                        I = mm + kk - ll - 1
                        fg[1] = s(I)
                        Call(Cal2(fg, cs))
                        s[I] = fg(1)
                        if  ( I <> kk ) :
                            fg[2] = - cs(2) * e(I - 1)
                            e[I - 1] = cs(1) * e(I - 1)
                        if  ( ( cs(1) <> 1 )  or  ( cs(2) <> 0 ) ) :
                            for J in vbForRange(1, n):
                                d = cs(1) * dblV(J, I) + cs(2) * dblV(J, mm)
                                dblV[J, mm] = - cs(2) * dblV(J, I) + cs(1) * dblV(J, mm)
                                dblV[J, I] = d
                else:
                    kk = ks + 1
                    fg[2] = e(kk - 1)
                    e[kk - 1] = 0
                    for I in vbForRange(kk, mm):
                        fg[1] = s(I)
                        Call(Cal2(fg, cs))
                        s[I] = fg(1)
                        fg[2] = - cs(2) * e(I)
                        e[I] = cs(1) * e(I)
                        if  ( ( cs(1) <> 1 )  or  ( cs(2) <> 0 ) ) :
                            for J in vbForRange(1, m):
                                d = cs(1) * dblU(J, I) + cs(2) * dblU(J, kk - 1)
                                dblU[J, kk - 1] = - cs(2) * dblU(J, I) + cs(1) * dblU(J, kk - 1)
                                dblU[J, I] = d
    _ret = True
    return _ret

def Cal1(dblA, e, s, dblV, m, n):
    I = Integer()

    J = Integer()

    p = Integer()

    q = Integer()

    d = Double()
    if  ( m >= n ) :
        I = n
    else:
        I = m
    for J in vbForRange(1, I - 1):
        dblA[J, J] = s(J)
        dblA[J, J + 1] = e(J)
    dblA[I, I] = s(I)
    if  ( m < n ) :
        dblA[I, I + 1] = e(I)
    for I in vbForRange(1, n - 1):
        for J in vbForRange(I + 1, n):
            d = dblV(I, J)
            dblV[I, J] = dblV(J, I)
            dblV[J, I] = d

def Cal2(fg, cs):
    R = Double()

    d = Double()
    if  ( ( Abs(fg(1)) + Abs(fg(2)) )  == 0 ) :
        cs[1] = 1
        cs[2] = 0
        d = 0
    else:
        d = Sqr(fg(1) * fg(1) + fg(2) * fg(2))
        if  ( Abs(fg(1)) > Abs(fg(2)) ) :
            d = Abs(d)
            if  ( fg(1) < 0 ) :
                d = - d
        if  ( Abs(fg(2)) >= Abs(fg(1)) ) :
            d = Abs(d)
            if  ( fg(2) < 0 ) :
                d = - d
        cs[1] = fg(1) / d
        cs[2] = fg(2) / d
    R = 1
    if  ( Abs(fg(1)) > Abs(fg(2)) ) :
        R = cs(2)
    else:
        if  ( cs(1) <> 0 ) :
            R = 1 / cs(1)
    fg[1] = d
    fg[2] = R

def MInv(m, n, dblA, dblAP, dblU, dblV, ka, eps):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()

    L = Integer()
    # 局部变量
    if not MUav(m, n, dblA, dblU, dblV, ka, eps):
        _ret = False
        return _ret
    J = n
    if  ( m < n ) :
        J = m
    J = J - 1
    k = 0
    while ( k <= J ):
        if GetInputState:
            DoEvents()
        if  ( dblA(k + 1, k + 1) == 0 ) :
            # VB2PY (UntranslatedCode) GoTo o_lable
            pass
        k = k + 1
    k = k - 1
    for I in vbForRange(0, n - 1):
        for J in vbForRange(0, m - 1):
            if GetInputState:
                DoEvents()
            dblAP[I + 1, J + 1] = 0
            for L in vbForRange(0, k):
                if GetInputState:
                    DoEvents()
                dblAP[I + 1, J + 1] = dblAP(I + 1, J + 1) + dblV(L + 1, I + 1) * dblU(J + 1, L + 1) / dblA(L + 1, L + 1)
    _ret = True
    return _ret

def LEMorbid(n, dblA, dblB, dblX, eps):
    _ret = None
    I = Integer()

    J = Integer()

    k = Integer()

    kk = Integer()

    q = Double()

    qq = Double()
    # 局部变量
    p = vbObjectInitialize((n, n,), Double)
    R = vbObjectInitialize((n,), Double)
    e = vbObjectInitialize((n, 1,), Double)
    x = vbObjectInitialize((n, 1,), Double)
    xx = vbObjectInitialize((n,), Double)
    I = 60
    for k in vbForRange(1, n):
        for J in vbForRange(1, n):
            p[k, J] = dblA(k, J)
    for k in vbForRange(1, n):
        x[k, 1] = dblB(k)
    for k in vbForRange(1, n):
        xx[k] = x(k, 1)
    # 全选主元高斯消去法
    if  ( not LEGauss(n, p, xx) ) :
        _ret = False
        return _ret
    for k in vbForRange(1, n):
        x[k, 1] = xx(k)
    q = 1 + eps
    while ( q >= eps ):
        if  ( I == 0 ) :
            _ret = False
            return _ret
        I = I - 1
        # 矩阵乘法
        Call(MMul(n, n, 1, dblA, x, e))
        for k in vbForRange(1, n):
            R[k] = dblB(k) - e(k, 1)
        for k in vbForRange(1, n):
            for J in vbForRange(1, n):
                p[k, J] = dblA(k, J)
        # 全选主元高斯消去法
        if  ( not LEGauss(n, p, R) ) :
            _ret = False
            return _ret
        q = 0
        for k in vbForRange(1, n):
            qq = Abs(R(k)) /  ( 1 + Abs(x(k, 1) + R(k)) )
            if  ( qq > q ) :
                q = qq
        for k in vbForRange(1, n):
            x[k, 1] = x(k, 1) + R(k)
    # 解赋值返回
    for k in vbForRange(1, n):
        dblX[k] = x(k, 1)
    _ret = True
    return _ret

# VB2PY (UntranslatedCode) Attribute VB_Name = "LEModule"
# VB2PY (UntranslatedCode) Option Explicit
