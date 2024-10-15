Attribute VB_Name = "algrithm"
Option Explicit
Public a(20, 20) As Integer, n As Integer, m As Integer, b(20) As Integer, c(20) As Integer
Public gama(20) As Integer, s(20) As Integer, theta(20) As Single
Public nonji(20) As Integer, ji(20) As Integer

'Sub SimpleForm(x() As Integer, gama As Integer, b As Integer, s As Integer, c As Integer, theta As Integer, a() As Integer)
Sub SimpleForm()
Dim Sol As Integer, ext As Integer, nl As Integer, k As Integer, l As Integer
Sol = -1
 While (Sol = -1)
    Call CalculateGama          '计算检验数
    'k = GotMaxGama
    Sol = GotSolution()         '是否已达到最优解
    If Sol = 0 Then
        Call OutputSolution     '输出结果
        MsgBox "最优结果", vbOKOnly
        Exit Sub
    End If
    k = GotMaxGama
    If NoLimit(k) = -1 Then         '该问题是否无界
        MsgBox "问题无界", vbOKOnly
        Exit Sub
    End If
    If Sol = -2 Then   '该问题有解否
        MsgBox "no solution", vbOKOnly
        Exit Sub
    End If

    l = CalculateTheta(k)       '选取退出变量
    Call ChangeGroup(l, k)           '变换方程组
    
 Wend
    
End Sub

Sub CalculateGama()
Dim j As Integer, r As Integer, i As Integer, t As Integer
For j = 0 To n - m - 1
    r = nonji(j)
    s(r) = 0
    For i = 0 To m - 1
        t = ji(i)
        s(r) = c(t) * a(t, r) + s(r)
    Next i
    gama(r) = c(r) - s(r)
    Debug.Print "gama(" & r & "): " & gama(r)
Next j
End Sub

Function GotSolution()
Dim j As Integer, r As Integer
Dim i As Integer
For j = 0 To n - m - 1
    r = nonji(j)
    If gama(r) > 0 Then
        GotSolution = -1          '未达到最优解
        Exit Function
    End If
Next j
GotSolution = 0
'若基变量中仍含有人工变量，无解.??????not ok?????????????

'sol=-2
End Function

Function NoLimit(k As Integer)
Dim i As Integer
 For i = 0 To m - 1
     If a(i, k) > 0 Then
        NoLimit = 0
         Exit Function
     End If
 Next i
NoLimit = -1    '问题无界

End Function

Sub OutputSolution()
Dim i As Integer, j As Integer, r As Integer, StrOut As String
For j = 0 To m - 1
    StrOut = StrOut & "x" & ji(j) & ": " & b(ji(i)) & " "
     
Next j
MsgBox "结果：" & vbCrLf & StrOut, vbOKOnly
End Sub

Function GotMaxGama()
'返回最大Gama的非基变量的下标,(即不是非基变量数组的序号,而是该数组元素的值)

Dim i As Integer, val As Integer, j As Integer, r As Integer
val = gama(0)
j = nonji(0)
For i = 1 To n - m - 1
    r = nonji(i)
    If gama(r) > val Then
        val = gama(r)
        j = r
    End If
Next i
GotMaxGama = j
End Function

Function CalculateTheta(k As Integer)     '计算theta(),并求其最小值
Dim i As Integer, j As Integer, r As Integer, min As Single
'k代表最大的检验数所对应的非基变量的下标
'k = GotMaxGama
min = b(0) / a(0, k)
j = 0
Debug.Print "theta(0): " & min
For i = 1 To m - 1
    If a(i, k) > 0 Then
        theta(i) = b(i) / a(i, k)
        Debug.Print "theta(" & i & "): " & theta(i)
        If theta(i) < min Then
            min = theta(i)
            j = i
        End If
    End If
Next i
CalculateTheta = j
'返回的是最小的THETA值对应的基变量的下标

End Function

Sub ChangeGroup(l As Integer, k As Integer)
'l代表最小的theta值所对应的基变量下标,为退出变量
'k代表最大的GAMA值所对应的非基变量下标，为引入变量
'a(nonji(l),ji(k))为主元素 /* 2000.06.01 */
'a(l,k)为主元素

Dim i As Integer, j As Integer, aik As Integer
'l = CalculateTheta
'k = GotMaxGama
For i = 0 To m - 1
    aik = a(i, k)
    For j = 0 To n - 1
     '   aik = a(i, k)
        If i <> l Then
            a(i, j) = a(i, j) - aik * a(l, j) / a(l, k)
        Else
            a(i, j) = a(i, j) / a(l, k)
        End If
        Debug.Print "a(" & i & "," & j & "): " & a(i, j)
    Next j
    b(i) = b(i) - aik * b(l) / a(l, k)
Next i
Call ChangeVar(l, k)
End Sub

Sub ChangeVar(l As Integer, k As Integer)
'由于l代表最小的theta值所对应的基变量下标在数组中的序号!!!!! , ji(l)为退出变量,
'k代表最大的GAMA值所对应的非基变量下标(不是序号），为引入变量 ,须先在nonji()中找到对应的数组序号
Dim tt As Integer, i As Integer
Dim dd As String
'tt = ji(k)
'ji(k) = nonji(l)
'nonji(l) = tt
dd = "ji: "
For tt = 0 To m - 1
    dd = dd & ji(tt)
Next
Debug.Print dd
dd = "feiji: "
For tt = 0 To n - m - 1
    dd = dd & nonji(tt)
Next
Debug.Print dd

tt = ji(l)
ji(l) = k
'For i = 0 To m - 1
'    If ji(i) = l Then
'        ji(i) = k
'        Exit Sub
'    End If
'Next i
For i = 0 To n - m - 1
    If nonji(i) = k Then
        nonji(i) = tt
        'Exit Sub
    End If
Next i

dd = "ji: "
For tt = 0 To m - 1
    dd = dd & ji(tt)
Next
Debug.Print dd
dd = "feiji: "
For tt = 0 To n - m - 1
    dd = dd & nonji(tt)
Next
Debug.Print dd
End Sub

