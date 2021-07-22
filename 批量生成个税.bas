Attribute VB_Name = "模块1"
'用于对付纳税人批量补缴个人所得税，可以根据一个模板文件与指定的
'月份区间进行自动生成申报表
'珍爱生命，远离金税三
' lianjie 2017-06-28，2018-01-19,2018-07-02,2018-10-29,2018-11-23
'2018-12-28,2019-03-05,2019-08-25


'全局变量，用于文件名与申报月份的存储 格式：2015-05
Dim mouths(0 To 100) As String
Dim pth As String '当前运行的路径


Sub main_fuck()

    '定义变量
    Dim MyFile As String, WJM As String, GZBM As String
    Dim XLapp As New Excel.Application
    Dim Xlbook As Excel.Workbook
    Dim m_file As String
    Dim tmp_m() As String
    Dim fmth As String
    Dim lmth As String
    Dim fday As String '一个月的开头一天
    Dim lday As String '一个月的最后一天
    Dim n As Integer
    Dim fdate_rg As String
    Dim ldate_rg As String
    Dim jcfy_rg As String '减除费用
    Dim max_row As Integer
    Dim title As String
        
    '参数设置
    ChDir (ThisWorkbook.Path) '切换回当前目录
    pth = ThisWorkbook.Path
    fdate_rg = "H11:H"
    ldate_rg = "I11:I"
    jcfy_rg = "Y11:Y"
    m_file = ""
    fmth = ""
    lmth = ""
    max_row = 10
    TABLETITLE = "扣缴个人所得税申报表（适用于综合所得预扣预缴）"
    
    m_file = Sheets("Sheet1").cells(2, 1)
    fmth = Sheets("Sheet1").cells(2, 2)
    lmth = Sheets("Sheet1").cells(2, 3)
    If (m_file = "" Or fmth = "" Or lmth = "") Then
       res = MsgBox("请填写正确的参数", vbOKOnly)
       Exit Sub
    End If
    If is_illegal(fmth, lmth) Then
       res = MsgBox("请填写正确的参数", vbOKOnly)
       Exit Sub
    End If
    
    If Dir(pth & "\" & m_file, vbDirectory) <> Empty Then
        '先生成一堆文件
        Call make_many_files(pth & "\" & m_file, fmth, lmth)
    Else
        res = MsgBox("模板文件不存在!", vbOKOnly)
        Exit Sub
    End If
    Application.DisplayAlerts = False '取消覆盖保存时的提示
    Application.ScreenUpdating = False '取消屏幕刷新
    WJM = m_file
    GZBM = "扣缴个人所得税报告表"
    MyFile = ThisWorkbook.Path & "\" & WJM
    Set Xlbook = XLapp.Workbooks.Open(MyFile)
    title = Xlbook.Sheets(1).cells(1, 1)
    Xlbook.Close True
    Set XLapp = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    '2019-01-01及以后的表格
    If title = TABLETITLE Then
        For n = 0 To 100
            If mouths(n) = "" Then
                Exit For
            End If
            '有关信息设置
            Application.DisplayAlerts = False '取消覆盖保存时的提示
            Application.ScreenUpdating = False '取消屏幕刷新
            WJM = mouths(n)
            GZBM = "扣缴个人所得税申报表（适用于综合所得预扣预缴）"
            MyFile = ThisWorkbook.Path & "\" & WJM & ".xls"
            Set Xlbook = XLapp.Workbooks.Open(MyFile)
            If Xlbook.Sheets(1).Name <> GZBM Then
                res = MsgBox("表格模板有问题啊!", vbOKOnly)
                Xlbook.Close True
                Set XLapp = Nothing
                Application.ScreenUpdating = True
                Application.DisplayAlerts = True
                Exit Sub
            End If
            If Application.Version > 11 Then 'execl 2003
                Xlbook.CheckCompatibility = False '该死的兼容性检查
            End If
            XLapp.Visible = False
            fday = mouths(n) & "-01"
            tmp_m() = Split(mouths(n), "-")
            If tmp_m(1) = "01" Or tmp_m(1) = "03" Or tmp_m(1) = "05" Or tmp_m(1) = "07" Or _
            tmp_m(1) = "08" Or tmp_m(1) = "10" Or tmp_m(1) = "12" Then
                lday = mouths(n) & "-31"
            ElseIf tmp_m(1) = "02" Then '2月份，看是不是润年
                If tmp_m(0) Mod 4 = 0 And tmp_m(0) Mod 100 <> 0 Then '闰年
                    lday = mouths(n) & "-29"
                Else
                    If tmp_m(0) Mod 100 = 0 And tmp_m(0) Mod 400 = 0 Then '闰年
                        lday = mouths(n) & "-29"
                    Else '非闰年
                        lday = mouths(n) & "-28"
                    End If
                End If
            Else '剩下的是小月了
                    lday = mouths(n) & "-30"
            End If
            With Xlbook.Sheets(GZBM)
                .Range("M3") = fday '把月头一天写入对应单元格
                .Range("R3") = lday '把月未一天写入对应单元格
            End With
            Xlbook.Close True
            Set Xlbook = Nothing
            Set XLapp = Nothing
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
        Next n
    Else
        '处理2019-01-01以前的表格
        For n = 0 To 100
            If mouths(n) = "" Then
                Exit For
            End If
            '有关信息设置
            Application.DisplayAlerts = False '取消覆盖保存时的提示
            Application.ScreenUpdating = False '取消屏幕刷新
            WJM = mouths(n)
            GZBM = "扣缴个人所得税报告表"
            MyFile = ThisWorkbook.Path & "\" & WJM & ".xls"
            Set Xlbook = XLapp.Workbooks.Open(MyFile)
            If Xlbook.Sheets(1).Name <> GZBM Then
                res = MsgBox("表格模板有问题啊!", vbOKOnly)
                Xlbook.Close True
                Set XLapp = Nothing
                Application.ScreenUpdating = True
                Application.DisplayAlerts = True
                Exit Sub
            End If
            If Application.Version > 11 Then 'execl 2003
                Xlbook.CheckCompatibility = False '该死的兼容性检查
            End If
            XLapp.Visible = False
            fday = mouths(n) & "-01"
            tmp_m() = Split(mouths(n), "-")
            If tmp_m(1) = "01" Or tmp_m(1) = "03" Or tmp_m(1) = "05" Or tmp_m(1) = "07" Or _
            tmp_m(1) = "08" Or tmp_m(1) = "10" Or tmp_m(1) = "12" Then
                lday = mouths(n) & "-31"
            ElseIf tmp_m(1) = "02" Then '2月份，看是不是润年
                If tmp_m(0) Mod 4 = 0 And tmp_m(0) Mod 100 <> 0 Then '闰年
                        lday = mouths(n) & "-29"
                Else
                    If tmp_m(0) Mod 100 = 0 And tmp_m(0) Mod 400 = 0 Then '闰年
                        lday = mouths(n) & "-29"
                    Else '非闰年
                        lday = mouths(n) & "-28"
                    End If
                End If
            Else '剩下的是小月了
                lday = mouths(n) & "-30"
            End If
            If n = 0 Then
                Do
                        max_row = max_row + 1
                        td = Xlbook.Sheets(GZBM).Range("G" & CStr(max_row))
                Loop Until td = ""
                fdate_rg = fdate_rg & CStr(max_row - 1)
                ldate_rg = ldate_rg & CStr(max_row - 1)
                jcfy_rg = jcfy_rg & CStr(max_row - 1)
            End If
            If max_row > 11 Then
                With Xlbook.Sheets(GZBM)
                    .Range("M3") = fday '把月头一天写入对应单元格
                    .Range("P3") = lday '把月未一天写入对应单元格
                    .Range(fdate_rg) = fday
                    .Range(ldate_rg) = lday
                End With
                If fday >= "2018-10-01" Then
                    With Xlbook.Sheets(GZBM)
                        .Range(jcfy_rg) = 5000 '修正减除费用
                    End With
                Else
                    With Xlbook.Sheets(GZBM)
                        .Range(jcfy_rg) = 3500 '修正减除费用
                    End With
                End If
            Else
                res = MsgBox("表格模板里面数据不完整!", vbOKOnly)
                Xlbook.Close True
                Set Xlbook = Nothing
                Set XLapp = Nothing
                Application.ScreenUpdating = True
                Application.DisplayAlerts = True
                Exit For
            End If
            Xlbook.Close True
            Set Xlbook = Nothing
            Set XLapp = Nothing
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
        Next n
    End If
End Sub

'根据第一个月与最后一个月生成每一个月的文件,以main_file为模板创建副本
'make_many_files("201504.xls", "2015-05", "2017-04")
Private Function make_many_files(main_file As String, first_mouth As String, last_mouth As String)
    Dim fso As Object
    Dim y As Integer '年
    Dim m As Integer '月
    Dim n As Integer
    Dim yf As Integer
    Dim yl As Integer
    Dim arr1() As String
    Dim arr2() As String

    arr1() = Split(first_mouth, "-")
    arr2() = Split(last_mouth, "-")
    n = 0
    j = Int(arr1(1))
    
    For i = Int(arr1(0)) To Int(arr2(0)) 'arr1(0)开始的年度，arr2(0)结束的年度
        Do
            If j < 10 Then
                mouths(n) = CStr(i) & "-0" & CStr(j)
            Else
                mouths(n) = CStr(i) & "-" & CStr(j)
            End If
            n = n + 1
            If i = Int(arr2(0)) And j = Int(arr2(1)) Then
                Exit Do
            End If
            j = j + 1
        Loop Until j > 12
        j = 1
    Next i
    
    Set fso = CreateObject("Scripting.FilesyStemObject")
    For n = 0 To 100
        If mouths(n) = "" Then
            Exit For
        End If
        fso.CopyFile main_file, pth & "\" & mouths(n) & ".xls"
    Next n
    Set fso = Nothing
End Function

Private Function is_illegal(first_mouth As String, last_mouth As String) As Integer
    Dim n1 As Long
    Dim n2 As Long
    Dim arr1() As String
    Dim arr2() As String
    arr1() = Split(first_mouth, "-")
    If Int(arr1(1)) > 12 Then arr1(1) = "12" '防止月份填错
    n1 = Int(arr1(0) & arr1(1))
    arr2() = Split(last_mouth, "-")
    If Int(arr2(1)) > 12 Then arr2(1) = "12"
    n2 = Int(arr2(0) & arr2(1))
    If n2 < n1 Or Int(arr2(0)) >= 2019 And Int(arr1(0)) <= 2018 Then
        is_illegal = 1
    Else
        is_illegal = 0
    End If
End Function
