Attribute VB_Name = "ģ��1"
'���ڶԸ���˰���������ɸ�������˰�����Ը���һ��ģ���ļ���ָ����
'�·���������Զ������걨��
'�䰮������Զ���˰��
' lianjie 2017-06-28��2018-01-19,2018-07-02,2018-10-29,2018-11-23
'2018-12-28,2019-03-05,2019-08-25


'ȫ�ֱ����������ļ������걨�·ݵĴ洢 ��ʽ��2015-05
Dim mouths(0 To 100) As String
Dim pth As String '��ǰ���е�·��


Sub main_fuck()

    '�������
    Dim MyFile As String, WJM As String, GZBM As String
    Dim XLapp As New Excel.Application
    Dim Xlbook As Excel.Workbook
    Dim m_file As String
    Dim tmp_m() As String
    Dim fmth As String
    Dim lmth As String
    Dim fday As String 'һ���µĿ�ͷһ��
    Dim lday As String 'һ���µ����һ��
    Dim n As Integer
    Dim fdate_rg As String
    Dim ldate_rg As String
    Dim jcfy_rg As String '��������
    Dim max_row As Integer
    Dim title As String
        
    '��������
    ChDir (ThisWorkbook.Path) '�л��ص�ǰĿ¼
    pth = ThisWorkbook.Path
    fdate_rg = "H11:H"
    ldate_rg = "I11:I"
    jcfy_rg = "Y11:Y"
    m_file = ""
    fmth = ""
    lmth = ""
    max_row = 10
    TABLETITLE = "�۽ɸ�������˰�걨���������ۺ�����Ԥ��Ԥ�ɣ�"
    
    m_file = Sheets("Sheet1").cells(2, 1)
    fmth = Sheets("Sheet1").cells(2, 2)
    lmth = Sheets("Sheet1").cells(2, 3)
    If (m_file = "" Or fmth = "" Or lmth = "") Then
       res = MsgBox("����д��ȷ�Ĳ���", vbOKOnly)
       Exit Sub
    End If
    If is_illegal(fmth, lmth) Then
       res = MsgBox("����д��ȷ�Ĳ���", vbOKOnly)
       Exit Sub
    End If
    
    If Dir(pth & "\" & m_file, vbDirectory) <> Empty Then
        '������һ���ļ�
        Call make_many_files(pth & "\" & m_file, fmth, lmth)
    Else
        res = MsgBox("ģ���ļ�������!", vbOKOnly)
        Exit Sub
    End If
    Application.DisplayAlerts = False 'ȡ�����Ǳ���ʱ����ʾ
    Application.ScreenUpdating = False 'ȡ����Ļˢ��
    WJM = m_file
    GZBM = "�۽ɸ�������˰�����"
    MyFile = ThisWorkbook.Path & "\" & WJM
    Set Xlbook = XLapp.Workbooks.Open(MyFile)
    title = Xlbook.Sheets(1).cells(1, 1)
    Xlbook.Close True
    Set XLapp = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    '2019-01-01���Ժ�ı��
    If title = TABLETITLE Then
        For n = 0 To 100
            If mouths(n) = "" Then
                Exit For
            End If
            '�й���Ϣ����
            Application.DisplayAlerts = False 'ȡ�����Ǳ���ʱ����ʾ
            Application.ScreenUpdating = False 'ȡ����Ļˢ��
            WJM = mouths(n)
            GZBM = "�۽ɸ�������˰�걨���������ۺ�����Ԥ��Ԥ�ɣ�"
            MyFile = ThisWorkbook.Path & "\" & WJM & ".xls"
            Set Xlbook = XLapp.Workbooks.Open(MyFile)
            If Xlbook.Sheets(1).Name <> GZBM Then
                res = MsgBox("���ģ�������Ⱑ!", vbOKOnly)
                Xlbook.Close True
                Set XLapp = Nothing
                Application.ScreenUpdating = True
                Application.DisplayAlerts = True
                Exit Sub
            End If
            If Application.Version > 11 Then 'execl 2003
                Xlbook.CheckCompatibility = False '�����ļ����Լ��
            End If
            XLapp.Visible = False
            fday = mouths(n) & "-01"
            tmp_m() = Split(mouths(n), "-")
            If tmp_m(1) = "01" Or tmp_m(1) = "03" Or tmp_m(1) = "05" Or tmp_m(1) = "07" Or _
            tmp_m(1) = "08" Or tmp_m(1) = "10" Or tmp_m(1) = "12" Then
                lday = mouths(n) & "-31"
            ElseIf tmp_m(1) = "02" Then '2�·ݣ����ǲ�������
                If tmp_m(0) Mod 4 = 0 And tmp_m(0) Mod 100 <> 0 Then '����
                    lday = mouths(n) & "-29"
                Else
                    If tmp_m(0) Mod 100 = 0 And tmp_m(0) Mod 400 = 0 Then '����
                        lday = mouths(n) & "-29"
                    Else '������
                        lday = mouths(n) & "-28"
                    End If
                End If
            Else 'ʣ�µ���С����
                    lday = mouths(n) & "-30"
            End If
            With Xlbook.Sheets(GZBM)
                .Range("M3") = fday '����ͷһ��д���Ӧ��Ԫ��
                .Range("R3") = lday '����δһ��д���Ӧ��Ԫ��
            End With
            Xlbook.Close True
            Set Xlbook = Nothing
            Set XLapp = Nothing
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
        Next n
    Else
        '����2019-01-01��ǰ�ı��
        For n = 0 To 100
            If mouths(n) = "" Then
                Exit For
            End If
            '�й���Ϣ����
            Application.DisplayAlerts = False 'ȡ�����Ǳ���ʱ����ʾ
            Application.ScreenUpdating = False 'ȡ����Ļˢ��
            WJM = mouths(n)
            GZBM = "�۽ɸ�������˰�����"
            MyFile = ThisWorkbook.Path & "\" & WJM & ".xls"
            Set Xlbook = XLapp.Workbooks.Open(MyFile)
            If Xlbook.Sheets(1).Name <> GZBM Then
                res = MsgBox("���ģ�������Ⱑ!", vbOKOnly)
                Xlbook.Close True
                Set XLapp = Nothing
                Application.ScreenUpdating = True
                Application.DisplayAlerts = True
                Exit Sub
            End If
            If Application.Version > 11 Then 'execl 2003
                Xlbook.CheckCompatibility = False '�����ļ����Լ��
            End If
            XLapp.Visible = False
            fday = mouths(n) & "-01"
            tmp_m() = Split(mouths(n), "-")
            If tmp_m(1) = "01" Or tmp_m(1) = "03" Or tmp_m(1) = "05" Or tmp_m(1) = "07" Or _
            tmp_m(1) = "08" Or tmp_m(1) = "10" Or tmp_m(1) = "12" Then
                lday = mouths(n) & "-31"
            ElseIf tmp_m(1) = "02" Then '2�·ݣ����ǲ�������
                If tmp_m(0) Mod 4 = 0 And tmp_m(0) Mod 100 <> 0 Then '����
                        lday = mouths(n) & "-29"
                Else
                    If tmp_m(0) Mod 100 = 0 And tmp_m(0) Mod 400 = 0 Then '����
                        lday = mouths(n) & "-29"
                    Else '������
                        lday = mouths(n) & "-28"
                    End If
                End If
            Else 'ʣ�µ���С����
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
                    .Range("M3") = fday '����ͷһ��д���Ӧ��Ԫ��
                    .Range("P3") = lday '����δһ��д���Ӧ��Ԫ��
                    .Range(fdate_rg) = fday
                    .Range(ldate_rg) = lday
                End With
                If fday >= "2018-10-01" Then
                    With Xlbook.Sheets(GZBM)
                        .Range(jcfy_rg) = 5000 '������������
                    End With
                Else
                    With Xlbook.Sheets(GZBM)
                        .Range(jcfy_rg) = 3500 '������������
                    End With
                End If
            Else
                res = MsgBox("���ģ���������ݲ�����!", vbOKOnly)
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

'���ݵ�һ���������һ��������ÿһ���µ��ļ�,��main_fileΪģ�崴������
'make_many_files("201504.xls", "2015-05", "2017-04")
Private Function make_many_files(main_file As String, first_mouth As String, last_mouth As String)
    Dim fso As Object
    Dim y As Integer '��
    Dim m As Integer '��
    Dim n As Integer
    Dim yf As Integer
    Dim yl As Integer
    Dim arr1() As String
    Dim arr2() As String

    arr1() = Split(first_mouth, "-")
    arr2() = Split(last_mouth, "-")
    n = 0
    j = Int(arr1(1))
    
    For i = Int(arr1(0)) To Int(arr2(0)) 'arr1(0)��ʼ����ȣ�arr2(0)���������
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
    If Int(arr1(1)) > 12 Then arr1(1) = "12" '��ֹ�·����
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
