
'0521_1回目のテスト内容

Attribute VB_Name = "shiwake_itaku"
Option Explicit

Sub �d�����L�����_���񂽂񌟕i_�ϑ�()

Application.ScreenUpdating = False

Dim imanobook As String
imanobook = ActiveWorkbook.Name

Dim imanosheet As String
imanosheet = ActiveSheet.Name

Workbooks(imanobook).Worksheets("�V�[��").Activate
    Range("A1:F20").ClearContents
Worksheets("���񂽂񌟕i").Activate

 Dim Sh As Worksheet
Set Sh = Worksheets("�V�[��")
Dim sh2 As Worksheet
Set sh2 = Worksheets("�V�[���ԍ�")

Dim orgPrinter As String


'Dim s_handan As Integer
'If Workbooks(imanobook).Worksheets("�ݒ�").Cells(11, 1) = "��ԕ\���o��" Then
's_handan = 1
'End If
'Dim teiban_handan As Integer

Dim a As Long
Dim i As Long


Dim hiduke As String
Dim kingaku As Long
Dim kingaku_kakutei As String
Dim namae As String
Dim kaisya_name As String
Dim target_kaisya As String

Dim lotno As String
Dim taitoru As String
Dim ichiba As String

Dim shouhinbango As Long
Dim shori As Long
Dim Path As String
Dim buf As String, Cnt As Long

Dim bango As Long

Dim shouhinbango_kakutei As String

Dim siwake_user As Long
Dim useridR As Long

Dim syouhin_itaku As Long
Dim houzin_kozinR As Long



Dim forudaNO As Long

Dim jougen As Long
Dim lotno_kakutei As String

Dim WshNetworkObject As Object
Set WshNetworkObject = CreateObject("WScript.Network")
Dim computamei As String
Dim com2 As String
Dim duplicate As Integer
Dim keikoku As String
Dim shouhinbango_full As Long
Dim user_id As String
Dim memo1 As String
Dim memo2 As String

Dim kingaku_kakutei_uchizei As String
Dim full_flag As Integer
Dim moto_kingaku As Long

Dim nen As String
    nen = Right(Year(Now), 2)
Dim tsuki As String
    If Month(Now) < 10 Then
    tsuki = "0" & Month(Now)
    Else
    tsuki = Month(Now)
    End If
Dim hi As String
    If Day(Now) < 10 Then
    hi = "0" & Day(Now)
    Else
    hi = Day(Now)
    End If
Dim ji As String
    If Hour(Now) < 10 Then
    ji = "0" & Day(Now)
    Else
    ji = Hour(Now)
    End If
Dim fun As String
    If Minute(Now) < 10 Then
    fun = "0" & Day(Now)
    Else
    fun = Minute(Now)
    End If
Dim byou As String
    If Second(Now) < 10 Then
    byou = "0" & Day(Now)
    Else
    byou = Second(Now)
    End If
    If Len(byou) = 3 Then
    byou = Left(byou, 2)
    End If
    
Dim sheetmei As String
Dim sheetmei_syurui As String
Dim fairumei As String
Dim xxxx As String
Dim fairumei2 As String
Dim keta As Long
Dim sheetmei_free As String
Dim sheetmei_bag As String


For siwake_user = 1 To Workbooks("�d�����p_�V�[���o��_�ϑ��p").Worksheets("���񂽂񌟕i").Cells(11, 1000).End(xlToLeft).Column

Select Case Workbooks("�d�����p_�V�[���o��_�ϑ��p").Worksheets("���񂽂񌟕i").Cells(1, siwake_user)

    Case "���[�U�[ID"
    useridR = siwake_user
    
    End Select

Next siwake_user
    
'���i�ꗗebay_�ϑ��p�A�ڋq���V�[�g
For syouhin_itaku = 1 To Workbooks("���i�ꗗebay_�ϑ��p").Worksheets("�ڋq���").Cells(11, 1000).End(xlToLeft).Column

Select Case Workbooks("���i�ꗗebay_�ϑ��p").Worksheets("�ڋq���").Cells(1, syouhin_itaku)

    Case "�@�lor�l"
    houzin_kozinR = syouhin_itaku

End Select

Next syouhin_itaku

    With WshNetworkObject
    computamei = .ComputerName
    End With
        
    com2 = Left(computamei, 1) & Right(computamei, 1)

    orgPrinter = Application.ActivePrinter
    
  On Error Resume Next
  
  If Workbooks("���i�ꗗebay_�ϑ��p") Is Nothing Then
MsgBox "���i�ꗗebay_�ϑ��p�@���J���Ă��������I"
GoTo L1
End If
 
    If Len(Worksheets("���񂽂񌟕i").Cells(2, 1)) = 0 Then
MsgBox "���z������܂���I"
GoTo L1
End If

    If Len(Worksheets("���񂽂񌟕i").Cells(12, 1)) = 0 Then
MsgBox "���[�U�[ID���擾���Ă��������I"
GoTo L1
End If
        
    forudaNO = forudaNO����(computamei)
    jougen = forudaNO + 49999
    
'    Path = "D:\JP Dropbox\�o�i�O�f�[�^\�o�b�O�ŐV" & forudaNO & "\"
'    buf = Dir(Path & "*")
'    Do While buf <> ""

    Path = "D:\JP Dropbox\�d��\�X�^�b�t_���D\�e�X�g�V�[��" & forudaNO & "\"
    buf = Dir(Path & "*")
    Do While buf <> ""
        
    bango = Replace(buf, ".csv", "")

    '�g�p���Ă����t�ԍ��܂ł��p�X��CSV�ɏo��
    '�e�X�g�̂��ߏC�����K�v
'    Path = "D:\JP Dropbox\�o�i�O�f�[�^\�o�b�O�ŐV" & forudaNO & "\"
'    Path = "C:\Users\jp_bu\Desktop\�e�X�g\�e�X�g�V�[��" & forudaNO & "\"
'    buf = Dir(Path & "*")
        
        If bango < jougen Then
        shouhinbango_full = bango + 1
        Else
        MsgBox ("�t�H���_�̏���ɒB���܂����B���̃t�H���_�ł�蒼���Ă��������B")
        GoTo L1
        End If
        
        buf = Dir()
        
    Loop
    
    
       duplicate = shouhinbango_duplicate_check(shouhinbango_full)
       
       If duplicate = 1 Then
       keikoku = MsgBox(shouhinbango_full & Chr(13) & Chr(13) & "���i�ԍ��̏d���̉\���������ł��B" & Chr(13) & Chr(13) & "everything��" & shouhinbango_full & "���`�F�b�N���Ă�������" & Chr(13) & _
         "���ɓ����ԍ����o�^����Ă�����d���ł��B" & Chr(13) & Chr(13) & "�l�b�g�����`�F�b�N���ĊǗ��҂ɕ񍐂��Ă��������B", vbCritical)
         Worksheets(imanosheet).Activate
       Exit Sub
       End If
       
       Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(19, 1) = shouhinbango_full
    
    If shouhinbango_full > 999999 Then
    shouhinbango_kakutei = shouhinbango_full
    Else
    shouhinbango_kakutei = "0" & shouhinbango_full
    End If
    
    '���z
    kingaku = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(2, 1)

    If kingaku = 0 Then
    kingaku_kakutei = "00000000"
    ElseIf kingaku < 10 Then
    kingaku_kakutei = "0000000" & kingaku
    ElseIf kingaku > 9 And kingaku < 100 Then
    kingaku_kakutei = "000000" & kingaku
    ElseIf kingaku > 99 And kingaku < 1000 Then
    kingaku_kakutei = "00000" & kingaku
    ElseIf kingaku > 999 And kingaku < 10000 Then
    kingaku_kakutei = "0000" & kingaku '00009999
    ElseIf kingaku > 9999 And kingaku < 100000 Then
    kingaku_kakutei = "000" & kingaku
    ElseIf kingaku > 99999 And kingaku < 1000000 Then
    kingaku_kakutei = "00" & kingaku
    ElseIf kingaku > 999999 And kingaku < 10000000 Then
    kingaku_kakutei = "0" & kingaku
    ElseIf kingaku > 9999999 And kingaku < 100000000 Then
    kingaku_kakutei = kingaku
    Else
    MsgBox ("�d�����z���������������I")
    End If
    
            
    If Len(kingaku) = 0 Then
    kingaku_kakutei = "00000000"
    End If
    
    '���i�ꗗebay_�ϑ��p�Ŏ擾����(�@�lor�l)
    For i = 1 To Workbooks("���i�ꗗebay_�ϑ��p").workshhets("�ڋq���").Cells(Rows.Count, 1).End(xlUp).Row
        
    
    
    
    
    namae = shouhinbango_kakutei & "_" & kingaku_kakutei & "_" & computamei



        '����1
        memo1 = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(8, 1)

        '����2
        memo2 = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(9, 1)

        '���[�U�[ID
        user_id = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(12, 1)

        '���i�ԍ�
'        shouhinbango_full = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(19, 1)

    
    
        Workbooks(imanobook).Worksheets("�V�[���ԍ�").Cells(1, 1) = shouhinbango_full
    Call sh2.PrintOut(ActivePrinter:="Brother QL-800")
    
    Workbooks(imanobook).Worksheets("�V�[���ԍ�").Select
    Range("A1:F20").ClearContents
    
    
'---�V�[��
    '1�s�ڋ�
    '���i�ԍ�
    Workbooks(imanobook).Worksheets("�V�[��").Cells(2, 1) = shouhinbango_full
    '�o�[�R�[�h
    Workbooks(imanobook).Worksheets("�V�[��").Cells(3, 1) = "*" & shouhinbango_full & "*"
    '���z
    Workbooks(imanobook).Worksheets("�V�[��").Cells(4, 1) = kingaku & "�~"
    '���[�U�[ID
    Workbooks(imanobook).Worksheets("�V�[��").Cells(4, 2) = user_id
    
    
    '��Ж�
    If Len(Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(12, 1)) <> 0 Then
        target_kaisya = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(12, 1)
        
        For i = 2 To Workbooks("���i�ꗗebay_�ϑ��p").Worksheets("�ڋq���").Cells(Rows.Count, 1).End(xlUp).Row
            If Workbooks("���i�ꗗebay_�ϑ��p").Worksheets("�ڋq���").Cells(i, 1) = target_kaisya Then
                kaisya_name = Workbooks("���i�ꗗebay_�ϑ��p").Worksheets("�ڋq���").Cells(i, 2)
                Exit For
            End If
        Next i
        
    
    Workbooks(imanobook).Worksheets("�V�[��").Cells(4, 3) = kaisya_name
    
    End If
    
    
    '����1
    Workbooks(imanobook).Worksheets("�V�[��").Cells(6, 1) = memo1
    '����2
    Workbooks(imanobook).Worksheets("�V�[��").Cells(7, 1) = memo2
    
    
    
    Call Sh.PrintOut(ActivePrinter:="Brother QL-800")
    
    
    Workbooks(imanobook).Worksheets("�V�[��").Select
    Range("A2:F20").ClearContents
'
'         Workbooks(imanobook).Worksheets("�V�[���ԍ�").Cells(1, 1) = shouhinbango_full
'    Call Sh2.PrintOut(ActivePrinter:="Brother QL-800")
'
'    Workbooks(imanobook).Worksheets("�V�[���ԍ�").Select
'    Range("A1:F20").ClearContents
        
             
    '�e�X�g�p
'    Workbooks(imanobook).Sheets("csv").Select
'    Sheets("csv").Copy
'    ChDir "D:\JP Dropbox\�o�i�O�f�[�^\���i�ꗗ�o�^�p"
'    ActiveWorkbook.SaveAs Filename:="D:\JP Dropbox\�o�i�O�f�[�^\���i�ꗗ�o�^�p\" & namae & ".csv", FileFormat:=xlCSV, _
'        CreateBackup:=False
'
'    Workbooks(namae).Close False

    Workbooks(imanobook).Sheets("csv").Select
    Sheets("csv").Copy
    ChDir "D:\JP Dropbox\�d��\�X�^�b�t_���D\�e�X�g�V�[��\���i�ꗗ�o�^�p�e�X�g"
    ActiveWorkbook.SaveAs Filename:="D:\JP Dropbox\�d��\�X�^�b�t_���D\�e�X�g�V�[��\���i�ꗗ�o�^�p�e�X�g\" & namae & ".csv", FileFormat:=xlCSV, _
        CreateBackup:=False

    Workbooks(namae).Close False
    
    
    Workbooks(imanobook).Worksheets("���񂽂񌟕i").Activate
    Range("A2:A10").ClearContents
    Range("A12").ClearContents
    
    '�e�X�g�p
'    Name "D:\JP Dropbox\�o�i�O�f�[�^\�o�b�O�ŐV" & forudaNO & "\" & bango & ".csv" As "D:\JP Dropbox\�o�i�O�f�[�^\�o�b�O�ŐV" & forudaNO & "\" & bango + 1 & ".csv"
    Name "D:\JP Dropbox\�d��\�X�^�b�t_���D\�e�X�g�V�[��" & forudaNO & "\" & bango & ".csv" As "D:\JP Dropbox\�d��\�X�^�b�t_���D\�e�X�g�V�[��" & forudaNO & "\" & bango + 1 & ".csv"
        
        
    Workbooks(imanobook).Worksheets("���񂽂񌟕i").Activate
    Worksheets("���񂽂񌟕i").Cells(1, 1).Select
    
    


    '�A���t�H���_�Ƀf�[�^���A�E�g����
    
    fairumei = Replace(Workbooks("���i�ꗗebay_�ϑ��p").Name, ".xlsx", "")
    
    If shouhinbango_full > 999999 Then
        keta = Left(shouhinbango_full, 1) * 100
        shouhinbango = shouhinbango_full - (keta * 10000)
        Else
        keta = ""
        shouhinbango = shouhinbango_full
        End If
        
        sheetmei_free = "N�t���[" & keta
        sheetmei_bag = "�o�b�O" & keta

    sheetmei_syurui = Left(sheetmei_free, Len(sheetmei_free) - 3)
    
    fairumei2 = nen & tsuki & hi & ji & fun & byou & "_itakuID_" & "�d�����V�[���f�[�^"
    '�e�X�g�p�̂��ߏC�����K�v
'    Workbooks.Add.SaveAs Filename:="D:\JP Dropbox\�A�E�g�C��\�A��\" & fairumei2 & ".xlsx"
    Workbooks.Add.SaveAs Filename:="C:\Users\jp_bu\Desktop\�e�X�g\�e�X�g\" & fairumei2 & ".xlsx"

    Workbooks(fairumei2).Worksheets("Sheet1").Cells(1, 1) = shouhinbango_full '���i�ԍ�
    Workbooks(fairumei2).Worksheets("Sheet1").Cells(1, 2) = sheetmei_free '�V�[�g��
    Workbooks(fairumei2).Worksheets("Sheet1").Cells(1, 3) = shouhinbango  '�s��
    Workbooks(fairumei2).Worksheets("Sheet1").Cells(1, 4) = sheetmei_syurui '���
    Workbooks(fairumei2).Worksheets("Sheet1").Cells(1, 5) = user_id '���[�U�[ID
    
    Workbooks(fairumei2).Close True
    
    
    MsgBox "�f�[�^���A�E�g���܂����B"


    
    '�����v�Z
    Application.ScreenUpdating = True
    
L1:

Workbooks(imanobook).Worksheets("���񂽂񌟕i").Activate
    
    
    
End Sub




Sub �d�����L�����_�A�����i_�ϑ��p()


    Application.ScreenUpdating = False
    

Dim imanobook As String
imanobook = ActiveWorkbook.Name

Dim imanosheet As String
imanosheet = ActiveSheet.Name

Dim Sh As Worksheet
Set Sh = Worksheets("�V�[��")
Dim sh2 As Worksheet
Set sh2 = Worksheets("�V�[���ԍ�")

Dim orgPrinter As String
    
Dim i As Long
Dim x As Long

Dim gyou As Long
Dim cont As Long

Dim namae As String

Dim kingaku As Long

Dim shouhinbango As Long
Dim shori As Long
Dim Path As String
Dim buf As String, Cnt As Long

Dim bango As Long
Dim shouhinbango_kakutei As String
Dim kingaku_kakutei As String


Dim forudaNO As Long

Dim clear_handan As Integer
Dim jougen As Long

Dim WshNetworkObject As Object
Set WshNetworkObject = CreateObject("WScript.Network")
Dim computamei As String
Dim com2 As String
Dim duplicate As Integer
Dim keikoku As String
Dim shouhinbango_full As Long

Dim shori_handan As Variant

Dim shori_houhou As Integer
Dim shori_handan2 As Variant
Dim handan1 As Variant

Dim kingaku_kakutei_uchizei As String
Dim moto_kingaku As Long

Dim user_id As String
Dim memo1 As String
Dim memo2 As String

Dim target_kaisya As String

Dim kaisya_name As String

    
'    If Len(Workbooks(imanobook).Worksheets("�ݒ�").Cells(12, 1)) <> 0 Then
'    uchizei = 1 ''''���Ńt���O
'    End If
    
    
    shori_handan = InputBox("���v���z" & Chr(10) & Chr(10) & WorksheetFunction.Sum(Range("A26:A100")) * 100 & "�~" _
    & Chr(10) & Chr(10) & "1  �V�[���o��" & Chr(10) & "2  �V�[���o��(�ԍ��V�[���Ȃ��j" & Chr(10) & "3  ���v���z��������߂�", Default:=1)
    
'    If shori_handan = 3 Then
'    GoTo L1
'    End If
'
'    If Len(shori_handan) = 0 Then
'    GoTo L1
'    End If
    
    gyou = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(10000, 1).End(xlUp).Row
    
    If gyou > 100 Then
    MsgBox ("�s�������������ł��B��蒼���Ă��������B")
        GoTo L1
    End If
    
    For x = 26 To gyou ''''''''''�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@''''''''''''''''''��U�S���`�F�b�N
    If Len(Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(x, 1)) = 0 Then
    MsgBox (x & "�@�s�ڂɋ��z�������Ă��܂���B")
    GoTo L1
    End If
    

    Next x
    
    
    
    With WshNetworkObject
    computamei = .ComputerName
    End With
        
    com2 = Left(computamei, 1) & Right(computamei, 1)
    
    forudaNO = forudaNO����(computamei)
    jougen = forudaNO + 49999
    
    
'    If shori_handan = 2 Then
'    kingaku = InputBox("���z�́H�i�����āj")
'    kingaku = kingaku * 100
'    maisuu = InputBox("�����́H", Default:=2)
'    gyou = maisuu + 17
'    End If
    
     '�e�X�g�p�̂��ߏC�����K�v
'    Path = "D:\JP Dropbox\�o�i�O�f�[�^\�o�b�O�ŐV" & forudaNO & "\"
'    buf = Dir(Path & "*")
'    Do While buf <> ""
    
    Path = "D:\JP Dropbox\�d��\�X�^�b�t_���D\�e�X�g�V�[��" & forudaNO & "\"
    buf = Dir(Path & "*")
    Do While buf <> ""
        
        bango = Replace(buf, ".csv", "")
        If bango < jougen Then
        shouhinbango_full = bango + 1
        Else
        MsgBox ("�t�H���_�̏���ɒB���܂����B���̃t�H���_�ł�蒼���Ă��������B")
        GoTo L1
        
        End If
                
        buf = Dir()
    Loop
    
    
    For x = 26 To gyou
    
    
        If shouhinbango_full > 999999 Then
    shouhinbango_kakutei = shouhinbango_full
    Else
    shouhinbango_kakutei = "0" & shouhinbango_full
    End If
    
        If x = 26 Then
        
       duplicate = shouhinbango_duplicate_check(shouhinbango_full)
       
       If duplicate = 1 Then
       keikoku = MsgBox(shouhinbango_full & Chr(13) & Chr(13) & "���i�ԍ��̏d���̉\���������ł��B" & Chr(13) & Chr(13) & "everything��" & shouhinbango_full & "���`�F�b�N���Ă�������" & Chr(13) & _
         "���ɓ����ԍ����o�^����Ă�����d���ł��B" & Chr(13) & Chr(13) & "�l�b�g�����`�F�b�N���ĊǗ��҂ɕ񍐂��Ă��������B", vbCritical)
       Exit Sub
       End If
        
        End If
            
'    If Len(Workbooks(imanobook).Worksheets("�ݒ�").Cells(8, 1)) <> 0 Then
'    hitocode3 = Workbooks(imanobook).Worksheets("�ݒ�").Cells(8, 1)
'    Else
'    MsgBox "�o�C���[�������͂���Ă��܂���B" & Chr(10) & "��U�I�����܂��B"
'    GoTo L1
'    End If
'
'    buyerNO_kakutei = �o�C���[�R�[�h����2(imanobook, hitocode3)
'    buyer_kakutei = hitocode3
    
    
    
'    If uchizei = 1 Then
    
'    If shori_houhou = 1 Then
'    kingaku = Round((Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(x, 1) / shouhizei), 0)
'    moto_kingaku = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(x, 1)
'    Else
'    kingaku = Round(((Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(x, 1) * 100) / shouhizei), 0)
'    moto_kingaku = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(x, 1) * 100
'    End If
'
'    Else
'
'    If shori_houhou = 1 Then
'    kingaku = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(x, 1)
'
'    Else
'    kingaku = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(x, 1) * 100
'    End If
'
'    End If
    
    If kingaku = 0 Then
    kingaku_kakutei = "00000000"
    ElseIf kingaku < 10 Then
    kingaku_kakutei = "0000000" & kingaku
    ElseIf kingaku > 9 And kingaku < 100 Then
    kingaku_kakutei = "000000" & kingaku
    ElseIf kingaku > 99 And kingaku < 1000 Then
    kingaku_kakutei = "00000" & kingaku
    ElseIf kingaku > 999 And kingaku < 10000 Then
    kingaku_kakutei = "0000" & kingaku '00009999
    ElseIf kingaku > 9999 And kingaku < 100000 Then
    kingaku_kakutei = "000" & kingaku
    ElseIf kingaku > 99999 And kingaku < 1000000 Then
    kingaku_kakutei = "00" & kingaku
    ElseIf kingaku > 999999 And kingaku < 10000000 Then
    kingaku_kakutei = "0" & kingaku
    ElseIf kingaku > 9999999 And kingaku < 100000000 Then
    kingaku_kakutei = kingaku
    Else
    MsgBox ("�d�����z��������������")
    End If
    
            
    If Len(kingaku) = 0 Then
    kingaku_kakutei = "00000000"
    End If
    
    namae = shouhinbango_kakutei & "_" & kingaku_kakutei & com2 & "_" & computamei
    
'    If Len(kingaku_kakutei) <> 8 Then
'    MsgBox ("�d�����z�����������ł��B��蒼���Ă��������B")
'    GoTo L1
'    End If
'
'    teigaku_kakutei = "0000"
    
        
'    If Len(Workbooks(imanobook).Worksheets("�ݒ�").Cells(9, 1)) <> 0 Then
'    ichiba = Workbooks(imanobook).Worksheets("�ݒ�").Cells(9, 1)
'    Else
'    ichiba = "xxx"
'    End If
'
'    If Len(ichiba) <> 3 Then
'    ichiba = "xxx"
'    End If
    
'    If Len(Workbooks(imanobook).Worksheets("�ݒ�").Cells(10, 1)) <> 0 Then
'    hiduke = Workbooks(imanobook).Worksheets("�ݒ�").Cells(10, 1)
'    Else
'    hiduke = "19000101"
'    End If
'
'    If Len(Workbooks(imanobook).Worksheets("�ݒ�").Cells(10, 1)) <> 8 Then
'    hiduke = "19000101"
'    End If
'
'    If Len(buyerNO_kakutei) = 0 Then
'    buyerNO_kakutei = "J"
'    MsgBox ("�o�C���[�R�[�h������܂���BJ�ɂ��Ă����܂��B")
'    End If
         
    
'    namae = shouhinbango_kakutei & buyerNO_kakutei & kingaku_kakutei & ichiba & teigaku_kakutei & com2 & hiduke & "_" & computamei
'    namae = shouhinbango_kakutei & "_" & buyerNO_kakutei & "_" & kingaku_kakutei & "_" & ichiba & "_" & teigaku_kakutei & "_" & com2 & "_" & hiduke & "_" & computamei
'    namae2 = shouhinbango
'    namae3 = shouhinbango_kakutei & buyer_kakutei & kingaku_kakutei
    
            
'---�V�[��

    Workbooks(imanobook).Worksheets("�V�[���ԍ�").Cells(1, 1) = shouhinbango_full
    '1�s�ڋ�
    '���i�ԍ�
    Workbooks(imanobook).Worksheets("�V�[��").Cells(2, 1) = shouhinbango_full
    '�o�[�R�[�h
    Workbooks(imanobook).Worksheets("�V�[��").Cells(3, 1) = "*" & shouhinbango_full & "*"
    '���z
    Workbooks(imanobook).Worksheets("�V�[��").Cells(4, 1) = kingaku & "�~"
    '���[�U�[ID
    Workbooks(imanobook).Worksheets("�V�[��").Cells(4, 2) = user_id
    
    
    '��Ж�
    If Len(Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(12, 1)) <> 0 Then
        target_kaisya = Workbooks(imanobook).Worksheets("���񂽂񌟕i").Cells(12, 1)
        
        For i = 2 To Workbooks("���i�ꗗebay_�ϑ��p").Worksheets("�ڋq���").Cells(Rows.Count, 1).End(xlUp).Row
            If Workbooks("���i�ꗗebay_�ϑ��p").Worksheets("�ڋq���").Cells(i, 1) = target_kaisya Then
                kaisya_name = Workbooks("���i�ꗗebay_�ϑ��p").Worksheets("�ڋq���").Cells(i, 2)
                Exit For
            End If
        Next i
        
    
    Workbooks(imanobook).Worksheets("�V�[��").Cells(4, 3) = kaisya_name
    
    End If
    
    
    '����1
    Workbooks(imanobook).Worksheets("�V�[��").Cells(6, 1) = memo1
    '����2
    Workbooks(imanobook).Worksheets("�V�[��").Cells(7, 1) = memo2

'    Workbooks(imanobook).Worksheets("�V�[���ԍ�").Cells(1, 1) = shouhinbango_full
'    Workbooks(imanobook).Worksheets("�V�[��").Cells(2, 1) = shouhinbango_full
'    Workbooks(imanobook).Worksheets("�V�[��").Cells(3, 1) = "*" & shouhinbango_full & "*"
'
'    Workbooks(imanobook).Worksheets("�V�[��").Cells(4, 1) = kingaku & "�~"
'    Workbooks(imanobook).Worksheets("�V�[��").Cells(4, 2) = teigaku
'    Workbooks(imanobook).Worksheets("�V�[��").Cells(4, 3) = buyer_kakutei
'
'    Workbooks(imanobook).Worksheets("�V�[��").Cells(5, 1) = hiduke '���t
'    If uchizei = 1 Then
'    Workbooks(imanobook).Worksheets("�V�[��").Cells(4, 1) = kingaku & "�~ (" & moto_kingaku & "�~)"
'    Workbooks(imanobook).Worksheets("�V�[��").Cells(5, 3) = ichiba & "���Ŕ���" '�s��
'    Else
'    Workbooks(imanobook).Worksheets("�V�[��").Cells(5, 3) = ichiba '�s��
'    End If
'    Workbooks(imanobook).Worksheets("�V�[��").Cells(6, 1) = lotno_kakutei
'    Workbooks(imanobook).Worksheets("�V�[��").Cells(7, 1) = sonota
    
    Call Sh.PrintOut(ActivePrinter:="Brother QL-800")
    
    Workbooks(imanobook).Worksheets("�V�[��").Select
    Range("A1:F20").ClearContents
    
    If shori_handan = 1 Then
    Call sh2.PrintOut(ActivePrinter:="Brother QL-800")
    
    Workbooks(imanobook).Worksheets("�V�[���ԍ�").Select
    Range("A1:F20").ClearContents
    End If
    
    
     '�e�X�g�p�̂��ߏC�����K�v
'    Workbooks(imanobook).Sheets("csv").Select
'    Sheets("csv").Copy
'    ChDir "D:\JP Dropbox\�o�i�O�f�[�^\���i�ꗗ�o�^�p"
'    ActiveWorkbook.SaveAs Filename:="D:\JP Dropbox\�o�i�O�f�[�^\���i�ꗗ�o�^�p\" & namae & ".csv", FileFormat:=xlCSV, _
'        CreateBackup:=False

    Workbooks(imanobook).Sheets("csv").Select
    Sheets("csv").Copy
    ChDir "D:\JP Dropbox\�d��\�X�^�b�t_���D\�e�X�g�V�[��\���i�ꗗ�o�^�p�e�X�g"
    ActiveWorkbook.SaveAs Filename:="D:\JP Dropbox\�d��\�X�^�b�t_���D\�e�X�g\�e�X�g�V�[��\���i�ꗗ�o�^�p�e�X�g\" & namae & ".csv", FileFormat:=xlCSV, _
        CreateBackup:=False

            
    Workbooks(namae).Close False
    
    shouhinbango_full = shouhinbango_full + 1
    cont = cont + 1
    
    Next x
    
    '�e�X�g�p
'    Name "D:\JP Dropbox\�o�i�O�f�[�^\�o�b�O�ŐV" & forudaNO & "\" & bango & ".csv" As "D:\JP Dropbox\�o�i�O�f�[�^\�o�b�O�ŐV" & forudaNO & "\" & bango + cont & ".csv"
     Name "D:\JP Dropbox\�d��\�X�^�b�t_���D\�e�X�g�V�[��" & forudaNO & "\" & bango & ".csv" As "D:\JP Dropbox\�d��\�X�^�b�t_���D\�e�X�g�V�[��" & forudaNO & "\" & bango + 1 & ".csv"

    
    
    clear_handan = InputBox("1  �N���A����" & Chr(10) & "2  �N���A���Ȃ�", Default:=1)
    
    If clear_handan = 1 Then
    Call �N���A�d����_���񂽂񌟕i
    End If
            
    Workbooks(imanobook).Worksheets("���񂽂񌟕i").Activate
    Worksheets("���񂽂񌟕i").Cells(1, 1).Select
    

  '�����v�Z
Application.ScreenUpdating = True


    
L1:

Workbooks(imanobook).Worksheets("���񂽂񌟕i").Activate
    
    
    
End Sub


Function shouhinbango_duplicate_check_shiwake_set(ByVal shouhinbango_kakutei_suuji As Long) As Integer


Dim shouhinbango_kakutei As String
Dim bango_moto As String
Dim bango As String
Dim Path As String
Dim buf As String
Dim flag As Integer
Dim shouhinbango_full As Long
Dim imanobook As String
imanobook = ActiveWorkbook.Name

shouhinbango_full = shouhinbango_kakutei_suuji
    If shouhinbango_full < 1000000 Then
   shouhinbango_kakutei = "0" & shouhinbango_full
   Else
   shouhinbango_kakutei = shouhinbango_full
   End If
      
    
    If flag <> 1 Then
    
    If Workbooks(imanobook).Worksheets("�ݒ�").Cells(30, 1) = shouhinbango_full Then
    flag = 1
    End If
          
    Path = "D:\JP Dropbox\�o�i�O�f�[�^\���i�ꗗ�o�^�p_�ς�\"
    buf = Dir(Path & "*")
    
    Do While buf <> ""
        
        bango_moto = Replace(buf, ".csv", "")
        bango = Left(bango_moto, 7)
        
        If shouhinbango_kakutei = bango Then
        shouhinbango_duplicate_check_shiwake_set = 1
        flag = 1
        Exit Do
        End If
        
        buf = Dir()
    Loop
    
    End If
    
    If flag <> 1 Then
    Path = "D:\JP Dropbox\�o�i�O�f�[�^\���i�ꗗ�o�^�p\"
    buf = Dir(Path & "*")
    
    Do While buf <> ""
        
        bango_moto = Replace(buf, ".csv", "")
        bango = Left(bango_moto, 7)
        
        If shouhinbango_kakutei = bango Then
        shouhinbango_duplicate_check_shiwake_set = 1
        Exit Do
        End If
        
        buf = Dir()
    Loop
    
    End If
    

End Function

