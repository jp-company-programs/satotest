
'0521_1回目のテスト内容

'0521_2回目のテスト内容

Attribute VB_Name = "shiwake_itaku"
Option Explicit

Sub 仕分け記入印刷_かんたん検品_委託()

Application.ScreenUpdating = False

Dim imanobook As String
imanobook = ActiveWorkbook.Name

Dim imanosheet As String
imanosheet = ActiveSheet.Name

Workbooks(imanobook).Worksheets("シール").Activate
    Range("A1:F20").ClearContents
Worksheets("かんたん検品").Activate

 Dim Sh As Worksheet
Set Sh = Worksheets("シール")
Dim sh2 As Worksheet
Set sh2 = Worksheets("シール番号")

Dim orgPrinter As String


'Dim s_handan As Integer
'If Workbooks(imanobook).Worksheets("設定").Cells(11, 1) = "定番表を出す" Then
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


For siwake_user = 1 To Workbooks("仕分け用_シール出し_委託用").Worksheets("かんたん検品").Cells(11, 1000).End(xlToLeft).Column

Select Case Workbooks("仕分け用_シール出し_委託用").Worksheets("かんたん検品").Cells(1, siwake_user)

    Case "ユーザーID"
    useridR = siwake_user
    
    End Select

Next siwake_user
    
'商品一覧ebay_委託用、顧客情報シート
For syouhin_itaku = 1 To Workbooks("商品一覧ebay_委託用").Worksheets("顧客情報").Cells(11, 1000).End(xlToLeft).Column

Select Case Workbooks("商品一覧ebay_委託用").Worksheets("顧客情報").Cells(1, syouhin_itaku)

    Case "法人or個人"
    houzin_kozinR = syouhin_itaku

End Select

Next syouhin_itaku

    With WshNetworkObject
    computamei = .ComputerName
    End With
        
    com2 = Left(computamei, 1) & Right(computamei, 1)

    orgPrinter = Application.ActivePrinter
    
  On Error Resume Next
  
  If Workbooks("商品一覧ebay_委託用") Is Nothing Then
MsgBox "商品一覧ebay_委託用　を開いてください！"
GoTo L1
End If
 
    If Len(Worksheets("かんたん検品").Cells(2, 1)) = 0 Then
MsgBox "金額がありません！"
GoTo L1
End If

    If Len(Worksheets("かんたん検品").Cells(12, 1)) = 0 Then
MsgBox "ユーザーIDを取得してください！"
GoTo L1
End If
        
    forudaNO = forudaNOつける(computamei)
    jougen = forudaNO + 49999
    
'    Path = "D:\JP Dropbox\出品前データ\バッグ最新" & forudaNO & "\"
'    buf = Dir(Path & "*")
'    Do While buf <> ""

    Path = "D:\JP Dropbox\仕事\スタッフ_香織\テストシール" & forudaNO & "\"
    buf = Dir(Path & "*")
    Do While buf <> ""
        
    bango = Replace(buf, ".csv", "")

    '使用している受付番号までをパスのCSVに出力
    'テストのため修正が必要
'    Path = "D:\JP Dropbox\出品前データ\バッグ最新" & forudaNO & "\"
'    Path = "C:\Users\jp_bu\Desktop\テスト\テストシール" & forudaNO & "\"
'    buf = Dir(Path & "*")
        
        If bango < jougen Then
        shouhinbango_full = bango + 1
        Else
        MsgBox ("フォルダの上限に達しました。他のフォルダでやり直してください。")
        GoTo L1
        End If
        
        buf = Dir()
        
    Loop
    
    
       duplicate = shouhinbango_duplicate_check(shouhinbango_full)
       
       If duplicate = 1 Then
       keikoku = MsgBox(shouhinbango_full & Chr(13) & Chr(13) & "商品番号の重複の可能性が高いです。" & Chr(13) & Chr(13) & "everythingで" & shouhinbango_full & "をチェックしてください" & Chr(13) & _
         "既に同じ番号が登録されていたら重複です。" & Chr(13) & Chr(13) & "ネット環境をチェックして管理者に報告してください。", vbCritical)
         Worksheets(imanosheet).Activate
       Exit Sub
       End If
       
       Workbooks(imanobook).Worksheets("かんたん検品").Cells(19, 1) = shouhinbango_full
    
    If shouhinbango_full > 999999 Then
    shouhinbango_kakutei = shouhinbango_full
    Else
    shouhinbango_kakutei = "0" & shouhinbango_full
    End If
    
    '金額
    kingaku = Workbooks(imanobook).Worksheets("かんたん検品").Cells(2, 1)

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
    MsgBox ("仕入金額がおかしいかも！")
    End If
    
            
    If Len(kingaku) = 0 Then
    kingaku_kakutei = "00000000"
    End If
    
    '商品一覧ebay_委託用で取得する(法人or個人)
    For i = 1 To Workbooks("商品一覧ebay_委託用").workshhets("顧客情報").Cells(Rows.Count, 1).End(xlUp).Row
        
    
    
    
    
    namae = shouhinbango_kakutei & "_" & kingaku_kakutei & "_" & computamei



        'メモ1
        memo1 = Workbooks(imanobook).Worksheets("かんたん検品").Cells(8, 1)

        'メモ2
        memo2 = Workbooks(imanobook).Worksheets("かんたん検品").Cells(9, 1)

        'ユーザーID
        user_id = Workbooks(imanobook).Worksheets("かんたん検品").Cells(12, 1)

        '商品番号
'        shouhinbango_full = Workbooks(imanobook).Worksheets("かんたん検品").Cells(19, 1)

    
    
        Workbooks(imanobook).Worksheets("シール番号").Cells(1, 1) = shouhinbango_full
    Call sh2.PrintOut(ActivePrinter:="Brother QL-800")
    
    Workbooks(imanobook).Worksheets("シール番号").Select
    Range("A1:F20").ClearContents
    
    
'---シール
    '1行目空白
    '商品番号
    Workbooks(imanobook).Worksheets("シール").Cells(2, 1) = shouhinbango_full
    'バーコード
    Workbooks(imanobook).Worksheets("シール").Cells(3, 1) = "*" & shouhinbango_full & "*"
    '金額
    Workbooks(imanobook).Worksheets("シール").Cells(4, 1) = kingaku & "円"
    'ユーザーID
    Workbooks(imanobook).Worksheets("シール").Cells(4, 2) = user_id
    
    
    '会社名
    If Len(Workbooks(imanobook).Worksheets("かんたん検品").Cells(12, 1)) <> 0 Then
        target_kaisya = Workbooks(imanobook).Worksheets("かんたん検品").Cells(12, 1)
        
        For i = 2 To Workbooks("商品一覧ebay_委託用").Worksheets("顧客情報").Cells(Rows.Count, 1).End(xlUp).Row
            If Workbooks("商品一覧ebay_委託用").Worksheets("顧客情報").Cells(i, 1) = target_kaisya Then
                kaisya_name = Workbooks("商品一覧ebay_委託用").Worksheets("顧客情報").Cells(i, 2)
                Exit For
            End If
        Next i
        
    
    Workbooks(imanobook).Worksheets("シール").Cells(4, 3) = kaisya_name
    
    End If
    
    
    'メモ1
    Workbooks(imanobook).Worksheets("シール").Cells(6, 1) = memo1
    'メモ2
    Workbooks(imanobook).Worksheets("シール").Cells(7, 1) = memo2
    
    
    
    Call Sh.PrintOut(ActivePrinter:="Brother QL-800")
    
    
    Workbooks(imanobook).Worksheets("シール").Select
    Range("A2:F20").ClearContents
'
'         Workbooks(imanobook).Worksheets("シール番号").Cells(1, 1) = shouhinbango_full
'    Call Sh2.PrintOut(ActivePrinter:="Brother QL-800")
'
'    Workbooks(imanobook).Worksheets("シール番号").Select
'    Range("A1:F20").ClearContents
        
             
    'テスト用
'    Workbooks(imanobook).Sheets("csv").Select
'    Sheets("csv").Copy
'    ChDir "D:\JP Dropbox\出品前データ\商品一覧登録用"
'    ActiveWorkbook.SaveAs Filename:="D:\JP Dropbox\出品前データ\商品一覧登録用\" & namae & ".csv", FileFormat:=xlCSV, _
'        CreateBackup:=False
'
'    Workbooks(namae).Close False

    Workbooks(imanobook).Sheets("csv").Select
    Sheets("csv").Copy
    ChDir "D:\JP Dropbox\仕事\スタッフ_香織\テストシール\商品一覧登録用テスト"
    ActiveWorkbook.SaveAs Filename:="D:\JP Dropbox\仕事\スタッフ_香織\テストシール\商品一覧登録用テスト\" & namae & ".csv", FileFormat:=xlCSV, _
        CreateBackup:=False

    Workbooks(namae).Close False
    
    
    Workbooks(imanobook).Worksheets("かんたん検品").Activate
    Range("A2:A10").ClearContents
    Range("A12").ClearContents
    
    'テスト用
'    Name "D:\JP Dropbox\出品前データ\バッグ最新" & forudaNO & "\" & bango & ".csv" As "D:\JP Dropbox\出品前データ\バッグ最新" & forudaNO & "\" & bango + 1 & ".csv"
    Name "D:\JP Dropbox\仕事\スタッフ_香織\テストシール" & forudaNO & "\" & bango & ".csv" As "D:\JP Dropbox\仕事\スタッフ_香織\テストシール" & forudaNO & "\" & bango + 1 & ".csv"
        
        
    Workbooks(imanobook).Worksheets("かんたん検品").Activate
    Worksheets("かんたん検品").Cells(1, 1).Select
    
    


    '連続フォルダにデータをアウトする
    
    fairumei = Replace(Workbooks("商品一覧ebay_委託用").Name, ".xlsx", "")
    
    If shouhinbango_full > 999999 Then
        keta = Left(shouhinbango_full, 1) * 100
        shouhinbango = shouhinbango_full - (keta * 10000)
        Else
        keta = ""
        shouhinbango = shouhinbango_full
        End If
        
        sheetmei_free = "Nフリー" & keta
        sheetmei_bag = "バッグ" & keta

    sheetmei_syurui = Left(sheetmei_free, Len(sheetmei_free) - 3)
    
    fairumei2 = nen & tsuki & hi & ji & fun & byou & "_itakuID_" & "仕分けシールデータ"
    'テスト用のため修正が必要
'    Workbooks.Add.SaveAs Filename:="D:\JP Dropbox\アウトイン\連続\" & fairumei2 & ".xlsx"
    Workbooks.Add.SaveAs Filename:="C:\Users\jp_bu\Desktop\テスト\テスト\" & fairumei2 & ".xlsx"

    Workbooks(fairumei2).Worksheets("Sheet1").Cells(1, 1) = shouhinbango_full '商品番号
    Workbooks(fairumei2).Worksheets("Sheet1").Cells(1, 2) = sheetmei_free 'シート名
    Workbooks(fairumei2).Worksheets("Sheet1").Cells(1, 3) = shouhinbango  '行数
    Workbooks(fairumei2).Worksheets("Sheet1").Cells(1, 4) = sheetmei_syurui '種類
    Workbooks(fairumei2).Worksheets("Sheet1").Cells(1, 5) = user_id 'ユーザーID
    
    Workbooks(fairumei2).Close True
    
    
    MsgBox "データをアウトしました。"


    
    '自動計算
    Application.ScreenUpdating = True
    
L1:

Workbooks(imanobook).Worksheets("かんたん検品").Activate
    
    
    
End Sub




Sub 仕分け記入印刷_連続検品_委託用()


    Application.ScreenUpdating = False
    

Dim imanobook As String
imanobook = ActiveWorkbook.Name

Dim imanosheet As String
imanosheet = ActiveSheet.Name

Dim Sh As Worksheet
Set Sh = Worksheets("シール")
Dim sh2 As Worksheet
Set sh2 = Worksheets("シール番号")

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

    
'    If Len(Workbooks(imanobook).Worksheets("設定").Cells(12, 1)) <> 0 Then
'    uchizei = 1 ''''内税フラグ
'    End If
    
    
    shori_handan = InputBox("合計金額" & Chr(10) & Chr(10) & WorksheetFunction.Sum(Range("A26:A100")) * 100 & "円" _
    & Chr(10) & Chr(10) & "1  シール出す" & Chr(10) & "2  シール出す(番号シールなし）" & Chr(10) & "3  合計金額見たから戻る", Default:=1)
    
'    If shori_handan = 3 Then
'    GoTo L1
'    End If
'
'    If Len(shori_handan) = 0 Then
'    GoTo L1
'    End If
    
    gyou = Workbooks(imanobook).Worksheets("かんたん検品").Cells(10000, 1).End(xlUp).Row
    
    If gyou > 100 Then
    MsgBox ("行数がおかしいです。やり直してください。")
        GoTo L1
    End If
    
    For x = 26 To gyou ''''''''''　　　　　　　　　　　　　　　''''''''''''''''''一旦全部チェック
    If Len(Workbooks(imanobook).Worksheets("かんたん検品").Cells(x, 1)) = 0 Then
    MsgBox (x & "　行目に金額が入っていません。")
    GoTo L1
    End If
    

    Next x
    
    
    
    With WshNetworkObject
    computamei = .ComputerName
    End With
        
    com2 = Left(computamei, 1) & Right(computamei, 1)
    
    forudaNO = forudaNOつける(computamei)
    jougen = forudaNO + 49999
    
    
'    If shori_handan = 2 Then
'    kingaku = InputBox("金額は？（略して）")
'    kingaku = kingaku * 100
'    maisuu = InputBox("枚数は？", Default:=2)
'    gyou = maisuu + 17
'    End If
    
     'テスト用のため修正が必要
'    Path = "D:\JP Dropbox\出品前データ\バッグ最新" & forudaNO & "\"
'    buf = Dir(Path & "*")
'    Do While buf <> ""
    
    Path = "D:\JP Dropbox\仕事\スタッフ_香織\テストシール" & forudaNO & "\"
    buf = Dir(Path & "*")
    Do While buf <> ""
        
        bango = Replace(buf, ".csv", "")
        If bango < jougen Then
        shouhinbango_full = bango + 1
        Else
        MsgBox ("フォルダの上限に達しました。他のフォルダでやり直してください。")
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
       keikoku = MsgBox(shouhinbango_full & Chr(13) & Chr(13) & "商品番号の重複の可能性が高いです。" & Chr(13) & Chr(13) & "everythingで" & shouhinbango_full & "をチェックしてください" & Chr(13) & _
         "既に同じ番号が登録されていたら重複です。" & Chr(13) & Chr(13) & "ネット環境をチェックして管理者に報告してください。", vbCritical)
       Exit Sub
       End If
        
        End If
            
'    If Len(Workbooks(imanobook).Worksheets("設定").Cells(8, 1)) <> 0 Then
'    hitocode3 = Workbooks(imanobook).Worksheets("設定").Cells(8, 1)
'    Else
'    MsgBox "バイヤー名が入力されていません。" & Chr(10) & "一旦終了します。"
'    GoTo L1
'    End If
'
'    buyerNO_kakutei = バイヤーコードつける2(imanobook, hitocode3)
'    buyer_kakutei = hitocode3
    
    
    
'    If uchizei = 1 Then
    
'    If shori_houhou = 1 Then
'    kingaku = Round((Workbooks(imanobook).Worksheets("かんたん検品").Cells(x, 1) / shouhizei), 0)
'    moto_kingaku = Workbooks(imanobook).Worksheets("かんたん検品").Cells(x, 1)
'    Else
'    kingaku = Round(((Workbooks(imanobook).Worksheets("かんたん検品").Cells(x, 1) * 100) / shouhizei), 0)
'    moto_kingaku = Workbooks(imanobook).Worksheets("かんたん検品").Cells(x, 1) * 100
'    End If
'
'    Else
'
'    If shori_houhou = 1 Then
'    kingaku = Workbooks(imanobook).Worksheets("かんたん検品").Cells(x, 1)
'
'    Else
'    kingaku = Workbooks(imanobook).Worksheets("かんたん検品").Cells(x, 1) * 100
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
    MsgBox ("仕入金額がおかしいかも")
    End If
    
            
    If Len(kingaku) = 0 Then
    kingaku_kakutei = "00000000"
    End If
    
    namae = shouhinbango_kakutei & "_" & kingaku_kakutei & com2 & "_" & computamei
    
'    If Len(kingaku_kakutei) <> 8 Then
'    MsgBox ("仕入金額がおかしいです。やり直してください。")
'    GoTo L1
'    End If
'
'    teigaku_kakutei = "0000"
    
        
'    If Len(Workbooks(imanobook).Worksheets("設定").Cells(9, 1)) <> 0 Then
'    ichiba = Workbooks(imanobook).Worksheets("設定").Cells(9, 1)
'    Else
'    ichiba = "xxx"
'    End If
'
'    If Len(ichiba) <> 3 Then
'    ichiba = "xxx"
'    End If
    
'    If Len(Workbooks(imanobook).Worksheets("設定").Cells(10, 1)) <> 0 Then
'    hiduke = Workbooks(imanobook).Worksheets("設定").Cells(10, 1)
'    Else
'    hiduke = "19000101"
'    End If
'
'    If Len(Workbooks(imanobook).Worksheets("設定").Cells(10, 1)) <> 8 Then
'    hiduke = "19000101"
'    End If
'
'    If Len(buyerNO_kakutei) = 0 Then
'    buyerNO_kakutei = "J"
'    MsgBox ("バイヤーコードがありません。Jにしておきます。")
'    End If
         
    
'    namae = shouhinbango_kakutei & buyerNO_kakutei & kingaku_kakutei & ichiba & teigaku_kakutei & com2 & hiduke & "_" & computamei
'    namae = shouhinbango_kakutei & "_" & buyerNO_kakutei & "_" & kingaku_kakutei & "_" & ichiba & "_" & teigaku_kakutei & "_" & com2 & "_" & hiduke & "_" & computamei
'    namae2 = shouhinbango
'    namae3 = shouhinbango_kakutei & buyer_kakutei & kingaku_kakutei
    
            
'---シール

    Workbooks(imanobook).Worksheets("シール番号").Cells(1, 1) = shouhinbango_full
    '1行目空白
    '商品番号
    Workbooks(imanobook).Worksheets("シール").Cells(2, 1) = shouhinbango_full
    'バーコード
    Workbooks(imanobook).Worksheets("シール").Cells(3, 1) = "*" & shouhinbango_full & "*"
    '金額
    Workbooks(imanobook).Worksheets("シール").Cells(4, 1) = kingaku & "円"
    'ユーザーID
    Workbooks(imanobook).Worksheets("シール").Cells(4, 2) = user_id
    
    
    '会社名
    If Len(Workbooks(imanobook).Worksheets("かんたん検品").Cells(12, 1)) <> 0 Then
        target_kaisya = Workbooks(imanobook).Worksheets("かんたん検品").Cells(12, 1)
        
        For i = 2 To Workbooks("商品一覧ebay_委託用").Worksheets("顧客情報").Cells(Rows.Count, 1).End(xlUp).Row
            If Workbooks("商品一覧ebay_委託用").Worksheets("顧客情報").Cells(i, 1) = target_kaisya Then
                kaisya_name = Workbooks("商品一覧ebay_委託用").Worksheets("顧客情報").Cells(i, 2)
                Exit For
            End If
        Next i
        
    
    Workbooks(imanobook).Worksheets("シール").Cells(4, 3) = kaisya_name
    
    End If
    
    
    'メモ1
    Workbooks(imanobook).Worksheets("シール").Cells(6, 1) = memo1
    'メモ2
    Workbooks(imanobook).Worksheets("シール").Cells(7, 1) = memo2

'    Workbooks(imanobook).Worksheets("シール番号").Cells(1, 1) = shouhinbango_full
'    Workbooks(imanobook).Worksheets("シール").Cells(2, 1) = shouhinbango_full
'    Workbooks(imanobook).Worksheets("シール").Cells(3, 1) = "*" & shouhinbango_full & "*"
'
'    Workbooks(imanobook).Worksheets("シール").Cells(4, 1) = kingaku & "円"
'    Workbooks(imanobook).Worksheets("シール").Cells(4, 2) = teigaku
'    Workbooks(imanobook).Worksheets("シール").Cells(4, 3) = buyer_kakutei
'
'    Workbooks(imanobook).Worksheets("シール").Cells(5, 1) = hiduke '日付
'    If uchizei = 1 Then
'    Workbooks(imanobook).Worksheets("シール").Cells(4, 1) = kingaku & "円 (" & moto_kingaku & "円)"
'    Workbooks(imanobook).Worksheets("シール").Cells(5, 3) = ichiba & "内税買い" '市場
'    Else
'    Workbooks(imanobook).Worksheets("シール").Cells(5, 3) = ichiba '市場
'    End If
'    Workbooks(imanobook).Worksheets("シール").Cells(6, 1) = lotno_kakutei
'    Workbooks(imanobook).Worksheets("シール").Cells(7, 1) = sonota
    
    Call Sh.PrintOut(ActivePrinter:="Brother QL-800")
    
    Workbooks(imanobook).Worksheets("シール").Select
    Range("A1:F20").ClearContents
    
    If shori_handan = 1 Then
    Call sh2.PrintOut(ActivePrinter:="Brother QL-800")
    
    Workbooks(imanobook).Worksheets("シール番号").Select
    Range("A1:F20").ClearContents
    End If
    
    
     'テスト用のため修正が必要
'    Workbooks(imanobook).Sheets("csv").Select
'    Sheets("csv").Copy
'    ChDir "D:\JP Dropbox\出品前データ\商品一覧登録用"
'    ActiveWorkbook.SaveAs Filename:="D:\JP Dropbox\出品前データ\商品一覧登録用\" & namae & ".csv", FileFormat:=xlCSV, _
'        CreateBackup:=False

    Workbooks(imanobook).Sheets("csv").Select
    Sheets("csv").Copy
    ChDir "D:\JP Dropbox\仕事\スタッフ_香織\テストシール\商品一覧登録用テスト"
    ActiveWorkbook.SaveAs Filename:="D:\JP Dropbox\仕事\スタッフ_香織\テスト\テストシール\商品一覧登録用テスト\" & namae & ".csv", FileFormat:=xlCSV, _
        CreateBackup:=False

            
    Workbooks(namae).Close False
    
    shouhinbango_full = shouhinbango_full + 1
    cont = cont + 1
    
    Next x
    
    'テスト用
'    Name "D:\JP Dropbox\出品前データ\バッグ最新" & forudaNO & "\" & bango & ".csv" As "D:\JP Dropbox\出品前データ\バッグ最新" & forudaNO & "\" & bango + cont & ".csv"
     Name "D:\JP Dropbox\仕事\スタッフ_香織\テストシール" & forudaNO & "\" & bango & ".csv" As "D:\JP Dropbox\仕事\スタッフ_香織\テストシール" & forudaNO & "\" & bango + 1 & ".csv"

    
    
    clear_handan = InputBox("1  クリアする" & Chr(10) & "2  クリアしない", Default:=1)
    
    If clear_handan = 1 Then
    Call クリア仕分け_かんたん検品
    End If
            
    Workbooks(imanobook).Worksheets("かんたん検品").Activate
    Worksheets("かんたん検品").Cells(1, 1).Select
    

  '自動計算
Application.ScreenUpdating = True


    
L1:

Workbooks(imanobook).Worksheets("かんたん検品").Activate
    
    
    
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
    
    If Workbooks(imanobook).Worksheets("設定").Cells(30, 1) = shouhinbango_full Then
    flag = 1
    End If
          
    Path = "D:\JP Dropbox\出品前データ\商品一覧登録用_済み\"
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
    Path = "D:\JP Dropbox\出品前データ\商品一覧登録用\"
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

