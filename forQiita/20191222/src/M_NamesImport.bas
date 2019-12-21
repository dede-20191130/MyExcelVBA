Attribute VB_Name = "M_NamesImport"
'@Folder("VBAProject")
Option Explicit

'******************************************************************************************
'*関数名    ：NamesImport
'*機能      ：csvファイルの名前定義を読み取り、指定したブックの名前定義設定に新しく取得した名前定義をインポートする。
'*引数(1)   ：csvファイルのパス
'*引数(2)   ：インポート対象ブックオブジェクト
'******************************************************************************************
Public Sub NamesImport(ByVal prmCsvFilePath As String, _
                       ByRef prmWB As Workbook)
    
    '定数
    Const FUNC_NAME As String = "NamesInput"
    
    '変数
    Dim namesArr As Variant
    Dim cnt As Long
    
    On Error GoTo ErrorHandler
    '---以下に処理を記述---
    
    'csv読み込み
    namesArr = NamesImport_Sub_CsvToArray(prmCsvFilePath)
    
    '指定したブックの名前定義設定に新しく取得した名前定義を挿入する
    For cnt = LBound(namesArr, 1) To UBound(namesArr, 1)
        '参照形式ごとに処理を分岐
        Select Case Application.ReferenceStyle
        Case xlA1
            prmWB.Names.Add namesArr(cnt, 0), namesArr(cnt, 1), True
        Case Else
            prmWB.Names.Add Name:=namesArr(cnt, 0), RefersToR1C1:=namesArr(cnt, 2), Visible:=True
        End Select
        
        'コメントを設定
        prmWB.Names(namesArr(cnt, 0)).Comment = namesArr(cnt, 3)
    Next
    
    MsgBox "名前定義のインポートを完了しました。"

ExitHandler:
    
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生しましたので終了します" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号" & Err.Number & Chr(13) & Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Sub

'******************************************************************************************
'*関数名    ：NamesImport_Sub_CsvToArray
'*機能      ：csvファイルの内容を二次元配列に格納
'*引数(1)   ：csvファイルパス
'*戻り値    ：二次元配列
'******************************************************************************************
Private Function NamesImport_Sub_CsvToArray(ByVal prmCsvFilePath As String) As Variant
    
    '定数
    Const FUNC_NAME As String = "NamesInput_Sub_CsvToArray"
    
    '変数
    Dim strArr() As String
    Dim lineArr As Variant
    Dim iFile As Long
    Dim buf As String
    Dim cnt_1 As Long
    Dim cnt_2 As Long
    Dim tempNum As Long
    
    On Error GoTo ErrorHandler
    
    '---以下に処理を記述---
    
    '初期化
    cnt_1 = 0
    cnt_2 = 0
    
    'ファイル番号取得
    iFile = FreeFile
    
    'csvファイルの行数取得
    With CreateObject("Scripting.FileSystemObject").OpenTextFile(prmCsvFilePath, 8)
        tempNum = .Line
    End With
    
    'Redim
    ReDim strArr(tempNum - 2, 3)
    
    'csvを開き、データを行数分配列に格納
    Open prmCsvFilePath For Input As #iFile
    Do Until EOF(iFile)
        Line Input #iFile, buf
        'ヘッダー行は無視する
        If cnt_1 <> 0 Then
            lineArr = Split(buf, DELIMITER)
            For cnt_2 = 0 To UBound(lineArr)
                strArr(cnt_1 - 1, cnt_2) = lineArr(cnt_2)
            Next cnt_2
        End If
        
        cnt_1 = cnt_1 + 1
            
    Loop
    Close #iFile
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '下記処理はエラーになる：VBAでは二次元配列に対してのRedimは最後の次元の要素数のみ許容
    '    '使用していない配列をリサイズする
    '    tempNum = cnt_1 - 1
    '    ReDim Preserve strArr(tempNum, 3)
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※

    '戻り値設定
    NamesImport_Sub_CsvToArray = strArr
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生しましたので終了します" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号" & Err.Number & Chr(13) & Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Function


