Attribute VB_Name = "M_NamesExport"
Option Explicit

'******************************************************************************************
'*関数名    ：NamesExport
'*機能      ：ブックの名前定義情報をエクスポートしたcsvを出力する。
'*            引数のフォルダに、引数のブックに関連付いた名前定義をエクスポートしたcsvファイルを作成する
'*引数(1)   ：出力先フォルダパス
'*引数(2)   ：名前定義出力対象のブックオブジェクト
'******************************************************************************************
Public Sub NamesExport(ByVal prmFolderPath As String, _
                       ByVal prmWB As Workbook)
    
    '定数
    Const FUNC_NAME As String = "NamesOutput"
    Const ForWriting As Long = 2                 '新規書き込み
    
    '変数
    Dim outPutCsvPath As String
    
    On Error GoTo ErrorHandler
    '---以下に処理を記述---
    
    With CreateObject("Scripting.FileSystemObject")
    
        '出力csvのパスを設定
        outPutCsvPath = (prmFolderPath & "\" & .GetBaseName(prmWB.Name) & "_Names.csv")
    
        'csvファイルを作成する
        'すでに同名ファイルが存在する場合は上書きの可否を問う
        If .FileExists(outPutCsvPath) Then
            If MsgBox("上書きします。" & vbLf & "よろしいですか。", vbYesNo) = vbNo Then
                GoTo ExitHandler
            End If
        End If
        
        .CreateTextFile (outPutCsvPath)
        
        'csvファイルを開く
        With .OpenTextFile(outPutCsvPath, ForWriting, True)
            '名前定義の書き込み
            .Write NamesExport_Sub_PullNamesData(prmWB)
            .Close
        End With
        
    End With
    
    MsgBox "エクスポートが完了しました。"
    
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
'*関数名    ：NamesExport_Sub_PullNamesData
'*機能      ：引数のブックの名前定義データをcsv形式文字列として取得
'*引数(1)   ：対象のブックオブジェクト
'*戻り値    ：csv形式文字列
'******************************************************************************************
Private Function NamesExport_Sub_PullNamesData(ByVal prmWB As Workbook) As String
    
    '定数
    Const FUNC_NAME As String = "NamesOutput_Sub_PullNamesData"
    
    '変数
    Dim cnt As Long
    Dim tempStr As String
    
    On Error GoTo ErrorHandler
    '戻り値初期値
    NamesExport_Sub_PullNamesData = ""
    
    '---以下に処理を記述---
    
    'csvヘッダーを設定
    tempStr = """名前""" & DELIMITER & _
              """参照範囲（A1）""" & DELIMITER & _
              """参照範囲（R1C1）""" & DELIMITER & _
              """コメント"""
    tempStr = tempStr & vbCrLf
    
    '定義された名前の数だけループ
    For cnt = 1 To prmWB.Names.Count
        
        'ブック範囲の名前定義のみ抽出
        If TypeName(prmWB.Names(cnt).Parent) = "Workbook" Then
        
            '情報取得
            tempStr = tempStr & prmWB.Names(cnt).Name & DELIMITER & _
                                                      prmWB.Names(cnt).RefersTo & DELIMITER & _
                                                      prmWB.Names(cnt).RefersToR1C1 & DELIMITER & _
                                                      prmWB.Names(cnt).Comment
                                                      
            '改行設定
            tempStr = tempStr & vbCrLf
        End If
    Next
    
    '最後の改行コードを除去
    tempStr = Mid(tempStr, 1, Len(tempStr) - 2)
    
    '戻り値設定
    NamesExport_Sub_PullNamesData = tempStr
    
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


