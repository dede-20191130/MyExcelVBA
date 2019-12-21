Attribute VB_Name = "M_Common"
'@Folder("VBAProject")
Option Explicit



'******************************************************************************************
'*関数名    ：ReturnCurrentBookNames
'*機能      ：現在開かれているエクセルブックの名前の結合文字列（カンマ区切り）を返す
'*引数(1)   ：無し
'*戻り値    ：結合文字列
'******************************************************************************************
Public Function ReturnCurrentBookNames() As String
    
    '定数
    Const FUNC_NAME As String = "ReturnCurrentBookNames"
    
    '変数
    Dim tempStr As String
    Dim cnt As Long
    
    On Error GoTo ErrorHandler
    
    '---以下に処理を記述---
    
    'ブックの名前取得
    For cnt = 1 To Workbooks.Count
        tempStr = tempStr & Workbooks(cnt).Name & ","
    Next
    
    '最後のカンマ除去
    tempStr = Mid(tempStr, 1, Len(tempStr) - 1)
    
    '戻り値設定
    ReturnCurrentBookNames = tempStr
    
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



