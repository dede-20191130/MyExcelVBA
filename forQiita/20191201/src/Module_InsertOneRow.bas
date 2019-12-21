Attribute VB_Name = "Module_InsertOneRow"
Option Explicit


'******************************************************************************************
'*関数名    ：insertOneRow
'*機能      ：一行飛ばしで空行挿入
'*引数(1)   ：参照元の範囲
'*引数(2)   ：移動先の表の一番上のデータのセルオブジェクト
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function insertOneRow(ByVal myRange As Range, _
                             ByVal topCell As Range) As String
    
    '定数
    Const FUNC_NAME As String = "insertOneRow"
    
    '変数
    Dim delta As Long                   '一番上のデータのセルとアクティブセルの行番号の差
    
    On Error GoTo ErrorHandler
    '戻り値初期値
    insertOneRow = ""
    
    '---以下に処理を記述---
    
    '変数の値を取得
    delta = Application.ThisCell.Row - topCell.Row
    
    'もしdeltaが範囲のセル数以上なら処理を終了
    If delta >= myRange.Count * 2 Then Exit Function
    
    'もし一番上のデータのセルと入力セルの行番号の差が偶数ならば値を格納する
    If delta Mod 2 = 0 Then
        '戻り値を格納
        insertOneRow = myRange(delta / 2 + 1)
    End If

    
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

