Attribute VB_Name = "Tools"
Option Explicit

'******************************************************************************************
'*関数名    ：invalidateEvents
'*機能      ：ブックのイベントをフォーム表示している間無効にする
'*引数(1)   ：無し
'******************************************************************************************
Public Sub invalidateEvents()
    
    '定数
    Const FUNC_NAME As String = "invalidateEvents"
    
    '変数
    
    On Error GoTo ErrorHandler
    '---以下に処理を記述---
    
    'イベント無効化
    Application.EnableEvents = False
    
    'フォーム表示
    F_invalidateEvents.Show vbModeless

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

