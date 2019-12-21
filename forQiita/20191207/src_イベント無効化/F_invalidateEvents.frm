VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_invalidateEvents 
   Caption         =   "InvalidateEvents"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "F_invalidateEvents.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "F_invalidateEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'******************************************************************************************
'*関数名    ：CommandButton_Close_Click
'*機能      ：閉じるボタン押下時処理
'*引数(1)   ：
'******************************************************************************************
Private Sub CommandButton_Close_Click()
    
    '定数
    Const FUNC_NAME As String = "CommandButton_Close_Click"
    
    '変数
    
    On Error GoTo ErrorHandler
    '---以下に処理を記述---
    
    'フォームを閉じる
    Unload F_invalidateEvents
    
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
'*関数名    ：UserForm_QueryClose
'*機能      ：〇〇○
'*引数(1)   ：〇〇○
'******************************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    '定数
    Const FUNC_NAME As String = "UserForm_QueryClose"
    
    '変数
    
    On Error GoTo ErrorHandler
    '---以下に処理を記述---
    
    'イベント有効化
    Application.EnableEvents = True

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


