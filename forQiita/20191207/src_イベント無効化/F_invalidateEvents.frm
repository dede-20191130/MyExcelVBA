VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_invalidateEvents 
   Caption         =   "InvalidateEvents"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "F_invalidateEvents.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "F_invalidateEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'******************************************************************************************
'*�֐���    �FCommandButton_Close_Click
'*�@�\      �F����{�^������������
'*����(1)   �F
'******************************************************************************************
Private Sub CommandButton_Close_Click()
    
    '�萔
    Const FUNC_NAME As String = "CommandButton_Close_Click"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    '---�ȉ��ɏ������L�q---
    
    '�t�H�[�������
    Unload F_invalidateEvents
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

        MsgBox "�G���[���������܂����̂ŏI�����܂�" & _
                vbLf & _
                "�֐����F" & FUNC_NAME & _
                vbLf & _
                "�G���[�ԍ�" & Err.Number & Chr(13) & Err.Description, vbCritical
        
        GoTo ExitHandler
        
End Sub



'******************************************************************************************
'*�֐���    �FUserForm_QueryClose
'*�@�\      �F�Z�Z��
'*����(1)   �F�Z�Z��
'******************************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    '�萔
    Const FUNC_NAME As String = "UserForm_QueryClose"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    '---�ȉ��ɏ������L�q---
    
    '�C�x���g�L����
    Application.EnableEvents = True

ExitHandler:

    Exit Sub
    
ErrorHandler:

        MsgBox "�G���[���������܂����̂ŏI�����܂�" & _
                vbLf & _
                "�֐����F" & FUNC_NAME & _
                vbLf & _
                "�G���[�ԍ�" & Err.Number & Chr(13) & Err.Description, vbCritical
        
        GoTo ExitHandler
        
End Sub


