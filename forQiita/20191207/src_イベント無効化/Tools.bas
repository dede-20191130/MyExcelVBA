Attribute VB_Name = "Tools"
Option Explicit

'******************************************************************************************
'*�֐���    �FinvalidateEvents
'*�@�\      �F�u�b�N�̃C�x���g���t�H�[���\�����Ă���Ԗ����ɂ���
'*����(1)   �F����
'******************************************************************************************
Public Sub invalidateEvents()
    
    '�萔
    Const FUNC_NAME As String = "invalidateEvents"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    '---�ȉ��ɏ������L�q---
    
    '�C�x���g������
    Application.EnableEvents = False
    
    '�t�H�[���\��
    F_invalidateEvents.Show vbModeless

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

