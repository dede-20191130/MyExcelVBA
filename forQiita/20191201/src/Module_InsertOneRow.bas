Attribute VB_Name = "Module_InsertOneRow"
Option Explicit


'******************************************************************************************
'*�֐���    �FinsertOneRow
'*�@�\      �F��s��΂��ŋ�s�}��
'*����(1)   �F�Q�ƌ��͈̔�
'*����(2)   �F�ړ���̕\�̈�ԏ�̃f�[�^�̃Z���I�u�W�F�N�g
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function insertOneRow(ByVal myRange As Range, _
                             ByVal topCell As Range) As String
    
    '�萔
    Const FUNC_NAME As String = "insertOneRow"
    
    '�ϐ�
    Dim delta As Long                   '��ԏ�̃f�[�^�̃Z���ƃA�N�e�B�u�Z���̍s�ԍ��̍�
    
    On Error GoTo ErrorHandler
    '�߂�l�����l
    insertOneRow = ""
    
    '---�ȉ��ɏ������L�q---
    
    '�ϐ��̒l���擾
    delta = Application.ThisCell.Row - topCell.Row
    
    '����delta���͈͂̃Z�����ȏ�Ȃ珈�����I��
    If delta >= myRange.Count * 2 Then Exit Function
    
    '������ԏ�̃f�[�^�̃Z���Ɠ��̓Z���̍s�ԍ��̍��������Ȃ�Βl���i�[����
    If delta Mod 2 = 0 Then
        '�߂�l���i�[
        insertOneRow = myRange(delta / 2 + 1)
    End If

    
ExitHandler:

    Exit Function
    
ErrorHandler:

        MsgBox "�G���[���������܂����̂ŏI�����܂�" & _
                vbLf & _
                "�֐����F" & FUNC_NAME & _
                vbLf & _
                "�G���[�ԍ�" & Err.Number & Chr(13) & Err.Description, vbCritical
        
        GoTo ExitHandler
        
End Function

