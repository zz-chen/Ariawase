Attribute VB_Name = "Module2"
Option Explicit

Sub �`�惁�j���[�ǉ�()
'
    Dim MnuBar As CommandBar
    Dim Ctrl As CommandBarControl
    Dim MnuItem As CommandBarControl
    Dim SubMnuItem As CommandBarControl
'
    With CommandBars("Worksheet Menu Bar")
        
        .Reset
        
        Set Ctrl = .Controls.Add(Type:=msoControlPopup)
        
        With Ctrl
            .Caption = "�X�P�W���[���\�`��c�[��(&T)"
'
            Set MnuItem = .Controls.Add(Type:=msoControlButton)
            With MnuItem
                .Caption = "�����ݒ�(&I)"
                .OnAction = "�X�P�W���[�������ݒ���"
            End With
'
            Set MnuItem = .Controls.Add(Type:=msoControlButton)
            With MnuItem
                .Caption = "�\���ѕ`��(&E)"
                .OnAction = "�X�P�W���[���\�`��"
            End With
'
        End With
'
    End With
'
End Sub
Sub Auto_Open()

  �`�惁�j���[�ǉ�

End Sub

Sub AUTO_Close()

  ReturnToExcel

End Sub

Sub ReturnToExcel()
'
    Dim CmdBar As CommandBar
'
    CommandBars("Worksheet Menu Bar").Reset
'
    For Each CmdBar In CommandBars
        If CmdBar.Name = "UserMenuBar" Then
            CmdBar.Delete
        End If
    Next
'
End Sub

