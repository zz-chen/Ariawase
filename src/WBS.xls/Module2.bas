Attribute VB_Name = "Module2"
Option Explicit

Sub 描画メニュー追加()
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
            .Caption = "スケジュール表描画ツール(&T)"
'
            Set MnuItem = .Controls.Add(Type:=msoControlButton)
            With MnuItem
                .Caption = "書式設定(&I)"
                .OnAction = "スケジュール書式設定画面"
            End With
'
            Set MnuItem = .Controls.Add(Type:=msoControlButton)
            With MnuItem
                .Caption = "予実績描画(&E)"
                .OnAction = "スケジュール表描画"
            End With
'
        End With
'
    End With
'
End Sub
Sub Auto_Open()

  描画メニュー追加

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

