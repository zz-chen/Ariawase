VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 書式設定画面 
   Caption         =   "書式の設定"
   ClientHeight    =   3465
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4200
   OleObjectBlob   =   "書式設定画面.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "書式設定画面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
On Error GoTo Date_Error
Worksheets("マクロ").Range("基準日") = DateSerial(基準年, 基準月, 基準日)
Worksheets("マクロ").Range("基準日2") = DateSerial(基準年2, 基準月2, 基準日2)

If 日次 = True Then
    Worksheets("マクロ").Range("日程表タイプ") = "D"
    Worksheets("マクロ").Range("列毎の日数") = 1
Else
    Worksheets("マクロ").Range("日程表タイプ") = "W"
    Worksheets("マクロ").Range("列毎の日数") = 7
End If

If (Worksheets("マクロ").Range("基準日2") - Worksheets("マクロ").Range("基準日")) _
    / Worksheets("マクロ").Range("列毎の日数") < 19 Then
    書式設定_Error = 5
    GoTo Exit_Sub
End If

If (Worksheets("マクロ").Range("基準日2") - Worksheets("マクロ").Range("基準日")) _
    / Worksheets("マクロ").Range("列毎の日数") + Worksheets("マクロ").Range("表開始列") > 255 Then
    書式設定_Error = 6
    GoTo Exit_Sub
End If

If 予定線率 > 90 Or 予定線率 < 10 Then
    書式設定_Error = 2
    GoTo Exit_Sub
End If

If 実績線率 > 90 Or 実績線率 < 10 Then
    書式設定_Error = 3
    GoTo Exit_Sub
End If

If 実績線率 < 予定線率 Then
    書式設定_Error = 0
Else
    書式設定_Error = 4
    GoTo Exit_Sub
End If

Worksheets("マクロ").Range("予定線率") = 予定線率 / 100
Worksheets("マクロ").Range("実績線率") = 実績線率 / 100

'2004/12/05追加 begin

If 日数進捗 = True Then
    Worksheets("マクロ").Range("進捗指標単位") = "Day"
Else
    Worksheets("マクロ").Range("進捗指標単位") = "%"
End If

'2004/12/05追加 end

Unload Me
Exit Sub

Date_Error:
書式設定_Error = 1
Exit_Sub:
Unload Me
End Sub


Private Sub CommandButton1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub CommandButton1_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
End Sub

Private Sub CommandButton1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Unload Me
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label5_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label5_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub Label5_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label8_Click()

End Sub
Option Explicit


Private Sub UserForm_Click()

End Sub
