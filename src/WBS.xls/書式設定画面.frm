VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �����ݒ��� 
   Caption         =   "�����̐ݒ�"
   ClientHeight    =   3465
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4200
   OleObjectBlob   =   "�����ݒ���.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�����ݒ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
On Error GoTo Date_Error
Worksheets("�}�N��").Range("���") = DateSerial(��N, ���, ���)
Worksheets("�}�N��").Range("���2") = DateSerial(��N2, ���2, ���2)

If ���� = True Then
    Worksheets("�}�N��").Range("�����\�^�C�v") = "D"
    Worksheets("�}�N��").Range("�񖈂̓���") = 1
Else
    Worksheets("�}�N��").Range("�����\�^�C�v") = "W"
    Worksheets("�}�N��").Range("�񖈂̓���") = 7
End If

If (Worksheets("�}�N��").Range("���2") - Worksheets("�}�N��").Range("���")) _
    / Worksheets("�}�N��").Range("�񖈂̓���") < 19 Then
    �����ݒ�_Error = 5
    GoTo Exit_Sub
End If

If (Worksheets("�}�N��").Range("���2") - Worksheets("�}�N��").Range("���")) _
    / Worksheets("�}�N��").Range("�񖈂̓���") + Worksheets("�}�N��").Range("�\�J�n��") > 255 Then
    �����ݒ�_Error = 6
    GoTo Exit_Sub
End If

If �\����� > 90 Or �\����� < 10 Then
    �����ݒ�_Error = 2
    GoTo Exit_Sub
End If

If ���ѐ��� > 90 Or ���ѐ��� < 10 Then
    �����ݒ�_Error = 3
    GoTo Exit_Sub
End If

If ���ѐ��� < �\����� Then
    �����ݒ�_Error = 0
Else
    �����ݒ�_Error = 4
    GoTo Exit_Sub
End If

Worksheets("�}�N��").Range("�\�����") = �\����� / 100
Worksheets("�}�N��").Range("���ѐ���") = ���ѐ��� / 100

'2004/12/05�ǉ� begin

If �����i�� = True Then
    Worksheets("�}�N��").Range("�i���w�W�P��") = "Day"
Else
    Worksheets("�}�N��").Range("�i���w�W�P��") = "%"
End If

'2004/12/05�ǉ� end

Unload Me
Exit Sub

Date_Error:
�����ݒ�_Error = 1
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
