Attribute VB_Name = "Module1"
Option Explicit

Public �񍐓� As Date       '�񍐓����Z�b�g
Public ��� As Date       '�X�P�W���[���\�̂e�q�n�l���t
Public ���2 As Date      '�X�P�W���[���\�̂s�n���t
Public �����\�^�C�v         'D�F�����CW�F�T���CM�F����
Public Const �\�J�n�s = 1
Public Const �\�J�n�� = 1
Public �\�I���s
Public �\�I����
Public �����J�n��
Public Const �����J�n�s = 1
Public �����I����
Public �����I���s
Public �J�n����
Public �I������
Public ��
Public �s
Public ���эs
Public ���ѓ� As Date
Public �񖈂̓���
Public Const �w�i���� = ""
Public WBS���� As Boolean
Public �\�͈͊O As Boolean
Public �O�sX
Public �O�sY
Public �񍐓�X
Public �\�J�n��X
Public �\�I����X
Public �\�����
Public ���ѐ���
Public �����ݒ�_Error
Public �����ݒ芮�� As Boolean


'***********
Sub ������()
'***********

�񍐓� = Worksheets("�}�N��").Range("�񍐓�")
��� = Worksheets("�}�N��").Range("���")
���2 = Worksheets("�}�N��").Range("���2")

If Worksheets("�}�N��").Range("�����\�^�C�v") = "D" Then
   �񖈂̓��� = 1
Else
   �񖈂̓��� = 7
End If

�\����� = Worksheets("�}�N��").Range("�\�����")
���ѐ��� = Worksheets("�}�N��").Range("���ѐ���")

'2004/12/05�ǉ� begin

If Worksheets("�}�N��").Range("�i���w�W�P��") = "Day" Then
    Worksheets("�X�P�W���[���\").Range("�i���w�W���o��") = "�i���w�W�i�����j"
    Worksheets("�X�P�W���[���\").Range("�i���w�W").Select
            Selection.NumberFormatLocal = "+#,###;[��]-#,###"
Else
    Worksheets("�X�P�W���[���\").Range("�i���w�W���o��") = "�i���w�W�i���j"
    Worksheets("�X�P�W���[���\").Range("�i���w�W").Select
            Selection.NumberFormatLocal = "0%"
End If

'2004/12/05�ǉ� end

�\�I���s = Range("�`��̈�").Rows.Count
�\�I���� = Range("�`��̈�").Columns.Count
�����I���s = Range("�`��̈�").Rows.Count
�����I���� = Range("�`��̈�").Columns.Count

End Sub

'**********************
Sub �X�P�W���[���\�`��()
Attribute �X�P�W���[���\�`��.VB_Description = "�}�N���L�^�� : 2004/3/5  ���[�U�[�� :  "
Attribute �X�P�W���[���\�`��.VB_ProcData.VB_Invoke_Func = " \n14"
'**********************

If ActiveSheet.Name <> "�X�P�W���[���\" Then
    MsgBox "�V�[�g�u�X�P�W���[���\�v���J���Ă�����s���Ă��������B"
    Exit Sub
Else
End If

������

Dim REPLY

Do While REPLY <> 6
    On Error GoTo Error_Exit
    �񍐓� = InputBox("�񍐓���yy/mm/dd�`���œ��͂��Ă�������", "�񍐓��̓���", �񍐓�)
    If �񍐓� > ��� And �񍐓� < ��� + (�\�I���� - �\�J�n��) * �񖈂̓��� Then
        Range("�񍐓�") = �񍐓�
        REPLY = MsgBox("�񍐓� " & �񍐓� & " �̃X�P�W���[���\���쐬���܂����H", vbQuestion + vbYesNo)
    Else
        MsgBox "�񍐓����`��͈͊O�ł��B"
    End If
Loop

'�I�[�g�V�F�C�v�C���̏���

On Error Resume Next
ActiveSheet.Rectangles.Delete
ActiveSheet.Lines.Delete

'�\�J�n��w���W�����

�����J�n�� = �\�J�n��

Call GetCellXYext(Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����J�n�� + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, �O�sX, �O�sY, Pos:=1)

�\�J�n��X = �O�sX

'�\�I����w���W�����

�����J�n�� = �\�I����

Call GetCellXYext(Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����J�n�� + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, �O�sX, �O�sY, Pos:=1)

�\�I����X = �O�sX

'�񍐓��w���W�����

�����J�n�� = Application.WorksheetFunction.RoundUp((�񍐓� - (��� - 1 * �񖈂̓���)) / �񖈂̓���, 0)

Call GetCellXYext(Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����J�n�� + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, �O�sX, �O�sY, Pos:=1)
�񍐓�X = �O�sX

'�`��
For �s = �\�J�n�s To �\�I���s Step 1
    �s����
Next �s

Exit Sub

Error_Exit:
MsgBox ("���t���K��O�ł��B���j���[�����蒼���Ă��������B")
End Sub

'***********
Sub �s����()
'***********

'�\����Ȃ�
If Worksheets("�X�P�W���[���\").Range("�J�n��").Cells(�s, 1) = 0 Or _
   Worksheets("�X�P�W���[���\").Range("�I����").Cells(�s, 1) = 0 Then
    
    �󔒍s�̈�ȕ`��
    GoTo exit�s����

End If


�\��:

�����J�n�� = Application.WorksheetFunction.RoundUp _
            ((Worksheets("�X�P�W���[���\").Range("�J�n��").Cells(�s, 1) - (��� - 1 * �񖈂̓���)) / �񖈂̓���, 0)

�����I���� = Application.WorksheetFunction.RoundUp _
            ((Worksheets("�X�P�W���[���\").Range("�I����").Cells(�s, 1) - (��� - 1 * �񖈂̓���)) / �񖈂̓���, 0)

�\�͈͊O = False

If �����I���� < �\�J�n�� Then
    
    �\�͈͊O = True
    �����J�n�� = 1
    �����I���� = 1

Else

    If �����J�n�� > �\�I���� Then
        
        �\�͈͊O = True
        �����J�n�� = �\�I����
        �����I���� = �\�I����
        
    Else
        
        �����J�n�� = Application.WorksheetFunction.Max(�����J�n��, �\�J�n��)
        �����I���� = Application.WorksheetFunction.Min(�����I����, �\�I����)

    End If

End If

�\��`��

����:

If Worksheets("�X�P�W���[���\").Range("���ъJ�n��").Cells(�s, 1) = 0 Then GoTo exit�s����

���ѓ� = 0

If Worksheets("�X�P�W���[���\").Range("���яI����").Cells(�s, 1) <> 0 Then
   
   ���ѓ� = Application.WorksheetFunction.Max( _
            Worksheets("�X�P�W���[���\").Range("���яI����").Cells(�s, 1), _
            Worksheets("�X�P�W���[���\").Range("�I����").Cells(�s, 1))

Else

   If Worksheets("�}�N��").Range("�i���w�W�P��") = "Day" Then
   
      If Worksheets("�X�P�W���[���\").Range("�I����").Cells(�s, 1) < �񍐓� Then
      
        ���ѓ� = Worksheets("�X�P�W���[���\").Range("�I����").Cells(�s, 1) + _
                 Worksheets("�X�P�W���[���\").Range("�i���w�W").Cells(�s, 1)
        
      Else
      
        ���ѓ� = �񍐓� + Worksheets("�X�P�W���[���\").Range("�i���w�W").Cells(�s, 1)
    
      End If

   Else
   
      ���ѓ� = Worksheets("�X�P�W���[���\").Range("�J�n��").Cells(�s, 1) - 1 + _
               Application.WorksheetFunction.RoundUp( _
                 ( _
                   Worksheets("�X�P�W���[���\").Range("�I����").Cells(�s, 1) + 1 - _
                   Worksheets("�X�P�W���[���\").Range("�J�n��").Cells(�s, 1) _
                 ) _
                 * Abs(Worksheets("�X�P�W���[���\").Range("�i���w�W").Cells(�s, 1)) _
                 , 0)
   End If
   
End If

If Worksheets("�}�N��").Range("�i���w�W�P��") = "Day" Then
   
   �����J�n�� = Application.WorksheetFunction.RoundUp _
               ((Worksheets("�X�P�W���[���\").Range("���ъJ�n��").Cells(�s, 1) - _
               (��� - 1 * �񖈂̓���)) / �񖈂̓���, 0)
               
Else

   �����J�n�� = Application.WorksheetFunction.RoundUp _
               ((Worksheets("�X�P�W���[���\").Range("�J�n��").Cells(�s, 1) - _
               (��� - 1 * �񖈂̓���)) / �񖈂̓���, 0)
               
End If

�����I���� = Application.WorksheetFunction.RoundUp _
            ((���ѓ� - (��� - 1 * �񖈂̓���)) / �񖈂̓���, 0)

�\�͈͊O = False

If �����I���� < �\�J�n�� Then
    
    �\�͈͊O = True
    �����J�n�� = 0
    �����I���� = 0

Else
    
    If �����J�n�� > �\�I���� Then
    
        �\�͈͊O = True
        �����J�n�� = �\�I����
        �����I���� = �\�I����
          
    Else
    
        �����J�n�� = Application.WorksheetFunction.Max(�����J�n��, �\�J�n��)
        �����I���� = Application.WorksheetFunction.Min(�����I����, �\�I����)
    
    End If

End If

WBS���� = False

If Worksheets("�X�P�W���[���\").Range("���яI����").Cells(�s, 1) <> 0 Then
   WBS���� = True
End If

���ѕ`��

exit�s����:

End Sub
Sub �󔒍s�̈�ȕ`��()
On Error Resume Next
'
Dim Rng As Range
Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim Y2 As Single
Dim X3 As Single
Dim Y3 As Single
Dim X4 As Single
Dim Y4 As Single
    
Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����J�n�� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X1, Y1, Pos:=1)
'
Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����I���� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X2, Y2, Pos:=2)

Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����I���� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X3, Y3, Pos:=3)

Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����J�n�� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X4, Y4, Pos:=4)
              
With ActiveSheet.Shapes.AddLine(�O�sX, �O�sY, �񍐓�X, Y2).Select
    Selection.ShapeRange.Fill.Transparency = 0#
    Selection.ShapeRange.Line.Weight = 2.25
    Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
    Selection.ShapeRange.Line.Style = msoLineSingle
    Selection.ShapeRange.Line.Transparency = 0#
    Selection.ShapeRange.Line.Visible = msoTrue
    Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
    Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
End With
   
   �O�sX = �񍐓�X
   �O�sY = Y2
   
With ActiveSheet.Shapes.AddLine(�O�sX, �O�sY, �񍐓�X, Y3).Select
    Selection.ShapeRange.Fill.Transparency = 0#
    Selection.ShapeRange.Line.Weight = 2.25
    Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
    Selection.ShapeRange.Line.Style = msoLineSingle
    Selection.ShapeRange.Line.Transparency = 0#
    Selection.ShapeRange.Line.Visible = msoTrue
    Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
    Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
End With
     
End Sub
'*************
Sub �\��`��()
'*************

On Error Resume Next
'
Dim Rng As Range
Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim Y2 As Single
Dim X3 As Single
Dim Y3 As Single
Dim X4 As Single
Dim Y4 As Single
    
Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����J�n�� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X1, Y1, Pos:=1)
'
Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����I���� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X2, Y2, Pos:=2)

Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����I���� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X3, Y3, Pos:=3)

Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����J�n�� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X4, Y4, Pos:=4)
     
If �\�͈͊O Then

Else

    With ActiveSheet.Shapes.AddShape(msoShapeRectangle, X1, Y1 + (Y3 - Y2) * (1 - �\�����) / 2, X2 - X1, (Y3 - Y2) * �\�����)
         .Fill.Solid
         .Fill.ForeColor.SchemeColor = 44
    End With

End If

Set Rng = Nothing

�\���ȕ`��:

With ActiveSheet.Shapes.AddLine(�O�sX, �O�sY, �񍐓�X, Y2).Select
    Selection.ShapeRange.Fill.Transparency = 0#
    Selection.ShapeRange.Line.Weight = 2.25
    Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
    Selection.ShapeRange.Line.Style = msoLineSingle
    Selection.ShapeRange.Line.Transparency = 0#
    Selection.ShapeRange.Line.Visible = msoTrue
    Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
    Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
End With

�O�sX = �񍐓�X
�O�sY = Y2

If Worksheets("�X�P�W���[���\").Range("���ъJ�n��").Cells(�s, 1) = 0 Then

    With ActiveSheet.Shapes.AddLine(�O�sX, �O�sY, Application.WorksheetFunction.Min(X1, �񍐓�X), Y1 + (Y3 - Y2) * (1 - �\�����) / 2).Select
    Selection.ShapeRange.Fill.Transparency = 0#
    Selection.ShapeRange.Line.Weight = 2.25
    Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
    Selection.ShapeRange.Line.Style = msoLineSingle
    Selection.ShapeRange.Line.Transparency = 0#
    Selection.ShapeRange.Line.Visible = msoTrue
    Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
    Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
   End With
   
   �O�sX = Application.WorksheetFunction.Min(X1, �񍐓�X)
   �O�sY = Y1 + (Y3 - Y2) * (1 - �\�����) / 2
   
   If �\�͈͊O And �����J�n�� = �\�J�n�� Then
      
   Else
    
      With ActiveSheet.Shapes.AddLine(�O�sX, �O�sY, Application.WorksheetFunction.Min(X4, �񍐓�X), Y1 + (Y3 - Y2) * (1 - �\�����) / 2 + (Y3 - Y2) * �\�����).Select
        Selection.ShapeRange.Fill.Transparency = 0#
        Selection.ShapeRange.Line.Weight = 2.25
        Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
        Selection.ShapeRange.Line.Style = msoLineSingle
        Selection.ShapeRange.Line.Transparency = 0#
        Selection.ShapeRange.Line.Visible = msoTrue
        Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
        Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
       End With
       
    End If
         
   �O�sX = Application.WorksheetFunction.Min(X4, �񍐓�X)
   �O�sY = Y1 + (Y3 - Y2) * (1 - �\�����) / 2 + (Y3 - Y2) * �\�����
End If

End Sub

'************
Sub ���ѕ`��()
'************
  
On Error Resume Next

Dim Rng As Range
Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim Y2 As Single
Dim X3 As Single
Dim Y3 As Single
Dim X4 As Single
Dim Y4 As Single
    
Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����J�n�� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X1, Y1, Pos:=1)

Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����I���� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X2, Y2, Pos:=2)

Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����I���� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X3, Y3, Pos:=3)

Call GetCellXYext(�s - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Row, _
     �����J�n�� - 1 + Worksheets("�X�P�W���[���\").Range("�`��̈�").Column, X4, Y4, Pos:=4)
     
If �\�͈͊O Then

Else

    If WBS���� Then
        With ActiveSheet.Shapes.AddShape(msoShapeRectangle, X1, Y1 + (Y3 - Y2) * 0.4, X2 - X1, (Y3 - Y2) * ���ѐ���)
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.SchemeColor = 12
        .Fill.Transparency = 0#
        .Line.Weight = 0.75
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Transparency = 0#
        .Line.Visible = msoTrue
        .Line.ForeColor.SchemeColor = 64
        .Line.BackColor.RGB = RGB(255, 255, 255)
        End With
    Else
        With ActiveSheet.Shapes.AddShape(msoShapeRectangle, X1, Y1 + (Y3 - Y2) * 0.4, X2 - X1, (Y3 - Y2) * ���ѐ���)
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.SchemeColor = 65
        .Fill.Transparency = 0#
        .Line.Weight = 1.5
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Transparency = 0#
        .Line.Visible = msoTrue
        .Line.ForeColor.SchemeColor = 64
        .Line.BackColor.RGB = RGB(255, 255, 255)
        End With
    End If
    
End If

���ш�ȕ`��:

With ActiveSheet.Shapes.AddLine(�O�sX, �O�sY, �񍐓�X, Y1).Select
    Selection.ShapeRange.Fill.Transparency = 0#
    Selection.ShapeRange.Line.Weight = 2.25
    Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
    Selection.ShapeRange.Line.Style = msoLineSingle
    Selection.ShapeRange.Line.Transparency = 0#
    Selection.ShapeRange.Line.Visible = msoTrue
    Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
    Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
End With

�O�sX = �񍐓�X
�O�sY = Y1

If WBS���� = True Then
    With ActiveSheet.Shapes.AddLine _
        (�O�sX, �O�sY, Application.WorksheetFunction.Max(X2, �񍐓�X), Y1 + (Y3 - Y2) * ���ѐ���).Select
        Selection.ShapeRange.Fill.Transparency = 0#
        Selection.ShapeRange.Line.Weight = 2.25
        Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
        Selection.ShapeRange.Line.Style = msoLineSingle
        Selection.ShapeRange.Line.Transparency = 0#
        Selection.ShapeRange.Line.Visible = msoTrue
        Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
        Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
    End With
    
    �O�sX = Application.WorksheetFunction.Max(X2, �񍐓�X)
    �O�sY = Y1 + (Y3 - Y2) * ���ѐ���
    
Else
    With ActiveSheet.Shapes.AddLine(�O�sX, �O�sY, X2, Y1 + (Y3 - Y2) * ���ѐ���).Select
        Selection.ShapeRange.Fill.Transparency = 0#
        Selection.ShapeRange.Line.Weight = 2.25
        Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
        Selection.ShapeRange.Line.Style = msoLineSingle
        Selection.ShapeRange.Line.Transparency = 0#
        Selection.ShapeRange.Line.Visible = msoTrue
        Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
        Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
        Selection.ShapeRange.Line.BeginArrowheadLength = msoArrowheadLengthMedium
        Selection.ShapeRahnge.Line.BeginArrowheadWidth = msoArrowheadWidthMedium
        Selection.ShapeRange.Line.BeginArrowheadStyle = msoArrowheadNone
        Selection.ShapeRange.Line.EndArrowheadLength = msoArrowheadLengthMedium
        Selection.ShapeRange.Line.EndArrowheadWidth = msoArrowheadWidthMedium
        Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadNone
    End With
    
    �O�sX = X2
    �O�sY = Y1 + (Y3 - Y2) * ���ѐ���

End If

If �\�͈͊O And �����J�n�� = �\�I���� _
    Or _
   �\�͈͊O And Worksheets("�X�P�W���[���\").Range("���яI����").Cells(�s, 1) = 0 Then

Else

    If WBS���� = True Then
        With ActiveSheet.Shapes.AddLine _
        (�O�sX, �O�sY, _
        Application.WorksheetFunction.Max(X2, �񍐓�X), Y1 + (Y3 - Y2) * ���ѐ��� + (Y3 - Y2) * �\�����).Select
        Selection.ShapeRange.Fill.Transparency = 0#
        Selection.ShapeRange.Line.Weight = 2.25
        Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
        Selection.ShapeRange.Line.Style = msoLineSingle
        Selection.ShapeRange.Line.Transparency = 0#
        Selection.ShapeRange.Line.Visible = msoTrue
        Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
        Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
        End With
    Else
        With ActiveSheet.Shapes.AddLine(�O�sX, �O�sY, X2, Y1 + (Y3 - Y2) * ���ѐ��� + (Y3 - Y2) * �\�����).Select
        Selection.ShapeRange.Fill.Transparency = 0#
        Selection.ShapeRange.Line.Weight = 2.25
        Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
        Selection.ShapeRange.Line.Style = msoLineSingle
        Selection.ShapeRange.Line.Transparency = 0#
        Selection.ShapeRange.Line.Visible = msoTrue
        Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
        Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
        End With
    End If

End If
    
If WBS���� = True Then

    �O�sX = Application.WorksheetFunction.Max(X2, �񍐓�X)
    �O�sY = Y1 + (Y3 - Y2) * ���ѐ��� + (Y3 - Y2) * �\�����
Else
    �O�sX = X2
    �O�sY = Y1 + (Y3 - Y2) * ���ѐ��� + (Y3 - Y2) * �\�����
End If

End Sub

'�`����W�����߂�
'************************************************
Sub GetCellXYext(RowNo, ColNo, X, Y, Optional Pos)
'************************************************
    Dim R As Long
    Dim C As Integer
'
    Dim dx As Single
    Dim dy As Single
'
    R = RowNo - 1
    C = ColNo - 1
'
    If R = 0 Then
        Y = 0
    Else
        Y = ActiveSheet.Rows("1:" & R).Height
    End If
'
    If C = 0 Then
        X = 0
    Else
        X = ActiveSheet.Columns("A:" & GetColName(C)).Width
    End If
'
    dx = ActiveSheet.Cells(RowNo, ColNo).Width
    dy = ActiveSheet.Cells(RowNo, ColNo).Height
'
    If IsMissing(Pos) Then Pos = 1
'
    Select Case Pos
        Case 2
            X = X + dx
        Case 3
            X = X + dx
            Y = Y + dy
        Case 4
            Y = Y + dy
    End Select
'
End Sub
'*****************************************************
'***��ԍ��iColNo�j����񖼂𓾂�
'*****************************************************
Function GetColName(ColNo)
'
    Dim Adrs As String
    Dim Pos As Integer
'
    Adrs = ActiveSheet.Columns(ColNo).Address(False, False)
'
    Pos = InStr(Adrs, ":")
'
    GetColName = Mid(Adrs, Pos + 1)
'
End Function
'************************
Sub �X�P�W���[�������ݒ���()
'************************

If ActiveSheet.Name <> "�X�P�W���[���\" Then
    MsgBox "�V�[�g�u�X�P�W���[���\�v���J���Ă�����s���Ă��������B"
    Exit Sub
Else
End If

�����ݒ芮�� = False

Do Until �����ݒ芮�� = True

�����ݒ���.��N _
    = Application.WorksheetFunction.Fixed(Year(Worksheets("�}�N��").Range("���")), 0, True)
�����ݒ���.��� _
    = Application.WorksheetFunction.Fixed(Month(Worksheets("�}�N��").Range("���")), 0, True)
�����ݒ���.��� _
    = Application.WorksheetFunction.Fixed(Day(Worksheets("�}�N��").Range("���")), 0, True)
    
�����ݒ���.��N2 _
    = Application.WorksheetFunction.Fixed(Year(Worksheets("�}�N��").Range("���2")), 0, True)
�����ݒ���.���2 _
    = Application.WorksheetFunction.Fixed(Month(Worksheets("�}�N��").Range("���2")), 0, True)
�����ݒ���.���2 _
    = Application.WorksheetFunction.Fixed(Day(Worksheets("�}�N��").Range("���2")), 0, True)
    
If Worksheets("�}�N��").Range("�����\�^�C�v") = "D" Then
    �����ݒ���.���� = 1
Else
    �����ݒ���.�T�� = 7
End If

�����ݒ���.�\����� = Worksheets("�}�N��").Range("�\�����") * 100
�����ݒ���.���ѐ��� = Worksheets("�}�N��").Range("���ѐ���") * 100

'2004/12/05�ǉ� begin

If Worksheets("�}�N��").Range("�i���w�W�P��") = "Day" Then
    �����ݒ���.�����i�� = 1
Else
    �����ݒ���.�S�����i�� = 1
End If

'2004/12/05�ǉ� end

�����ݒ�_Error = 9

�����ݒ���.Show
       
Select Case �����ݒ�_Error
        
        Case 0
        �����ݒ芮�� = True
        Case 1
        MsgBox "�\�J�n���܂��͕\�I�������K��O�ł��B"
        Case 2
        MsgBox "�\��������K��O�ł��B"
        Case 3
        MsgBox "���ѐ������K��O�ł��B"
        Case 4
        MsgBox "���ѐ����͗\�������菬�������Ă��������B"
        Case 5
        MsgBox "�X�P�W���[���\�̍ŏ��񐔂�20�������Ƃ��Ă��܂��B�񐔂�20�ȏ�ɂȂ�\�����Ԃ�ݒ肵�Ă��������"
        Case 6
        MsgBox "�X�P�W���[���\�̗񐔂�Excel�̍ő�l�i256�j�𒴂��܂��B�񐔂�256�ȉ��ɂȂ�\�����Ԃ�ݒ肵�Ă��������B"
        Case 9
        Exit Sub
End Select
    
Loop

Application.ScreenUpdating = False

Dim ������, �V����, ���ԍ�

�V���� = Application.WorksheetFunction.RoundUp( _
        (Worksheets("�}�N��").Range("���2") - Worksheets("�}�N��").Range("���")) _
       / Worksheets("�}�N��").Range("�񖈂̓���"), 0)
������ = Worksheets("�}�N��").Range("�\�I����") - Worksheets("�}�N��").Range("�\�J�n��")

���ԍ� = �V���� - ������

If ���ԍ� = 0 Then
Else
    If ���ԍ� > 0 Then
        Range(Columns(GetColName(Worksheets("�}�N��").Range("�\�J�n��") + 1)), _
              Columns(GetColName(Worksheets("�}�N��").Range("�\�J�n��") + ���ԍ�))).Select
        Selection.Insert Shift:=xlToRight
    Else
        Range(Columns(GetColName(Worksheets("�}�N��").Range("�\�J�n��") + 1)), _
              Columns(GetColName(Worksheets("�}�N��").Range("�\�J�n��") + ���ԍ� * -1))).Select
        Selection.Delete Shift:=xlToLeft
    End If
End If

������

On Error Resume Next

ActiveSheet.Rectangles.Delete
ActiveSheet.Lines.Delete

Worksheets("�}�N��").Range("���") = ���

Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Range(Cells(4, 1), Cells(4, �\�I����)).Select
    With Selection.Interior
        .ColorIndex = 0
        .PatternColorIndex = 1
    End With
    
Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Range(Cells(4, 1), Cells(4, �\�I����)).ClearContents

Dim ��

For �� = �\�J�n�� To �\�I����
    
    Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(3, ��) = _
    "'" & Application.WorksheetFunction.Fixed(Day(��� - �񖈂̓��� + �� * �񖈂̓���), 0, True)
    
    If �� <> Application.WorksheetFunction.Fixed(Month(��� - �񖈂̓��� + �� * �񖈂̓���), 0, True) Then
        �� = Application.WorksheetFunction.Fixed(Month(��� - �񖈂̓��� + �� * �񖈂̓���), 0, True)
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(1, ��) = _
        "'" & Application.WorksheetFunction.Fixed(Month(��� - �񖈂̓��� + �� * �񖈂̓���), 0, True) & "��"
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Range(Cells(1, ��), Cells(2, ��)).Select
            With Selection.Interior
                .ColorIndex = 0
                .PatternColorIndex = 1
            End With
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlDouble
                .Weight = xlThick
                .ColorIndex = 5
            End With
    Else
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(1, ��) = Null
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(1, ��).Borders(xlEdgeLeft).LineStyle = xlNone
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(2, ��).Borders(xlEdgeLeft).LineStyle = xlNone
    End If
    
   If �񖈂̓��� = 1 Then
       
       Select Case Weekday(��� - �񖈂̓��� + �� * �񖈂̓���)
        Case 1
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(4, ��) = "��"
        Case 2
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(4, ��) = "��"
        Case 3
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(4, ��) = "��"
        Case 4
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(4, ��) = "��"
        Case 5
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(4, ��) = "��"
        Case 6
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(4, ��) = "��"
        Case 7
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(4, ��) = "�y"
        End Select
    
        Select Case Weekday(��� - �񖈂̓��� + �� * �񖈂̓���)
        Case 1
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(4, ��).Select
            With Selection.Interior
                .Pattern = xlGray16
                .PatternColorIndex = 1
            End With
        Case 7
        Worksheets("�X�P�W���[���\").Range("�X�P�W���[���\").Cells(4, ��).Select
            With Selection.Interior
                .Pattern = xlGray16
                .PatternColorIndex = 1
            End With
        End Select
    Else
    End If
    
Next ��

Application.ScreenUpdating = True

End Sub


