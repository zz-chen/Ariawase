Attribute VB_Name = "Module1"
Option Explicit

Public 報告日 As Date       '報告日をセット
Public 基準日 As Date       'スケジュール表のＦＲＯＭ日付
Public 基準日2 As Date      'スケジュール表のＴＯ日付
Public 日程表タイプ         'D：日次，W：週次，M：月次
Public Const 表開始行 = 1
Public Const 表開始列 = 1
Public 表終了行
Public 表終了列
Public 走査開始列
Public Const 走査開始行 = 1
Public 走査終了列
Public 走査終了行
Public 開始日列
Public 終了日列
Public 列
Public 行
Public 実績行
Public 実績日 As Date
Public 列毎の日数
Public Const 背景文字 = ""
Public WBS完了 As Boolean
Public 表範囲外 As Boolean
Public 前行X
Public 前行Y
Public 報告日X
Public 表開始列X
Public 表終了列X
Public 予定線率
Public 実績線率
Public 書式設定_Error
Public 書式設定完了 As Boolean


'***********
Sub 初期化()
'***********

報告日 = Worksheets("マクロ").Range("報告日")
基準日 = Worksheets("マクロ").Range("基準日")
基準日2 = Worksheets("マクロ").Range("基準日2")

If Worksheets("マクロ").Range("日程表タイプ") = "D" Then
   列毎の日数 = 1
Else
   列毎の日数 = 7
End If

予定線率 = Worksheets("マクロ").Range("予定線率")
実績線率 = Worksheets("マクロ").Range("実績線率")

'2004/12/05追加 begin

If Worksheets("マクロ").Range("進捗指標単位") = "Day" Then
    Worksheets("スケジュール表").Range("進捗指標見出し") = "進捗指標（日数）"
    Worksheets("スケジュール表").Range("進捗指標").Select
            Selection.NumberFormatLocal = "+#,###;[赤]-#,###"
Else
    Worksheets("スケジュール表").Range("進捗指標見出し") = "進捗指標（％）"
    Worksheets("スケジュール表").Range("進捗指標").Select
            Selection.NumberFormatLocal = "0%"
End If

'2004/12/05追加 end

表終了行 = Range("描画領域").Rows.Count
表終了列 = Range("描画領域").Columns.Count
走査終了行 = Range("描画領域").Rows.Count
走査終了列 = Range("描画領域").Columns.Count

End Sub

'**********************
Sub スケジュール表描画()
Attribute スケジュール表描画.VB_Description = "マクロ記録日 : 2004/3/5  ユーザー名 :  "
Attribute スケジュール表描画.VB_ProcData.VB_Invoke_Func = " \n14"
'**********************

If ActiveSheet.Name <> "スケジュール表" Then
    MsgBox "シート「スケジュール表」を開いてから実行してください。"
    Exit Sub
Else
End If

初期化

Dim REPLY

Do While REPLY <> 6
    On Error GoTo Error_Exit
    報告日 = InputBox("報告日をyy/mm/dd形式で入力してください", "報告日の入力", 報告日)
    If 報告日 > 基準日 And 報告日 < 基準日 + (表終了列 - 表開始列) * 列毎の日数 Then
        Range("報告日") = 報告日
        REPLY = MsgBox("報告日 " & 報告日 & " のスケジュール表を作成しますか？", vbQuestion + vbYesNo)
    Else
        MsgBox "報告日が描画範囲外です。"
    End If
Loop

'オートシェイプ，線の消去

On Error Resume Next
ActiveSheet.Rectangles.Delete
ActiveSheet.Lines.Delete

'表開始列Ｘ座標を入手

走査開始列 = 表開始列

Call GetCellXYext(Worksheets("スケジュール表").Range("描画領域").Row, _
     走査開始列 + Worksheets("スケジュール表").Range("描画領域").Column, 前行X, 前行Y, Pos:=1)

表開始列X = 前行X

'表終了列Ｘ座標を入手

走査開始列 = 表終了列

Call GetCellXYext(Worksheets("スケジュール表").Range("描画領域").Row, _
     走査開始列 + Worksheets("スケジュール表").Range("描画領域").Column, 前行X, 前行Y, Pos:=1)

表終了列X = 前行X

'報告日Ｘ座標を入手

走査開始列 = Application.WorksheetFunction.RoundUp((報告日 - (基準日 - 1 * 列毎の日数)) / 列毎の日数, 0)

Call GetCellXYext(Worksheets("スケジュール表").Range("描画領域").Row, _
     走査開始列 + Worksheets("スケジュール表").Range("描画領域").Column, 前行X, 前行Y, Pos:=1)
報告日X = 前行X

'描画
For 行 = 表開始行 To 表終了行 Step 1
    行走査
Next 行

Exit Sub

Error_Exit:
MsgBox ("日付が規定外です。メニューからやり直してください。")
End Sub

'***********
Sub 行走査()
'***********

'予定日なし
If Worksheets("スケジュール表").Range("開始日").Cells(行, 1) = 0 Or _
   Worksheets("スケジュール表").Range("終了日").Cells(行, 1) = 0 Then
    
    空白行の稲妻描画
    GoTo exit行走査

End If


予定:

走査開始列 = Application.WorksheetFunction.RoundUp _
            ((Worksheets("スケジュール表").Range("開始日").Cells(行, 1) - (基準日 - 1 * 列毎の日数)) / 列毎の日数, 0)

走査終了列 = Application.WorksheetFunction.RoundUp _
            ((Worksheets("スケジュール表").Range("終了日").Cells(行, 1) - (基準日 - 1 * 列毎の日数)) / 列毎の日数, 0)

表範囲外 = False

If 走査終了列 < 表開始列 Then
    
    表範囲外 = True
    走査開始列 = 1
    走査終了列 = 1

Else

    If 走査開始列 > 表終了列 Then
        
        表範囲外 = True
        走査開始列 = 表終了列
        走査終了列 = 表終了列
        
    Else
        
        走査開始列 = Application.WorksheetFunction.Max(走査開始列, 表開始列)
        走査終了列 = Application.WorksheetFunction.Min(走査終了列, 表終了列)

    End If

End If

予定描画

実績:

If Worksheets("スケジュール表").Range("実績開始日").Cells(行, 1) = 0 Then GoTo exit行走査

実績日 = 0

If Worksheets("スケジュール表").Range("実績終了日").Cells(行, 1) <> 0 Then
   
   実績日 = Application.WorksheetFunction.Max( _
            Worksheets("スケジュール表").Range("実績終了日").Cells(行, 1), _
            Worksheets("スケジュール表").Range("終了日").Cells(行, 1))

Else

   If Worksheets("マクロ").Range("進捗指標単位") = "Day" Then
   
      If Worksheets("スケジュール表").Range("終了日").Cells(行, 1) < 報告日 Then
      
        実績日 = Worksheets("スケジュール表").Range("終了日").Cells(行, 1) + _
                 Worksheets("スケジュール表").Range("進捗指標").Cells(行, 1)
        
      Else
      
        実績日 = 報告日 + Worksheets("スケジュール表").Range("進捗指標").Cells(行, 1)
    
      End If

   Else
   
      実績日 = Worksheets("スケジュール表").Range("開始日").Cells(行, 1) - 1 + _
               Application.WorksheetFunction.RoundUp( _
                 ( _
                   Worksheets("スケジュール表").Range("終了日").Cells(行, 1) + 1 - _
                   Worksheets("スケジュール表").Range("開始日").Cells(行, 1) _
                 ) _
                 * Abs(Worksheets("スケジュール表").Range("進捗指標").Cells(行, 1)) _
                 , 0)
   End If
   
End If

If Worksheets("マクロ").Range("進捗指標単位") = "Day" Then
   
   走査開始列 = Application.WorksheetFunction.RoundUp _
               ((Worksheets("スケジュール表").Range("実績開始日").Cells(行, 1) - _
               (基準日 - 1 * 列毎の日数)) / 列毎の日数, 0)
               
Else

   走査開始列 = Application.WorksheetFunction.RoundUp _
               ((Worksheets("スケジュール表").Range("開始日").Cells(行, 1) - _
               (基準日 - 1 * 列毎の日数)) / 列毎の日数, 0)
               
End If

走査終了列 = Application.WorksheetFunction.RoundUp _
            ((実績日 - (基準日 - 1 * 列毎の日数)) / 列毎の日数, 0)

表範囲外 = False

If 走査終了列 < 表開始列 Then
    
    表範囲外 = True
    走査開始列 = 0
    走査終了列 = 0

Else
    
    If 走査開始列 > 表終了列 Then
    
        表範囲外 = True
        走査開始列 = 表終了列
        走査終了列 = 表終了列
          
    Else
    
        走査開始列 = Application.WorksheetFunction.Max(走査開始列, 表開始列)
        走査終了列 = Application.WorksheetFunction.Min(走査終了列, 表終了列)
    
    End If

End If

WBS完了 = False

If Worksheets("スケジュール表").Range("実績終了日").Cells(行, 1) <> 0 Then
   WBS完了 = True
End If

実績描画

exit行走査:

End Sub
Sub 空白行の稲妻描画()
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
    
Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査開始列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X1, Y1, Pos:=1)
'
Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査終了列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X2, Y2, Pos:=2)

Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査終了列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X3, Y3, Pos:=3)

Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査開始列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X4, Y4, Pos:=4)
              
With ActiveSheet.Shapes.AddLine(前行X, 前行Y, 報告日X, Y2).Select
    Selection.ShapeRange.Fill.Transparency = 0#
    Selection.ShapeRange.Line.Weight = 2.25
    Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
    Selection.ShapeRange.Line.Style = msoLineSingle
    Selection.ShapeRange.Line.Transparency = 0#
    Selection.ShapeRange.Line.Visible = msoTrue
    Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
    Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
End With
   
   前行X = 報告日X
   前行Y = Y2
   
With ActiveSheet.Shapes.AddLine(前行X, 前行Y, 報告日X, Y3).Select
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
Sub 予定描画()
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
    
Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査開始列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X1, Y1, Pos:=1)
'
Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査終了列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X2, Y2, Pos:=2)

Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査終了列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X3, Y3, Pos:=3)

Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査開始列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X4, Y4, Pos:=4)
     
If 表範囲外 Then

Else

    With ActiveSheet.Shapes.AddShape(msoShapeRectangle, X1, Y1 + (Y3 - Y2) * (1 - 予定線率) / 2, X2 - X1, (Y3 - Y2) * 予定線率)
         .Fill.Solid
         .Fill.ForeColor.SchemeColor = 44
    End With

End If

Set Rng = Nothing

予定稲妻描画:

With ActiveSheet.Shapes.AddLine(前行X, 前行Y, 報告日X, Y2).Select
    Selection.ShapeRange.Fill.Transparency = 0#
    Selection.ShapeRange.Line.Weight = 2.25
    Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
    Selection.ShapeRange.Line.Style = msoLineSingle
    Selection.ShapeRange.Line.Transparency = 0#
    Selection.ShapeRange.Line.Visible = msoTrue
    Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
    Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
End With

前行X = 報告日X
前行Y = Y2

If Worksheets("スケジュール表").Range("実績開始日").Cells(行, 1) = 0 Then

    With ActiveSheet.Shapes.AddLine(前行X, 前行Y, Application.WorksheetFunction.Min(X1, 報告日X), Y1 + (Y3 - Y2) * (1 - 予定線率) / 2).Select
    Selection.ShapeRange.Fill.Transparency = 0#
    Selection.ShapeRange.Line.Weight = 2.25
    Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
    Selection.ShapeRange.Line.Style = msoLineSingle
    Selection.ShapeRange.Line.Transparency = 0#
    Selection.ShapeRange.Line.Visible = msoTrue
    Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
    Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
   End With
   
   前行X = Application.WorksheetFunction.Min(X1, 報告日X)
   前行Y = Y1 + (Y3 - Y2) * (1 - 予定線率) / 2
   
   If 表範囲外 And 走査開始列 = 表開始列 Then
      
   Else
    
      With ActiveSheet.Shapes.AddLine(前行X, 前行Y, Application.WorksheetFunction.Min(X4, 報告日X), Y1 + (Y3 - Y2) * (1 - 予定線率) / 2 + (Y3 - Y2) * 予定線率).Select
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
         
   前行X = Application.WorksheetFunction.Min(X4, 報告日X)
   前行Y = Y1 + (Y3 - Y2) * (1 - 予定線率) / 2 + (Y3 - Y2) * 予定線率
End If

End Sub

'************
Sub 実績描画()
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
    
Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査開始列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X1, Y1, Pos:=1)

Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査終了列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X2, Y2, Pos:=2)

Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査終了列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X3, Y3, Pos:=3)

Call GetCellXYext(行 - 1 + Worksheets("スケジュール表").Range("描画領域").Row, _
     走査開始列 - 1 + Worksheets("スケジュール表").Range("描画領域").Column, X4, Y4, Pos:=4)
     
If 表範囲外 Then

Else

    If WBS完了 Then
        With ActiveSheet.Shapes.AddShape(msoShapeRectangle, X1, Y1 + (Y3 - Y2) * 0.4, X2 - X1, (Y3 - Y2) * 実績線率)
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
        With ActiveSheet.Shapes.AddShape(msoShapeRectangle, X1, Y1 + (Y3 - Y2) * 0.4, X2 - X1, (Y3 - Y2) * 実績線率)
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

実績稲妻描画:

With ActiveSheet.Shapes.AddLine(前行X, 前行Y, 報告日X, Y1).Select
    Selection.ShapeRange.Fill.Transparency = 0#
    Selection.ShapeRange.Line.Weight = 2.25
    Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
    Selection.ShapeRange.Line.Style = msoLineSingle
    Selection.ShapeRange.Line.Transparency = 0#
    Selection.ShapeRange.Line.Visible = msoTrue
    Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
    Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
End With

前行X = 報告日X
前行Y = Y1

If WBS完了 = True Then
    With ActiveSheet.Shapes.AddLine _
        (前行X, 前行Y, Application.WorksheetFunction.Max(X2, 報告日X), Y1 + (Y3 - Y2) * 実績線率).Select
        Selection.ShapeRange.Fill.Transparency = 0#
        Selection.ShapeRange.Line.Weight = 2.25
        Selection.ShapeRange.Line.DashStyle = msoLineSquareDot
        Selection.ShapeRange.Line.Style = msoLineSingle
        Selection.ShapeRange.Line.Transparency = 0#
        Selection.ShapeRange.Line.Visible = msoTrue
        Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
        Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
    End With
    
    前行X = Application.WorksheetFunction.Max(X2, 報告日X)
    前行Y = Y1 + (Y3 - Y2) * 実績線率
    
Else
    With ActiveSheet.Shapes.AddLine(前行X, 前行Y, X2, Y1 + (Y3 - Y2) * 実績線率).Select
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
    
    前行X = X2
    前行Y = Y1 + (Y3 - Y2) * 実績線率

End If

If 表範囲外 And 走査開始列 = 表終了列 _
    Or _
   表範囲外 And Worksheets("スケジュール表").Range("実績終了日").Cells(行, 1) = 0 Then

Else

    If WBS完了 = True Then
        With ActiveSheet.Shapes.AddLine _
        (前行X, 前行Y, _
        Application.WorksheetFunction.Max(X2, 報告日X), Y1 + (Y3 - Y2) * 実績線率 + (Y3 - Y2) * 予定線率).Select
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
        With ActiveSheet.Shapes.AddLine(前行X, 前行Y, X2, Y1 + (Y3 - Y2) * 実績線率 + (Y3 - Y2) * 予定線率).Select
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
    
If WBS完了 = True Then

    前行X = Application.WorksheetFunction.Max(X2, 報告日X)
    前行Y = Y1 + (Y3 - Y2) * 実績線率 + (Y3 - Y2) * 予定線率
Else
    前行X = X2
    前行Y = Y1 + (Y3 - Y2) * 実績線率 + (Y3 - Y2) * 予定線率
End If

End Sub

'描画座標を求める
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
'***列番号（ColNo）から列名を得る
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
Sub スケジュール書式設定画面()
'************************

If ActiveSheet.Name <> "スケジュール表" Then
    MsgBox "シート「スケジュール表」を開いてから実行してください。"
    Exit Sub
Else
End If

書式設定完了 = False

Do Until 書式設定完了 = True

書式設定画面.基準年 _
    = Application.WorksheetFunction.Fixed(Year(Worksheets("マクロ").Range("基準日")), 0, True)
書式設定画面.基準月 _
    = Application.WorksheetFunction.Fixed(Month(Worksheets("マクロ").Range("基準日")), 0, True)
書式設定画面.基準日 _
    = Application.WorksheetFunction.Fixed(Day(Worksheets("マクロ").Range("基準日")), 0, True)
    
書式設定画面.基準年2 _
    = Application.WorksheetFunction.Fixed(Year(Worksheets("マクロ").Range("基準日2")), 0, True)
書式設定画面.基準月2 _
    = Application.WorksheetFunction.Fixed(Month(Worksheets("マクロ").Range("基準日2")), 0, True)
書式設定画面.基準日2 _
    = Application.WorksheetFunction.Fixed(Day(Worksheets("マクロ").Range("基準日2")), 0, True)
    
If Worksheets("マクロ").Range("日程表タイプ") = "D" Then
    書式設定画面.日次 = 1
Else
    書式設定画面.週次 = 7
End If

書式設定画面.予定線率 = Worksheets("マクロ").Range("予定線率") * 100
書式設定画面.実績線率 = Worksheets("マクロ").Range("実績線率") * 100

'2004/12/05追加 begin

If Worksheets("マクロ").Range("進捗指標単位") = "Day" Then
    書式設定画面.日数進捗 = 1
Else
    書式設定画面.百分率進捗 = 1
End If

'2004/12/05追加 end

書式設定_Error = 9

書式設定画面.Show
       
Select Case 書式設定_Error
        
        Case 0
        書式設定完了 = True
        Case 1
        MsgBox "表開始日または表終了日が規定外です。"
        Case 2
        MsgBox "予定線幅が規定外です。"
        Case 3
        MsgBox "実績線幅が規定外です。"
        Case 4
        MsgBox "実績線幅は予定線幅より小さくしてください。"
        Case 5
        MsgBox "スケジュール表の最小列数は20を下限としています。列数が20以上になる表示期間を設定してください｡"
        Case 6
        MsgBox "スケジュール表の列数がExcelの最大値（256）を超えます。列数が256以下になる表示期間を設定してください。"
        Case 9
        Exit Sub
End Select
    
Loop

Application.ScreenUpdating = False

Dim 旧期間, 新期間, 期間差

新期間 = Application.WorksheetFunction.RoundUp( _
        (Worksheets("マクロ").Range("基準日2") - Worksheets("マクロ").Range("基準日")) _
       / Worksheets("マクロ").Range("列毎の日数"), 0)
旧期間 = Worksheets("マクロ").Range("表終了列") - Worksheets("マクロ").Range("表開始列")

期間差 = 新期間 - 旧期間

If 期間差 = 0 Then
Else
    If 期間差 > 0 Then
        Range(Columns(GetColName(Worksheets("マクロ").Range("表開始列") + 1)), _
              Columns(GetColName(Worksheets("マクロ").Range("表開始列") + 期間差))).Select
        Selection.Insert Shift:=xlToRight
    Else
        Range(Columns(GetColName(Worksheets("マクロ").Range("表開始列") + 1)), _
              Columns(GetColName(Worksheets("マクロ").Range("表開始列") + 期間差 * -1))).Select
        Selection.Delete Shift:=xlToLeft
    End If
End If

初期化

On Error Resume Next

ActiveSheet.Rectangles.Delete
ActiveSheet.Lines.Delete

Worksheets("マクロ").Range("基準日") = 基準日

Worksheets("スケジュール表").Range("スケジュール表").Range(Cells(4, 1), Cells(4, 表終了列)).Select
    With Selection.Interior
        .ColorIndex = 0
        .PatternColorIndex = 1
    End With
    
Worksheets("スケジュール表").Range("スケジュール表").Range(Cells(4, 1), Cells(4, 表終了列)).ClearContents

Dim 月

For 列 = 表開始列 To 表終了列
    
    Worksheets("スケジュール表").Range("スケジュール表").Cells(3, 列) = _
    "'" & Application.WorksheetFunction.Fixed(Day(基準日 - 列毎の日数 + 列 * 列毎の日数), 0, True)
    
    If 月 <> Application.WorksheetFunction.Fixed(Month(基準日 - 列毎の日数 + 列 * 列毎の日数), 0, True) Then
        月 = Application.WorksheetFunction.Fixed(Month(基準日 - 列毎の日数 + 列 * 列毎の日数), 0, True)
        Worksheets("スケジュール表").Range("スケジュール表").Cells(1, 列) = _
        "'" & Application.WorksheetFunction.Fixed(Month(基準日 - 列毎の日数 + 列 * 列毎の日数), 0, True) & "月"
        Worksheets("スケジュール表").Range("スケジュール表").Range(Cells(1, 列), Cells(2, 列)).Select
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
        Worksheets("スケジュール表").Range("スケジュール表").Cells(1, 列) = Null
        Worksheets("スケジュール表").Range("スケジュール表").Cells(1, 列).Borders(xlEdgeLeft).LineStyle = xlNone
        Worksheets("スケジュール表").Range("スケジュール表").Cells(2, 列).Borders(xlEdgeLeft).LineStyle = xlNone
    End If
    
   If 列毎の日数 = 1 Then
       
       Select Case Weekday(基準日 - 列毎の日数 + 列 * 列毎の日数)
        Case 1
        Worksheets("スケジュール表").Range("スケジュール表").Cells(4, 列) = "日"
        Case 2
        Worksheets("スケジュール表").Range("スケジュール表").Cells(4, 列) = "月"
        Case 3
        Worksheets("スケジュール表").Range("スケジュール表").Cells(4, 列) = "火"
        Case 4
        Worksheets("スケジュール表").Range("スケジュール表").Cells(4, 列) = "水"
        Case 5
        Worksheets("スケジュール表").Range("スケジュール表").Cells(4, 列) = "木"
        Case 6
        Worksheets("スケジュール表").Range("スケジュール表").Cells(4, 列) = "金"
        Case 7
        Worksheets("スケジュール表").Range("スケジュール表").Cells(4, 列) = "土"
        End Select
    
        Select Case Weekday(基準日 - 列毎の日数 + 列 * 列毎の日数)
        Case 1
        Worksheets("スケジュール表").Range("スケジュール表").Cells(4, 列).Select
            With Selection.Interior
                .Pattern = xlGray16
                .PatternColorIndex = 1
            End With
        Case 7
        Worksheets("スケジュール表").Range("スケジュール表").Cells(4, 列).Select
            With Selection.Interior
                .Pattern = xlGray16
                .PatternColorIndex = 1
            End With
        End Select
    Else
    End If
    
Next 列

Application.ScreenUpdating = True

End Sub


