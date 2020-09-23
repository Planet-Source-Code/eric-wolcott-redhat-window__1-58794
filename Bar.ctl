VERSION 5.00
Begin VB.UserControl Bar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin RedHatSkin.Bar_Button Bar_Button1 
      Height          =   1170
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      _extentx        =   1508
      _extenty        =   2064
      hold_caption    =   "Test"
   End
End
Attribute VB_Name = "Bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const Color_1 = "55,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,12895428,16119285"

Public Event Clicked(Index As Integer)
Public Event MouseOver(Index As Integer)
Public Event MouseOff(Index As Integer)

Function Button_Enable(Index As Integer, NewValue As Boolean)
Bar_Button1(Index).Enabled = NewValue
End Function

Function Button_Caption(Index As Integer, NewValue As String)
Bar_Button1(Index).Caption = NewValue
End Function

Function Button_Icon(Index As Integer, NewValue As Integer)
Bar_Button1(Index).Icon = NewValue
End Function

Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, X As Integer, Y As Integer, Optional Gray As Boolean = True, Optional Brightness As Integer = 0) As Integer
    If ColorPallet = "" Then Exit Function
    Dim PixCount
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
        If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
            If Colors(Count) = &HFF00FF And Brightness <> 0 Then
            
            Else
                If Gray = True Then
                UserControl.Line (X + CurrentColumn, Y + CurrentRow)-(X + CurrentColumn + Legnth, Y + CurrentRow), AdjustBrightness(Colors(Count), Brightness)
                Else
                UserControl.Line (X + CurrentColumn, Y + CurrentRow)-(X + CurrentColumn + Legnth, Y + CurrentRow), MakeGrey(Colors(Count))
                End If
            End If
        CurrentRow = CurrentRow + 1
    Next
    LoadBmpMenuLines = CurrentColumn
End Function

Function LoadGUI()
LoadBmpMenuLines UserControl.ScaleWidth, Color_1, 0, 0
UserControl.Height = 56 * 15
Dim X
For X = 0 To Bar_Button1.Count - 1
Bar_Button1(X).Height = UserControl.ScaleHeight - 2
Next
End Function

Private Sub Bar_Button1_Clicked(Index As Integer)
RaiseEvent Clicked(Index)
End Sub

Private Sub Bar_Button1_MouseOff(Index As Integer)
RaiseEvent MouseOff(Index)
End Sub

Private Sub Bar_Button1_MouseOver(Index As Integer)
RaiseEvent MouseOver(Index)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 0
End Sub

Private Sub UserControl_Resize()
LoadGUI
End Sub

Private Sub UserControl_Show()
LoadGUI
End Sub

Function SubClassMe(SubClass As Boolean)
Dim X
Select Case SubClass
Case True
    For X = 0 To Bar_Button1.Count - 1
    Bar_Button1(X).SubClassMe True
    Next
Case False
    For X = 0 To Bar_Button1.Count - 1
    Bar_Button1(X).SubClassMe False
    Next
End Select
End Function

Function AddButton(Caption As String, Icon As Integer, Enabled As Boolean)
Dim X: X = Bar_Button1.Count
If X = 1 Then
    Bar_Button1(X - 1).Caption = Caption
    Bar_Button1(X - 1).Icon = Icon
    Bar_Button1(X - 1).Enabled = Enabled
    Bar_Button1(X - 1).Visible = True
    Load Bar_Button1(X)
Else
    Bar_Button1(X - 1).Left = Bar_Button1(X - 2).Left + Bar_Button1(X - 2).Width
    Bar_Button1(X - 1).Top = 0
    Bar_Button1(X - 1).Caption = Caption
    Bar_Button1(X - 1).Icon = Icon
    Bar_Button1(X - 1).Enabled = Enabled
    Bar_Button1(X - 1).Visible = True
    Load Bar_Button1(X)
End If
End Function
