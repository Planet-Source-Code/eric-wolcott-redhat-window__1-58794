VERSION 5.00
Begin VB.UserControl TextBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "TextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const Color_1 = "23,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,14606046,10066329"

Private Const Color_2 = "23,10066329,14606046,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,10066329"

Private Const Color_3 = "23,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329"

Public Event Change()

Property Get Text() As String
Text = Text1.Text
End Property

Public Property Let Text(NewValue As String)
Text1.Text = NewValue
End Property


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
UserControl.Cls
LoadBmpMenuLines 1, Color_1, 0, 0
LoadBmpMenuLines UserControl.ScaleWidth, Color_2, 2, 0
LoadBmpMenuLines 1, Color_1, UserControl.ScaleWidth - 1, 0
UserControl.Height = 24 * 15

Text1.Left = 3
Text1.Width = UserControl.ScaleWidth - 6
End Function

Private Sub Text1_Change()
RaiseEvent Change
End Sub

Private Sub UserControl_Resize()
LoadGUI
End Sub

Private Sub UserControl_Show()
LoadGUI
End Sub


