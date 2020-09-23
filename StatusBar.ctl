VERSION 5.00
Begin VB.UserControl StatusBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   3795
   End
End
Attribute VB_Name = "StatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const Color_1 = "19,12895428,16119285,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390"

Private Const Color_2 = "19,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1," & _
"16777215,10066329,-1,-1,-1,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,-1,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,16777215,10066329,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1"

Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, X As Integer, Y As Integer, Optional Gray As Boolean = True, Optional Brightness As Integer = 0) As Integer
    If ColorPallet = "" Then Exit Function
    Dim PixCount
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
        If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
            If Colors(Count) <> -1 Then
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
LoadBmpMenuLines 1, Color_2, UserControl.ScaleWidth - 18, 0
Label1.Width = UserControl.Width
UserControl.Height = 20 * 15
End Function

Public Sub Caption(NewValue As String)
Label1.Caption = NewValue
End Sub

Private Sub UserControl_Resize()
LoadGUI
End Sub

Private Sub UserControl_Show()
LoadGUI
End Sub
