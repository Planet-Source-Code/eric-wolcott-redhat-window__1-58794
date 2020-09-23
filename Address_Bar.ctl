VERSION 5.00
Begin VB.UserControl Address_Bar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin RedHatSkin.TextBox TextBox1 
      Height          =   360
      Left            =   1230
      TabIndex        =   0
      Top             =   90
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   375
      TabIndex        =   1
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "Address_Bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const Color_1 = "38,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,12895428,16777215"

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
UserControl.Height = 39 * 15
TextBox1.Width = UserControl.ScaleWidth - TextBox1.Left - 5
End Function


Private Sub UserControl_Resize()
LoadGUI
End Sub

Private Sub UserControl_Show()
LoadGUI
End Sub
