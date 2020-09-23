VERSION 5.00
Begin VB.UserControl TitleBarBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   3825
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   3825
   End
End
Attribute VB_Name = "TitleBarBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const Color_1 = "19,0,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608"

Private Const Color_2 = "19,0,14068374,11297860,11297860,11297860,11297860,11297860,11232067,11232067,11232067,11364166,11627854,11891542,12089436,12287843,12749426,13144955,13804430,14332833,789259"

Private Const Color_3 = "19,0,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608,6765608"

Public Event MouseDown(Button, Shift, X, Y)

Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, X As Integer, Y As Integer, Optional Gray As Boolean = True, Optional Brightness As Integer) As Integer
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
LoadBmpMenuLines 1, Color_1, 0, 0
LoadBmpMenuLines UserControl.ScaleWidth, Color_2, 1, 0
LoadBmpMenuLines 1, Color_1, UserControl.ScaleWidth - 1, 0
UserControl.Height = 20 * 15
End Function

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    LoadGUI
End Sub

Private Sub UserControl_Show()
    LoadGUI
End Sub

Function Caption(newCaption As String)
Label1.Caption = newCaption
Label2.Caption = newCaption
End Function
