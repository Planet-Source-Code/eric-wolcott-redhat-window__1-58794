VERSION 5.00
Begin VB.UserControl Skin 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   ControlContainer=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   Begin RedHatSkin.TitleBarButton TitleBarButton1 
      Height          =   300
      Left            =   4575
      TabIndex        =   1
      Top             =   0
      Width           =   30
      _ExtentX        =   582
      _ExtentY        =   529
      Hold_Button     =   3
   End
   Begin RedHatSkin.TitleBarButton TitleBarButton2 
      Height          =   300
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Width           =   30
      _ExtentX        =   556
      _ExtentY        =   529
      Hold_Button     =   1
   End
   Begin RedHatSkin.TitleBarButton TitleBarButton3 
      Height          =   300
      Left            =   4260
      TabIndex        =   3
      Top             =   0
      Width           =   30
      _ExtentX        =   556
      _ExtentY        =   529
      Hold_Button     =   2
   End
   Begin RedHatSkin.TitleBarButton TitleBarButton4 
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   529
   End
   Begin RedHatSkin.TitleBarBar TitleBarBar1 
      Height          =   300
      Left            =   330
      TabIndex        =   0
      Top             =   -15
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   529
   End
End
Attribute VB_Name = "Skin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const HTBOTTOMRIGHT = 17
Private Const HTBOTTOM = 15
Private Const HTBOTTOMLEFT = 16
Private Const HTLEFT = 10
Private Const HTRIGHT = 11
Private Const HTTOP = 12
Private Const HTTOPLEFT = 13
Private Const HTTOPRIGHT = 14

Private Const Color_1 = "15,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,12089951,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,10049852,13278344,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,10049852,13278344,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,10049852,13278344,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852"
Private Const Color_2 = "5,0,16711935,16711935,16711935,16711935,16711935,13278344,0,0,16711935,16711935,16711935,12089951,13278344,13278344,0,16711935,16711935,12089951,12089951,12089951,10049852,0,16711935,12089951,12089951,12089951,10049852,0,16711935,12089951,12089951,12089951,12089951,10049852,0,10049852,12089951,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951," & _
"10049852,0,10049852,12089951,10049852,10049852,10049852,0"
Private Const Color_3 = "5,10592673,16777215,15132390,15132390,12105912,0"
Private Const Color_4 = "5,12089951,13278344,13278344,13278344,12089951,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,10049852,13278344,12089951,12089951,10049852,0,13278344,12089951,12089951,12089951,10049852,0,12089951,12089951,12089951,10049852,0,16711935,12089951,12089951,12089951,10049852,0,16711935,12089951,10049852,10049852,0,16711935,16711935,10049852,0,0,16711935," & _
"16711935,16711935,0,16711935,16711935,16711935,16711935,16711935"
Private Const Color_5 = "15,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,12089951,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,13278344,10049852,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,10049852,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,12089951,0,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,10049852,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"

Public Event MouseDown(Button, Shift, X, Y)
Public Event MouseMove(Button, Shift, X, Y)
Private Resize As Boolean

Public IsGrey As Boolean

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
On Error Resume Next
TitleBarBar1.Caption Parent.Caption
ResetGUI
TitleBarButton4.Left = 0
TitleBarButton4.Top = 0
TitleBarBar1.Left = TitleBarButton4.Width - 5
TitleBarBar1.Top = 0
TitleBarButton1.Top = 0
TitleBarButton1.Left = UserControl.ScaleWidth - TitleBarButton1.Width '- 20
TitleBarButton3.Top = 0
TitleBarButton3.Left = TitleBarButton1.Left - TitleBarButton3.Width '- 19
TitleBarButton2.Top = 0
TitleBarButton2.Left = TitleBarButton3.Left - TitleBarButton2.Width '- 18
TitleBarBar1.Top = 0
TitleBarBar1.Left = TitleBarButton4.Width
TitleBarBar1.Width = TitleBarButton2.Left - TitleBarButton4.Width
DrawBorders
End Function

Function ResetGUI()
UserControl.Width = Parent.Width
UserControl.Height = Parent.Height
UserControl.Cls
End Function

Function DrawBorders()
UserControl.Line (0, 0)-(0, UserControl.ScaleHeight - 6), 0
UserControl.Line (1, 0)-(1, UserControl.ScaleHeight - 6), 16777215
UserControl.Line (2, 0)-(2, UserControl.ScaleHeight - 6), 15132390
UserControl.Line (3, 0)-(3, UserControl.ScaleHeight - 6), 15132390
UserControl.Line (4, 0)-(4, UserControl.ScaleHeight - 6), 15132390
UserControl.Line (5, 0)-(5, UserControl.ScaleHeight - 6), 10592673

UserControl.Line (UserControl.ScaleWidth - 6, 0)-(UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6), 10592673
UserControl.Line (UserControl.ScaleWidth - 5, 0)-(UserControl.ScaleWidth - 5, UserControl.ScaleHeight - 6), 16777215
UserControl.Line (UserControl.ScaleWidth - 4, 0)-(UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 6), 15132390
UserControl.Line (UserControl.ScaleWidth - 3, 0)-(UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 6), 15132390
UserControl.Line (UserControl.ScaleWidth - 2, 0)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 6), 12105912
UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 6), 0

LoadBmpMenuLines 1, Color_1, 0, UserControl.ScaleHeight - 15 - 6, IsGrey
LoadBmpMenuLines 1, Color_2, 0, UserControl.ScaleHeight - 6, IsGrey
LoadBmpMenuLines UserControl.ScaleWidth - 44, Color_3, 22, UserControl.ScaleHeight - 6, IsGrey
LoadBmpMenuLines 1, Color_4, UserControl.ScaleWidth - 22, UserControl.ScaleHeight - 6, IsGrey
LoadBmpMenuLines 1, Color_5, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 15 - 6, IsGrey
End Function

Private Sub TitleBarButton1_Click()
SubClassMe False
Unload Parent
End Sub

Private Sub TitleBarButton2_Click()
Parent.WindowState = 1
End Sub

Private Sub TitleBarButton3_Click()
If Parent.WindowState = 2 Then
Parent.WindowState = 0
Else
Parent.WindowState = 2
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If X < 7 Then
        ReleaseCapture
        If UserControl.ScaleHeight - Y < 22 Then
            Call SendMessage(Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMLEFT, 0)
        Else
            Call SendMessage(Parent.hwnd, WM_NCLBUTTONDOWN, HTLEFT, 0)
        End If
    ElseIf UserControl.ScaleWidth - X < 7 Then
        ReleaseCapture
        If UserControl.ScaleHeight - Y < 22 Then
            Call SendMessage(Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0)
        Else
            Call SendMessage(Parent.hwnd, WM_NCLBUTTONDOWN, HTRIGHT, 0)
        End If
    ElseIf UserControl.ScaleHeight - Y < 7 Then
        ReleaseCapture
        If X < 22 Then
            Call SendMessage(Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMLEFT, 0)
        ElseIf UserControl.ScaleWidth - X < 22 Then
            Call SendMessage(Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0)
        Else
            Call SendMessage(Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOM, 0)
        End If
    End If
End Sub

Private Sub TitleBarBar1_MouseDown(Button As Variant, Shift As Variant, X As Variant, Y As Variant)
        ReleaseCapture
        Call SendMessage(Parent.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If X < 7 Then
        If UserControl.ScaleHeight - Y < 22 Then
        Screen.MousePointer = 6
        Else
        Screen.MousePointer = 9
        End If
    ElseIf UserControl.ScaleWidth - X < 7 Then
        If UserControl.ScaleHeight - Y < 22 Then
            Screen.MousePointer = 8
        Else
            Screen.MousePointer = 9
        End If
    ElseIf UserControl.ScaleHeight - Y < 7 Then
        If X < 22 Then
        Screen.MousePointer = 6
        ElseIf UserControl.ScaleWidth - X < 22 Then
        Screen.MousePointer = 8
        Else
        Screen.MousePointer = 7
        End If
    Else
        Screen.MousePointer = 0
    End If
End Sub

Private Sub UserControl_Resize()
'LoadGUI
End Sub

Private Sub UserControl_Show()
LoadGUI
    Parent.ScaleMode = 3
    Parent.BackColor = &HFF00FF
    UserControl.Parent.BorderStyle = 0
    'UserControl.Parent.ShowInTaskbar = True
TransMyForm
End Sub

Function SubClassMe(TurnOn As Boolean)
Select Case TurnOn
Case True
TitleBarButton4.SubClass False
TitleBarButton3.SubClass False
TitleBarButton2.SubClass False
TitleBarButton1.SubClass False
Case False
TitleBarButton4.SubClass True
TitleBarButton3.SubClass True
TitleBarButton2.SubClass True
TitleBarButton1.SubClass True
End Select
End Function

Function TransMyForm()
    Dim Ret As Long
    Ret = GetWindowLong(Parent.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Parent.hwnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Parent.hwnd, &HFF00FF, 0, LWA_COLORKEY
End Function
