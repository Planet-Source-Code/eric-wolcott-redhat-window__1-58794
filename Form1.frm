VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   1080
   ClientTop       =   1440
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   578
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RedHatSkin.Skin Skin1 
      Height          =   8670
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   15293
      Begin RedHatSkin.StatusBar StatusBar1 
         Height          =   300
         Left            =   90
         TabIndex        =   12
         Top             =   8265
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   529
      End
      Begin RedHatSkin.Address_Bar Address_Bar1 
         Height          =   585
         Left            =   90
         TabIndex        =   11
         Top             =   1140
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   1032
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6540
         Left            =   105
         ScaleHeight     =   6540
         ScaleWidth      =   9510
         TabIndex        =   2
         Top             =   1725
         Width           =   9510
         Begin RedHatSkin.ctlIcon ctlIcon1 
            Height          =   1200
            Index           =   0
            Left            =   225
            TabIndex        =   3
            Top             =   285
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   2117
            Hold_Picture    =   "Form1.frx":0000
            Hold_Caption    =   "Font\nOptions"
            DoubleLine      =   -1  'True
         End
         Begin RedHatSkin.ctlIcon ctlIcon1 
            Height          =   1200
            Index           =   1
            Left            =   1620
            TabIndex        =   4
            Top             =   285
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   2117
            Hold_Picture    =   "Form1.frx":14D2
            Hold_Caption    =   $"Form1.frx":297C
         End
         Begin RedHatSkin.ctlIcon ctlIcon1 
            Height          =   1200
            Index           =   2
            Left            =   2910
            TabIndex        =   5
            Top             =   285
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   2117
            Hold_Picture    =   "Form1.frx":2986
            Hold_Caption    =   "Network\nOptions"
            DoubleLine      =   -1  'True
         End
         Begin RedHatSkin.ctlIcon ctlIcon1 
            Height          =   1200
            Index           =   3
            Left            =   4275
            TabIndex        =   6
            Top             =   285
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   2117
            Hold_Picture    =   "Form1.frx":4448
            Hold_Caption    =   "PC\nConfig"
            DoubleLine      =   -1  'True
         End
         Begin RedHatSkin.ctlIcon ctlIcon1 
            Height          =   1200
            Index           =   4
            Left            =   5580
            TabIndex        =   7
            Top             =   285
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   2117
            Hold_Picture    =   "Form1.frx":591A
            Hold_Caption    =   "Window\nFX"
            DoubleLine      =   -1  'True
         End
         Begin RedHatSkin.ctlIcon ctlIcon1 
            Height          =   1200
            Index           =   5
            Left            =   6765
            TabIndex        =   8
            Top             =   285
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   2117
            Hold_Picture    =   "Form1.frx":717C
            Hold_Caption    =   "Font\nOptions"
            DoubleLine      =   -1  'True
         End
         Begin RedHatSkin.ctlIcon ctlIcon1 
            Height          =   1200
            Index           =   6
            Left            =   7845
            TabIndex        =   9
            Top             =   285
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   2117
            Hold_Picture    =   "Form1.frx":86F2
            Hold_Caption    =   "Menu\nOptions"
            DoubleLine      =   -1  'True
         End
         Begin RedHatSkin.ctlIcon ctlIcon1 
            Height          =   1200
            Index           =   7
            Left            =   240
            TabIndex        =   10
            Top             =   1740
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   2117
            Hold_Picture    =   "Form1.frx":9DF4
            Hold_Caption    =   "Media\nOptions"
            DoubleLine      =   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   3390
            Left            =   5865
            Picture         =   "Form1.frx":B166
            Top             =   3915
            Width           =   3315
         End
      End
      Begin RedHatSkin.Bar Bar1 
         Height          =   840
         Left            =   90
         TabIndex        =   1
         Top             =   300
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   1482
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bar1_Clicked(Index As Integer)
    StatusBar1.Caption "ToolBar Button #" & Index & " was clicked."
End Sub

Private Sub Bar1_MouseOff(Index As Integer)
    StatusBar1.Caption "ToolBar Button #" & Index & " no longer has mouse over."
End Sub

Private Sub Bar1_MouseOver(Index As Integer)
    StatusBar1.Caption "ToolBar Button #" & Index & " has mouse over."
End Sub

Private Sub ctlIcon1_DblClick(Index As Integer)
    StatusBar1.Caption "Icon #" & Index & " has been Double Clicked."
End Sub

Private Sub ctlIcon1_GotFocus(Index As Integer)
    ctlIcon1(Index).SelectMe
End Sub

Private Sub ctlIcon1_LostFocus(Index As Integer)
    ctlIcon1(Index).Clear
    ctlIcon1(Index).LoadGUI
End Sub

Private Sub ctlIcon1_MouseDown(Index As Integer, Button As Variant, Shift As Variant, X As Variant, Y As Variant)
    StatusBar1.Caption "Icon #" & Index & " has button(" & Button & ") down."
End Sub

Private Sub ctlIcon1_MouseUP(Index As Integer, Button As Variant, Shift As Variant, X As Variant, Y As Variant)
    StatusBar1.Caption "Icon #" & Index & " has button(" & Button & ") up."
End Sub

Private Sub Form_Load()
    Skin1.Top = 0
    Skin1.Left = 0
    Skin1.SubClassMe True
    Bar1.AddButton "Back", 0, True
    Bar1.AddButton "Foward", 1, True
    Bar1.AddButton "Up", 2, True
    Bar1.AddButton "Refresh", 3, True
    Bar1.AddButton "Home", 4, True
    Bar1.SubClassMe True
    StatusBar1.Caption "8 Objects Loaded"
End Sub

Private Sub Form_Resize()
    Skin1.IsGrey = True
    Skin1.LoadGUI
    Bar1.Top = 300
    Bar1.Left = 90
    Bar1.Width = Me.Width - 180
    Address_Bar1.Width = Bar1.Width
    Picture1.Width = Me.Width - 200
    StatusBar1.Left = 90
    StatusBar1.Width = Me.Width - 180
    StatusBar1.Top = Me.Height - StatusBar1.Height - 90 '1545
    Picture1.Height = Me.Height - Picture1.Top - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Skin1.SubClassMe False
    Bar1.SubClassMe False
End Sub


