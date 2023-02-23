VERSION 5.00
Begin VB.Form frmTestConfigurations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color dialog configurations"
   ClientHeight    =   7056
   ClientLeft      =   816
   ClientTop       =   1200
   ClientWidth     =   14124
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7056
   ScaleWidth      =   14124
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame20 
      Caption         =   "Style Box"
      Height          =   1300
      Left            =   11254
      TabIndex        =   58
      Top             =   432
      Width           =   2600
      Begin VB.CommandButton Command20 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   60
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture20 
         Appearance      =   0  'Flat
         BackColor       =   &H00A8D59F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   59
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame19 
      Caption         =   "Style Box Big"
      Height          =   1300
      Left            =   11254
      TabIndex        =   55
      Top             =   2076
      Width           =   2600
      Begin VB.PictureBox Picture19 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   57
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   56
         Top             =   456
         Width           =   1300
      End
   End
   Begin VB.Frame Frame18 
      Caption         =   "Style Box Complete && Big"
      Height          =   1300
      Left            =   11254
      TabIndex        =   52
      Top             =   3744
      Width           =   2600
      Begin VB.PictureBox Picture18 
         Appearance      =   0  'Flat
         BackColor       =   &H00781012&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   54
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   53
         Top             =   456
         Width           =   1300
      End
   End
   Begin VB.Frame Frame17 
      Caption         =   "Style Box Simple"
      Height          =   1300
      Left            =   11254
      TabIndex        =   49
      Top             =   5412
      Width           =   2600
      Begin VB.CommandButton Command17 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   51
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture17 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   50
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "Remember position"
      Height          =   1300
      Left            =   5736
      TabIndex        =   45
      Top             =   5412
      Width           =   2600
      Begin VB.PictureBox Picture16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   47
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   46
         Top             =   456
         Width           =   1300
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "Simple && Big"
      Height          =   1300
      Left            =   312
      TabIndex        =   42
      Top             =   5412
      Width           =   2600
      Begin VB.CommandButton Command15 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   44
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture15 
         Appearance      =   0  'Flat
         BackColor       =   &H00496F24&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   43
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Keep open (floating window)"
      Height          =   1300
      Left            =   3024
      TabIndex        =   39
      Top             =   5412
      Width           =   2600
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   396
         Left            =   1080
         TabIndex        =   48
         Top             =   720
         Width           =   1300
      End
      Begin VB.PictureBox Picture14 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3FCFC&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   41
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Show"
         Height          =   396
         Left            =   1080
         TabIndex        =   40
         Top             =   300
         Width           =   1300
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Set a ""Context"" for user settings"
      Height          =   1300
      Left            =   8494
      TabIndex        =   36
      Top             =   5412
      Width           =   2600
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   38
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   37
         Top             =   456
         Width           =   1300
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Change color in one touch"
      Height          =   1300
      Left            =   8494
      TabIndex        =   33
      Top             =   3744
      Width           =   2600
      Begin VB.CommandButton Command12 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   35
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H006F246F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   34
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Simple"
      Height          =   1300
      Left            =   3024
      TabIndex        =   30
      Top             =   3744
      Width           =   2600
      Begin VB.CommandButton Command11 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   32
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         BackColor       =   &H006F5624&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   31
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Complete && Big"
      Height          =   1300
      Left            =   312
      TabIndex        =   27
      Top             =   3744
      Width           =   2600
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   29
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   28
         Top             =   456
         Width           =   1300
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Use HSL color system"
      Height          =   1300
      Left            =   5736
      TabIndex        =   24
      Top             =   3744
      Width           =   2600
      Begin VB.CommandButton Command9 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   26
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   25
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Complete"
      Height          =   1300
      Left            =   312
      TabIndex        =   21
      Top             =   2076
      Width           =   2600
      Begin VB.CommandButton Command8 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   23
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   22
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Compact"
      Height          =   1300
      Left            =   3024
      TabIndex        =   18
      Top             =   2076
      Width           =   2600
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   20
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   19
         Top             =   456
         Width           =   1300
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Hue selection"
      Height          =   1300
      Left            =   5736
      TabIndex        =   15
      Top             =   432
      Width           =   2600
      Begin VB.CommandButton Command6 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   17
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H0059C5F9&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   16
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Hide parameters section"
      Height          =   1300
      Left            =   5736
      TabIndex        =   12
      Top             =   2076
      Width           =   2600
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4CE33&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   14
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   13
         Top             =   456
         Width           =   1300
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Hide recent colors"
      Height          =   1300
      Left            =   8494
      TabIndex        =   9
      Top             =   2076
      Width           =   2600
      Begin VB.CommandButton Command4 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   11
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H0081FAF7&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   10
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Saturation selection"
      Height          =   1300
      Left            =   8494
      TabIndex        =   6
      Top             =   432
      Width           =   2600
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H002404FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   8
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   7
         Top             =   456
         Width           =   1300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Big"
      Height          =   1300
      Left            =   3024
      TabIndex        =   3
      Top             =   408
      Width           =   2600
      Begin VB.CommandButton Command2 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   5
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H0061DC7C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   4
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Standard"
      Height          =   1300
      Left            =   312
      TabIndex        =   0
      Top             =   408
      Width           =   2600
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   708
         ScaleWidth      =   756
         TabIndex        =   2
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   1
         Top             =   456
         Width           =   1300
      End
   End
End
Attribute VB_Name = "frmTestConfigurations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private WithEvents mDlg12 As ColorDialog
Attribute mDlg12.VB_VarHelpID = -1
Private WithEvents mDlg14 As ColorDialog
Attribute mDlg14.VB_VarHelpID = -1

Private Sub cmdHide_Click()
    If Not mDlg14 Is Nothing Then
        mDlg14.Hide
        Set mDlg14 = Nothing
    End If
End Sub

Private Sub Command1_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.Color = Picture1.BackColor
    If oDlg.Show Then
        Picture1.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command11_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SetSimple
    oDlg.RememberPosition = True
    
    oDlg.Color = Picture11.BackColor
    If oDlg.Show Then
        Picture11.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command12_Click()
    Set mDlg12 = New ColorDialog
  '  mDlg12.SetSimple
    mDlg12.RememberPosition = True
    
    mDlg12.Color = Picture12.BackColor
    mDlg12.Show
    Set mDlg12 = Nothing
End Sub

Private Sub Command13_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.RememberPosition = True
    oDlg.Context = "Command13"
    
    oDlg.Color = Picture13.BackColor
    If oDlg.Show Then
        Picture13.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command14_Click()
    Set mDlg14 = New ColorDialog
    mDlg14.SetSimple
    mDlg14.RememberPosition = True
    mDlg14.Modeless = True
    
    mDlg14.Color = Picture14.BackColor
    mDlg14.Show
End Sub

Private Sub Command15_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SetSimple
    oDlg.RememberPosition = True
    oDlg.SizeBig = True
    
    oDlg.Color = Picture15.BackColor
    If oDlg.Show Then
        Picture15.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command16_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.RememberPosition = True
    
    oDlg.Color = Picture16.BackColor
    If oDlg.Show Then
        Picture16.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command17_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.Style = cdStyleBox
    oDlg.SetSimple
    oDlg.RememberPosition = True
    
    oDlg.Color = Picture17.BackColor
    If oDlg.Show Then
        Picture17.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command18_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SetComplete
    oDlg.Style = cdStyleBox
    oDlg.SizeBig = True
    oDlg.RecentColorsColumns = 2
    
    oDlg.Color = Picture18.BackColor
    oDlg.FixedPalette = False
    If oDlg.Show Then
        Picture18.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command19_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.Style = cdStyleBox
    oDlg.SizeBig = True
    
    oDlg.Color = Picture19.BackColor
    If oDlg.Show Then
        Picture19.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command2_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SizeBig = True
    
    oDlg.Color = Picture2.BackColor
    If oDlg.Show Then
        Picture2.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command20_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.Style = cdStyleBox
    
    oDlg.Color = Picture20.BackColor
    If oDlg.Show Then
        Picture20.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command3_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SliderOptionsAvailable = cdSliderOptionsNone
    oDlg.SliderParameter = cdParameterSaturation
    
    oDlg.Color = Picture3.BackColor
    If oDlg.Show Then
        Picture3.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command4_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.RecentColorsColumns = 0

    oDlg.Color = Picture4.BackColor
    If oDlg.Show Then
        Picture4.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command5_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.ColorValuesSectionVisible = False

    oDlg.Color = Picture5.BackColor
    If oDlg.Show Then
        Picture5.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command6_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SliderOptionsAvailable = cdSliderOptionsNone
    oDlg.SliderParameter = cdParameterHue

    oDlg.Color = Picture6.BackColor
    If oDlg.Show Then
        Picture6.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command7_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SetCompact
    
    oDlg.Color = Picture7.BackColor
    If oDlg.Show Then
        Picture7.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command8_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SetComplete
    
    oDlg.Color = Picture8.BackColor
    If oDlg.Show Then
        Picture8.BackColor = oDlg.Color
    End If
    Set oDlg = Nothing
End Sub

Private Sub Command9_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.ColorSystem = cdColorSystemHSL
    
    oDlg.Color = Picture9.BackColor
    If oDlg.Show Then
        Picture9.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command10_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SetComplete
    oDlg.SizeBig = True
    oDlg.RecentColorsColumns = 2
    
    oDlg.Color = Picture10.BackColor
    oDlg.FixedPalette = False
    If oDlg.Show Then
        Picture10.BackColor = oDlg.Color
    End If
End Sub


Private Sub Form_Load()
    Dim ctl As Control
    Dim iRgn As Long
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is PictureBox Then
            iRgn = CreateRoundRectRgn(0, 0, Me.ScaleX(ctl.Width, Me.ScaleMode, vbPixels), Me.ScaleY(ctl.Height, Me.ScaleMode, vbPixels), 12, 12)
            SetWindowRgn ctl.hWnd, iRgn, True
            DeleteObject iRgn
        End If
    Next
End Sub

Private Sub mDlg12_Change()
    Picture12.BackColor = mDlg12.Color
End Sub

Private Sub mDlg14_Change()
    Picture14.BackColor = mDlg14.Color
End Sub

