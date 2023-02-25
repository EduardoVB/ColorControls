VERSION 5.00
Begin VB.Form frmTestMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test ColorDialog"
   ClientHeight    =   4524
   ClientLeft      =   2208
   ClientTop       =   2160
   ClientWidth     =   8616
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4524
   ScaleWidth      =   8616
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Default"
      Height          =   1320
      Left            =   420
      TabIndex        =   17
      Top             =   240
      Width           =   2600
      Begin VB.CommandButton Command1 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   19
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
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
         TabIndex        =   18
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Style Box"
      Height          =   1320
      Left            =   3132
      TabIndex        =   14
      Top             =   240
      Width           =   2600
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
         TabIndex        =   16
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   15
         Top             =   456
         Width           =   1300
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Drop down (custom made)"
      Height          =   1320
      Left            =   5844
      TabIndex        =   11
      Top             =   1908
      Width           =   2600
      Begin VB.CommandButton Command5 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   13
         Top             =   456
         Width           =   1300
      End
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
         TabIndex        =   12
         Top             =   336
         Width           =   756
      End
      Begin VB.Label Label1 
         Caption         =   "Only with non-modal forms"
         Height          =   252
         Left            =   120
         TabIndex        =   20
         Top             =   1040
         Width           =   2352
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Big Size"
      Height          =   1320
      Left            =   5844
      TabIndex        =   8
      Top             =   264
      Width           =   2600
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
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
      Begin VB.CommandButton Command6 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   9
         Top             =   456
         Width           =   1300
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Compact"
      Height          =   1320
      Left            =   3132
      TabIndex        =   5
      Top             =   1908
      Width           =   2600
      Begin VB.CommandButton Command7 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   7
         Top             =   456
         Width           =   1300
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
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
         TabIndex        =   6
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Complete"
      Height          =   1320
      Left            =   420
      TabIndex        =   2
      Top             =   1908
      Width           =   2600
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
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
      Begin VB.CommandButton Command8 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   3
         Top             =   456
         Width           =   1300
      End
   End
   Begin VB.CommandButton cmdTestProperties 
      Caption         =   "Test all properties"
      Height          =   456
      Left            =   3480
      TabIndex        =   1
      Top             =   3780
      Width           =   2712
   End
   Begin VB.CommandButton cmdTestConfigurations 
      Caption         =   "Check more options"
      Height          =   456
      Left            =   480
      TabIndex        =   0
      Top             =   3780
      Width           =   2712
   End
End
Attribute VB_Name = "frmTestMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Sub cmdTestConfigurations_Click()
    frmTestConfigurations.Show , Me
End Sub

Private Sub cmdTestProperties_Click()
    frmTestProperties.Show , Me
End Sub

Private Sub Command1_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.Color = Picture1.BackColor
    If oDlg.Show Then
        Picture1.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command2_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.Style = cdStyleBox
    oDlg.Color = Picture2.BackColor
    If oDlg.Show Then
        Picture2.BackColor = oDlg.Color
    End If

End Sub

Private Sub Command5_Click()
    Dim iFrm As New frmDropDown
    
    iFrm.Color = Picture5.BackColor
    iFrm.SetTransparency
    iFrm.Move Me.Left + Frame5.Left + Command5.Left, Me.Top + Frame5.Top + Command5.Top + Command5.Height + (Me.Height - Me.ScaleHeight)
    iFrm.Show , Me
    Do While IsFormLoaded(iFrm)
        DoEvents
    Loop
    If iFrm.ColorSet Then
        Picture5.BackColor = iFrm.Color
    End If
End Sub

Private Sub Command6_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SizeBig = True
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

Private Function IsFormLoaded(nForm As Form) As Boolean
    Dim frm As Form
    
    For Each frm In Forms
        If frm Is nForm Then
            IsFormLoaded = True
            Exit For
        End If
    Next
End Function

