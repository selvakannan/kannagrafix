VERSION 5.00
Begin VB.MDIForm FormMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808080&
   Caption         =   "kannagrafix automatic digital product  selling ; ADPS"
   ClientHeight    =   8280
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picProgBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   923
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7905
      Width           =   13845
   End
   Begin VB.PictureBox picLeftPane 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7905
      Left            =   0
      ScaleHeight     =   525
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   515
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7755
      Begin kannagrafix.jcbutton cmdOpen 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "contributors"
         CaptionEffects  =   0
      End
      Begin kannagrafix.jcbutton cmdSave 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "updates"
         CaptionEffects  =   0
      End
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2250
         Left            =   360
         ScaleHeight     =   150
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   600
         TabIndex        =   4
         Top             =   12000
         Visible         =   0   'False
         Width           =   9000
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000002&
         X1              =   5
         X2              =   142
         Y1              =   232
         Y2              =   232
      End
      Begin VB.Label lblCoordinates 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(X, Y)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   7560
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label lblImgSize 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Size: WidthxHeight"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D1B499&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   1845
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         X1              =   5
         X2              =   142
         Y1              =   160
         Y2              =   160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         X1              =   5
         X2              =   142
         Y1              =   104
         Y2              =   104
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   13845
      TabIndex        =   5
      Top             =   0
      Width           =   13845
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuDonate 
         Caption         =   "Support PhotoDemon with a small donation (thank you!)"
      End
      Begin VB.Menu MnuHelpSepBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCheckUpdates 
         Caption         =   "Check for &Updates..."
      End
      Begin VB.Menu MnuVisitWebsite 
         Caption         =   "&Visit the PhotoDemon Website"
      End
      Begin VB.Menu MnuEmailAuthor 
         Caption         =   "Submit Feedback..."
      End
      Begin VB.Menu MnuBugReport 
         Caption         =   "Submit Bug Report..."
      End
      Begin VB.Menu MnuHelpSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "&About PhotoDemon"
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'ShellExecute is preferable to VB's 'Shell' command; I use it for two items in the "Help" menu - sending
' me an email, and opening the kannagrafix website (currently just art4artist.com)
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Sub cmdOpen_Click()
   
    'Before we can display the "About" form, we need to paint the kannagrafix logo to it.
    Dim logoWidth As Long, logoHeight As Long
    Dim logoAspectRatio As Double
    
    logoWidth = FormMain.picLogo.ScaleWidth
    logoHeight = FormMain.picLogo.ScaleHeight
    logoAspectRatio = CDbl(logoWidth) / CDbl(logoHeight)
    
    FormAbout.Visible = False
   ' SetStretchBltMode FormAbout.hDC, STRETCHBLT_HALFTONE
   ' StretchBlt FormAbout.hDC, 0, 0, FormAbout.ScaleWidth, FormAbout.ScaleWidth / logoAspectRatio, FormMain.picLogo.hDC, 0, 0, logoWidth, logoHeight, vbSrcCopy
  '  FormAbout.Picture = FormAbout.Image
    
    'With the painting done, we can now display the form.
    FormAbout.Show 1, FormMain
End Sub

Private Sub cmdRedo_Click()

End Sub

Private Sub cmdSave_Click()
  'GitHub requires a login for submitting Issues; check for that first
    Dim msgReturn As VbMsgBoxResult
    
    msgReturn = MsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, kannagrafix needs to know where to send it." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbQuestion + vbApplicationModal + vbYesNo, "Thanks for making " & PROGRAMNAME & " better")
    
    'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the tannerhelland.com contact form
    If msgReturn = vbYes Then
        'Shell a browser window with the GitHub issue report form
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL
    Else
        'Shell a browser window with the tannerhelland.com kannagrafix contact form
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL
    End If

End Sub


Private Sub MnuBugReport_Click()
    
    'GitHub requires a login for submitting Issues; check for that first
    Dim msgReturn As VbMsgBoxResult
    
    msgReturn = MsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, kannagrafix needs to know where to send it." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbQuestion + vbApplicationModal + vbYesNo, "Thanks for making " & PROGRAMNAME & " better")
    
    'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the tannerhelland.com contact form
    If msgReturn = vbYes Then
        'Shell a browser window with the GitHub issue report form
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL
    Else
        'Shell a browser window with the tannerhelland.com kannagrafix contact form
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL
    End If

End Sub
Private Sub MnuEmailAuthor_Click()
    
    'Shell a browser window with the tannerhelland.com contact form
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL

End Sub
Private Sub MnuVisitWebsite_Click()
    'Nothing special here - just launch the default web browser with kannagrafix's page on tannerhelland.com
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL
End Sub


Private Sub MnuDonate_Click()
    'Launch the default web browser with the tannerhelland.com donation page
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL
End Sub


