VERSION 5.00
Begin VB.Form frm_apache_php 
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   2295
   End
   Begin kannagrafix.jcbutton jcbutton1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "BANKING"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frm_apache_php"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private dblWordID As Double

Private Sub jcbutton1_Click()
        ShellExecute FormMain.HWnd, "Open", "http://localhost/bank/", "", 0, SW_SHOWNORMAL

End Sub
