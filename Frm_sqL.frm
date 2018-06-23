VERSION 5.00
Begin VB.Form frm_apache_sqL 
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   12315
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Text            =   "Text4"
      Top             =   360
      Width           =   1935
   End
   Begin kannagrafix.jcbutton jcbutton1 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   360
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
      Caption         =   $"Frm_sqL.frx":0000
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frm_apache_sqL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
