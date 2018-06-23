VERSION 5.00
Begin VB.MDIForm FormMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808080&
   Caption         =   "kannagrafix automatic digital product  selling ; ADPS"
   ClientHeight    =   10320
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   16305
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
      ScaleWidth      =   1087
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9945
      Width           =   16305
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
      Height          =   9945
      Left            =   0
      ScaleHeight     =   661
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   515
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7755
      Begin kannagrafix.jcbutton jcbutton3 
         Height          =   375
         Left            =   3000
         TabIndex        =   36
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
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
         Caption         =   "Remote server database connect"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton22 
         Height          =   375
         Left            =   1920
         TabIndex        =   32
         Top             =   2520
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "vb6_Access_Projects"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton21 
         Height          =   375
         Left            =   1920
         TabIndex        =   31
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "HTML_CSS_Projects"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton20 
         Height          =   375
         Left            =   1920
         TabIndex        =   30
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "APACHE_SQL_Projects"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton19 
         Height          =   375
         Left            =   1920
         TabIndex        =   29
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "APACHE_PHP_Projects"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton18 
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   3960
         Width           =   1815
         _ExtentX        =   3201
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
         Caption         =   "Everything"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton5 
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "pictures"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton9 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
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
         Caption         =   "NOTE PAD"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton2 
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Developers links;"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton1 
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
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
         Caption         =   "send mail"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
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
         TabIndex        =   2
         Top             =   12000
         Visible         =   0   'False
         Width           =   9000
      End
      Begin kannagrafix.jcbutton jcbutton9 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
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
         Caption         =   "DOUBLE KILLER"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton9 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   1815
         _ExtentX        =   3201
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
         Caption         =   "Localhost-Dashboard"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton9 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
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
         Caption         =   "xampp"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton9 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   1815
         _ExtentX        =   3201
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
         Caption         =   "xampp-phpmyadmin"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton9 
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   1815
         _ExtentX        =   3201
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
         Caption         =   "xampp-database"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton9 
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   1815
         _ExtentX        =   3201
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
         Caption         =   "HELP PAD"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton6 
         Height          =   375
         Left            =   4560
         TabIndex        =   17
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "VB6"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton7 
         Height          =   375
         Left            =   4560
         TabIndex        =   18
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "PHP"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton8 
         Height          =   375
         Left            =   4560
         TabIndex        =   19
         Top             =   2520
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "HTML_JAVA"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton10 
         Height          =   375
         Left            =   4560
         TabIndex        =   20
         Top             =   2880
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "LocalHost"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton11 
         Height          =   375
         Left            =   4560
         TabIndex        =   21
         Top             =   3240
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Downloads"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton12 
         Height          =   375
         Left            =   4560
         TabIndex        =   22
         Top             =   3600
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "pictures"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton13 
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         Top             =   3960
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "pictures"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton14 
         Height          =   375
         Left            =   4560
         TabIndex        =   24
         Top             =   4320
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "pictures"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton15 
         Height          =   375
         Left            =   4560
         TabIndex        =   25
         Top             =   4680
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "pictures"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton16 
         Height          =   375
         Left            =   4560
         TabIndex        =   26
         Top             =   5040
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "pictures"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton17 
         Height          =   375
         Left            =   4560
         TabIndex        =   27
         Top             =   5400
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "pictures"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label Label4 
         Caption         =   "MySQL • ODBC • ADO • XML • PHP • XOOPS"
         Height          =   495
         Left            =   240
         TabIndex        =   35
         Top             =   7680
         Width           =   6135
      End
      Begin VB.Label Label3 
         Caption         =   "driver :myODBC 3.51.06.exe  mysql control centre 0.9.2.exe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   7080
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   "connecting to aMySQL database using the MySQL ODBCdriver.  dll; Microsoft Remote Data Object 2.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   33
         Top             =   6000
         Width           =   7215
      End
      Begin VB.Label Label1 
         Caption         =   "WIN FOLDERS:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   16
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "WIN PROGRAMS:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   16305
      TabIndex        =   3
      Top             =   0
      Width           =   16305
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   16305
      TabIndex        =   15
      Top             =   0
      Width           =   16305
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
Option Explicit
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private dblWordID As Double
'// runs word


Private Sub cmdRedo_Click()

End Sub

Private Sub cmdSave_Click()
  'GitHub requires a login for submitting Issues; check for that first
    Dim msgReturn As VbMsgBoxResult
    
    msgReturn = MsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, kannagrafix needs to know where to send it." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbQuestion + vbApplicationModal + vbYesNo, "Thanks for making " & "KANNAGRAFIX" & " better")
    
    'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the tannerhelland.com contact form
    If msgReturn = vbYes Then
        'Shell a browser window with the GitHub issue report form
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL
    Else
        'Shell a browser window with the tannerhelland.com kannagrafix contact form
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL
    End If

End Sub


Private Sub jcbutton1_Click()
Main_Form.Show

End Sub

Private Sub jcbutton10_Click()

 
ShellExecute Me.HWnd, "Explore", "C:\xampp\htdocs\", "", "C:\", 1

End Sub

Private Sub jcbutton11_Click()
ShellExecute Me.HWnd, "Explore", "C:\Users\user\Downloads", "", "C:\", 1

End Sub

Private Sub jcbutton18_Click()
'"C:\Program Files\Everything\Everything.exe"
 dblWordID = Shell("C:\Program Files\Everything\Everything.exe", vbNormalFocus)
     'AppActivate dblWordID

End Sub

Private Sub jcbutton19_Click()
  
   
    frm_apache_php.Show 1, FormMain
End Sub

Private Sub jcbutton2_Click()
   
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

Private Sub jcbutton20_Click()

    frm_apache_sqL.Show 1, FormMain
End Sub

Private Sub jcbutton21_Click()
    FRM_HTML_CSS.Show 1, FormMain

End Sub

Private Sub jcbutton22_Click()
    Frm_VB6_Access.Show 1, FormMain

End Sub

Private Sub jcbutton3_Click()
MainFrm.Show

End Sub

Private Sub jcbutton4_Click()

End Sub

Private Sub jcbutton5_Click()
 ShellExecute Me.HWnd, "Explore", "e:\pictures", "", "e:\", 1


End Sub
Function pathOfFile(fileName As String) As String
        Dim posn As Integer
        posn = InStrRev(fileName, "\")
        If posn > 0 Then
            pathOfFile = Left$(fileName, posn)
        Else
            pathOfFile = ""
        End If
    End Function

Private Sub jcbutton6_Click()
  ShellExecute Me.HWnd, "Explore", "e:\vb6", "", "e:\", 1

End Sub

Private Sub jcbutton7_Click()
 ShellExecute Me.HWnd, "Explore", "e:\php", "", "e:\", 1

End Sub

Private Sub jcbutton8_Click()
 ShellExecute Me.HWnd, "Explore", "e:\html_java", "", "e:\", 1

End Sub

Private Sub jcbutton9_Click(Index As Integer)

  Select Case Index
        Case 0
             dblWordID = Shell("notepad", vbNormalFocus)
    
        Case 1
             dblWordID = Shell("doublekiller/doublekiller.EXE", vbNormalFocus)
     

        Case 2
                   ShellExecute FormMain.HWnd, "Open", "http://localhost/dashboard/", "", 0, SW_SHOWNORMAL
Case 3
 dblWordID = Shell("C:\xampp\xampp-control.exe", vbNormalFocus)
    
'C:\xampp\xampp-control.exe
Case 4
        ShellExecute FormMain.HWnd, "Open", "http://localhost/phpmyadmin/", "", 0, SW_SHOWNORMAL

Case 5
        ShellExecute FormMain.HWnd, "Open", "http://localhost/phpmyadmin/server_databases.php", "", 0, SW_SHOWNORMAL

'http://localhost/phpmyadmin/server_databases.php
Case 6

ShellExecute Me.HWnd, "open", App.Path & "\STOCK\HELP SCS.htm", vbNullString, vbNullString, SW_SHOWNORMAL

        End Select
End Sub

Private Sub MnuBugReport_Click()
    
    'GitHub requires a login for submitting Issues; check for that first
    Dim msgReturn As VbMsgBoxResult
    
    msgReturn = MsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, kannagrafix needs to know where to send it." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbQuestion + vbApplicationModal + vbYesNo, "Thanks for making " & "KANNAGRAFIX" & " better")
    
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


