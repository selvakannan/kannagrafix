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
   Begin VB.PictureBox RightBar 
      Align           =   4  'Align Right
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   9945
      Left            =   8940
      ScaleHeight     =   663
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   491
      TabIndex        =   31
      Top             =   0
      Width           =   7365
      Begin VB.ListBox lstFiles 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   8460
         ItemData        =   "VBP_FormMain.frx":0000
         Left            =   240
         List            =   "VBP_FormMain.frx":0002
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.FileListBox lstFilters1 
         Height          =   2040
         Left            =   360
         Pattern         =   "*.exe"
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.FileListBox lstFilters 
         Height          =   2040
         Left            =   360
         Pattern         =   "*.exe"
         TabIndex        =   32
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
      End
   End
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
      ScaleWidth      =   555
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8355
      Begin VB.TextBox Text1 
         Height          =   3015
         Left            =   1680
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   5760
         Width           =   5415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Left            =   1920
         TabIndex        =   41
         Top             =   1080
         Width           =   735
      End
      Begin kannagrafix.jcbutton jcbutton1 
         Height          =   495
         Left            =   600
         TabIndex        =   40
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
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
         Caption         =   ""
         PictureNormal   =   "VBP_FormMain.frx":0004
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   2880
         TabIndex        =   36
         Top             =   3240
         Width           =   495
      End
      Begin VB.DirListBox Dir1 
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   5160
         Visible         =   0   'False
         Width           =   855
      End
      Begin kannagrafix.jcbutton jcbutton12 
         Height          =   375
         Left            =   4560
         TabIndex        =   30
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
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
         Caption         =   "icon editor"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton28 
         Height          =   375
         Left            =   6360
         TabIndex        =   25
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
         Caption         =   "Folder Size"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton27 
         Height          =   375
         Left            =   5040
         TabIndex        =   24
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "Edit THIS VBP"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton26 
         Height          =   375
         Left            =   1560
         TabIndex        =   23
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
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
         Caption         =   "ADMIN"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton25 
         Height          =   375
         Left            =   2280
         TabIndex        =   22
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "C-PANEL"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton24 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "art4artist/mw19"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton23 
         Height          =   375
         Left            =   7200
         TabIndex        =   20
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "Dir Lister"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton4 
         Height          =   375
         Left            =   4560
         TabIndex        =   19
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
         Caption         =   "Delete Empty Folders"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton22 
         Height          =   375
         Left            =   6360
         TabIndex        =   18
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
         Caption         =   "vb6_Access_Projects"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton21 
         Height          =   375
         Left            =   4560
         TabIndex        =   17
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
         Caption         =   "HTML_CSS_Projects"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton20 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3840
         Width           =   2055
         _ExtentX        =   3625
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
         Left            =   120
         TabIndex        =   15
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
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
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
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
         Caption         =   ""
         PictureNormal   =   "VBP_FormMain.frx":015E
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "everything search"
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton5 
         Height          =   375
         Left            =   4560
         TabIndex        =   10
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
         Caption         =   "kannagrafix-master"
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
         Left            =   5760
         TabIndex        =   4
         Top             =   360
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
         TabIndex        =   5
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
         Caption         =   "Localhost-Dashboard"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton9 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
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
         Caption         =   ""
         PictureNormal   =   "VBP_FormMain.frx":0378
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton9 
         Height          =   615
         Index           =   4
         Left            =   1440
         TabIndex        =   7
         Top             =   2280
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
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
         Caption         =   ""
         PictureNormal   =   "VBP_FormMain.frx":1452
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton9 
         Height          =   615
         Index           =   5
         Left            =   720
         TabIndex        =   8
         Top             =   2280
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
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
         Caption         =   ""
         PictureNormal   =   "VBP_FormMain.frx":252C
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin kannagrafix.jcbutton jcbutton9 
         Height          =   375
         Index           =   6
         Left            =   6360
         TabIndex        =   9
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
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
         TabIndex        =   12
         Top             =   720
         Width           =   495
         _ExtentX        =   873
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
         Left            =   7320
         TabIndex        =   13
         Top             =   2160
         Width           =   855
         _ExtentX        =   1508
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
      Begin kannagrafix.jcbutton jcbutton10 
         Height          =   375
         Left            =   2040
         TabIndex        =   29
         Top             =   1800
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "eBag"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label Label7 
         Caption         =   "computer name;"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "WEB SITE DESIGNING FOR CART  SHOPS, PAYMENT GATEWAY SETTINGS"
         Height          =   2415
         Left            =   1320
         TabIndex        =   39
         Top             =   4920
         Width           =   5895
      End
      Begin VB.Label Label5 
         Caption         =   "MARKET PLACE FOR ORIGINAL ARTWORKS"
         Height          =   495
         Left            =   2400
         TabIndex        =   38
         Top             =   4440
         Width           =   5295
      End
      Begin VB.Label Label4 
         Caption         =   "Agency for MAXICAB RE-SALES"
         Height          =   375
         Left            =   2520
         TabIndex        =   37
         Top             =   3960
         Width           =   5295
      End
      Begin VB.Label Label3 
         Caption         =   "OFFLINE;"
         Height          =   375
         Left            =   5040
         TabIndex        =   28
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "XAMPP;"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "ONLINE;"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   1695
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
      TabIndex        =   11
      Top             =   0
      Width           =   16305
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "open"
      End
      Begin VB.Menu mnuopenprojects 
         Caption         =   "&open projects"
      End
      Begin VB.Menu mnulinks 
         Caption         =   "&Web Links"
      End
      Begin VB.Menu mnuremotedata 
         Caption         =   "&Remote server database connect"
      End
      Begin VB.Menu mnumail 
         Caption         =   "&Send Mail"
      End
   End
   Begin VB.Menu mnuapps 
      Caption         =   "&Installers"
      Begin VB.Menu mnuvb6 
         Caption         =   "vb6 setup"
      End
      Begin VB.Menu filter 
         Caption         =   "default"
         Index           =   0
      End
   End
   Begin VB.Menu Filters 
      Caption         =   "&MY apps"
      Begin VB.Menu filter1 
         Caption         =   "default"
         Index           =   0
      End
   End
   Begin VB.Menu mnuwinapps 
      Caption         =   "&win Apps"
      Begin VB.Menu mnuresmon 
         Caption         =   "Resource Moniter"
      End
      Begin VB.Menu mnunote 
         Caption         =   "notepad"
      End
      Begin VB.Menu mnuodbc 
         Caption         =   "ODBC Connectivity"
      End
      Begin VB.Menu mnuinternetoptions 
         Caption         =   "internet options"
      End
      Begin VB.Menu mnucalc 
         Caption         =   "calculator"
      End
      Begin VB.Menu mnucharmap 
         Caption         =   "charmap"
      End
      Begin VB.Menu mnuremote 
         Caption         =   "Remote Desktop connection"
      End
      Begin VB.Menu mnusqlnet 
         Caption         =   "SQL Server Client Network Utility"
      End
      Begin VB.Menu mnufire 
         Caption         =   "Windows Firewall"
      End
      Begin VB.Menu mnusysconfig 
         Caption         =   "System information"
      End
      Begin VB.Menu mnunetwork 
         Caption         =   "Network Connections"
      End
      Begin VB.Menu mnudevice 
         Caption         =   "Device Manager"
      End
      Begin VB.Menu mnudownloads 
         Caption         =   "downloads"
      End
   End
   Begin VB.Menu mnugra 
      Caption         =   "Graphics Editors"
      Begin VB.Menu mnupscs2 
         Caption         =   "Photoshop cs2"
      End
      Begin VB.Menu mnufast 
         Caption         =   "faststone"
      End
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
         Caption         =   "about system"
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
Private sAppName As String, sAppPath As String
Private dblWordID As Double
'// runs word
''# preparation (in a separate module)

''# use (anywhere)
'rundll32 Shell32.dll,Control_RunDLL INETCPL.CPL
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



Private Sub Dir1_Click()
MsgBox Dir1
End Sub

Private Sub filter1_Click(Index As Integer)
Dim path As String
    
    path = App.path
    
    If Right(path, 1) <> "\" Then path = path & "\"
    path = path & "filters\"
    
    'filter(Index).Tag = Path
   'MsgBox Path & filter1(Index).Tag
        dblWordID = Shell(path & filter1(Index).Tag, vbNormalFocus)
End Sub

Private Sub jcbutton1_Click()
        ShellExecute FormMain.HWnd, "Open", "https://mail.google.com/mail/u/0/?tab=wm", "", 0, SW_SHOWNORMAL

End Sub

Private Sub jcbutton10_Click()

 
                   ShellExecute FormMain.HWnd, "Open", "http://localhost/eBag/", "", 0, SW_SHOWNORMAL


End Sub

Private Sub jcbutton11_Click()

End Sub

Private Sub jcbutton12_Click()

        ShellExecute FormMain.HWnd, "Open", "Filters\TinyGFX32.exe", "", 0, SW_SHOWNORMAL

End Sub



Private Sub jcbutton18_Click()
'"C:\Program Files\Everything\Everything.exe"
 dblWordID = Shell(GetSystemDrive & "\Program Files\Everything\Everything.exe", vbNormalFocus)
     'AppActivate dblWordID

End Sub

Private Sub jcbutton19_Click()
  
   
    frm_apache_php.Show 1, FormMain
End Sub

Private Sub jcbutton2_Click()
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

Private Sub jcbutton23_Click()
 dblWordID = Shell(GetSystemDrive & "\Program Files (x86)\DirLister\DirLister.exe", vbNormalFocus)

'C:\Program Files (x86)\DirLister
End Sub

Private Sub jcbutton24_Click()
        ShellExecute FormMain.HWnd, "Open", "https://art4artist.com/mw19", "", 0, SW_SHOWNORMAL

End Sub

Private Sub jcbutton25_Click()
        ShellExecute FormMain.HWnd, "Open", "https://art4artist.com/cpanel", "", 0, SW_SHOWNORMAL

End Sub

Private Sub jcbutton26_Click()
        ShellExecute FormMain.HWnd, "Open", "https://art4artist.com/admin", "", 0, SW_SHOWNORMAL

End Sub

Private Sub jcbutton27_Click()
'dblWordID = Shell("kgmain.vbp", vbNormalFocus)
        ShellExecute FormMain.HWnd, "Open", "kgmain.vbp", "", 0, SW_SHOWNORMAL

End Sub

Private Sub jcbutton28_Click()
 dblWordID = Shell(GetSystemDrive & "\Program Files (x86)\FolderSize\FolderSize.exe", vbNormalFocus)

End Sub

Private Sub jcbutton3_Click()

End Sub

Private Sub jcbutton4_Click()
dblWordID = Shell("EmptyFolderCleanerPortable/EmptyFolderCleanerPortable.EXE", vbNormalFocus)
End Sub

Private Sub jcbutton5_Click()
 ShellExecute Me.HWnd, "Explore", "e:\kannagrafix-master", "", "e:\", 1


End Sub
Function pathOfFile(FileName As String) As String
        Dim posn As Integer
        posn = InStrRev(FileName, "\")
        If posn > 0 Then
            pathOfFile = Left$(FileName, posn)
        Else
            pathOfFile = ""
        End If
    End Function

Private Sub jcbutton6_Click()
  ShellExecute Me.HWnd, "Explore", "e:\kannagrafix-master\kg_vb6", "", "e:\", 1

End Sub

Private Sub jcbutton7_Click()
 ShellExecute Me.HWnd, "Explore", "e:\kannagrafix-master\Kg_PHP", "", "e:\", 1

End Sub

Private Sub jcbutton8_Click()

End Sub

Private Sub jcbutton9_Click(Index As Integer)

  Select Case Index
        Case 0
             
    
        Case 1
            ' dblWordID = Shell("doublekiller/doublekiller.EXE", vbNormalFocus)
     

        Case 2
                   ShellExecute FormMain.HWnd, "Open", "http://localhost/dashboard/", "", 0, SW_SHOWNORMAL
Case 3
sAppName = "XAMPP Control Panel v3.2.2   [ Compiled: Nov 12th 2015 ]"
            

sAppPath = GetSystemDrive & "\xampp\xampp-control.exe"
  For x = 0 To Dir1.ListCount - 1
         lstFiles.AddItem Right(Dir1.List(x), Len(Dir1.List(x)) - 16)
    Next x

 'check if application is running
    If IsTaskRunning(sAppName) Then
       ' MsgBox "Application '" & sAppName & "' is running!"
           lstFiles.Visible = True

    Else
        dblWordID = Shell(GetSystemDrive & "\xampp\xampp-control.exe", vbNormalFocus)
       lstFiles.Visible = True

    End If

'C:\xampp\xampp-control.exe
Case 4
        ShellExecute FormMain.HWnd, "Open", "http://localhost/phpmyadmin/", "", 0, SW_SHOWNORMAL

Case 5
        ShellExecute FormMain.HWnd, "Open", "http://localhost/phpmyadmin/server_databases.php", "", 0, SW_SHOWNORMAL

'http://localhost/phpmyadmin/server_databases.php
Case 6

ShellExecute Me.HWnd, "open", App.path & "HELP SCS.htm", vbNullString, vbNullString, SW_SHOWNORMAL

        End Select
End Sub

Private Sub lstFiles_Click()
Dim X1 As String
'Dim X2 As String

    ''# use (anywhere)
'MsgBox FindWindowHandle("XAMPP Control Panel v3.2.2   [ Compiled: Nov 12th 2015 ]")
sAppName = "XAMPP Control Panel v3.2.2   [ Compiled: Nov 12th 2015 ]"
            

sAppPath = GetSystemDrive & "\xampp\xampp-control.exe"

 'check if application is running
    If IsTaskRunning(sAppName) Then
       X1 = lstFiles.ListIndex
   ShellExecute FormMain.HWnd, "Open", "http://localhost/" & lstFiles.List(X1), "", 0, SW_SHOWNORMAL


    Else
        dblWordID = Shell(GetSystemDrive & "\xampp\xampp-control.exe", vbNormalFocus)
   ShellExecute FormMain.HWnd, "Open", "http://localhost/" & lstFiles.List(X1), "", 0, SW_SHOWNORMAL

    End If
           lstFiles.Visible = True


End Sub

Private Sub MDIForm_Load()
 Dim path As String
    Dim i As Integer
    Dim Title As String
    Dir1.path = GetSystemDrive & "\xampp\htdocs"
    path = App.path
       getip
  Dim intInc As Integer
    Dim strDisplay As String
    
  

    If Right(path, 1) <> "\" Then path = path & "\"
    path = path & "filters\installers"
    
    FormMain.lstFilters.path = path
    
    For i = 0 To FormMain.lstFilters.ListCount - 1
        If i <> 0 Then
            Load FormMain.filter(i)
        End If
        
        FormMain.filter(i).Visible = True
        
        FormMain.filter(i).Tag = FormMain.lstFilters.List(i)
        
        Title = Mid(FormMain.lstFilters.List(i), 1, InStr(1, FormMain.lstFilters.List(i), ".exe") - 1)
        Title = UCase(Left(Title, 1)) & Mid(Title, 2)
        FormMain.filter(i).Caption = Title
    Next i

 Dim Path1 As String
    Path1 = App.path
    
    If Right(Path1, 1) <> "\" Then Path1 = Path1 & "\"
    Path1 = Path1 & "filters"
    
    FormMain.lstFilters1.path = Path1
    
    For i = 0 To FormMain.lstFilters1.ListCount - 1
        If i <> 0 Then
            Load FormMain.filter1(i)
        End If
        
        FormMain.filter1(i).Visible = True
        
        FormMain.filter1(i).Tag = FormMain.lstFilters1.List(i)
        
        Title = Mid(FormMain.lstFilters1.List(i), 1, InStr(1, FormMain.lstFilters1.List(i), ".exe") - 1)
        Title = UCase(Left(Title, 1)) & Mid(Title, 2)
        FormMain.filter1(i).Caption = Title
    Next i
End Sub

Private Sub Filter_Click(Index As Integer)
    
Dim path As String
    
    path = App.path
    
    If Right(path, 1) <> "\" Then path = path & "\"
    path = path & "filters\installers\"
    
    'filter(Index).Tag = Path
    ' MsgBox
              dblWordID = Shell(path & filter(Index).Tag, vbNormalFocus)


 
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload MenuFrm
Unload MainFrm

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

Private Sub mnucalc_Click()
dblWordID = Shell("calc", vbNormalFocus)

End Sub

Private Sub mnucomp_Click()

End Sub

Private Sub mnucharmap_Click()
dblWordID = Shell("charmap", vbNormalFocus)

End Sub

Private Sub mnudevice_Click()
dblWordID = Shell("rundll32 Shell32.dll,Control_RunDLL hdwwiz.cpl", vbNormalFocus)

End Sub

Private Sub mnudownloads_Click()
ShellExecute Me.HWnd, "Explore", GetSystemDrive & "\Users\user\Downloads", "", GetSystemDrive & "\", 1

End Sub

Private Sub MnuEmailAuthor_Click()
    
    'Shell a browser window with the tannerhelland.com contact form
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL

End Sub

Private Sub mnufast_Click()
dblWordID = Shell(GetSystemDrive & "\Program Files (x86)\FastStone Image Viewer\FSViewer.exe", vbNormalFocus)

End Sub

Private Sub mnufire_Click()
dblWordID = Shell("rundll32 Shell32.dll,Control_RunDLL firewall.cpl", vbNormalFocus)

End Sub

Private Sub MnuHelpSepBar1_Click()
    Dim intInc As Integer
    Dim strDisplay As String

  Text1 = ""
    ' Here we start printing
    For intInc = 1 To 35
        strDisplay = strDisplay & Environ(intInc)
        strDisplay = strDisplay & Space(5)
    Next intInc
    
    Text1 = strDisplay
End Sub

Private Sub mnuinternetoptions_Click()
'INETCPL.CPL
'dblWordID = Shell("inetcpl.cpl", vbNormalFocus)
'ShellExecute FormMain.HWnd, "Open", GetWinPath & "\system32\inetcpl.cpl", "", 0, SW_SHOWNORMAL
dblWordID = Shell("rundll32 Shell32.dll,Control_RunDLL INETCPL.CPL", vbNormalFocus)

End Sub

Private Sub mnulinks_Click()
  
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

Private Sub mnumail_Click()
Main_Form.Show

End Sub


Private Sub mnunetwork_Click()
dblWordID = Shell("rundll32 Shell32.dll,Control_RunDLL ncpa.cpl", vbNormalFocus)

End Sub

Private Sub mnunote_Click()
dblWordID = Shell("notepad", vbNormalFocus)
End Sub

Private Sub mnuodbc_Click()
ShellExecute FormMain.HWnd, "Open", GetWinPath & "\system32\odbcad32.exe", "", 0, SW_SHOWNORMAL

End Sub

Private Sub mnuopen_Click()
 Dim FileName As String

 FileName = GetOpenName("Open...")
 
       ShellExecute FormMain.HWnd, "Open", FileName, "", 0, SW_SHOWNORMAL
End Sub


Private Sub mnupscs2_Click()
dblWordID = Shell(GetSystemDrive & "\Program Files (x86)\Adobe\Adobe Photoshop CS2\Photoshop.exe", vbNormalFocus)
'Shell "C:\Program Files\Adobe\Photoshop CS\Photoshop.exe"

End Sub

Private Sub mnuremote_Click()
dblWordID = Shell("mstsc", vbNormalFocus)

End Sub

Private Sub mnuremotedata_Click()
MainFrm.Show

End Sub

Private Sub mnuresmon_Click()
dblWordID = Shell("resmon", vbNormalFocus)

End Sub

Private Sub mnuslide_Click()
dblWordID = Shell("calc", vbNormalFocus)

End Sub

Private Sub mnusqlnet_Click()
dblWordID = Shell("cliconfg", vbNormalFocus)

End Sub

Private Sub mnusysconfig_Click()
dblWordID = Shell("msinfo32", vbNormalFocus)

End Sub

Private Sub mnuvb6_Click()
'E:\kannagrafix-master\VB6
              dblWordID = Shell(App.path & "\VB6\setup.exe", vbNormalFocus)
Dim path As String
    
    path = App.path
    
    If Right(path, 1) <> "\" Then path = path & "\"
    path = path & "vb6\"
    
    'filter(Index).Tag = Path
    ' MsgBox Path & "SETUP.EXE"
              dblWordID = Shell(path & "SETUP.EXE", vbNormalFocus)

End Sub

Private Sub MnuVisitWebsite_Click()
    'Nothing special here - just launch the default web browser with kannagrafix's page on tannerhelland.com
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL
End Sub


Private Sub MnuDonate_Click()
    'Launch the default web browser with the tannerhelland.com donation page
        ShellExecute FormMain.HWnd, "Open", "https://github.com/selvakannan/kannagrafix/tree/master", "", 0, SW_SHOWNORMAL
End Sub


