VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8925
   ClientLeft      =   360
   ClientTop       =   300
   ClientWidth     =   11985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   595
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6480
      Top             =   2880
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   8910
      ScaleHeight     =   136.284
      ScaleMode       =   0  'User
      ScaleWidth      =   161.999
      TabIndex        =   81
      Top             =   2400
      Width           =   2400
   End
   Begin VB.TextBox SendRmstxt 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1980
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox SendGms 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1950
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1950
      Visible         =   0   'False
      Width           =   8205
   End
   Begin VB.TextBox SendCMSTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1950
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   7800
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   10
      Top             =   9240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   7320
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   9
      Top             =   9240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   6840
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   8
      Top             =   9240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   5880
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   7
      Top             =   9360
      Visible         =   0   'False
      Width           =   420
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5760
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   5640
      Top             =   2520
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   5160
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5760
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   4680
      Top             =   2520
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4200
      Top             =   2520
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2385
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":34005
      Left            =   8640
      List            =   "frmMain.frx":34007
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   2640
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1455
      Left            =   120
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   450
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":34009
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox MainViewPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6195
      Left            =   120
      MousePointer    =   99  'Custom
      ScaleHeight     =   413
      ScaleMode       =   0  'User
      ScaleWidth      =   542
      TabIndex        =   20
      Top             =   2280
      Width           =   8130
      Begin VB.Timer anti_CE 
         Interval        =   9000
         Left            =   5040
         Top             =   720
      End
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   3000
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   1680
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   5
      Left            =   18000
      TabIndex        =   97
      Top             =   13560
      Width           =   750
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   8940
      TabIndex        =   29
      Top             =   6540
      Width           =   750
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   9015
      Left            =   8280
      TabIndex        =   96
      Top             =   -720
      Width           =   105
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      _Version        =   327682
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   12240
      TabIndex        =   95
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   55
      Top             =   960
      Width           =   540
   End
   Begin VB.Label imgAsignarSkill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   11280
      TabIndex        =   57
      Top             =   840
      Width           =   225
   End
   Begin VB.Label lblclan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Dolwur AO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8760
      TabIndex        =   65
      Top             =   960
      Width           =   2355
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   9960
      TabIndex        =   72
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   9015
      Left            =   0
      TabIndex        =   69
      Top             =   0
      Width           =   90
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   9015
      Left            =   11925
      TabIndex        =   68
      Top             =   240
      Width           =   135
   End
   Begin VB.Image DropGold 
      Height          =   255
      Left            =   10320
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label lblclan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Dolwur AO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   8880
      TabIndex        =   67
      Top             =   960
      Width           =   2355
   End
   Begin VB.Label lblclan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Dolwur AO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   8760
      TabIndex        =   66
      Top             =   960
      Width           =   2355
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6390
      TabIndex        =   64
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   7350
      TabIndex        =   63
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblhelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   62
      ToolTipText     =   " "
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label IconosegD 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   9840
      TabIndex        =   61
      Top             =   5760
      Width           =   435
   End
   Begin VB.Image ImagePARTY 
      Height          =   255
      Left            =   3120
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   12480
      TabIndex        =   60
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label IconoSeg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   8760
      TabIndex        =   59
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajo"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   58
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgEstadisticas 
      Height          =   345
      Left            =   10320
      MouseIcon       =   "frmMain.frx":34086
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   1290
   End
   Begin VB.Label Labelgm4 
      BackStyle       =   0  'Transparent
      Caption         =   "Panel 1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   56
      Top             =   8280
      Width           =   615
   End
   Begin VB.Image cmdInfo 
      Height          =   405
      Left            =   10560
      MouseIcon       =   "frmMain.frx":341D8
      MousePointer    =   99  'Custom
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image CmdLanzar 
      Height          =   495
      Left            =   8760
      MouseIcon       =   "frmMain.frx":3432A
      MousePointer    =   99  'Custom
      Top             =   5160
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "zoom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8400
      TabIndex        =   54
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "zoom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   8400
      TabIndex        =   53
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "zoom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   52
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   10860
      TabIndex        =   51
      Top             =   960
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lvllbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   10965
      TabIndex        =   50
      Top             =   960
      Width           =   225
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   10800
      TabIndex        =   49
      Top             =   6420
      Width           =   90
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   8640
      TabIndex        =   48
      Top             =   8040
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   8655
      TabIndex        =   47
      Top             =   8010
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   8655
      TabIndex        =   46
      Top             =   8010
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   8625
      TabIndex        =   45
      Top             =   8010
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   8655
      TabIndex        =   44
      Tag             =   "5"
      Top             =   8010
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   8640
      TabIndex        =   43
      Top             =   7680
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   8655
      TabIndex        =   42
      Top             =   7680
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   8655
      TabIndex        =   41
      Top             =   7680
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   8655
      TabIndex        =   40
      Top             =   7680
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   8655
      TabIndex        =   39
      Top             =   7680
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   8655
      TabIndex        =   38
      Top             =   7320
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   8640
      TabIndex        =   37
      Top             =   7320
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   8640
      TabIndex        =   36
      Top             =   7320
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   8670
      TabIndex        =   35
      Top             =   7320
      Width           =   1350
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   8880
      TabIndex        =   34
      Top             =   6915
      Width           =   975
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   8895
      TabIndex        =   33
      Top             =   6915
      Width           =   975
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   8895
      TabIndex        =   32
      Top             =   6915
      Width           =   975
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   8865
      TabIndex        =   31
      Top             =   6915
      Width           =   975
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   8895
      TabIndex        =   30
      Top             =   6915
      Width           =   975
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   8955
      TabIndex        =   28
      Top             =   6540
      Width           =   750
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   8925
      TabIndex        =   27
      Top             =   6540
      Width           =   750
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   8925
      TabIndex        =   26
      Top             =   6540
      Width           =   750
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   8940
      TabIndex        =   25
      Top             =   6540
      Width           =   750
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   8640
      TabIndex        =   24
      Top             =   7320
      Width           =   1350
   End
   Begin VB.Image STAShp 
      Height          =   150
      Left            =   8760
      Picture         =   "frmMain.frx":3447C
      Top             =   6570
      Width           =   1305
   End
   Begin VB.Image MANShp 
      Height          =   150
      Left            =   8760
      Picture         =   "frmMain.frx":34F0E
      Top             =   6930
      Width           =   1305
   End
   Begin VB.Image Hpshp 
      Height          =   150
      Left            =   8760
      Picture         =   "frmMain.frx":359A0
      Top             =   7365
      Width           =   1305
   End
   Begin VB.Image COMIDAsp 
      Height          =   150
      Left            =   8760
      Picture         =   "frmMain.frx":36432
      Top             =   7710
      Width           =   1305
   End
   Begin VB.Image AGUAsp 
      Height          =   150
      Left            =   8760
      Picture         =   "frmMain.frx":36EC4
      Top             =   8010
      Width           =   1305
   End
   Begin VB.Image CMSG 
      Height          =   255
      Left            =   10440
      MousePointer    =   99  'Custom
      Top             =   6660
      Width           =   1155
   End
   Begin VB.Image imgPMSG 
      Height          =   180
      Left            =   10320
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invisible"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7440
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Labelgm3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cr de 5 Seg"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Labelgm44 
      BackStyle       =   0  'Transparent
      Caption         =   "Panel 2"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   17
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Minimizar 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   11520
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Cerrar 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10560
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblmapaname 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Catacumbas de Ullathorpe"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8850
      TabIndex        =   14
      Top             =   8640
      Width           =   2535
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   225
      Index           =   1
      Left            =   11280
      Top             =   2520
      Width           =   195
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   255
      Index           =   0
      Left            =   11280
      Top             =   2880
      Width           =   195
   End
   Begin VB.Image imgClanes 
      Height          =   255
      Left            =   10320
      Top             =   8160
      Width           =   1275
   End
   Begin VB.Image imgOpciones 
      Height          =   285
      Left            =   10320
      Top             =   7320
      Width           =   1155
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   10320
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8520
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13470
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   180
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      MouseIcon       =   "frmMain.frx":37956
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label lblFPS 
      BackStyle       =   0  'Transparent
      Caption         =   "101"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   435
   End
   Begin VB.Image xz 
      Height          =   255
      Index           =   0
      Left            =   13320
      Top             =   600
      Width           =   255
   End
   Begin VB.Image xzz 
      Height          =   315
      Index           =   1
      Left            =   13200
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      MouseIcon       =   "frmMain.frx":37AA8
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Left            =   8520
      TabIndex        =   2
      Top             =   8730
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   210
      Left            =   8160
      TabIndex        =   1
      Top             =   8730
      Width           =   210
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa 1 [50,50]"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9435
      TabIndex        =   0
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   6225
      Left            =   120
      Top             =   2280
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Image InvEqu 
      Height          =   4485
      Left            =   8370
      Picture         =   "frmMain.frx":37BFA
      Top             =   1560
      Width           =   3300
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   10845
      TabIndex        =   73
      Top             =   960
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   10860
      TabIndex        =   74
      Top             =   960
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   10860
      TabIndex        =   75
      Top             =   960
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   10875
      TabIndex        =   76
      Top             =   960
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lvllbl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   10950
      TabIndex        =   77
      Top             =   960
      Width           =   225
   End
   Begin VB.Label lvllbl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   10965
      TabIndex        =   78
      Top             =   960
      Width           =   225
   End
   Begin VB.Label lvllbl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   10965
      TabIndex        =   79
      Top             =   960
      Width           =   225
   End
   Begin VB.Label lvllbl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   10980
      TabIndex        =   80
      Top             =   960
      Width           =   225
   End
   Begin VB.Label lblarmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4590
      TabIndex        =   82
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblarmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   83
      Top             =   9600
      Width           =   855
   End
   Begin VB.Label lblarmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   4860
      TabIndex        =   84
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label lblhelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   85
      Top             =   9360
      Width           =   855
   End
   Begin VB.Label lblhelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   86
      Top             =   9360
      Width           =   855
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   4110
      TabIndex        =   87
      ToolTipText     =   " "
      Top             =   9600
      Width           =   855
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   4140
      TabIndex        =   88
      ToolTipText     =   " "
      Top             =   9600
      Width           =   855
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   89
      ToolTipText     =   " "
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   2910
      TabIndex        =   90
      ToolTipText     =   " "
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label lblclan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Dolwur AO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   3
      Left            =   8880
      TabIndex        =   91
      Top             =   960
      Width           =   2355
   End
   Begin VB.Label lblclan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Dolwur AO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   4
      Left            =   8760
      TabIndex        =   92
      Top             =   960
      Width           =   2355
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "zoom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   3
      Left            =   8400
      TabIndex        =   93
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "zoom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   4
      Left            =   8400
      TabIndex        =   94
      Top             =   720
      Width           =   3255
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' x Auto Pots
Private Enum eVentanas

    vHechizos = 1
    vInventario = 2

End Enum

Private Panel              As Byte
Private LastPanel          As Byte
Private Const InvalidSlot  As Byte = 255
' x Auto Pots

' x button
Private mouse_Down         As Boolean
Private mouse_UP           As Boolean
' x button

Public n As Byte

Public Pulsacion_Fisica As Boolean

Private MouseInvBoton As Long

Public Attack As Boolean
Private Last_I      As Long
Public WithEvents dragInventory As clsGrapchicalInventory
Attribute dragInventory.VB_VarHelpID = -1

Dim Ancho As Integer
Dim alto As Integer
Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Private clsFormulario As clsFormMovementManager

Private cBotonDiamArriba As clsGraphicalButton
Private cBotonDiamAbajo As clsGraphicalButton
Private cBotonMapa As clsGraphicalButton
Private cBotonGrupo As clsGraphicalButton
Private cBotonOpciones As clsGraphicalButton
Private cBotonEstadisticas As clsGraphicalButton
Private cBotonClanes As clsGraphicalButton
Private cBotonAsignarSkill As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public picSkillStar As Picture

Private cmsgSupr As Boolean
Private bCMSG As Boolean
Private btmpCMSG As Boolean
Private sPartyChat As String

Private Sub anti_CE_Timer()
Call BuscarEngine
End Sub

Private Sub Cerrar_Click()
Call Audio.PlayWave(SND_CLICK)
If MsgBox("?Desea cerrar Desterium  AO?", vbYesNo + vbQuestion, "Desterium  AO") = vbYes Then
  prgRun = False
Else
            Exit Sub
        End If
End Sub
Private Sub DropGold_Click()
    Inventario.SelectGold
    If UserGLD > 0 Then
        If Not Comerciando Then FrmCantidad.Show , frmMain
    End If
    
End Sub

Private Sub Form_Load()

'frmConnect.font.Name = "Tahoma"
' frmConnect.font.Size = 85

 Dim CursorDir As String
    Dim Cursor As Long
    
    'Drag And Drop
    Set dragInventory = Inventario
    
    CursorDir = App.path & "\Recursos\Cursor.ani" 'normal1.ani
    hSwapCursor = SetClassLong(frmMain.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
    hSwapCursor = SetClassLong(frmMain.PicInv.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
    hSwapCursor = SetClassLong(frmMain.hlst.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
    
    'Consola Inteligente
    Detectar RecTxt.hwnd, Me.hwnd
    
    If NoRes Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me, 120
    End If

    'Me.Picture = LoadPicture(DirGraficos & "VentanaPrincipal.JPG")
    
    InvEqu.Picture = LoadPicture(DirGraficos & "CentroInventario.JPG")
    
    Call LoadButtons
    
    Me.Left = 0
    Me.Top = 0
    
        lblmapaname.Visible = True
    Coord.Visible = False
    
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    Dim i As Integer
    
    GrhPath = DirGraficos

    Set cBotonDiamArriba = New clsGraphicalButton
    Set cBotonDiamAbajo = New clsGraphicalButton
    Set cBotonGrupo = New clsGraphicalButton
    Set cBotonOpciones = New clsGraphicalButton
    Set cBotonEstadisticas = New clsGraphicalButton
    Set cBotonClanes = New clsGraphicalButton
    Set cBotonAsignarSkill = New clsGraphicalButton
    Set cBotonMapa = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton

    'Set picSkillStar = LoadPicture(GrhPath & "BotonAsignarSkills.bmp")

    'If SkillPoints > 0 Then imgAsignarSkill.Picture = picSkillStar
    
    imgAsignarSkill.MouseIcon = picMouseIcon
    lblDropGold.MouseIcon = picMouseIcon
    lblCerrar.MouseIcon = picMouseIcon
    lblMinimizar.MouseIcon = picMouseIcon
    
    For i = 0 To 3
        picSM(i).MouseIcon = picMouseIcon
    Next i
End Sub

Public Sub LightSkillStar(ByVal bTurnOn As Boolean)
    If bTurnOn Then
       ' imgAsignarSkill.Picture = picSkillStar
    Else
       ' Set imgAsignarSkill.Picture = Nothing
    End If
End Sub

Private Sub cmdMoverHechi_Click(index As Integer)
    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub
        Dim sTemp As String
    
        Select Case index
            Case 1 'subir
                If hlst.ListIndex = 0 Then Exit Sub
            Case 0 'bajar
                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        End Select
    
        Call WriteMoveSpell(index = 1, hlst.ListIndex + 1)
        
        Select Case index
            Case 1 'subir
                sTemp = hlst.List(hlst.ListIndex - 1)
                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex - 1
            Case 0 'bajar
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1
        End Select
    End If
End Sub

Public Sub ControlSM(ByVal index As Byte, ByVal Mostrar As Boolean)
Dim GrhIndex As Long
Dim SR As RECT
Dim DR As RECT

GrhIndex = GRH_INI_SM + index + SM_CANT * (CInt(Mostrar) + 1)

With GrhData(GrhIndex)
    SR.Left = .sX
    SR.Right = SR.Left + .pixelWidth
    SR.Top = .sY
    SR.Bottom = SR.Top + .pixelHeight
    
    DR.Left = 0
    DR.Right = .pixelWidth
    DR.Top = 0
    DR.Bottom = .pixelHeight
End With

Select Case index

    Case eSMType.sSafemode
        If Mostrar Then
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, True)
            picSM(index).ToolTipText = "Seguro activado."
            frmMain.IconoSeg.Caption = ""
        Else
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, True)
            picSM(index).ToolTipText = "Seguro desactivado."
            frmMain.IconoSeg.Caption = "X"
        End If
        
    Case eSMType.DragMode
        If Mostrar Then
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_DRAG_DESACTIVADO, 255, 0, 0, True, False, True)
            frmMain.IconosegD.Caption = "X"
        Else
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_DRAG_ACTIVADO, 0, 255, 0, True, False, True)
            frmMain.IconosegD.Caption = ""
        End If

End Select

SMStatus(index) = Mostrar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 18/11/2009
'18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
'***************************************************
#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If
    
    
     
If Pulsacion_Fisica = False Then
Exit Sub
End If
Pulsacion_Fisica = True
    
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) And (Not SendRmstxt.Visible) And (Not SendGms.Visible) Then
    
    
        If KeyCode = vbKeyF5 Then
        frmRetos.Show , frmMain
        'Call AddtoRichTextBox(frmMain.RecTxt, "Sistema Desactivado", 0, 200, 200, False, False, True)
        End If
        
                      If KeyCode = vbKeyF1 Then
        WriteLevel
        End If
        
                      If KeyCode = vbKeyF2 Then
                      If MsgBox("?Est?s seguro de resetear tu personaje?", vbYesNo + vbQuestion, "Desterium  AO") = vbYes Then
        WriteReset
               Call ShowConsoleMsg("Reloguea para ver los cambios.", , , , True)
        End If
        End If
        
              If KeyCode = vbKeyF3 Then
        WritePartyJoin
        End If
        
        If KeyCode = vbKeyEnd Then
        WriteMeditate
        End If
        
                If KeyCode = vbKeyF7 Then
                   ' Call ShowConsoleMsg("[BETA] El sistema de PARTY esta desactivado.", 255, 0, 0)
                    Call Protocol.writeRequestPartyForm
                    End If
        
        
        If KeyCode = vbKeyF9 Then
        'Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        Exit Sub
        End If
        
        'Comandos Gms.
             
             
             If esGM(UserCharIndex) = True Then
            If KeyCode = vbKeyQ Then
              If SendTxt.Visible Or SendGms.Visible Then Exit Sub
                If Not FrmCantidad.Visible Then
                ShowConsoleMsg "Escriba un mensaje global.", 0, 255, 255
                    SendRmstxt.Visible = True
                    SendRmstxt.SetFocus
                    End If
                End If
             
                If KeyCode = vbKeyG Then
                If SendTxt.Visible Or SendRmstxt.Visible Then Exit Sub
                    If Not FrmCantidad.Visible Then
                    ShowConsoleMsg "Escriba un mensaje a los Game Masters.", 0, 255, 255
                    SendGms.Visible = True
                    SendGms.SetFocus
                    End If
                End If
                End If
             
        If KeyCode = vbKeyI Then
        Call ParseUserCommand("/INVISIBLE")
        End If

        If KeyCode = vbKeyW Then
        Call ParseUserCommand("/TRABAJANDO")
        End If
        If KeyCode = vbKeyP Then
        Call ParseUserCommand("/PANELGM")
        End If
        
        If KeyCode = vbKeyF10 Then
            'Intervalo permite usar este sistema?
            If Not FotoD_CanSend Then
            Call AddtoRichTextBox(frmMain.RecTxt, "Haz alcanzado el m?ximo de envio de 1 FotoDenuncia por minuto. Esper? unos instantes y volve a intentar.", 0, 200, 200, False, False, True)
            Exit Sub
            End If
            'Aca guardamos el string que nos devuelve FotoD_Capturar.
            Dim nString    As String
            FotoD_Capturar nString
            'Si el string da nullo, es por que nadie esta insultando.
            If nString = vbNullString Then
            Call AddtoRichTextBox(frmMain.RecTxt, "Nadie te esta insultando. Las FotoDenuncias solo sirven para denunciar agravios.", 0, 200, 200, False, False, True)
            Else 'Si no, enviamos.
            Call AddtoRichTextBox(frmMain.RecTxt, "La FotoDenuncia fue sacada correctamente.", 0, 200, 200, False, False, True)
            WriteDenounce "[FOTODENUNCIAS] : " & nString
            End If
            End If
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    Audio.MusicActivated = Not Audio.MusicActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    Audio.SoundActivated = Not Audio.SoundActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFxs)
                    Audio.SoundEffectsActivated = Not Audio.SoundEffectsActivated
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
                    Call WriteCombatModeToggle
                    Iscombate = Not Iscombate
                
                    Case vbKeyMultiply:
                    If frmMain.IconoSeg.Visible Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("Escribe /SEG para quitar el seguro", .red, .green, .blue, .bold, .italic)
                    End With
                    'Call AddtoRichTextBox(frmMain.RecTxt, "Escribe /SEG para quitar el seguro", 0, 200, 200, False, False, True)
                    Else
                    Call WriteSafeToggle
                    End If
                
                Case vbKeyZ:
                    If DialogosClanes.Activo = False Then
                        Call ShowConsoleMsg("Consola flotante de clanes activada.", 255, 200, 200)
                        DialogosClanes.Activo = True
                    Else
                        Call ShowConsoleMsg("Consola flotante de clanes desactivada.", 255, 200, 200)
                        DialogosClanes.Activo = False
                    End If
                
                
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                                                       Case CustomKeys.BindedKey(eKeyType.mKeyRETOS)
                     frmRetos.Show , frmMain
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem(1)
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
               ' Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    'Call WriteSafeToggle

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle
            End Select
        Else
            Select Case KeyCode
                'Custom messages!
                Case vbKey0 To vbKey9
                    Dim CustomMessage As String
                    
                    CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)
                    If LenB(CustomMessage) <> 0 Then
                        ' No se pueden mandar mensajes personalizados de clan o privado!
                        If UCase(Left(CustomMessage, 5)) <> "/CMSG" And _
                            Left(CustomMessage, 1) <> "\" Then
                            
                            Call ParseUserCommand(CustomMessage)
                        End If
                    End If
            End Select
        End If
    End If
    
Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
                    If (Not Comerciando) And (Not MirandoAsignarSkills) And _
              (Not frmMSG.Visible) And (Not MirandoForo) And _
              (Not frmEstadisticas.Visible) And (Not FrmCantidad.Visible) Then
            End If
   If bCMSG = True Then Exit Sub 'Si est? activado el cmsgimg lo cancelamos
                SendTxt.Visible = True 'Mostramos el Sendtxt
                SendTxt.SetFocus 'Lo priorizamos
                cmsgSupr = True 'Activamos que fue con la tecla suprimir con lo que fue abierta
                bCMSG = True 'Activamos que se puso el CMSGimg
            
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call ScreenCapture
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            'If UserMinMAN = UserMaxMAN Then Exit Sub
            
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
                
        Call WriteMeditate
      
        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
          writeRequestPartyForm
        
        
        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If macrotrabajo.Enabled Then
                Call DesactivarMacroTrabajo
            Else
                Call ActivarMacroTrabajo
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            If frmMain.macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteQuit
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
                    If Shift <> 0 Then Exit Sub
       
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                    If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                        If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                    Else
                        If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
                    End If
     
                   If TrainingMacro.Enabled Then DesactivarMacroHechizos
                   If macrotrabajo.Enabled Then DesactivarMacroTrabajo
               Call WriteAttack
               Attack = True
           charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).Heading).Started = 1
           charlist(UserCharIndex).Escudo.ShieldWalk(charlist(UserCharIndex).Heading).Started = 1
             
             If Iscombate = False Then
             Attack = False
              charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).Heading).Started = 0
           charlist(UserCharIndex).Escudo.ShieldWalk(charlist(UserCharIndex).Heading).Started = 0
           End If
        
 Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If SendCMSTXT.Visible Then Exit Sub
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And _
              (Not frmMSG.Visible) And (Not MirandoForo) And _
              (Not frmEstadisticas.Visible) And (Not FrmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
            
    End Select
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewPic.Left
    MouseY = Y - MainViewPic.Top
   
      'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewShp.Width Then
        MouseX = MainViewPic.Width
    End If
 
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewShp.Height Then
        MouseY = MainViewShp.Height
    End If
    
     Ancho = lblmapaname.Left + lblmapaname.Width
    alto = lblmapaname.Top + lblmapaname.Height
    If X > lblmapaname.Left And X < Ancho And Y > lblmapaname.Top And Y < alto Then
        lblmapaname.Visible = False
        Coord.Visible = True
    Else
        lblmapaname.Visible = True
        Coord.Visible = False
    End If
    
    Ancho = lvllbl(0).Left + lvllbl(0).Width
    alto = lvllbl(0).Top + lvllbl(0).Height
    If X > lvllbl(0).Left And X < Ancho And Y > lvllbl(0).Top And Y < alto Then
        lvllbl(0).Visible = False
        lvllbl(1).Visible = False
        lvllbl(2).Visible = False
        lvllbl(3).Visible = False
        lvllbl(4).Visible = False
        lblporclvl(0).Visible = True
         lblporclvl(1).Visible = True
          lblporclvl(2).Visible = True
           lblporclvl(3).Visible = True
            lblporclvl(4).Visible = True
    Else
        lvllbl(0).Visible = True
         lvllbl(1).Visible = True
          lvllbl(2).Visible = True
          lvllbl(3).Visible = True
          lvllbl(4).Visible = True
        lblporclvl(0).Visible = False
        lblporclvl(1).Visible = False
        lblporclvl(2).Visible = False
        lblporclvl(3).Visible = False
        lblporclvl(4).Visible = False
    End If

    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewShp.Height Then
        MouseY = MainViewShp.Height
    End If
    Inventario.uMoveItem = False
    Inventario.sMoveItem = False
    
      If SendTxt.Visible Then
        SendTxt.SetFocus
    End If
    
End Sub


Private Sub CMSG_Click()
Call Audio.PlayWave(SND_CLICK)
    If Not CharTieneClan Then
    Call AddtoRichTextBox(frmMain.RecTxt, "?No perteneces a ning?n clan!", 0, 200, 200, False, False, True)
      If bCMSG = False Then
      cmsgSupr = False
    Exit Sub
    End If
Else
    bCMSG = Not bCMSG
    If bCMSG Then
    cmsgSupr = False
        CMSG.Picture = LoadPicture(App.path & "\Recursos\CMSG.jpg")
    Call AddtoRichTextBox(frmMain.RecTxt, "Todo lo que digas sera escuchado por tu clan.", 0, 200, 200, False, False)
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de ser escuchado por tu clan. ", 0, 200, 200, False, False)
        CMSG.Picture = LoadPicture("")
    End If
    End If
End Sub

Private Sub IconoSeg_Click()
WriteSafeToggle
End Sub

Private Sub IconosegD_Click()
'Sistema Deniega el Item
WriteDragToggle
End Sub


Private Sub Image1_Click()
    FrmRanking.Show vbModeless, frmMain
End Sub

Private Sub Image3_Click()
frmQuests.Show , frmMain
End Sub

Private Sub Image4_Click()
FrmMercado.Show , frmMain
End Sub

Private Sub ImagePARTY_dblclick()
 Call ParseUserCommand("/Onlineparty")
End Sub
Private Sub imageparty_click()
Call Audio.PlayWave(SND_CLICK)
Call AddtoRichTextBox(frmMain.RecTxt, "Presiona doble click para ver la experiencia repartida en tu party.", 0, 200, 200, False, False)
End Sub




Private Sub imgEstadisticas_Click()

Call Audio.PlayWave(SND_CLICK)

 Dim i As Integer
    If SkillPoints > 0 Then
    imgAsignarSkill.Visible = True
    Else
    imgAsignarSkill.Visible = False
    imgAsignarSkill.Enabled = False
    End If

    
    LlegaronSkills = False
    Call WriteRequestSkills
    Call FlushBuffer
    
    Do While Not LlegaronSkills
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    LlegaronSkills = False
    
    For i = 1 To NUMSKILLS
        frmSkills3.text1(i).Caption = UserSkills(i)
    Next i
    
    Alocados = SkillPoints
     LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
    Call WriteRequestFame
    Call FlushBuffer
    Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    frmSkills3.Iniciar_Labels
    frmSkills3.Show , frmMain
    frmSkills3.lbldatos.Caption = "Nivel: " & UserLvl & " Experiencia: " & UserExp & "/" & UserPasarNivel
    Alocados = SkillPoints
    frmSkills3.puntos.Caption = SkillPoints
    frmSkills3.Show , frmMain

    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
End Sub

Private Sub imgPMSG_Click()
Call Audio.PlayWave(SND_CLICK)
'----Boton partys Style TDS by IRuleDK----
PMSG = False 'Nos fijamos que no este activado con la tecla suprimir
If PMSGimg = False Then 'Si no hab?a apretado el bot?n -> lo activamos y le ponemos la imagen estilo TDS
PMSGimg = True
imgPMSG.Picture = LoadPicture(App.path & "\Recursos\Pmsg.jpg") 'Grafico del bot?n estilo tds
Call AddtoRichTextBox(frmMain.RecTxt, "Todo lo que digas sera escuchado por tu party. ", 255, 200, 200, False, False)
Else 'si ya estaba apretado lo desactivamos
PMSGimg = False 'desactivamos el boton
imgPMSG.Picture = LoadPicture("") 'lo ponemos normal sacandole la imagen verde
Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de ser escuchado por tu party. ", 255, 200, 200, False, False)
Call ControlSM(eSMType.mWork, True)
End If
End Sub

Private Sub Label1_Click()
Call ParseUserCommand("/invisible")
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    LastPressed.ToggleToNormal
        Inventario.uMoveItem = False
    Inventario.sMoveItem = False
End Sub

Private Sub Label3_Click()

End Sub



Private Sub Label5_Click()
Call WriteWorking
End Sub

Private Sub Label6_Click()
 Call AddtoRichTextBox(frmMain.RecTxt, "Presiona doble click para crear una party.", 0, 200, 200, False, False)
End Sub
Private Sub Label6_dblClick()
 Call ParseUserCommand("/crearparty")
End Sub

Private Sub Labelgm1_Click()
Call ParseUserCommand("/telep yo 1 50 50")
End Sub

Private Sub Labelgm2_Click()
If MsgBox("Esta todo listo para empezar la daga rusa?", vbYesNo, "Daga rusa") = vbYes Then
Call ParseUserCommand("/RMSG Luego de la cuenta envien los interesados en la Daga Rusa")
Call ParseUserCommand("/cr 5")
End If
End Sub

Private Sub Labelgm3_Click()
If MsgBox("Queres mandar un conteo de 5 segs flaco?", vbYesNo, "Cuenta Regresiva") = vbYes Then
Call ParseUserCommand("/cr 5")
End If
End Sub

Private Sub Labelgm4_Click()
frmPanelGm.Show , frmMain
End Sub

Private Sub Labelgm44_Click()
frmPanelGMS.Show , frmMain
End Sub

Private Sub Labelgm5_Click()
Call ParseUserCommand("/online")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub imgClanes_Click()
Call Audio.PlayWave(SND_CLICK)
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub imgAsignarSkill_Click()
    Dim i As Integer
    If SkillPoints > 0 Then
    imgAsignarSkill.Visible = True
    Else
    imgAsignarSkill.Visible = False
    imgAsignarSkill.Enabled = False
    End If

    
    LlegaronSkills = False
    Call WriteRequestSkills
    Call FlushBuffer
    
    Do While Not LlegaronSkills
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    LlegaronSkills = False
    
    For i = 1 To NUMSKILLS
        frmSkills3.text1(i).Caption = UserSkills(i)
    Next i
    
    Alocados = SkillPoints
     LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
    Call WriteRequestFame
    Call FlushBuffer
    Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    frmSkills3.Iniciar_Labels
    frmSkills3.Show , frmMain
    frmSkills3.lbldatos.Caption = "Nivel: " & UserLvl & " Experiencia: " & UserExp & "/" & UserPasarNivel
    Alocados = SkillPoints
    frmSkills3.puntos.Caption = SkillPoints
    frmSkills3.Show , frmMain

    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False

End Sub

Private Sub imgGrupo_Click()
Call Audio.PlayWave(SND_CLICK)
    Call writeRequestPartyForm
End Sub

Private Sub imgInvScrollDown_Click()
    Call Inventario.ScrollInventory(True)
End Sub

Private Sub imgInvScrollUp_Click()
    Call Inventario.ScrollInventory(False)
End Sub

Private Sub imgOpciones_Click()
Call Audio.PlayWave(SND_CLICK)
    Call frmOpciones.Show(vbModeless, frmMain)
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
        Inventario.uMoveItem = False
    Inventario.sMoveItem = False
End Sub
Private Sub lblScroll_Click(index As Integer)
    Inventario.ScrollInventory (index = 0)
End Sub



Private Sub lblMinimizar_Click()
    Me.WindowState = 1
End Sub

Private Sub macrotrabajo_Timer()
    If Inventario.SelectedItem = 0 Then
        DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    If Not Application.IsAppActive() Then
        DesactivarMacroTrabajo
        Exit Sub
    End If
    
    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not frmHerrero.Visible) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
     If Not (frmCarp.Visible = True) Then Call UsarItem(1)
End Sub

Public Sub ActivarMacroTrabajo()
    If Iscombate Then
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
                        Call ShowConsoleMsg("No puedes trabajar en modo combate.", .red, .green, .blue, .bold, .italic)
                    End With
     Exit Sub
   End If

    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Empiezas a trabajar", 0, 200, 200, False, False, True)
    Call ControlSM(eSMType.mWork, True)
  

End Sub

Public Sub DesactivarMacroTrabajo()

    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de trabajar", 0, 200, 200, False, False, True)
    Call ControlSM(eSMType.mWork, False)
 
End Sub

Private Sub Minimizar_Click()
Call Audio.PlayWave(SND_CLICK)
Me.WindowState = 1
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub
Private Sub MainViewPic_Click()
    Form_Click
      If SendTxt.Visible Then
        SendTxt.SetFocus
    End If
End Sub

Private Sub MainViewPic_DblClick()
    Form_DblClick
      If SendTxt.Visible Then
        SendTxt.SetFocus
    End If
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift

    Call ConvertCPtoTP(X, Y, tX, tY)
    
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseX = X
    MouseY = Y
    
    
    'LastPressed.ToggleToNormal
    
    Call ConvertCPtoTP(X, Y, tX, tY)
    
    If Inventario.sMoveItem And Not vbKeyShift Then
        General_Drop_X_Y tX, tY
        Inventario.uMoveItem = False
    Else
        If Inventario.sMoveItem And vbKeyShift Then
        FrmCantidad.Show , frmMain
        End If
    End If

      If SendTxt.Visible Then
        SendTxt.SetFocus
    End If

End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    clicX = X
    clicY = Y

   
     
End Sub


Private Sub mnuTirar_Click()
    Call TirarItem
    Inventario.uMoveItem = False
    Inventario.sMoveItem = False
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem(0)
End Sub

Private Sub PicMH_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar ?nicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, True)
End Sub

Private Sub lblmapaname_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblmapaname.Visible = False
    Coord.Visible = True
End Sub
Private Sub coord_click()
Call Audio.PlayWave(SND_CLICK)
 Call AddtoRichTextBox(frmMain.RecTxt, "Presiona doble click para abrir el mapa del mundo.", 0, 200, 200, False, False)
End Sub
Private Sub coord_dblclick()
Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub picSM_DblClick(index As Integer)
Select Case index
    Case eSMType.sResucitation
        Call WriteResuscitationToggle
        
    Case eSMType.sSafemode
        Call WriteSafeToggle
        
    Case eSMType.mSpells
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
 
    Case eSMType.mWork
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
        If macrotrabajo.Enabled Then
            Call DesactivarMacroTrabajo
        Else
            Call ActivarMacroTrabajo
        End If
End Select
End Sub



Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
        'Sistema de bot?n clanes estilo TDS by AmenO
 

    
    
   If KeyCode = CustomKeys.BindedKey(eKeyType.mKeyTalk) Then 'Si se apret? enter entonces:
        Call Dialogos.RemoveDialog(UserCharIndex)
                If PMSGimg = True Then 'Si est? activado el PMSGimg
                       sPartyChat = SendTxt.Text 'Mandamos lo que sea de Party
                   
                        '// Es mas rapido comprar byts que cadenas de letras :P
                       ' If sPartyChat <> "" Then
                   
                        If LenB(sPartyChat) <> 0 Then
                                Call ParseUserCommand("/PMSG " & sPartyChat)
                        End If
                        'Reiniciamos los valores
                       sPartyChat = vbNullString ' // Mejor vbnullstring que ""
                       SendTxt.Text = vbNullString ' // Mejor vbnullstring que ""
                End If
    End If
 
 
       If KeyCode = CustomKeys.BindedKey(eKeyType.mKeyTalk) Then 'Si se apret? enter entonces:
                Call Dialogos.RemoveDialog(UserCharIndex)
                If bCMSG = True Then 'Si est? activado el CMSGimg
                       stxtbuffercmsg = SendTxt.Text 'Mandamos lo que sea de CLAN
                       
                        '// Es mas rapido comprar byts que cadenas de letras :P
                       ' If stxtbuffercmsg <> "" Then
                       
                        If LenB(stxtbuffercmsg) <> 0 Then
                                Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
                        End If
 
                        'Reiniciamos los valores
                       stxtbuffercmsg = vbNullString ' // Mejor vbnullstring que ""
                        SendTxt.Text = vbNullString ' // Mejor vbnullstring que ""
 
                       If cmsgSupr = True Then 'Revisamos si fue con Suprimir
                                bCMSG = False 'Si fue as? desactivamos el cmsgimg
                       End If
 
                       KeyCode = 0
                       SendTxt.Visible = False
 
                       If PicInv.Visible Then
                               PicInv.SetFocus
                       Else
                               hlst.SetFocus
                       End If
 
                       Exit Sub
 
               End If
 
               If LenB(stxtbuffer) <> 0 Then
                       Call ParseUserCommand(stxtbuffer) ' Y si no hab?a nada de CMSG hacemos el proceso com?n para hablar
                End If
 
                stxtbuffer = vbNullString ' // Mejor vbnullstring que ""
               SendTxt.Text = vbNullString ' // Mejor vbnullstring que ""
                KeyCode = 0
                SendTxt.Visible = False
       
                If PicInv.Visible Then
                        PicInv.SetFocus
                Else
                        hlst.SetFocus
                End If
        End If
 
        '----Boton clanes Style TDS by AmenO----
End Sub

Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer

Call ModResolution.Check_All

End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
                Inventario.uMoveItem = False
                Inventario.sMoveItem = False
            Else
                If Inventario.amount(Inventario.SelectedItem) > 1 Then
                    If Not Comerciando Then FrmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem(ByRef SecondaryClick As Byte)
    If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub
    
    If Comerciando Then Exit Sub
    
    Dim ItemIndex As Integer
        
    ItemIndex = Inventario.SelectedItem
    
    If (ItemIndex > 0) And (ItemIndex < MAX_INVENTORY_SLOTS + 1) Then
        
        If Inventario.OBJType(ItemIndex) <> eObjType.otBarcos Then
            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)

                End With

                Exit Sub

            End If

        End If

        If Inventario.OBJType(ItemIndex) = eObjType.otPociones Then
            Call WriteUsePotions(ItemIndex, CByte(SecondaryClick))
        Else
            Call WriteUseItem(ItemIndex)

        End If

    End If

End Sub

Private Sub EquiparItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()
    If Not hlst.Visible Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    'Macros are disabled if focus is not on Argentum!
    If Not Application.IsAppActive() Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    If Comerciando Then Exit Sub
    
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.CastSpell, False) Then
        Call WriteCastSpell(hlst.ListIndex + 1)
        Call WriteWork(eSkill.Magia)
    End If
    
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    'If UsingSkill = Magia And Not MainTimer.Check(TimersIndex.CastSpell) Then Exit Sub
    
    'If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub
    
    Call WriteWorkLeftClick(tX, tY, UsingSkill)
    UsingSkill = 0
End Sub

Private Sub cmdLanzar_Click()

 If hlst.List(hlst.ListIndex) = "(Vacio)" Then Exit Sub

    If Iscombate = False Then
   With FontTypes(FontTypeNames.FONTTYPE_INFO)
    Call ShowConsoleMsg("??No puedes lanzar hechizos si no estas en modo combate!!", .red, .green, .blue, .bold, .italic)
   End With
    Exit Sub
    End If
    


    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("??Est?s muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If


End Sub


Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub cmdINFO_Click()
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
End Sub

Private Sub DespInv_Click(index As Integer)
    Inventario.ScrollInventory (index = 0)
End Sub
Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Inventario.uMoveItem Then
        PicInv.MousePointer = vbDefault
    End If
End Sub
Private Sub Form_Click()
    If Cartel Then Cartel = False

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                
                    If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                    
                   If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                          '  Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan r?pido.", .red, .green, .blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                         '       Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan r?pido.", .red, .green, .blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                  '  Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan r?pido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                   ' Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    'If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                ' Descastea
                If UsingSkill = Magia Or UsingSkill = Proyectiles Then
                    frmMain.MousePointer = vbDefault
                    UsingSkill = 0
                Else
                    
                    Call WriteRightClick(tX, tY)
                End If
                
                'Call AbrirMenuViewPort
            End If
        If MouseBoton = vbRightButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
        End If
End Sub
Private Sub Form_DblClick()
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/27/2007
'12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
'**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteDoubleClick(tX, tY)
    End If
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub lblDropGold_Click()

    Inventario.SelectGold
    If UserGLD > 0 Then
        If Not Comerciando Then FrmCantidad.Show , frmMain
    End If
    
End Sub

Private Sub Label4_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.path & "\Recursos\Centroinventario.JPG")

    Panel = eVentanas.vInventario

    If Panel <> LastPanel Then
        Call WriteSetMenu(Panel, 255)
        LastPanel = Panel
    End If

    ' Activo controles de inventario

    PicInv.Visible = True
            IconosegD.Visible = True
    IconoSeg.Visible = True
    Label6.Visible = True
    
    'imgInvScrollUp.Visible = True
    'imgInvScrollDown.Visible = True

    ' Desactivo controles de hechizo
    hlst.Visible = False
    cmdINFO.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    
End Sub
Private Sub label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
        Inventario.uMoveItem = False
    Inventario.sMoveItem = False
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.path & "\Recursos\Centrohechizos.JPG")
    
    Panel = eVentanas.vHechizos

    If Panel <> LastPanel Then

        Dim TempInv As Byte

        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
            TempInv = CByte(Inventario.SelectedItem)
        Else
            TempInv = 255 ' @@ Pasamos y tenemos ningun slot seleccionado entonces 255 ...

        End If

        Call WriteSetMenu(Panel, TempInv)
        LastPanel = Panel

    End If
    
    ' Activo controles de hechizos
    hlst.Visible = True
    cmdINFO.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    ' Desactivo controles de inventario
    PicInv.Visible = False
    IconosegD.Visible = False
    IconoSeg.Visible = False
    Label6.Visible = False
    'imgInvScrollUp.Visible = False
    'imgInvScrollDown.Visible = False

End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
        Inventario.uMoveItem = False
    Inventario.sMoveItem = False
End Sub

Private Sub picInv_DblClick()
        
    If (mouse_Down <> False) And (mouse_UP = True) Then Exit Sub
      
    mouse_UP = False
    ' x button
    
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo

        Inventario.uMoveItem = False
        
        If MouseInvBoton = vbRightButton Then Exit Sub
    
    

        Call UsarItem(0)


    
End Sub

Private Sub PicInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
                           
    '    / x button
    If (mouse_Down = False) Then Exit Sub
    mouse_Down = False
    mouse_UP = True
    '    / x button
    
    Call Audio.PlayWave(SND_CLICK)
    Inventario.uMoveItem = False
    MouseInvBoton = Button
End Sub

Private Sub RecTxt_Change()
On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If Not Application.IsAppActive() Then Exit Sub
    
   If SendTxt.Visible Then
        SendTxt.SetFocus
  ElseIf Me.SendRmstxt.Visible Then
        SendRmstxt.SetFocus
    ElseIf Me.SendGms.Visible Then
        SendGms.SetFocus
        ElseIf SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    ElseIf (Not Comerciando) And (Not MirandoAsignarSkills) And _
        (Not frmMSG.Visible) And (Not MirandoForo) And _
        (Not frmEstadisticas.Visible) And (Not FrmCantidad.Visible) Then
         
        If PicInv.Visible Then
            PicInv.SetFocus
        ElseIf hlst.Visible Then
            hlst.SetFocus
        End If
    End If
End Sub
Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If PicInv.Visible Then
        PicInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub
Private Function InGameArea() As Boolean
'***************************************************
'Author: NicoNZ
'Last Modification: 04/07/08
'Checks if last click was performed within or outside the game area.
'***************************************************
    If clicX < MainViewShp.Left Or clicX > MainViewShp.Left + MainViewShp.Width Then Exit Function
    If clicY < MainViewShp.Top Or clicY > MainViewShp.Top + MainViewShp.Height Then Exit Function
    
    InGameArea = True
End Function
Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - imped? se inserten caract?res no imprimibles
'**************************************************************

If Pulsacion_Fisica = False Then
Exit Sub
End If
Pulsacion_Fisica = True

    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = ""
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
        frmMain.SendTxt.SetFocus
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()
    'Clean input and output buffers
#If SeguridadAlkon Then
    Call ConnectionStablished(Socket1.PeerAddress)
#End If
    
    Second.Enabled = True

    Select Case EstadoLogin
Case E_MODO.BorrarPJ
       FrmRECBORR.Show vbModal
        Case E_MODO.RecuperarPJ
       FrmRECBORR.Show vbModal
        Case E_MODO.CrearNuevoPj
#If SeguridadAlkon Then
            Call mi(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
        
        Case E_MODO.Normal
#If SeguridadAlkon Then
            Call mi(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
        
        Case E_MODO.Dados
#If SeguridadAlkon Then
            Call mi(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If

            frmCrearPersonaje.Show vbModal
    End Select
End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long
    
    Second.Enabled = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    Do While i < Forms.Count - 1
        i = i + 1
        
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name And Forms(i).Name <> frmCrearPersonaje.Name Then
            Unload Forms(i)
        End If
    Loop
    
    On Local Error GoTo 0
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If
    
    frmMain.Visible = False
    
    pausa = False
    UserMeditar = False
    
#If SeguridadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    For i = 1 To MAX_INVENTORY_SLOTS
        
    Next i
    
    macrotrabajo.Enabled = False

    SkillPoints = 0
    Alocados = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim RD As String
    Dim data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub
    
#If SeguridadAlkon Then
    Call DataReceived(data)
#End If
    
    'Put data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim M As New frmMenuseFashion
            
            Load M
            M.SetCallback Me
            M.SetMenuId 1
            M.ListaInit 2, False
            
            If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
                M.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
            Else
                M.ListaSetItem 0, "<NPC>", True
            End If
            M.ListaSetItem 1, "Comerciar"
            
            M.ListaFin
            M.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
            Call UsarItem(0)
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        Call WriteLeftClick(tX, tY)
        
    Case 1 'Comerciar
        Call WriteLeftClick(tX, tY)
        Call WriteCommerceStart
    End Select
End Select
End Sub


'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"
    
    Second.Enabled = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    Do While i < Forms.Count - 1
        i = i + 1
        
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name And Forms(i).Name <> frmCrearPersonaje.Name Then
            Unload Forms(i)
        End If
    Loop
    On Local Error GoTo 0
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Debug.Print "Winsock Connect"
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
#If SeguridadAlkon Then
    Call ConnectionStablished(Winsock1.RemoteHostIP)
#End If
    
    Second.Enabled = True
    
    Select Case EstadoLogin
    Case E_MODO.BorrarPJ
       FrmRECBORR.Show vbModal
        Case E_MODO.RecuperarPJ
       FrmRECBORR.Show vbModal
        Case E_MODO.CrearNuevoPj
#If SeguridadAlkon Then
            Call mi(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
            
Case E_MODO.BorrarPersonaje
#If SeguridadAlkon Then
            Call mi(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
            

        Case E_MODO.Normal
#If SeguridadAlkon Then
            Call mi(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login

        Case E_MODO.Dados
#If SeguridadAlkon Then
            Call mi(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Audio.PlayMIDI("7.mid")
            frmCrearPersonaje.Show vbModal
            
#If SeguridadAlkon Then
            Call ProtectForm(frmCrearPersonaje)
#End If
    End Select
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim data() As Byte
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD
    
    data = StrConv(RD, vbFromUnicode)
    
#If SeguridadAlkon Then
    Call DataReceived(data)
#End If
    
    'Set data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close

    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub
#End If


Private Sub Winsock2_Connect()
#If SeguridadAlkon = 1 Then
    Call modURL.ProcessRequest
#End If
End Sub
Public Sub DesactivarMacroHechizos()
    TrainingMacro.Enabled = False
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True)
    Call ControlSM(eSMType.mSpells, False)
End Sub

Private Sub PicInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Position As Integer
Dim i As Long
Dim file_path As String
Dim data() As Byte
Dim bmpInfo As BITMAPINFO
Dim handle As Integer
Dim bmpData As StdPicture

    '    / x button
    mouse_Down = True
    mouse_UP = False
    '    / x button

If (Button = vbRightButton) And (Not Comerciando) Then

If Inventario.GrhIndex(Inventario.SelectedItem) < 1 Then
  'Call MsgBox("Primero debes seleccionar un item de tu inventario.", vbCritical + vbOKOnly)
  Exit Sub
  End If
  
If Inventario.GrhIndex(Inventario.SelectedItem) > 0 Then

        Last_I = Inventario.SelectedItem
        If Last_I > 0 And Last_I <= MAX_INVENTORY_SLOTS Then
                    
            Position = Search_GhID(Inventario.GrhIndex(Inventario.SelectedItem))
            
            If Position = 0 Then
                i = GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum
                Call Get_Bitmapp(DirGraficos, CStr(GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum) & ".BMP", bmpInfo, data)
                Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
                frmMain.ImageList1.ListImages.Add , CStr("g" & Inventario.GrhIndex(Inventario.SelectedItem)), Picture:=bmpData
                Position = frmMain.ImageList1.ListImages.Count
                Set bmpData = Nothing
            End If
            
            
            Inventario.uMoveItem = True
            
            Set PicInv.MouseIcon = frmMain.ImageList1.ListImages(Position).ExtractIcon
            frmMain.PicInv.MousePointer = vbCustom

            Exit Sub
        End If
    End If
End If
End Sub

Private Function Search_GhID(gh As Integer) As Integer

Dim i As Integer

For i = 1 To frmMain.ImageList1.ListImages.Count
    If frmMain.ImageList1.ListImages(i).Key = "g" & CStr(gh) Then
        Search_GhID = i
        Exit For
    End If
Next i

End Function

Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
Call Protocol.WriteDragInventory(originalSlot, newSlot, eMoveType.Inventory)
Inventario.uMoveItem = False
Inventario.sMoveItem = False
End Sub
Private Sub Label2_Click()
If UserLvl < 47 Then
Call ShowConsoleMsg("Nivel: " & UserLvl & " Experiencia: " & Format$(UserExp, "#,###") & "/" & Format$(UserPasarNivel, "#,###") & " " & "(" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%)", 0, 240, 240)
Else
Call AddtoRichTextBox(frmMain.RecTxt, "Nivel: " & UserLvl & " ^\\\\\\M?ximo///////^", 0, 200, 200, False, False, True)
End If
End Sub

Private Sub label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lvllbl(0).Visible = False
    lvllbl(1).Visible = False
    lvllbl(2).Visible = False
    lvllbl(3).Visible = False
    lvllbl(4).Visible = False
    lblporclvl(0).Visible = True
    lblporclvl(1).Visible = True
    lblporclvl(2).Visible = True
    lblporclvl(3).Visible = True
    lblporclvl(4).Visible = True
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
        
        If PicInv.Visible Then
            PicInv.SetFocus
        Else
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendCMSTXT.Text)
            CharAscii = Asc(mid$(SendCMSTXT.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendCMSTXT.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendCMSTXT.Text = tempstr
        End If
        
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Not GetAsyncKeyState(KeyCode) < 0 Then
Pulsacion_Fisica = False
Exit Sub
End If
Pulsacion_Fisica = True
End Sub
Private Sub SendRMSTXT_Change()
    stxtbufferrmsg = SendRmstxt.Text
End Sub
Private Sub sendgms_change()
stxtbufferrmsg = SendGms.Text
End Sub
Private Sub SendRMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbufferrmsg <> "" Then
            Call ParseUserCommand("/RMSG " & stxtbufferrmsg)
        End If
       ' frmMain.Label2 = ""
        stxtbufferrmsg = ""
        SendRmstxt.Text = ""
        KeyCode = 0
        Me.SendRmstxt.Visible = False
    End If
End Sub

Private Sub SendRMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub
Private Sub SendGms_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbufferrmsg <> "" Then
            Call ParseUserCommand("/GMSG " & stxtbufferrmsg)
        End If
       ' frmMain.Label2 = ""
        stxtbufferrmsg = ""
        SendGms.Text = ""
        KeyCode = 0
        Me.SendGms.Visible = False
    End If
End Sub

Private Sub SendGms_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub
