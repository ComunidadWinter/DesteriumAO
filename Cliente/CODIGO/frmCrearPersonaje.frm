VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Desterium AO"
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   374
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox DESCRIPCIONCLASE 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   12840
      Width           =   525
   End
   Begin VB.TextBox PIN 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   345
      Left            =   1560
      TabIndex        =   26
      Top             =   1860
      Width           =   2895
   End
   Begin VB.ComboBox lstAlienacion 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":4A919
      Left            =   120
      List            =   "frmCrearPersonaje.frx":4A923
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   12360
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.TextBox txtMail 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1560
      TabIndex        =   3
      Top             =   1470
      Width           =   2895
   End
   Begin VB.TextBox txtConfirmPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1110
      Width           =   2895
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.Timer tAnimacion 
      Left            =   120
      Top             =   0
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":4A936
      Left            =   1200
      List            =   "frmCrearPersonaje.frx":4A938
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2445
      Width           =   1425
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":4A93A
      Left            =   1560
      List            =   "frmCrearPersonaje.frx":4A944
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3360
      Width           =   1305
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":4A957
      Left            =   1200
      List            =   "frmCrearPersonaje.frx":4A959
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2910
      Width           =   1425
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":4A95B
      Left            =   6720
      List            =   "frmCrearPersonaje.frx":4A95D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   8040
      Width           =   1185
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image Picture1 
      Height          =   2655
      Left            =   4920
      Top             =   360
      Width           =   1935
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   5
      Left            =   14280
      Top             =   12270
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   4
      Left            =   14055
      Top             =   12270
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   3
      Left            =   13830
      Top             =   12270
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   2
      Left            =   13605
      Top             =   12270
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   1
      Left            =   13380
      Top             =   12270
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   5
      Left            =   14400
      Top             =   12225
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   4
      Left            =   14175
      Top             =   12225
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   3
      Left            =   13950
      Top             =   12225
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   2
      Left            =   13725
      Top             =   12225
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   5
      Left            =   14400
      Top             =   11940
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   4
      Left            =   14175
      Top             =   11940
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   3
      Left            =   13950
      Top             =   11940
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   2
      Left            =   13725
      Top             =   11940
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   5
      Left            =   14160
      Top             =   12135
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   4
      Left            =   13935
      Top             =   12135
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   3
      Left            =   13710
      Top             =   12135
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   2
      Left            =   13485
      Top             =   12135
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   5
      Left            =   14160
      Top             =   11850
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   4
      Left            =   13935
      Top             =   11850
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   3
      Left            =   13710
      Top             =   11850
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   2
      Left            =   13485
      Top             =   11850
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   1
      Left            =   13500
      Top             =   12225
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   1
      Left            =   13500
      Top             =   11940
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   1
      Left            =   13260
      Top             =   12135
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   1
      Left            =   13260
      Top             =   11850
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   5
      Left            =   14160
      Top             =   11565
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   4
      Left            =   13935
      Top             =   11565
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   3
      Left            =   13710
      Top             =   11565
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   2
      Left            =   13485
      Top             =   11565
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   1
      Left            =   13260
      Top             =   11565
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblEspecialidad 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   13560
      TabIndex        =   25
      Top             =   12600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2580
      TabIndex        =   24
      Top             =   13530
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   2580
      TabIndex        =   23
      Top             =   13230
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2580
      TabIndex        =   22
      Top             =   12945
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2580
      TabIndex        =   21
      Top             =   12645
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2580
      TabIndex        =   20
      Top             =   12360
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1770
      TabIndex        =   19
      Top             =   13530
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1770
      TabIndex        =   18
      Top             =   13230
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1770
      TabIndex        =   17
      Top             =   12945
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1770
      TabIndex        =   16
      Top             =   12645
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1770
      TabIndex        =   15
      Top             =   12360
      Width           =   225
   End
   Begin VB.Image imgAtributos 
      Height          =   270
      Left            =   9120
      Top             =   12840
      Width           =   975
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   2160
      Left            =   9360
      TabIndex        =   14
      Top             =   12600
      Width           =   2160
   End
   Begin VB.Image imgVolver 
      Height          =   525
      Left            =   120
      Top             =   4800
      Width           =   2250
   End
   Begin VB.Image imgCrear 
      Height          =   555
      Left            =   9720
      Top             =   4800
      Width           =   2130
   End
   Begin VB.Image imgalineacion 
      Height          =   240
      Left            =   5760
      Top             =   11760
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image imgGenero 
      Height          =   240
      Left            =   2760
      Top             =   12120
      Width           =   705
   End
   Begin VB.Image imgClase 
      Height          =   360
      Left            =   5520
      Top             =   10560
      Width           =   555
   End
   Begin VB.Image imgRaza 
      Height          =   255
      Left            =   4320
      Top             =   11760
      Width           =   570
   End
   Begin VB.Image imgPuebloOrigen 
      Height          =   225
      Left            =   10200
      Top             =   12120
      Width           =   1785
   End
   Begin VB.Image imgEspecialidad 
      Height          =   240
      Left            =   12210
      Top             =   12570
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image imgArcos 
      Height          =   225
      Left            =   13545
      Top             =   12660
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgArmas 
      Height          =   240
      Left            =   13530
      Top             =   12360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgEscudos 
      Height          =   255
      Left            =   13515
      Top             =   12060
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgVida 
      Height          =   225
      Left            =   13530
      Top             =   11790
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgMagia 
      Height          =   255
      Left            =   13485
      Top             =   11475
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgEvasion 
      Height          =   255
      Left            =   12720
      Top             =   11520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgConstitucion 
      Height          =   255
      Left            =   8400
      Top             =   14640
      Width           =   1080
   End
   Begin VB.Image imgCarisma 
      Height          =   360
      Left            =   8400
      Top             =   14280
      Width           =   765
   End
   Begin VB.Image imgInteligencia 
      Height          =   345
      Left            =   8280
      Top             =   13920
      Width           =   1245
   End
   Begin VB.Image imgAgilidad 
      Height          =   360
      Left            =   8280
      Top             =   13560
      Width           =   1215
   End
   Begin VB.Image imgFuerza 
      Height          =   360
      Left            =   8280
      Top             =   13200
      Width           =   795
   End
   Begin VB.Image imgF 
      Height          =   270
      Left            =   2520
      Top             =   12000
      Width           =   270
   End
   Begin VB.Image imgM 
      Height          =   270
      Left            =   1800
      Top             =   12000
      Width           =   270
   End
   Begin VB.Image imgD 
      Height          =   270
      Left            =   10200
      Top             =   12960
      Width           =   270
   End
   Begin VB.Image imgConfirmPasswd 
      Height          =   375
      Left            =   10560
      Top             =   12960
      Width           =   1320
   End
   Begin VB.Image imgPasswd 
      Height          =   375
      Left            =   10560
      Top             =   12480
      Width           =   1290
   End
   Begin VB.Image imgNombre 
      Height          =   360
      Left            =   4200
      Top             =   12120
      Width           =   1275
   End
   Begin VB.Image imgMail 
      Height          =   480
      Left            =   3720
      Top             =   12720
      Width           =   1395
   End
   Begin VB.Image imgTirarDados 
      Height          =   765
      Left            =   5640
      Top             =   12480
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   11760
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgDados 
      Height          =   1485
      Left            =   9840
      MouseIcon       =   "frmCrearPersonaje.frx":4A95F
      MousePointer    =   99  'Custom
      Top             =   600
      Width           =   1860
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   0
      Picture         =   "frmCrearPersonaje.frx":4AAB1
      Top             =   11640
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   9420
      TabIndex        =   13
      Top             =   1680
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   9420
      TabIndex        =   12
      Top             =   1320
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   9420
      TabIndex        =   11
      Top             =   2040
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   9420
      TabIndex        =   10
      Top             =   1020
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   9420
      TabIndex        =   9
      Top             =   690
      Width           =   225
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private cBotonPasswd As clsGraphicalButton
Private cBotonTirarDados As clsGraphicalButton
Private cBotonMail As clsGraphicalButton
Private cBotonNombre As clsGraphicalButton
Private cBotonConfirmPasswd As clsGraphicalButton
Private cBotonAtributos As clsGraphicalButton
Private cBotonD As clsGraphicalButton
Private cBotonM As clsGraphicalButton
Private cBotonF As clsGraphicalButton
Private cBotonFuerza As clsGraphicalButton
Private cBotonAgilidad As clsGraphicalButton
Private cBotonInteligencia As clsGraphicalButton
Private cBotonCarisma As clsGraphicalButton
Private cBotonConstitucion As clsGraphicalButton
Private cBotonEvasion As clsGraphicalButton
Private cBotonMagia As clsGraphicalButton
Private cBotonVida As clsGraphicalButton
Private cBotonEscudos As clsGraphicalButton
Private cBotonArmas As clsGraphicalButton
Private cBotonArcos As clsGraphicalButton
Private cBotonEspecialidad As clsGraphicalButton
Private cBotonPuebloOrigen As clsGraphicalButton
Private cBotonRaza As clsGraphicalButton
Private cBotonClase As clsGraphicalButton
Private cBotonGenero As clsGraphicalButton
Private cBotonAlineacion As clsGraphicalButton
Private cBotonVolver As clsGraphicalButton
Private cBotonCrear As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private picFullStar As Picture
Private picHalfStar As Picture
Private picGlowStar As Picture

Private Enum eHelp
    iePasswd
    ieTirarDados
    ieMail
    ieNombre
    ieConfirmPasswd
    ieAtributos
    ieD
    ieM
    ieF
    ieFuerza
    ieAgilidad
    ieInteligencia
    ieCarisma
    ieConstitucion
    ieEvasion
    ieMagia
    ieVida
    ieEscudos
    ieArmas
    ieArcos
    ieEspecialidad
    iePuebloOrigen
    ieRaza
    ieClase
    ieGenero
    ieAlineacion
End Enum

Private vHelp(25) As String
Private vEspecialidades() As String

Private Type tModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type

Private Type tModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    DañoArmas As Double
    DañoProyectiles As Double
    Escudo As Double
    Magia As Double
    Vida As Double
    Hit As Double
End Type

Private ModRaza() As tModRaza
Private ModClase() As tModClase

Private NroRazas As Integer
Private NroClases As Integer

Private Cargando As Boolean

Private currentGrh As Long
Private Dir As E_Heading

Private Sub Form_Load()
    'Me.Picture = LoadPicture(DirGraficos & "VentanaCrearPersonaje.jpg")
    
    Cargando = True
    Call LoadCharInfo
    Call CargarEspecialidades
    
    Call IniciarGraficos
    Call CargarCombos
    
    Call LoadHelp
    
    
    Call TirarDados
    
    Cargando = False
    
    'UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    UserHead = 0
    
    lstHogar.Text = "Ullathorpe"
    
#If SeguridadAlkon Then
    Call ProtectForm(Me)
#End If



End Sub

Private Sub CargarEspecialidades()

    ReDim vEspecialidades(1 To NroClases)
    
    vEspecialidades(eClass.Hunter) = "Ocultarse"
    vEspecialidades(eClass.Thief) = "Robar y Ocultarse"
    vEspecialidades(eClass.Assasin) = "Apuñalar"
    'vEspecialidades(eClass.Bandit) = "Combate Sin Armas"
    vEspecialidades(eClass.Druid) = "Domar"
    vEspecialidades(eClass.Pirat) = "Navegar"
End Sub

Private Sub IniciarGraficos()

    Dim GrhPath As String
    GrhPath = DirGraficos
    
    Set cBotonPasswd = New clsGraphicalButton
    Set cBotonTirarDados = New clsGraphicalButton
    Set cBotonMail = New clsGraphicalButton
    Set cBotonNombre = New clsGraphicalButton
    Set cBotonConfirmPasswd = New clsGraphicalButton
    Set cBotonAtributos = New clsGraphicalButton
    Set cBotonD = New clsGraphicalButton
    Set cBotonM = New clsGraphicalButton
    Set cBotonF = New clsGraphicalButton
    Set cBotonFuerza = New clsGraphicalButton
    Set cBotonAgilidad = New clsGraphicalButton
    Set cBotonInteligencia = New clsGraphicalButton
    Set cBotonCarisma = New clsGraphicalButton
    Set cBotonConstitucion = New clsGraphicalButton
    Set cBotonEvasion = New clsGraphicalButton
    Set cBotonMagia = New clsGraphicalButton
    Set cBotonVida = New clsGraphicalButton
    Set cBotonEscudos = New clsGraphicalButton
    Set cBotonArmas = New clsGraphicalButton
    Set cBotonArcos = New clsGraphicalButton
    Set cBotonEspecialidad = New clsGraphicalButton
    Set cBotonPuebloOrigen = New clsGraphicalButton
    Set cBotonRaza = New clsGraphicalButton
    Set cBotonClase = New clsGraphicalButton
    Set cBotonGenero = New clsGraphicalButton
    Set cBotonAlineacion = New clsGraphicalButton
    Set cBotonVolver = New clsGraphicalButton
    Set cBotonCrear = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    'Call cBotonPasswd.Initialize(imgPasswd, "", GrhPath & "BotonContraseña.jpg", _
                                    GrhPath & "BotonContraseña.jpg", Me, , , False, False)
                                    
    'Call cBotonTirarDados.Initialize(imgTirarDados, "", GrhPath & "BotonTirarDados.jpg", _
                                    GrhPath & "BotonTirarDados.jpg", Me, , , False, False)
                                    
    'Call cBotonMail.Initialize(imgMail, "", GrhPath & "BotonMailPj.jpg", _
                                    GrhPath & "BotonMailPj.jpg", Me, , , False, False)
                                    
    'Call cBotonNombre.Initialize(imgNombre, "", GrhPath & "BotonNombrePJ.jpg", _
                                    GrhPath & "BotonNombrePJ.jpg", Me, , , False, False)
                                    
    'Call cBotonConfirmPasswd.Initialize(imgConfirmPasswd, "", GrhPath & "BotonRepetirContraseña.jpg", _
                                    GrhPath & "BotonRepetirContraseña.jpg", Me, , , False, False)
                                    
    'Call cBotonAtributos.Initialize(imgAtributos, "", GrhPath & "BotonAtributos.jpg", _
                                    GrhPath & "BotonAtributos.jpg", Me, , , False, False)
                                    
    'Call cBotonD.Initialize(imgD, "", GrhPath & "BotonD.jpg", _
                                    GrhPath & "BotonD.jpg", Me, , , False, False)
                                    
    'Call cBotonM.Initialize(imgM, "", GrhPath & "BotonM.jpg", _
                                    GrhPath & "BotonM.jpg", Me, , , False, False)
                                    
    'Call cBotonF.Initialize(imgF, "", GrhPath & "BotonF.jpg", _
                                    GrhPath & "BotonF.jpg", Me, , , False, False)
                                    
    'Call cBotonFuerza.Initialize(imgFuerza, "", GrhPath & "BotonFuerza.jpg", _
                                    GrhPath & "BotonFuerza.jpg", Me, , , False, False)
                                    
   ' Call cBotonAgilidad.Initialize(imgAgilidad, "", GrhPath & "BotonAgilidad.jpg", _
                                    GrhPath & "BotonAgilidad.jpg", Me, , , False, False)
                                    
    'Call cBotonInteligencia.Initialize(imgInteligencia, "", GrhPath & "BotonInteligencia.jpg", _
                                    GrhPath & "BotonInteligencia.jpg", Me, , , False, False)
                                    
    'Call cBotonCarisma.Initialize(imgCarisma, "", GrhPath & "BotonCarisma.jpg", _
                                    GrhPath & "BotonCarisma.jpg", Me, , , False, False)
                                    
    'Call cBotonConstitucion.Initialize(imgConstitucion, "", GrhPath & "BotonConstitucion.jpg", _
                                    GrhPath & "BotonConstitucion.jpg", Me, , , False, False)
                                    
   ' Call cBotonEvasion.Initialize(imgEvasion, "", GrhPath & "BotonEvasion.jpg", _
                                    GrhPath & "BotonEvasion.jpg", Me, , , False, False)
                                    
   ' Call cBotonMagia.Initialize(imgMagia, "", GrhPath & "BotonMagia.jpg", _
                                    GrhPath & "BotonMagia.jpg", Me, , , False, False)
                                    
    'Call cBotonVida.Initialize(imgVida, "", GrhPath & "BotonVida.jpg", _
                                    GrhPath & "BotonVida.jpg", Me, , , False, False)
                                    
    'Call cBotonEscudos.Initialize(imgEscudos, "", GrhPath & "BotonEscudos.jpg", _
                                    GrhPath & "BotonEscudos.jpg", Me, , , False, False)
                                    
   ' Call cBotonArmas.Initialize(imgArmas, "", GrhPath & "BotonArmas.jpg", _
                                    GrhPath & "BotonArmas.jpg", Me, , , False, False)
                                    
   ' Call cBotonArcos.Initialize(imgArcos, "", GrhPath & "BotonArcos.jpg", _
                                    GrhPath & "BotonArcos.jpg", Me, , , False, False)
                                    
   ' Call cBotonEspecialidad.Initialize(imgEspecialidad, "", GrhPath & "BotonEspecialidad.jpg", _
                                    GrhPath & "BotonEspecialidad.jpg", Me, , , False, False)
                                    
    'Call cBotonPuebloOrigen.Initialize(imgPuebloOrigen, "", GrhPath & "BotonPuebloOrigen.jpg", _
                                    GrhPath & "BotonPuebloOrigen.jpg", Me, , , False, False)
                                    
    'Call cBotonRaza.Initialize(imgRaza, "", GrhPath & "BotonRaza.jpg", _
                                    GrhPath & "BotonRaza.jpg", Me, , , False, False)
                                    
    'Call cBotonClase.Initialize(imgClase, "", GrhPath & "BotonClase.jpg", _
                                    GrhPath & "BotonClase.jpg", Me, , , False, False)
                                    
    'Call cBotonGenero.Initialize(imgGenero, "", GrhPath & "BotonGenero.jpg", _
                                    GrhPath & "BotonGenero.jpg", Me, , , False, False)
                                    
   ' Call cBotonAlineacion.Initialize(imgalineacion, "", GrhPath & "BotonAlineacion.jpg", _
                                    GrhPath & "BotonAlineacion.jpg", Me, , , False, False)
                                    
   ' Call cBotonVolver.Initialize(imgVolver, "", GrhPath & "BotonVolverRollover.jpg", _
                                    GrhPath & "BotonVolverClick.jpg", Me)
                                    
    'Call cBotonCrear.Initialize(imgCrear, "", GrhPath & "BotonCrearPersonajeRollover.jpg", _
                                    GrhPath & "BotonCrearPersonajeClick.jpg", Me)

     
   ' Set picFullStar = LoadPicture(GrhPath & "EstrellaSimple.jpg")
    'Set picHalfStar = LoadPicture(GrhPath & "EstrellaMitad.jpg")
    'Set picGlowStar = LoadPicture(GrhPath & "EstrellaBrillante.jpg")

End Sub

Private Sub CargarCombos()
    Dim i As Integer
    
    lstProfesion.Clear
    
    For i = LBound(ListaClases) To NroClases
        lstProfesion.AddItem ListaClases(i)
    Next i
    
    lstHogar.Clear
    
    For i = LBound(Ciudades()) To UBound(Ciudades())
        lstHogar.AddItem Ciudades(i)
    Next i
    
    lstRaza.Clear
    
    For i = LBound(ListaRazas()) To NroRazas
        lstRaza.AddItem ListaRazas(i)
    Next i
    
    lstProfesion.ListIndex = 1
End Sub

Function CheckData() As Boolean
    If txtPasswd.Text <> txtConfirmPasswd.Text Then
        MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
        Exit Function
    End If
    If Len(txtPasswd.Text) < 4 Then
    MsgBox "¡¡La contraseña debe llevar más de 4 caracteres!!"
    Exit Function
End If
 
                         If Len(PIN.Text) < 3 Then
    MsgBox "¡¡La clave de PIN debe de tener más de 3 caracteres!!"
    Exit Function
End If

    
    If Not CheckMailString(txtMail.Text) Then
        MsgBox "Direccion de mail invalida."
        Exit Function
    End If

    If UserRaza = 0 Then
        MsgBox "Seleccione la raza del personaje."
        Exit Function
    End If
    
    If UserSexo = 0 Then
        MsgBox "Seleccione el sexo del personaje."
        Exit Function
    End If
    
    If UserClase = 0 Then
        MsgBox "Seleccione la clase del personaje."
        Exit Function
    End If
    
    If UserHogar = 0 Then
        MsgBox "Seleccione el hogar del personaje."
        Exit Function
    End If
    
    Dim i As Integer
    For i = 1 To NUMATRIBUTOS
        If UserAtributos(i) = 0 Then
            MsgBox "Los atributos del personaje son invalidos."
            Exit Function
        End If
    Next i
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    CheckData = True

End Function

Private Sub TirarDados()
    Call WriteThrowDices
    Call FlushBuffer
End Sub

Private Sub DirPJ_Click(index As Integer)
    Select Case index
        Case 0
            Dir = CheckDir(Dir + 1)
        Case 1
            Dir = CheckDir(Dir - 1)
    End Select
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearLabel
End Sub


Private Sub imgCrear_Click()

    Dim i As Integer
    Dim CharAscii As Byte
    
    UserName = txtNombre.Text
            
            If Len(txtNombre.Text) < 3 Then
    MsgBox "¡¡El nombre debe de tener mas de 3 carácteres!!"
    Exit Sub
End If
 
If Len(txtNombre.Text) >= 12 Then
    MsgBox "¡¡El nombre debe de tener menos de 12 caracteres!!"
    Exit Sub
End If
            
                        If Len(txtPasswd.Text) < 3 Then
    MsgBox "¡¡La contraseña debe de tener más de 3 caracteres!!"
    Exit Sub
End If

                        If Len(PIN.Text) < 3 Then
    MsgBox "¡¡La clave de PIN debe de tener más de 3 caracteres!!"
    Exit Sub
End If
            
            
Dim AllCr As Long
Dim CantidadEsp As Byte
Dim thiscr As String
 
Do
    AllCr = AllCr + 1
    If AllCr > Len(UserName) Then Exit Do
    thiscr = mid(UserName, AllCr, 1)
    If InStr(1, " ", UCase(thiscr)) = 1 Then
           CantidadEsp = CantidadEsp + 1
    End If
Loop
If CantidadEsp > 1 Then
     MsgBox "Nick inválido. El nombre no puede tener mas de un espacio."
     Exit Sub
End If
       UserName = Trim$(UserName)
     
    
    UserRaza = lstRaza.ListIndex + 1
    UserSexo = lstGenero.ListIndex + 1
    UserClase = lstProfesion.ListIndex + 1
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = Val(lblAtributos(i).Caption)
    Next i
    
    UserHogar = lstHogar.ListIndex + 1
    
    If Not CheckData Then Exit Sub
    
#If SeguridadAlkon Then
    UserPassword = MD5.GetMD5String(txtPasswd.Text)
    Call MD5.MD5Reset
#Else
    UserPassword = txtPasswd.Text
#End If
    
    For i = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, i, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Sub
        End If
    Next i
       If PIN = "" Then MsgBox "Escribe Un PIN": Exit Sub
    UserPin = PIN
    
    If PIN.Text = txtPasswd.Text Then
                MsgBox "¡¡Tu clave de PIN no puede ser igual a tu contraseña!!"
    Exit Sub
    End If
    
    UserEmail = txtMail.Text
    
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
#End If
    
    EstadoLogin = E_MODO.CrearNuevoPj
    
#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State <> sckConnected Then
#End If
        MsgBox "Error: Se ha perdido la conexion con el server."
        Unload Me
        
    Else
        Call Login
    End If
    
    bShowTutorial = True
End Sub

Private Sub imgDados_Click()
    Call Audio.PlayWave(SND_DICE)
            Call TirarDados
End Sub

Private Sub imgEspecialidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEspecialidad)
End Sub

Private Sub imgNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub imgPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.iePasswd)
End Sub

Private Sub imgConfirmPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)
End Sub

Private Sub imgAtributos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAtributos)
End Sub

Private Sub imgD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieD)
End Sub

Private Sub imgM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieM)
End Sub

Private Sub imgF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieF)
End Sub

Private Sub imgFuerza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieFuerza)
End Sub

Private Sub imgAgilidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAgilidad)
End Sub

Private Sub imgInteligencia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieInteligencia)
End Sub

Private Sub imgCarisma_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieCarisma)
End Sub

Private Sub imgConstitucion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConstitucion)
End Sub

Private Sub imgArcos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArcos)
End Sub

Private Sub imgArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArmas)
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEscudos)
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEvasion)
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMagia)
End Sub

Private Sub imgMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMail)
End Sub

Private Sub imgVida_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieVida)
End Sub

Private Sub imgTirarDados_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieTirarDados)
End Sub

Private Sub imgPuebloOrigen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.iePuebloOrigen)
End Sub

Private Sub imgRaza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieRaza)
End Sub

Private Sub imgClase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieClase)
End Sub

Private Sub imgGenero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieGenero)
End Sub

Private Sub imgalineacion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAlineacion)
End Sub



Private Sub imgVolver_Click()

    
    bShowTutorial = False
    
    Unload Me
    
    frmConnect.Show
End Sub
Private Sub lstGenero_Click()
    UserSexo = lstGenero.ListIndex + 1
End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
'    Image1.Picture = LoadPicture(App.path & "\Recursos\" & lstProfesion.Text & ".jpg")
'
    
    
    UserClase = lstProfesion.ListIndex + 1
    
    Select Case lstProfesion.Text
 
Case Is = "Mago"
       DESCRIPCIONCLASE.Text = "Mago : Es la clase por excelencia preferida por muchos de los jugadores. Si te interesan mucho los hechizos, encantamientos e invocaciones, esta clase es la indicada, ya que posee mucha maná. Pero la desventaja de esta clase es el hecho de tener poca vida como también poca evasión, lo que significa que las clases de cuerpo a cuerpo le van a acertar mucho más facilmente los golpes. Ventajas: Mucha maná y la mejor resistencia mágica. Desventajas: Poca evasión y pocos puntos de vida."
        Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Mago.gif")
Case Is = "Clerigo"
       DESCRIPCIONCLASE.Text = "Clérigo : Sabios y fuertes, es una clase que se especializa tanto en las artes mágicas como la lucha de cuerpo a cuerpo. Aunque sus golpes no sean tan fuertes puede combinar sus dos especialidades. Así como el druida, el clérigo tiene mucho maná, pero no tanto como el Mago. Ventajas: Buena defensa física y Mágica. Desventajas: Tiene un daño equitativo."
        Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Clerigo.gif")
Case Is = "Guerrero"
       DESCRIPCIONCLASE.Text = "Guerrero : Una clase que no usa magias sino que directamente usa la fuerza y combate cuerpo a cuerpo, sus golpes resultan ser increíblemente devastadores cuando éste se encuentra en su punto más elevado, posee una gran vida pero su desventaja es que no puede usar magias, esto hace que sea una presa fácil si no está acompañado. Aunque siempre puede contar con su velocidad con el arco y la flecha. Ventajas: Muchos puntos de vida, mucha fuerza, mucha defensa física. Desventajas: No tiene mana."
        Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Guerrero.gif")
Case Is = "Asesino"
       DESCRIPCIONCLASE.Text = "Asesino : Simplemente un ser sanguinario, la característica especial de esta clase es apuñalar mortalmente, esto significa que de un golpe podrías dejar a tu enemigo prácticamente muerto, su evasión es la mejor y no conoce el miedo contra quienes intenten pegarle. Ventajas: Mucha evasión y golpe mortal. Desventajas: Poco maná."
        Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Asesino.gif")
Case Is = "Ladron"
       DESCRIPCIONCLASE.Text = "Ladrón: Sigiloso delincuente. Quizás no sean poderosamente buenos con la espada, quizás no usen magias, no domen bestias, pero si te cruzas con ellos corre o ataca rápido dado que su gran habilidad es el hurto, con esta gran habilidad puede dejar a grandes guerreros totalmente indefensos, despojándolos de todas sus pertenecías mas preciadas. Son deshonrosos y delincuentes pero esto es solo una porción de como verdaderamente describirlos. Ventajas: Pueden ocultarse durante demasiado tiempo y robarte pertenencias u oro. Desventajas: Pocos puntos de vida, poca resistencia mágica y no tiene maná."
        Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Ladron.gif")
'Case Is = "Bandido"
       'DESCRIPCIONCLASE.Text = "Bandido: Sigiloso guerrero que puede ser letal si no lo ves a tiempo. Tiene casi tantos puntos de vida como un paladín pero su maná es mas limitada. El mismo cuenta con una habilidad para ocultarse que no la tienen ninguna otra clase y si te descuidas puedes recibir un golpe crítico y mortal de su parte. Ventajas: Puede ocultarse durante mucho tiempo y dar golpes críticos. Desventajas: Poca resistencia mágica y poco maná."
       ' Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\bandido.jpg")
Case Is = "Bardo"
       DESCRIPCIONCLASE.Text = "Bardo : Una clase muy práctica frente a las clases de cuerpo a cuerpo ya que posee una gran evasión lo que resulta difícil para enemigo cuando intente acertarle un golpe.  Sus ataques mágicos son muy poderosos aún más cuando usa un ítem especial que le da bonificación de poder mágico. Ventajas: mucho poder mágico y mucha evasión contra golpes. Desventajas: Poca resistencia mágica."
        Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Bardo.gif")
Case Is = "Druida"
       DESCRIPCIONCLASE.Text = "Druida : Amantes de la naturaleza, una clase que se especializa en domar criaturas y usarlas como mascotas, también usa conjuros para invocar otros tipos de criaturas que acuden a su ayuda, lo que permite que en su entrenamiento nunca esté solo. Tiene un hechizo especial que consiste en tomar la imagen de alguien o algo y transformarse, esto sirve para confundir a los enemigos, ideal para despistar en una batalla. Aunque no tiene tanta maná como el Mago, está entre las clases que más poseen. Ventajas: Mucho poder mágico y muy buena defensa mágica. Desventajas: Poca evasión."
        Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Druida.gif")
Case Is = "Paladin"
       DESCRIPCIONCLASE.Text = "Paladín : El paladín es una clase con mucha vida y una gran fuerza, ideal para encuentro contra otra clase que lucha cuerpo a cuerpo ya que sus golpes son algo más débiles que los del guerrero. Una de las desventajas es que su maná que es muy limitada y eso dificulta sus peleas contra clases mágicas. Ventajas: Buena Fuerza, buena defensa cuerpo a cuerpo y muchos puntos de vida. Desventajas: Poca resistencia mágica, poco mana."
        Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Paladin.gif")
Case Is = "Cazador"
       DESCRIPCIONCLASE.Text = "Cazador : Clase que no usa magia, pero que es muy hábil usando armas a distancias, tiene la habilidad de poder ocultarse entre las sombras cuando éste usa su armadura de cazador, tiene una bonificación de daño crítico para mejorar el rendimiento de su entrenamiento y esto hace que resulte fácil de entrenar. Sus fuertes ataques a distancia, hacen que cualquiera evite acercarse a él, aunque si está solo, el no tener magia le trae muchos problemas. Ventajas: Ataque a distancia de gran poder, habilidad de ocultarse. Desventajas: No tiene mana."
        Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Cazador.gif")
Case Is = "Trabajador"
       DESCRIPCIONCLASE.Text = "Trabajador : Tienen gran conocimiento en la construcción o extracción de bienes materiales para la subsistencia de las ciudades. Poseen gran cantidad de puntos de vida ya que los mismos se arriesgan a tareas en lugares peligrosos. Ventajas: Pueden construir o extraer materiales para la construcción o para la subsistencia. Desventajas: No tienen maná ni dominio en artes mágicas o de combate."
        Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\trabajador.gif")
Case Is = "Pirata"
       DESCRIPCIONCLASE.Text = "Pirata : Explorador sin miedo. Se dice que esta fabulosa clase dedico su vida a al conocimiento del mar por ende antiguos escritos dicen que mas de la mitad de las tierras de Desterium  fueron descubiertas por ellos aunque son muy aptos para la navegación también son comerciantes despreciable y grandes estafadores si te cruzas con ellos mas vale corre pues nunca andan solos. Ventajas: Amplio espacio en el inventario, al igual que vida y habilidades de navegación. Desventajas: No tiene mana."
        Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Pirata.gif")
End Select
    
    Call UpdateStats
    Call UpdateEspecialidad(UserClase)
End Sub

Private Sub UpdateEspecialidad(ByVal eClase As eClass)
    lblEspecialidad.Caption = vEspecialidades(eClase)
End Sub

Private Sub lstRaza_Click()
    UserRaza = lstRaza.ListIndex + 1
    
    Call UpdateStats
End Sub

Private Sub picHead_Click(index As Integer)
    ' No se mueve si clickea al medio
    If index = 2 Then Exit Sub
    
    Dim Counter As Integer
    Dim Head As Integer
    
    Head = UserHead
    
    
    UserHead = Head
    
    
End Sub



Private Sub PIN_CLICK()
MsgBox "Recuerda colocar datos que sólo tu sepas para evitar robos o pérdidas. También para poder acceder a funciones como borrar o recuperar el personaje, intercambiar tu personaje dentro del juego."
End Sub


Private Sub txtConfirmPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)
End Sub

Private Sub txtMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMail)
End Sub

Private Sub txtNombre_Change()
    txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Function CheckDir(ByRef Dir As E_Heading) As E_Heading

    If Dir > E_Heading.WEST Then Dir = E_Heading.NORTH
    If Dir < E_Heading.NORTH Then Dir = E_Heading.WEST
    
    CheckDir = Dir
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    If currentGrh > 0 Then _
        tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)

End Function

Private Sub LoadHelp()
    vHelp(eHelp.iePasswd) = "La contraseña que utilizarás para conectar tu personaje al juego."
    vHelp(eHelp.ieTirarDados) = "Presionando sobre los dados, se modificarán al azar los atributos de tu personaje, de esta manera puedes elegir los que más te parezcan para definir a tu personaje."
    vHelp(eHelp.ieMail) = "Es sumamente importante que ingreses una dirección de correo electrónico válida, ya que en el caso de perder la contraseña de tu personaje, se te enviará cuando lo requieras, a esa dirección."
    vHelp(eHelp.ieNombre) = "Sé cuidadoso al seleccionar el nombre de tu personaje. Argentum es un juego de rol, un mundo mágico y fantástico, y si seleccionás un nombre obsceno o con connotación política, los administradores borrarán tu personaje y no habrá ninguna posibilidad de recuperarlo."
    vHelp(eHelp.ieConfirmPasswd) = "La contraseña que utilizarás para conectar tu personaje al juego."
    vHelp(eHelp.ieAtributos) = "Son las cualidades que definen tu personaje. Generalmente se los llama ""Dados"". (Ver Tirar Dados)"
    vHelp(eHelp.ieD) = "Son los atributos que obtuviste al azar. Presioná la esfera roja para volver a tirarlos."
    vHelp(eHelp.ieM) = "Son los modificadores por raza que influyen en los atributos de tu personaje."
    vHelp(eHelp.ieF) = "Los atributos finales de tu personaje, de acuerdo a la raza que elegiste."
    vHelp(eHelp.ieFuerza) = "De ella dependerá qué tan potentes serán tus golpes, tanto con armas de cuerpo a cuerpo, a distancia o sin armas."
    vHelp(eHelp.ieAgilidad) = "Este atributo intervendrá en qué tan bueno seas, tanto evadiendo como acertando golpes, respecto de otros personajes como de las criaturas a las q te enfrentes."
    vHelp(eHelp.ieInteligencia) = "Influirá de manera directa en cuánto maná ganarás por nivel."
    vHelp(eHelp.ieCarisma) = "Será necesario tanto para la relación con otros personajes (entrenamiento en parties) como con las criaturas (domar animales)."
    vHelp(eHelp.ieConstitucion) = "Afectará a la cantidad de vida que podrás ganar por nivel."
    vHelp(eHelp.ieEvasion) = "Evalúa la habilidad esquivando ataques físicos."
    vHelp(eHelp.ieMagia) = "Puntúa la cantidad de maná que se tendrá."
    vHelp(eHelp.ieVida) = "Valora la cantidad de salud que se podrá llegar a tener."
    vHelp(eHelp.ieEscudos) = "Estima la habilidad para rechazar golpes con escudos."
    vHelp(eHelp.ieArmas) = "Evalúa la habilidad en el combate cuerpo a cuerpo con armas."
    vHelp(eHelp.ieArcos) = "Evalúa la habilidad en el combate a distancia con arcos. "
    vHelp(eHelp.ieEspecialidad) = ""
    vHelp(eHelp.iePuebloOrigen) = "Define el hogar de tu personaje. Sin embargo, el personaje nacerá en Nemahuak, la ciudad de los novatos."
    vHelp(eHelp.ieRaza) = "De la raza que elijas dependerá cómo se modifiquen los dados que saques. Podés cambiar de raza para poder visualizar cómo se modifican los distintos atributos."
    vHelp(eHelp.ieClase) = "La clase influirá en las características principales que tenga tu personaje, asi como en las magias e items que podrá utilizar. Las estrellas que ves abajo te mostrarán en qué habilidades se destaca la misma."
    vHelp(eHelp.ieGenero) = "Indica si el personaje será masculino o femenino. Esto influye en los items que podrá equipar."
    vHelp(eHelp.ieAlineacion) = "Indica si el personaje seguirá la senda del mal o del bien. (Actualmente deshabilitado)"
End Sub

Private Sub ClearLabel()
    LastPressed.ToggleToNormal
    lblHelp = ""
End Sub

Private Sub txtNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub txtPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.iePasswd)
End Sub

Public Sub UpdateStats()
    
    Call UpdateRazaMod
    Call UpdateStars
End Sub

Private Sub UpdateRazaMod()
    Dim SelRaza As Integer
    Dim i As Integer
    
    
    If lstRaza.ListIndex > -1 Then
    
        SelRaza = lstRaza.ListIndex + 1
        
        With ModRaza(SelRaza)
            lblModRaza(eAtributos.Fuerza).Caption = IIf(.Fuerza >= 0, "+", "") & .Fuerza
            lblModRaza(eAtributos.Agilidad).Caption = IIf(.Agilidad >= 0, "+", "") & .Agilidad
            lblModRaza(eAtributos.Inteligencia).Caption = IIf(.Inteligencia >= 0, "+", "") & .Inteligencia
            lblModRaza(eAtributos.Carisma).Caption = IIf(.Carisma >= 0, "+", "") & .Carisma
            lblModRaza(eAtributos.Constitucion).Caption = IIf(.Constitucion >= 0, "+", "") & .Constitucion
        End With
    End If
    
    ' Atributo total
    For i = 1 To NUMATRIBUTES
        lblAtributoFinal(i).Caption = Val(lblAtributos(i).Caption) + Val(lblModRaza(i))
    Next i
    
End Sub

Private Sub UpdateStars()
    Dim NumStars As Double
    
    If UserClase = 0 Then Exit Sub
    
    ' Estrellas de evasion
    NumStars = (2.454 + 0.073 * Val(lblAtributoFinal(eAtributos.Agilidad).Caption)) * ModClase(UserClase).Evasion
    Call SetStars(imgEvasionStar, NumStars * 2)
    
    ' Estrellas de magia
    NumStars = ModClase(UserClase).Magia * Val(lblAtributoFinal(eAtributos.Inteligencia).Caption) * 0.085
    Call SetStars(imgMagiaStar, NumStars * 2)
    
    ' Estrellas de vida
    NumStars = 0.24 + (Val(lblAtributoFinal(eAtributos.Constitucion).Caption) * 0.5 - ModClase(UserClase).Vida) * 0.475
    Call SetStars(imgVidaStar, NumStars * 2)
    
    ' Estrellas de escudo
    NumStars = 4 * ModClase(UserClase).Escudo
    Call SetStars(imgEscudosStar, NumStars * 2)
    
    ' Estrellas de armas
    NumStars = (0.509 + 0.01185 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * ModClase(UserClase).Hit * _
                ModClase(UserClase).DañoArmas + 0.119 * ModClase(UserClase).AtaqueArmas * _
                Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArmasStar, NumStars * 2)
    
    ' Estrellas de arcos
    NumStars = (0.4915 + 0.01265 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * _
                ModClase(UserClase).DañoProyectiles * ModClase(UserClase).Hit + 0.119 * ModClase(UserClase).AtaqueProyectiles * _
                Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArcoStar, NumStars * 2)
End Sub

Private Sub SetStars(ByRef ImgContainer As Object, ByVal NumStars As Integer)
    Dim FullStars As Integer
    Dim HasHalfStar As Boolean
    Dim index As Integer
    Dim Counter As Integer

    If NumStars > 0 Then
        
        If NumStars > 10 Then NumStars = 10
        
        FullStars = Int(NumStars / 2)
        
        ' Tienen brillo extra si estan todas
        If FullStars = 5 Then
            For index = 1 To FullStars
                ImgContainer(index).Picture = picGlowStar
            Next index
        Else
            ' Numero impar? Entonces hay que poner "media estrella"
            If (NumStars Mod 2) > 0 Then HasHalfStar = True
            
            ' Muestro las estrellas enteras
            If FullStars > 0 Then
                For index = 1 To FullStars
                    ImgContainer(index).Picture = picFullStar
                Next index
                
                Counter = FullStars
            End If
            
            ' Muestro la mitad de la estrella (si tiene)
            If HasHalfStar Then
                Counter = Counter + 1
                
                ImgContainer(Counter).Picture = picHalfStar
            End If
            
            ' Si estan completos los espacios, no borro nada
            If Counter <> 5 Then
                ' Limpio las que queden vacias
                For index = Counter + 1 To 5
                    Set ImgContainer(index).Picture = Nothing
                Next index
            End If
            
        End If
    Else
        ' Limpio todo
        For index = 1 To 5
            Set ImgContainer(index).Picture = Nothing
        Next index
    End If

End Sub

Private Sub LoadCharInfo()
    Dim SearchVar As String
    Dim i As Integer
    
    NroRazas = UBound(ListaRazas())
    NroClases = UBound(ListaClases())

    ReDim ModRaza(1 To NroRazas)
    ReDim ModClase(1 To NroClases)
    
    'Modificadores de Clase
    For i = 1 To NroClases
        With ModClase(i)
            SearchVar = ListaClases(i)
            
            .Evasion = Val(GetVar(IniPath & "CharInfo.dat", "MODEVASION", SearchVar))
            .AtaqueArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODATAQUEARMAS", SearchVar))
            .AtaqueProyectiles = Val(GetVar(IniPath & "CharInfo.dat", "MODATAQUEPROYECTILES", SearchVar))
            .DañoArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODDAÑOARMAS", SearchVar))
            .DañoProyectiles = Val(GetVar(IniPath & "CharInfo.dat", "MODDAÑOPROYECTILES", SearchVar))
            .Escudo = Val(GetVar(IniPath & "CharInfo.dat", "MODESCUDO", SearchVar))
            .Hit = Val(GetVar(IniPath & "CharInfo.dat", "HIT", SearchVar))
            .Magia = Val(GetVar(IniPath & "CharInfo.dat", "MODMAGIA", SearchVar))
            .Vida = Val(GetVar(IniPath & "CharInfo.dat", "MODVIDA", SearchVar))
        End With
    Next i
    
    'Modificadores de Raza
    For i = 1 To NroRazas
        With ModRaza(i)
            SearchVar = Replace(ListaRazas(i), " ", "")
        
            .Fuerza = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Inteligencia"))
            .Carisma = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Carisma"))
            .Constitucion = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Constitucion"))
        End With
    Next i

End Sub
