VERSION 5.00
Begin VB.Form FrmPjsMao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   Picture         =   "FrmPjsMao.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3600
      Top             =   2280
   End
   Begin VB.ListBox lstPjs 
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
      Height          =   2790
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2520
   End
   Begin VB.Label lblValor 
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
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   730
      TabIndex        =   1
      Top             =   3420
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   120
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Image ImgComprar 
      Height          =   375
      Left            =   120
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Image ImgOfrecer 
      Height          =   255
      Left            =   1680
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   1680
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPjsMao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Image1_Click()
Unload Me
'FrmMercado.SetFocus
End Sub


Private Sub Image2_Click()
    Call Audio.PlayWave(SND_CLICK)
    
    If lstPjs.ListIndex = -1 Then
        MsgBox "Selecciona un personaje para ver la información"
        Exit Sub
    End If
    
    Call WriteRequestInfoCharMAO(lstPjs.List(lstPjs.ListIndex))
    Unload Me
End Sub

Private Sub imgComprar_Click()
    Call Audio.PlayWave(SND_CLICK)
    
    If lstPjs.ListIndex = -1 Then
        MsgBox "Selecciona un personaje"
        Exit Sub
    End If
    
    If lblValor.Caption = "X CAMBIO" Then
        MsgBox "El personaje con el que intentas cambiar está en MODO CAMBIO"
        Exit Sub
    End If

    If lstPjs.List(lstPjs.ListIndex) <> vbNullString Then
        If MsgBox("¿Seguro que desea comprar el personaje " & lstPjs.List(lstPjs.ListIndex) & "?", vbYesNo + vbQuestion, "Mercado Desterium ") = vbYes Then
            Protocol.WriteBuyPj (UCase$(lstPjs.List(lstPjs.ListIndex)))
        End If
    End If
    
End Sub

Private Sub ImgOfrecer_Click()
    Call Audio.PlayWave(SND_CLICK)
    
    If lstPjs.ListIndex = -1 Then
        MsgBox "Selecciona un personaje"
        Exit Sub
    End If
    
    If Not lblValor.Caption = "X CAMBIO" Then
        MsgBox "El personaje con el que intentas cambiar no está en MODO CAMBIO"
        Exit Sub
    End If
    
    If lstPjs.List(lstPjs.ListIndex) <> vbNullString Then
        If MsgBox("¿Seguro que deseas ofertar tu personaje a cambio de este? Tu contraseña/pin/email pasarán a ser los datos del personaje recibido.", vbYesNo + vbQuestion, "Desterium  AO") = vbYes Then
            Call Protocol.WriteMercadoInvitation(UCase$(lstPjs.List(lstPjs.ListIndex)))
        Else
            Exit Sub
        End If
    End If

End Sub

Private Sub lstPjs_Click()
    Call WriteRequestTipoMAO(lstPjs.List(lstPjs.ListIndex))
End Sub
