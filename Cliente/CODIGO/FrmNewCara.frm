VERSION 5.00
Begin VB.Form FrmNewCara 
   BorderStyle     =   0  'None
   Caption         =   "Cambio de cara"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   LinkTopic       =   "Form4"
   Picture         =   "FrmNewCara.frx":0000
   ScaleHeight     =   2265
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicHead 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   2640
      Picture         =   "FrmNewCara.frx":CC26
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   600
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   2160
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Image ImgCerrar 
      Height          =   255
      Left            =   5040
      Top             =   1680
      Width           =   615
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   0
      Left            =   2760
      Picture         =   "FrmNewCara.frx":10104
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   1
      Left            =   3000
      Picture         =   "FrmNewCara.frx":10416
      Top             =   1320
      Width           =   240
   End
End
Attribute VB_Name = "FrmNewCara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IndexHead As Integer
Private Sub DirPJ_Click(index As Integer)
    If index = 0 Then
        Select Case UserRaza
            Case eRaza.Humano
                If UserSexo = eGenero.Hombre Then
                    If (IndexHead - 1) < LBound(HeadHombre.Humano()) Then Exit Sub
                    IndexHead = IndexHead - 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadHombre.Humano(IndexHead) & ".jpg")
                    
                Else
                    If (IndexHead - 1) < LBound(HeadMujer.Humano()) Then Exit Sub
                    IndexHead = IndexHead - 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadMujer.Humano(IndexHead) & ".jpg")
                End If
                
            Case eRaza.Elfo
                If UserSexo = eGenero.Hombre Then
                    If (IndexHead - 1) < LBound(HeadHombre.Elfo()) Then Exit Sub
                    IndexHead = IndexHead - 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadHombre.Elfo(IndexHead) & ".jpg")
                    
                Else
                    If (IndexHead - 1) < LBound(HeadMujer.Elfo()) Then Exit Sub
                    IndexHead = IndexHead - 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadMujer.Elfo(IndexHead) & ".jpg")
                End If
            Case eRaza.ElfoOscuro
                If UserSexo = eGenero.Hombre Then
                    If (IndexHead - 1) < LBound(HeadHombre.ElfoDrow()) Then Exit Sub
                    IndexHead = IndexHead - 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadHombre.ElfoDrow(IndexHead) & ".jpg")
                    
                Else
                    If (IndexHead - 1) < LBound(HeadMujer.ElfoDrow()) Then Exit Sub
                    IndexHead = IndexHead - 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadMujer.ElfoDrow(IndexHead) & ".jpg")
                End If
            Case eRaza.Gnomo
                If UserSexo = eGenero.Hombre Then
                    If (IndexHead - 1) < LBound(HeadHombre.Gnomo()) Then Exit Sub
                    IndexHead = IndexHead - 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadHombre.Gnomo(IndexHead) & ".jpg")
                    
                Else
                    If (IndexHead - 1) < LBound(HeadMujer.Gnomo()) Then Exit Sub
                    IndexHead = IndexHead - 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadMujer.Gnomo(IndexHead) & ".jpg")
                End If
            Case eRaza.Enano
                If UserSexo = eGenero.Hombre Then
                    If (IndexHead - 1) < LBound(HeadHombre.Enano()) Then Exit Sub
                    IndexHead = IndexHead - 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadHombre.Enano(IndexHead) & ".jpg")
                    
                Else
                    If (IndexHead - 1) < LBound(HeadMujer.Enano()) Then Exit Sub
                    IndexHead = IndexHead - 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadMujer.Enano(IndexHead) & ".jpg")
                End If
        End Select
    Else
        Select Case UserRaza
            Case eRaza.Humano
                If UserSexo = eGenero.Hombre Then
                    If (IndexHead + 1) > UBound(HeadHombre.Humano()) Then Exit Sub
                    IndexHead = IndexHead + 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadHombre.Humano(IndexHead) & ".jpg")
                    
                Else
                    If (IndexHead + 1) > UBound(HeadMujer.Humano()) Then Exit Sub
                    IndexHead = IndexHead + 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadMujer.Humano(IndexHead) & ".jpg")
                End If
                
            Case eRaza.Elfo
                If UserSexo = eGenero.Hombre Then
                    If (IndexHead + 1) > UBound(HeadHombre.Elfo()) Then Exit Sub
                    IndexHead = IndexHead + 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadHombre.Elfo(IndexHead) & ".jpg")
                    
                Else
                    If (IndexHead + 1) > UBound(HeadMujer.Elfo()) Then Exit Sub
                    IndexHead = IndexHead + 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadMujer.Elfo(IndexHead) & ".jpg")
                End If
            Case eRaza.ElfoOscuro
                If UserSexo = eGenero.Hombre Then
                    If (IndexHead + 1) > UBound(HeadHombre.ElfoDrow()) Then Exit Sub
                    IndexHead = IndexHead + 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadHombre.ElfoDrow(IndexHead) & ".jpg")
                    
                Else
                    If (IndexHead + 1) > UBound(HeadMujer.ElfoDrow()) Then Exit Sub
                    IndexHead = IndexHead + 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadMujer.ElfoDrow(IndexHead) & ".jpg")
                End If
            Case eRaza.Gnomo
                If UserSexo = eGenero.Hombre Then
                    If (IndexHead + 1) > UBound(HeadHombre.Gnomo()) Then Exit Sub
                    IndexHead = IndexHead - 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadHombre.Gnomo(IndexHead) & ".jpg")
                    
                Else
                    If (IndexHead + 1) > UBound(HeadMujer.Gnomo()) Then Exit Sub
                    IndexHead = IndexHead + 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadMujer.Gnomo(IndexHead) & ".jpg")
                End If
            Case eRaza.Enano
                If UserSexo = eGenero.Hombre Then
                    If (IndexHead + 1) > UBound(HeadHombre.Enano()) Then Exit Sub
                    IndexHead = IndexHead + 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadHombre.Enano(IndexHead) & ".jpg")
                    
                Else
                    If (IndexHead + 1) > UBound(HeadMujer.Enano()) Then Exit Sub
                    IndexHead = IndexHead + 1
                    Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & HeadMujer.Enano(IndexHead) & ".jpg")
                End If
        End Select
    End If
End Sub

Private Sub Form_Load()
    IndexHead = 1
End Sub

Private Sub Image1_Click()
    If UserSexo = eGenero.Hombre Then
        Select Case UserRaza
            Case eRaza.Humano
                WriteHead HeadHombre.Humano(IndexHead)
            Case eRaza.Elfo
                WriteHead HeadHombre.Elfo(IndexHead)
            Case eRaza.ElfoOscuro
                WriteHead HeadHombre.ElfoDrow(IndexHead)
            Case eRaza.Gnomo
                WriteHead HeadHombre.Gnomo(IndexHead)
            Case eRaza.Enano
                WriteHead HeadHombre.Enano(IndexHead)
        End Select
    Else
        Select Case UserRaza
            Case eRaza.Humano
                WriteHead HeadMujer.Humano(IndexHead)
            Case eRaza.Elfo
                WriteHead HeadMujer.Elfo(IndexHead)
            Case eRaza.ElfoOscuro
                WriteHead HeadMujer.ElfoDrow(IndexHead)
            Case eRaza.Gnomo
                WriteHead HeadMujer.Gnomo(IndexHead)
            Case eRaza.Enano
                WriteHead HeadMujer.Enano(IndexHead)
        End Select
    End If
    
    Unload Me
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

