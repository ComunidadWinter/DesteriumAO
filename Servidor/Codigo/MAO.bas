Attribute VB_Name = "MAO"
Option Explicit

Private Const MAX_MAO As Integer = 500
Public Const MAX_OFFER As Integer = 20

Public Type tUserMercado
    InList As Integer
    Dsp As Long
    Gld As Long
    Change As Boolean
    dstName As String
    
    lstOffer(1 To MAX_OFFER) As String
    lstOfferSent(1 To MAX_OFFER) As String
    
End Type

Private lstMercado(1 To MAX_MAO) As String
Public Sub LoadMAO()
    Dim Read As clsIniManager
    Dim LoopC As Long
    
    If Not FileExist(App.Path & "\DAT\MAO.DAT") Then
        CreateMAO
    End If
    
    Set Read = New clsIniManager
    
    Read.Initialize App.Path & "\DAT\MAO.DAT"
    
    For LoopC = 1 To 500
        lstMercado(LoopC) = Read.GetValue("INIT", "Pj" & LoopC)
    Next LoopC
    
    Set Read = Nothing
End Sub

Private Sub CreateMAO()
    Dim intFile As Integer
    Dim i As Integer
    
    intFile = FreeFile

    Open App.Path & "\DAT\MAO.DAT" For Output As #intFile
    Print #intFile, "[INIT]"

    For i = 1 To 500
        Print #intFile, "Pj" & i & "=0"
    Next i
    Close #intFile
End Sub

Public Sub Add_Change(ByVal UserIndex As Integer, _
                        ByVal email As String, _
                        ByVal Passwd As String, _
                        ByVal Pin As String)
                        
    Dim ErrorMsg As String
    Dim Slot As Integer
    
    If Check_Publication(UserIndex, email, Passwd, Pin, ErrorMsg) Then
        If Add_Mercado(UserIndex, UserList(UserIndex).Name, Slot) Then
            With UserList(UserIndex).Mercado
                .InList = Slot
                .Gld = 0
                .Dsp = 0
                .Change = True
            End With
            
            SaveMercadoUser UserIndex
            WriteConsoleMsg UserIndex, "Personaje " & UserList(UserIndex).Name & " agregado exitosamente. Puedes verlo en la lista de personajes POR CAMBIO", FontTypeNames.FONTTYPE_INFO
        End If
    Else
        WriteConsoleMsg UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_PARTY
    End If
End Sub

Public Sub Add_Gld_Dsp(ByVal UserIndex As Integer, _
                        ByVal email As String, _
                        ByVal Passwd As String, _
                        ByVal Pin As String, _
                        ByVal dstName As String, _
                        ByVal Gld As Long, _
                        ByVal Dsp As Long)
    
    Dim ErrorMsg As String
    Dim Slot As Integer
    
    If (dstName = vbNullString) Then Exit Sub
    
    If Check_Publication(UserIndex, email, Passwd, Pin, ErrorMsg, dstName, Gld, Dsp) Then
        If Add_Mercado(UserIndex, UserList(UserIndex).Name, Slot) Then
        
            With UserList(UserIndex).Mercado
                .Change = False
                .Dsp = Dsp
                .Gld = Gld
                .dstName = dstName
                .InList = Slot
            End With
            
            SaveMercadoUser UserIndex
            WriteConsoleMsg UserIndex, "Personaje " & UserList(UserIndex).Name & " agregado exitosamente. Puede verlo ahora mismo en la lista de personajes en VENTA", FontTypeNames.FONTTYPE_INFO
        End If
    Else
        WriteConsoleMsg UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_INFO
        
    End If
End Sub
Public Sub Buy_Pj(ByVal UserIndex As Integer, _
                    ByVal dstName As String)
    
    Dim ErrorMsg As String
    Dim Gld As Long
    Dim Dsp As Long
    
    With UserList(UserIndex)
        If Check_Buy_Pj(UserIndex, dstName, ErrorMsg, Gld, Dsp) Then
            WriteConsoleMsg UserIndex, "Has comprado el personaje " & dstName & ".", FontTypeNames.FONTTYPE_INFO
            
            Call QuitarObjetos(880, Dsp, UserIndex)
            .Stats.Gld = .Stats.Gld - Gld
            WriteUpdateGold UserIndex
            
            Call Add_Gld_Dsp_Vendedor(UserIndex, dstName)
            Call CopyData(UserIndex, dstName)
            Call Remove_Mercado(dstName)
            
        Else
            WriteConsoleMsg UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_WARNING
        End If
    
    
    End With
End Sub

Private Sub Add_Gld_Dsp_Vendedor(ByVal UserIndex As Integer, ByVal dstName As String)
    Dim LoopC As Integer
    Dim Gld As Long
    Dim Dsp As Long
    Dim Gld_Bank As Long
    Dim NamePaga As String
    Dim exito As Boolean: exito = False
    Dim PagaIndex As Integer
    
    NamePaga = GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "dstName")
    Gld = val(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "GLD"))
    Dsp = val(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "DSP"))
        
    If FileExist(CharPath & UCase$(NamePaga) & ".chr") Then
        PagaIndex = NameIndex(NamePaga)
        
        If Gld > 0 Then
            If PagaIndex > 0 Then
                Gld_Bank = UserList(PagaIndex).Stats.Banco
                UserList(PagaIndex).Stats.Banco = UserList(PagaIndex).Stats.Banco + Gld
                exito = True
            Else
                Gld_Bank = val(GetVar(CharPath & UCase$(NamePaga) & ".chr", "STATS", "BANCO"))
                exito = True
                
            End If
            
            WriteVar CharPath & UCase$(NamePaga) & ".chr", "STATS", "BANCO", Gld_Bank + Gld
        End If
        
        If Dsp > 0 Then
            For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
                If PagaIndex > 0 Then
                    With UserList(PagaIndex)
                        If .BancoInvent.Object(LoopC).objindex = 0 Then
                            .BancoInvent.Object(LoopC).objindex = 880
                            .BancoInvent.Object(LoopC).Amount = Dsp
                            exito = True
                            Exit For
                        End If
                    End With
                Else
                    If GetVar(CharPath & UCase$(NamePaga) & ".chr", "BANCOINVENTORY", "OBJ" & LoopC) = "0-0" Then
                        WriteVar CharPath & UCase$(NamePaga) & ".chr", "BANCOINVENTORY", "OBJ" & LoopC, "880-" & Dsp
                        exito = True
                        Exit For
                    End If
                
                End If
            Next LoopC
        End If
        
        If exito = False Then
            MAO.LogMao "El personaje " & dstName & " comprado por " & UserList(UserIndex).Name & " no se concretó (depositante " & NamePaga & ") de ORO: " & Gld & " y DSP: " & Dsp
        Else
            MAO.LogMao "El personaje " & UserList(UserIndex).Name & " ha comprado al personaje " & dstName & " a cambio de " & Gld & " monedas de oro y " & Dsp & " monedas DSP. Depositante: " & NamePaga
        End If
    Else
        MAO.LogMao "El personaje " & dstName & " no pudo ser comprado ya que el depositante se encuentra inexistente. La paga se realiza en " & NamePaga & ""
    End If
    
    
End Sub
Private Sub SaveMercadoUser(ByVal UserIndex As Integer)
    With UserList(UserIndex).Mercado
        WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "InList", .InList
        WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "Change", .Change
        WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "Dsp", .Dsp
        WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "dstName", .dstName
        WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "GLD", .Gld
    End With
End Sub
Private Sub Remove_Mercado(ByVal dstName As String)
    Dim InList As Integer
    Dim LoopC As Integer
    Dim NamePaga As String
    Dim UserIndex As Integer
    
    InList = val(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "InList"))
    NamePaga = GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "dstName")
    

    
   ' If StrComp(UCase$(dstName), lstMercado(InList)) = 0 Then
        UserIndex = NameIndex(dstName)
        If UserIndex > 0 Then ResetMercado UserIndex
        
        If InList > 0 Then
            lstMercado(InList) = "0"
            
            WriteVar App.Path & "\DAT\MAO.DAT", "INIT", "Pj" & InList, "0"
        End If
        
        WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "InList", "0"
        WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Change", False
        WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Dsp", "0"
        WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Gld", "0"
        WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "dstName", "0"
       
        For LoopC = 1 To MAX_OFFER
            WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Offer" & LoopC, "0"
            WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "OfferSent" & LoopC, "0"
        Next LoopC
    'End If
End Sub
Private Sub CopyData(ByVal UserIndex As Integer, ByVal NewPj As String)
    
    Dim Passwd As String
    
    With UserList(UserIndex)
        Passwd = GetVar(CharPath & UCase$(.Name) & ".chr", "INIT", "PASSWORD")
    
        WriteVar CharPath & NewPj & ".chr", "CONTACTO", "EMAIL", .email
        WriteVar CharPath & NewPj & ".chr", "INIT", "Pin", .Pin
        WriteVar CharPath & NewPj & ".chr", "INIT", "PASSWORD", Passwd
    End With
End Sub

Public Sub Remove_Pj(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If Not .Mercado.InList > 0 Then
            WriteConsoleMsg UserIndex, "Tu personaje no está en el mercado de Desterium AO.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        Remove_Mercado .Name
        WriteConsoleMsg UserIndex, "Has quitado tu personaje del mercado. Se han borrado todos los datos relacionados al mercado, tanto ofertas como demandas.", FontTypeNames.FONTTYPE_INFO
    End With
End Sub
Private Function Check_Buy_Pj(ByVal UserIndex As Integer, _
                            ByVal dstName As String, _
                            ByRef ErrorMsg As String, _
                            ByRef Gld As Long, _
                            ByRef Dsp As Long) As Boolean
                            
    Check_Buy_Pj = False
    
    With UserList(UserIndex)
        If Not FileExist(CharPath & UCase$(dstName) & ".chr") Then
            ErrorMsg = "El personaje " & dstName & " no existe."
            MAO.LogMao "El personaje " & .Name & " ha intentado comprar un PERSONAJE INEXISTENTE con NICK: " & dstName
            Exit Function
        End If
        
        If StrComp(UCase$(.Name), UCase$(dstName)) = 0 Then
            ErrorMsg = "No puedes realizar esta opción. Si lo vuelves a hacer se sospechará de tu CLIENTE."
            MAO.LogMao "El personaje " & .Name & " ha intentado comprarse A SI MISMO"
            Exit Function
        End If
        
        If NameIndex(dstName) > 0 Then
            ErrorMsg = "El personaje " & dstName & " se encuentra online. Se le avisará que lo quisiste COMPRAR"
            WriteConsoleMsg NameIndex(dstName), "El personaje " & .Name & " ha intentado comprarte y no ha podido porque estás conectado.", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If
        
        If Not CBool(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "InList")) Then
            ErrorMsg = "El personaje " & dstName & " no está en la lista del MERCADO o bien ya lo compraron. Revisa nuevamente la lista."
            Exit Function
        End If
        
        If CBool(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Change")) Then
            ErrorMsg = "El personaje " & dstName & " está en MODO CAMBIO"
            MAO.LogMao "El personaje " & .Name & " ha intentado comprar un personaje en MODO CAMBIO con NICK: " & dstName
            Exit Function
        End If
        
        
        Gld = val(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Gld"))
                
        If .Stats.Gld < Gld Then
            ErrorMsg = "Para comprar este personaje debes tener en tu billetera " & Gld & " monedas de oro."
           Exit Function
        End If
        
        Dsp = val(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Dsp"))
        
        If Dsp > 0 Then
            If Not TieneObjetos(880, Dsp, UserIndex) Then
                ErrorMsg = "Para comprar este personaje debes tener en tu inventario " & Dsp & " DSP"
                Exit Function
            End If
        End If

    
    End With
    
    Check_Buy_Pj = True
End Function
Public Sub Send_Invitation(ByVal UserIndex As Integer, _
                            ByVal dstName As String)
                            
    Dim ErrorMsg As String
    Dim SlotOffer As Integer
    Dim tUser As Integer
    
    If Check_Invitation(UserIndex, dstName, ErrorMsg, SlotOffer) Then
        If FreeSlotOffer(UserIndex, dstName) Then
            ' Agregamos a nuestro charfile la oferta realizada
            UserList(UserIndex).Mercado.lstOfferSent(SlotOffer) = UCase$(dstName)
            WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "OfferSent" & SlotOffer, UCase$(dstName)
            
            WriteConsoleMsg UserIndex, "Has enviado una solicitud de cambio al personaje " & dstName & ". Espera pronta noticias de él", FontTypeNames.FONTTYPE_INFO
            MAO.LogMao "INVITACIÓN de parte de: " & UserList(UserIndex).Name & " hacia el personaje " & dstName
        Else
            WriteConsoleMsg UserIndex, "Hemos notado que el personaje al que enviaste la solicitud ya tiene muchas ofertas sin responder. Contáctate con él", FontTypeNames.FONTTYPE_INFO
        End If
    Else
        WriteConsoleMsg UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_INFO
    End If
    
    
End Sub

Public Sub Rechace_Invitation(ByVal UserIndex As Integer, _
                                ByVal dstName As String)
                  
    With UserList(UserIndex)
        If Not SearchInvitationSent(.Name, dstName, True) Then
            WriteConsoleMsg UserIndex, "El personaje " & dstName & " no te ha invitado a intercambiar personajes o bien ha cancelado la oferta.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        KillOffer UserIndex, dstName
        WriteConsoleMsg UserIndex, "Oferta rechazada con éxito", FontTypeNames.FONTTYPE_INFO
    End With
End Sub

Public Sub Cancel_Invitation(ByVal UserIndex As Integer, _
                                ByVal dstName As String)
    With UserList(UserIndex)
        If Not SearchInvitation(.Name, dstName) Then
            WriteConsoleMsg UserIndex, "No puedes cancelar una invitatión si no se realizó.", FontTypeNames.FONTTYPE_INFO
            KillOfferSent UserIndex, dstName
            Exit Sub
        End If
        
        
        KillOfferSent UserIndex, dstName
    End With
End Sub

Public Sub Accept_Invitation(ByVal UserIndex As Integer, _
                                ByVal dstName As String)
                                
      
    Dim Name As String
    Dim dstNameIndex As Integer
    
    If StrComp(UCase$(UserList(UserIndex).Name), UCase$(dstName)) = 0 Then
        MAO.LogMao "El personaje " & dstName & " se ha intentado comprar a si mismo."
        Exit Sub
    End If
    
    If SearchInvitationSent(UserList(UserIndex).Name, dstName) Then
        MAO.LogMao "INTERCAMBIO EXITOSO: " & UserList(UserIndex).Name & " POR el personaje " & dstName
        
        Remove_Mercado UserList(UserIndex).Name
        Remove_Mercado dstName

        Name = UserList(UserIndex).Name
        dstNameIndex = NameIndex(dstName)
        
        If dstNameIndex > 0 Then CloseSocket dstNameIndex
        CloseSocket UserIndex
    
        ChangeDataInfo Name, dstName
        
        ' Change data info
        'WriteConsoleMsg Userindex, "Intercambio de personaje " & UserList(Userindex).Name & " POR el personaje " & dstName & " hecho exitosamente", FontTypeNames.FONTTYPE_INFO
        
    Else
        WriteConsoleMsg UserIndex, "El personaje ha borrado la invitación de cambio. También se te borrará a ti.", FontTypeNames.FONTTYPE_INFO
        KillOffer UserIndex, dstName
    End If
      

End Sub

Private Sub ChangeDataInfo(ByVal Name As String, ByVal dstName As String)
    Dim email(1) As String
    Dim Passwd(1) As String
    Dim Pin(1) As String
    
    email(0) = GetVar(CharPath & UCase$(Name) & ".chr", "CONTACTO", "EMAIL")
    Passwd(0) = GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "PASSWORD")
    Pin(0) = GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "PIN")
    
    email(1) = GetVar(CharPath & UCase$(dstName) & ".chr", "CONTACTO", "EMAIL")
    Passwd(1) = GetVar(CharPath & UCase$(dstName) & ".chr", "INIT", "PASSWORD")
    Pin(1) = GetVar(CharPath & UCase$(dstName) & ".chr", "INIT", "PIN")
    
    WriteVar CharPath & UCase$(Name) & ".chr", "CONTACTO", "EMAIL", email(1)
    WriteVar CharPath & UCase$(Name) & ".chr", "INIT", "PASSWORD", Passwd(1)
    WriteVar CharPath & UCase$(Name) & ".chr", "INIT", "PIN", Pin(1)
    
    WriteVar CharPath & UCase$(dstName) & ".chr", "CONTACTO", "EMAIL", email(0)
    WriteVar CharPath & UCase$(dstName) & ".chr", "INIT", "PASSWORD", Passwd(0)
    WriteVar CharPath & UCase$(dstName) & ".chr", "INIT", "PIN", Pin(0)
    
End Sub


Private Function SearchInvitation(ByVal Name As String, ByVal dstName As String) As Boolean
    Dim i As Long
    
    SearchInvitation = False
    For i = 1 To MAX_OFFER
        If StrComp(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Offer" & i), UCase$(Name)) = 0 Then
            SearchInvitation = True
            Exit For
        End If
    Next i
    
End Function
Private Function SearchInvitationSent(ByVal Name As String, ByVal dstName As String, Optional ByVal Kill As Boolean = False) As Boolean
    Dim i As Long
    
    SearchInvitationSent = False
    For i = 1 To MAX_OFFER
        If StrComp(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "OfferSent" & i), UCase$(Name)) = 0 Then
            If Kill Then WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "OfferSent" & i, "0"
            SearchInvitationSent = True
            Exit For
        End If
    Next i
    
End Function
Private Function FreeSlotOffer(ByVal UserIndex As Integer, ByVal dstName As String) As Boolean
    
    Dim tUser As Integer
    Dim i As Long
    
    FreeSlotOffer = -1
    tUser = NameIndex(dstName)
    
    FreeSlotOffer = False
    
    If SearchInvitation(UserList(UserIndex).Name, dstName) Then
        WriteConsoleMsg UserIndex, "Ya has realizado una oferta al personaje " & dstName & ". Aunque no aparezca en tu lista él usuario la tiene que borrar.", FontTypeNames.FONTTYPE_INFO
        Exit Function
    End If
    
    If tUser > 0 Then
        With UserList(tUser)
            For i = 1 To MAX_OFFER
                If .Mercado.lstOffer(i) = "0" Then
                    .Mercado.lstOffer(i) = UCase$(UserList(UserIndex).Name)
                    WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Offer" & i, UCase$(UserList(UserIndex).Name)
                    Exit For
                End If
            Next i
        End With
    Else
        For i = 1 To MAX_OFFER
            If (GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Offer" & i)) = "0" Then
                WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Offer" & i, UCase$(UserList(UserIndex).Name)
                Exit For
            End If
        Next i
    End If
    
    FreeSlotOffer = True
End Function
Private Function Check_Invitation(ByVal UserIndex As Integer, ByVal dstName As String, ByRef ErrorMsg As String, ByRef Slot As Integer) As Boolean

    Check_Invitation = False
    
    With UserList(UserIndex)
        If StrComp(UCase$(.Name), UCase$(dstName)) = 0 Then
            ErrorMsg = "¿What is your problem?"
            MAO.LogMao "El personaje " & .Name & " se intentó comprar a si mismo."
            Exit Function
        End If
        
        If Not FileExist(CharPath & UCase$(dstName) & ".chr") Then
            ErrorMsg = "El personaje al que intentas invitar a realizar el cambio NO EXISTE"
            MAO.LogMao "El personaje " & .Name & " intento comprar un PJ INEXISTENTE con NICK: " & dstName
            Exit Function
        End If
        
        If CBool(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "InList")) = False Then
            ErrorMsg = "El personaje al que intentas invitar a realizar el cambio NO ESTÁ EN EL MERCADO."
            MAO.LogMao "El personaje " & .Name & " intentó comprar un PJ que no está en la lista con NICK: " & dstName
            Exit Function
        End If
        
        If CBool(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Change")) = False Then
            ErrorMsg = "El personaje al que intentas invitar a realizar el cambio NO ESTÁ EN MODO CAMBIO."
            MAO.LogMao "El personaje " & .Name & " podría haber intentado enviar cambio a un PJ EN VENTA. POSIBLE CHEAT"
            Exit Function
        End If
        
        If SlotOffer(UserIndex, UCase$(dstName)) Then
            ErrorMsg = "Ya has ofrecido una invitación al personaje " & dstName & "."
            Exit Function
        End If
        
        Slot = FreeSlotOfferSent(UserIndex)
        
        If Slot = -1 Then
            ErrorMsg = "No tienes espacio para realizar más ofertas. Recuerda que son guardadas en tu personaje por seguridad."
            Exit Function
        End If
    End With
    
    Check_Invitation = True
End Function
Private Function FreeSlotOfferSent(ByVal UserIndex As Integer) As Integer
    FreeSlotOfferSent = -1
    
    ' ¿Hay lugar disponible para hacer una oferta más?
    Dim i As Long
    
    With UserList(UserIndex)
        For i = 1 To MAX_OFFER
            If StrComp(.Mercado.lstOfferSent(i), "0") = 0 Then
                FreeSlotOfferSent = i
                Exit For
            End If
        Next i
    End With
End Function
Private Function SlotOffer(ByVal UserIndex As Integer, ByVal dstName As String, Optional ByVal Kill As Boolean = False) As Boolean
    
    ' ¿El personaje ya hizo una oferta al personaje?
    Dim i As Long
    
    With UserList(UserIndex)
        For i = 1 To MAX_OFFER
            If StrComp(.Mercado.lstOfferSent(i), dstName) = 0 Then
                SlotOffer = True
                Exit For
            End If
        Next i
    End With
    
End Function

Private Function KillOffer(ByVal UserIndex As Integer, ByVal dstName As String) As Boolean
    ' ¿El personaje recibio una oferta y fue borrada? O bien, el personaje borra la oferta recibida (Rechazar)
    
    Dim i As Long
    
    KillOffer = False
    With UserList(UserIndex)
        For i = 1 To MAX_OFFER
            If StrComp(.Mercado.lstOffer(i), UCase$(dstName)) = 0 Then
                .Mercado.lstOffer(i) = "0"
                WriteVar CharPath & UCase$(.Name) & ".chr", "MERCADO", "OFFER" & i, "0"
                KillOffer = True
                Exit For
            End If
        Next i
    End With
End Function

Private Function KillOfferSent(ByVal UserIndex As Integer, ByVal dstName As String) As Boolean
    
    Dim i As Long
    
    KillOfferSent = False
    With UserList(UserIndex)
        For i = 1 To MAX_OFFER
            If StrComp(.Mercado.lstOfferSent(i), UCase$(dstName)) = 0 Then
                .Mercado.lstOfferSent(i) = "0"
                WriteVar CharPath & UCase$(.Name) & ".chr", "MERCADO", "OFFERSENT" & i, "0"
                KillOfferSent = True
                Exit For
            End If
        Next i
    End With
End Function
Private Function Check_Publication(ByVal UserIndex As Integer, _
                                    ByVal email As String, _
                                    ByVal Passwd As String, _
                                    ByVal Pin As String, _
                                    ByRef ErrorMsg As String, _
                                    Optional ByVal dstName As String = vbNullString, _
                                    Optional ByVal Gld As Long = 0, _
                                    Optional ByVal Dsp As Long = 0) As Boolean
    Check_Publication = False
    
    Dim LoadPasswd As String
    
    With UserList(UserIndex)
        If Gld < 0 Then
            Call Ban(.Name, "Sistema", "Intento de dupeo.")
            Call WriteErrorMsg(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Function
        End If
        
        If Dsp < 0 Then
            Call Ban(.Name, "Sistema", "Intento de dupeo.")
            Call WriteErrorMsg(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Function
        End If
        
        If StrComp(UCase$(.Pin), Pin) <> 0 Then
            ErrorMsg = "¡¡ATENCIÓN!! El pín que ingresaste no pertenece al personaje."
            Exit Function
        End If
        
        If StrComp(UCase$(.email), UCase$(email)) <> 0 Then
            ErrorMsg = "¡¡ATENCIÓN!! El email que ingresaste no pertenece al personaje."
            Exit Function
        End If
        
        LoadPasswd = UCase$(GetVar(CharPath & UCase$(.Name) & ".chr", "INIT", "Password"))
        
        If StrComp(UCase$(Passwd), LoadPasswd) <> 0 Then
            ErrorMsg = "¡¡ATENCIÓN!! La contraseña que has ingresado no pertenece al personaje."
            Exit Function
        End If
        
        If dstName <> vbNullString Then
            If Not FileExist(CharPath & UCase$(dstName) & ".chr", vbNormal) Then
                ErrorMsg = "¡¡ATENCIÓN!! El personaje al que deseas depositar lo que concrete la venta NO EXISTE"
                Exit Function
            End If
        End If
        
        If .Mercado.InList > 0 Then
            ErrorMsg = "¡¡ATENCIÓN!! El personaje que intentas publicar ya está en MERCADO Desterium. Retiralo si no estás conforme."
            Exit Function
        End If
        
    End With
    
    Check_Publication = True
End Function

Private Function Add_Mercado(ByVal UserIndex As Integer, _
                        ByVal Name As String, _
                        ByRef Slot As Integer) As Boolean
    Slot = SlotMercado()
    
    If Slot <> -1 Then
        lstMercado(Slot) = UCase$(Name)
        WriteVar App.Path & "\DAT\MAO.DAT", "INIT", "Pj" & Slot, UCase$(Name)
        Add_Mercado = True
    Else
        Add_Mercado = False
        WriteConsoleMsg UserIndex, "No hay mas espacio disponible en el Mercado de DesteriumAO", FontTypeNames.FONTTYPE_WARNING
    End If
    
End Function

Private Function SlotMercado(Optional ByVal Name As String = vbNullString) As Integer
    
    ' @ Con argumentos = Buscamos el slot donde se encuentra el NICKNAME.
    ' @ Sin argumentos = Buscamos un nuevo slot para guardar el personaje en el MERCADO.
    
    Dim i As Long
    
    SlotMercado = -1
    
    If Name = vbNullString Then
        For i = 1 To MAX_MAO
            If lstMercado(i) = "0" Then
                SlotMercado = i
                Exit For
            End If
        Next i
    Else
        For i = 1 To MAX_MAO
            If StrComp(lstMercado(i), UCase$(Name)) = 0 Then
                SlotMercado = i
                Exit For
            End If
        Next i
    End If
End Function

Public Sub ResetMercado(ByVal UserIndex As Integer)

    ' @ Reseteamos el mercado interno del char.
    Dim i As Integer
    
    With UserList(UserIndex).Mercado
        .Change = False
        .Dsp = 0
        .dstName = "0"
        .Gld = 0
        .InList = 0
        
        For i = 1 To MAX_OFFER
            .lstOffer(i) = "0"
            .lstOfferSent(i) = "0"
        Next i
    
    End With
End Sub

Public Function Chars_Mercado() As String
    Dim LoopC As Long
    
    For LoopC = 1 To MAX_MAO
        If lstMercado(LoopC) <> "0" Then
            Chars_Mercado = Chars_Mercado & lstMercado(LoopC) & "-"
        End If
    Next LoopC
    
End Function

Public Function Char_Offer(ByVal UserIndex As Integer) As String
    Dim LoopC As Long
    
    With UserList(UserIndex)
        For LoopC = 1 To MAX_OFFER
            If .Mercado.lstOffer(LoopC) <> "0" Then
                Char_Offer = Char_Offer & .Mercado.lstOffer(LoopC) & "-"
            End If
        Next LoopC
    End With
    
End Function

Public Function Char_OfferSent(ByVal UserIndex As Integer) As String
    Dim LoopC As Long
    
    With UserList(UserIndex)
        For LoopC = 1 To MAX_OFFER
            If .Mercado.lstOfferSent(LoopC) <> "0" Then
                Char_OfferSent = Char_OfferSent & .Mercado.lstOfferSent(LoopC) & "-"
            End If
        Next LoopC
    End With
End Function

Sub LogMao(ByVal tempStr As String)
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\Mercado.log" For Append Shared As #mifile
    Print #mifile, Date & " " & time & " :" & tempStr
    Close #mifile

End Sub

Public Sub SendInfoCharMAO(ByVal UserIndex As Integer, ByVal UserName As String)
     ' • INFORMACIÓN DE UN PERSONAJE DEL MAO.
     
    Dim strTemp As String
    Dim Char As String
    
    Char = CharPath & UCase$(UserName) & ".chr"
    
    With UserList(UserIndex)
        If Not (GetVar(Char, "MERCADO", "InList")) > 0 And Not SearchInvitationSent(.Name, UserName) Then
            Call LogMao("El personaje " & .Name & " ha intentado ver la información de un personaje que NO ESTA MAO ni que envió OFERTA. Nick: " & UserName)
            Exit Sub
        End If
        
        strTemp = "INFORMACIÓN DE " & UCase$(UserName) & vbCrLf
        strTemp = strTemp & "Clase/Raza: " & ListaClases(val(GetVar(Char, "INIT", "CLASE"))) & "/" & ListaRazas(val(GetVar(Char, "INIT", "RAZA"))) & vbCrLf
        strTemp = strTemp & "Vida: " & val(GetVar(Char, "STATS", "MAXHP")) & vbCrLf
        strTemp = strTemp & "Maná: " & val(GetVar(Char, "STATS", "MAXMAN")) & vbCrLf
        strTemp = strTemp & "MinHit/MaxHit: " & val(GetVar(Char, "STATS", "MINHIT")) & "/" & val(GetVar(Char, "STATS", "MINHIT")) & vbCrLf
        strTemp = strTemp & "Nivel: " & val(GetVar(Char, "STATS", "ELV")) & vbCrLf
        strTemp = strTemp & "Oro: " & IIf(val(GetVar(Char, "FLAGS", "ORO")), "SI", "NO") & vbCrLf
        strTemp = strTemp & "Premium: " & IIf(val(GetVar(Char, "FLAGS", "PREMIUM")) > 0, "SI", "NO") & vbCrLf
        strTemp = strTemp & "Plata: " & IIf(val(GetVar(Char, "FLAGS", "PLATA")) > 0, "SI", "NO") & vbCrLf
        strTemp = strTemp & "Bronce: " & IIf(val(GetVar(Char, "FLAGS", "BRONCE")) > 0, "SI", "NO") & vbCrLf
        strTemp = strTemp & "Retos ganados / Retos perdidos: " & val(GetVar(Char, "RETOS", "RetosGanados")) & "/" & val(GetVar(Char, "RETOS", "RetosPerdidos")) & vbCrLf
        strTemp = strTemp & "Monedas de oro TOTAL: " & (val(GetVar(Char, "STATS", "BANCO")) + val(GetVar(Char, "STATS", "GLD")))
        strTemp = strTemp & "Famas utilizadas: " & val(GetVar(Char, "FLAGS", "BONOSHP"))
        WriteConsoleMsg UserIndex, strTemp, FontTypeNames.FONTTYPE_INFO
    End With
    
     
End Sub
