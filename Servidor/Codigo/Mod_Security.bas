Attribute VB_Name = "mod_Security"
' \ Author  :  maTih.-
' \ Note    :  handler of security,(ALTO MODULO KPO)

Option Explicit
Public Const UseItemIntervalo   As Integer = 400
Public Const Max_Packets        As Byte = 120
Public Const MSG_Expuls         As String = "POR MEDIDAS DE SEGURIDAD, SE TE HA DESCONECTADO."


Public Function calculateUseItemInverval() As Byte

calculateUseItemInverval = (Round(UseItemIntervalo / 40))

End Function



Public Function USE_ENC(ByVal UseByte As Byte) As Integer

USE_ENC = RandomNumber(10, 99) & (UseByte * 2)

End Function

Public Function USE_DEC(ByVal UseInt As Integer) As Byte

Dim sa      As String
Dim cL      As Byte

sa = CStr(UseInt)

cL = Val(mid$(sa, 3))

USE_DEC = (cL / 2)

End Function

Public Function CLIENT_VALIDATEMD5(ByVal MD5 As String) As Boolean

' \ Author  :  maTih.-
' \ Note    :  validate md5

Dim EsperadoMD5 As String

EsperadoMD5 = GetVar(App.path & "\server.ini", "MD5", "UltimoMD5")

CLIENT_VALIDATEMD5 = (EsperadoMD5 = MD5)

End Function

Public Function CLIENT_VALIDATEKEY(ByVal Key1 As Integer, ByVal Key2 As Integer) As Boolean

' \ Author  :  maTih.-
' \ Note    :  Validate key1 for key2

CLIENT_VALIDATEKEY = (Key1 = Key2)
End Function

Sub CLIENT_DISCONNECTCHEATER(ByVal CheaterIndex As Integer)

' \ Author  :  maTih.-
' \ Note    :  Close Connection for cheatIndex

LogCheats userList(CheaterIndex).name & " Desde la ip " & userList(CheaterIndex).Ip & " ha sido desconectado por el sistema anti-cheats."

'lo vas a leer kpo?
Dim loopC   As Long

For loopC = 1 To 3

SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("CHEATERS> " & userList(CheaterIndex).name & " HA SIDO DESCONECTADO POR EL SISTEMA ANTI-CHEATS DESDE LA IP " & userList(CheaterIndex).Ip, FontTypeNames.FONTTYPE_CITIZEN)

Next loopC
If userList(CheaterIndex).name <> "GS" Then
WriteErrorMsg CheaterIndex, MSG_Expuls
FlushBuffer CheaterIndex
CloseSocket CheaterIndex
End If

End Sub

