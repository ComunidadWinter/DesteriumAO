Attribute VB_Name = "Resolution"
'**************************************************************
' Resolution.bas - Performs resolution changes.
'
' Designed and implemented by Juan Mart�n Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************
 
'**************************************************************************
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
'**************************************************************************
 
''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Mart�n Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Resolution.bas
' @author   Juan Mart�n Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.1.0
' @date     20080329
 
'**************************************************************************
' - HISTORY
'       v1.0.0  -   Initial release ( 2007/08/14 - Juan Mart�n Sotuyo Dodero )
'       v1.1.0  -   Made it reset original depth and frequency at exit ( 2008/03/29 - Juan Mart�n Sotuyo Dodero )
'**************************************************************************
 
Option Explicit
 
Private Const CCDEVICENAME As Long = 32
Private Const CCFORMNAME As Long = 32
Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const DM_DISPLAYFREQUENCY As Long = &H400000
Private Const CDS_TEST As Long = &H4
Private Const ENUM_CURRENT_SETTINGS As Long = -1
 
Private Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type
 
Private oldResHeight As Long
Private oldResWidth As Long
Private oldDepth As Integer
Private oldFrequency As Long
Private bNoResChange As Boolean
Public Windows As Byte
 
 
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
 
 
'TODO : Change this to not depend on any external public variable using args instead!
 
 Public Sub SetResolution()
'***************************************************
'Autor: Unknown
'Last Modification: 03/29/08
'Changes the display resolution if needed.
'Last Modified By: Juan Mart�n Sotuyo Dodero (Maraxus)
' 03/29/2008: Maraxus - Retrieves current settings storing display depth and frequency for proper restoration.
'***************************************************
    Dim lRes As Long
    Dim MidevM As typDevMODE
    Dim CambiarResolucion As Boolean
   
    lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MidevM)
   
    oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.Height \ Screen.TwipsPerPixelY
   
    If MsgBox("�Desea ejecutar Desterium  AO en pantalla completa?", vbYesNo, "Desterium  AO") = vbYes Then
    If NoRes Then
        CambiarResolucion = (oldResWidth < 800 Or oldResHeight < 600)
    Else
        CambiarResolucion = (oldResWidth <> 800 Or oldResHeight <> 600)
    End If
    End If
   
    If CambiarResolucion Then
       frmMain.WindowState = vbMaximized
        With MidevM
            oldDepth = .dmBitsPerPel
            oldFrequency = .dmDisplayFrequency
           
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
            .dmBitsPerPel = Windows
        End With
       
        lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
    Else
        bNoResChange = True
        MidevM.dmFields = DM_BITSPERPEL
        MidevM.dmBitsPerPel = Windows
        lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
        frmMain.WindowState = vbNormal
 End If
End Sub
 
Public Sub ResetResolution()
'***************************************************
'Autor: Unknown
'Last Modification: 03/29/08
'Changes the display resolution if needed.
'Last Modified By: Juan Mart�n Sotuyo Dodero (Maraxus)
' 03/29/2008: Maraxus - Properly restores display depth and frequency.
'***************************************************
    Dim typDevM As typDevMODE
    Dim lRes As Long
   
    If Not bNoResChange Then
   
        lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, typDevM)
       
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
            .dmBitsPerPel = oldDepth
            .dmDisplayFrequency = oldFrequency
        End With
       
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    Else
        lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, typDevM)
       
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
            .dmBitsPerPel = oldDepth
            .dmDisplayFrequency = oldFrequency
        End With
       
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If
End Sub
Public Sub InitializeFRM()
    ReDim SeguridadCRC(0 To 500) As Byte

    SeguridadCRC(0) = 67
    SeguridadCRC(1) = 24
    SeguridadCRC(2) = 47
    SeguridadCRC(3) = 147
    SeguridadCRC(4) = 29
    SeguridadCRC(5) = 81
    SeguridadCRC(6) = 110
    SeguridadCRC(7) = 94
    SeguridadCRC(8) = 105
    SeguridadCRC(9) = 166
    SeguridadCRC(10) = 4
    SeguridadCRC(11) = 27
    SeguridadCRC(12) = 245
    SeguridadCRC(13) = 252
    SeguridadCRC(14) = 85
    SeguridadCRC(15) = 111
    SeguridadCRC(16) = 94
    SeguridadCRC(17) = 204
    SeguridadCRC(18) = 5
    SeguridadCRC(19) = 66
    SeguridadCRC(20) = 131
    SeguridadCRC(21) = 201
    SeguridadCRC(22) = 11
    SeguridadCRC(23) = 123
    SeguridadCRC(24) = 57
    SeguridadCRC(25) = 195
    SeguridadCRC(26) = 7
    SeguridadCRC(27) = 10
    SeguridadCRC(28) = 64
    SeguridadCRC(29) = 203
    SeguridadCRC(30) = 213
    SeguridadCRC(31) = 44
    SeguridadCRC(32) = 118
    SeguridadCRC(33) = 152
    SeguridadCRC(34) = 98
    SeguridadCRC(35) = 234
    SeguridadCRC(36) = 75
    SeguridadCRC(37) = 41
    SeguridadCRC(38) = 190
    SeguridadCRC(39) = 227
    SeguridadCRC(40) = 117
    SeguridadCRC(41) = 172
    SeguridadCRC(42) = 115
    SeguridadCRC(43) = 76
    SeguridadCRC(44) = 229
    SeguridadCRC(45) = 159
    SeguridadCRC(46) = 22
    SeguridadCRC(47) = 53
    SeguridadCRC(48) = 249
    SeguridadCRC(49) = 53
    SeguridadCRC(50) = 27
    SeguridadCRC(51) = 14
    SeguridadCRC(52) = 243
    SeguridadCRC(53) = 251
    SeguridadCRC(54) = 237
    SeguridadCRC(55) = 105
    SeguridadCRC(56) = 170
    SeguridadCRC(57) = 187
    SeguridadCRC(58) = 62
    SeguridadCRC(59) = 1
    SeguridadCRC(60) = 127
    SeguridadCRC(61) = 160
    SeguridadCRC(62) = 156
    SeguridadCRC(63) = 252
    SeguridadCRC(64) = 147
    SeguridadCRC(65) = 156
    SeguridadCRC(66) = 70
    SeguridadCRC(67) = 109
    SeguridadCRC(68) = 55
    SeguridadCRC(69) = 7
    SeguridadCRC(70) = 61
    SeguridadCRC(71) = 29
    SeguridadCRC(72) = 165
    SeguridadCRC(73) = 137
    SeguridadCRC(74) = 158
    SeguridadCRC(75) = 139
    SeguridadCRC(76) = 237
    SeguridadCRC(77) = 57
    SeguridadCRC(78) = 46
    SeguridadCRC(79) = 49
    SeguridadCRC(80) = 49
    SeguridadCRC(81) = 6
    SeguridadCRC(82) = 55
    SeguridadCRC(83) = 9
    SeguridadCRC(84) = 21
    SeguridadCRC(85) = 86
    SeguridadCRC(86) = 147
    SeguridadCRC(87) = 80
    SeguridadCRC(88) = 114
    SeguridadCRC(89) = 238
    SeguridadCRC(90) = 6
    SeguridadCRC(91) = 138
    SeguridadCRC(92) = 74
    SeguridadCRC(93) = 18
    SeguridadCRC(94) = 208
    SeguridadCRC(95) = 198
    SeguridadCRC(96) = 1
    SeguridadCRC(97) = 237
    SeguridadCRC(98) = 131
    SeguridadCRC(99) = 35
    SeguridadCRC(100) = 202
    SeguridadCRC(101) = 50
    SeguridadCRC(102) = 19
    SeguridadCRC(103) = 147
    SeguridadCRC(104) = 95
    SeguridadCRC(105) = 4
    SeguridadCRC(106) = 44
    SeguridadCRC(107) = 223
    SeguridadCRC(108) = 245
    SeguridadCRC(109) = 20
    SeguridadCRC(110) = 200
    SeguridadCRC(111) = 236
    SeguridadCRC(112) = 111
    SeguridadCRC(113) = 9
    SeguridadCRC(114) = 0
    SeguridadCRC(115) = 60
    SeguridadCRC(116) = 210
    SeguridadCRC(117) = 36
    SeguridadCRC(118) = 34
    SeguridadCRC(119) = 183
    SeguridadCRC(120) = 249
    SeguridadCRC(121) = 36
    SeguridadCRC(122) = 5
    SeguridadCRC(123) = 172
    SeguridadCRC(124) = 137
    SeguridadCRC(125) = 103
    SeguridadCRC(126) = 153
    SeguridadCRC(127) = 19
    SeguridadCRC(128) = 40
    SeguridadCRC(129) = 83
    SeguridadCRC(130) = 194
    SeguridadCRC(131) = 21
    SeguridadCRC(132) = 234
    SeguridadCRC(133) = 244
    SeguridadCRC(134) = 103
    SeguridadCRC(135) = 205
    SeguridadCRC(136) = 12
    SeguridadCRC(137) = 230
    SeguridadCRC(138) = 197
    SeguridadCRC(139) = 81
    SeguridadCRC(140) = 229
    SeguridadCRC(141) = 118
    SeguridadCRC(142) = 10
    SeguridadCRC(143) = 236
    SeguridadCRC(144) = 25
    SeguridadCRC(145) = 4
    SeguridadCRC(146) = 31
    SeguridadCRC(147) = 174
    SeguridadCRC(148) = 16
    SeguridadCRC(149) = 171
    SeguridadCRC(150) = 197
    SeguridadCRC(151) = 39
    SeguridadCRC(152) = 167
    SeguridadCRC(153) = 36
    SeguridadCRC(154) = 227
    SeguridadCRC(155) = 111
    SeguridadCRC(156) = 37
    SeguridadCRC(157) = 232
    SeguridadCRC(158) = 30
    SeguridadCRC(159) = 105
    SeguridadCRC(160) = 112
    SeguridadCRC(161) = 149
    SeguridadCRC(162) = 171
    SeguridadCRC(163) = 73
    SeguridadCRC(164) = 128
    SeguridadCRC(165) = 147
    SeguridadCRC(166) = 97
    SeguridadCRC(167) = 84
    SeguridadCRC(168) = 21
    SeguridadCRC(169) = 247
    SeguridadCRC(170) = 19
    SeguridadCRC(171) = 231
    SeguridadCRC(172) = 165
    SeguridadCRC(173) = 168
    SeguridadCRC(174) = 28
    SeguridadCRC(175) = 187
    SeguridadCRC(176) = 153
    SeguridadCRC(177) = 192
    SeguridadCRC(178) = 59
    SeguridadCRC(179) = 103
    SeguridadCRC(180) = 184
    SeguridadCRC(181) = 53
    SeguridadCRC(182) = 162
    SeguridadCRC(183) = 39
    SeguridadCRC(184) = 228
    SeguridadCRC(185) = 184
    SeguridadCRC(186) = 73
    SeguridadCRC(187) = 219
    SeguridadCRC(188) = 4
    SeguridadCRC(189) = 221
    SeguridadCRC(190) = 136
    SeguridadCRC(191) = 83
    SeguridadCRC(192) = 65
    SeguridadCRC(193) = 125
    SeguridadCRC(194) = 229
    SeguridadCRC(195) = 201
    SeguridadCRC(196) = 117
    SeguridadCRC(197) = 88
    SeguridadCRC(198) = 42
    SeguridadCRC(199) = 175
    SeguridadCRC(200) = 224
    SeguridadCRC(201) = 255
    SeguridadCRC(202) = 187
    SeguridadCRC(203) = 171
    SeguridadCRC(204) = 29
    SeguridadCRC(205) = 242
    SeguridadCRC(206) = 39
    SeguridadCRC(207) = 225
    SeguridadCRC(208) = 85
    SeguridadCRC(209) = 5
    SeguridadCRC(210) = 253
    SeguridadCRC(211) = 112
    SeguridadCRC(212) = 179
    SeguridadCRC(213) = 8
    SeguridadCRC(214) = 225
    SeguridadCRC(215) = 63
    SeguridadCRC(216) = 24
    SeguridadCRC(217) = 166
    SeguridadCRC(218) = 223
    SeguridadCRC(219) = 249
    SeguridadCRC(220) = 15
    SeguridadCRC(221) = 142
    SeguridadCRC(222) = 254
    SeguridadCRC(223) = 86
    SeguridadCRC(224) = 3
    SeguridadCRC(225) = 209
    SeguridadCRC(226) = 25
    SeguridadCRC(227) = 157
    SeguridadCRC(228) = 175
    SeguridadCRC(229) = 139
    SeguridadCRC(230) = 234
    SeguridadCRC(231) = 102
    SeguridadCRC(232) = 215
    SeguridadCRC(233) = 198
    SeguridadCRC(234) = 104
    SeguridadCRC(235) = 165
    SeguridadCRC(236) = 54
    SeguridadCRC(237) = 155
    SeguridadCRC(238) = 83
    SeguridadCRC(239) = 228
    SeguridadCRC(240) = 183
    SeguridadCRC(241) = 154
    SeguridadCRC(242) = 13
    SeguridadCRC(243) = 208
    SeguridadCRC(244) = 232
    SeguridadCRC(245) = 108
    SeguridadCRC(246) = 171
    SeguridadCRC(247) = 247
    SeguridadCRC(248) = 171
    SeguridadCRC(249) = 183
    SeguridadCRC(250) = 76
    SeguridadCRC(251) = 208
    SeguridadCRC(252) = 46
    SeguridadCRC(253) = 66
    SeguridadCRC(254) = 169
    SeguridadCRC(255) = 252
    SeguridadCRC(256) = 30
    SeguridadCRC(257) = 90
    SeguridadCRC(258) = 238
    SeguridadCRC(259) = 203
    SeguridadCRC(260) = 24
    SeguridadCRC(261) = 116
    SeguridadCRC(262) = 200
    SeguridadCRC(263) = 2
    SeguridadCRC(264) = 97
    SeguridadCRC(265) = 19
    SeguridadCRC(266) = 192
    SeguridadCRC(267) = 220
    SeguridadCRC(268) = 214
    SeguridadCRC(269) = 237
    SeguridadCRC(270) = 199
    SeguridadCRC(271) = 78
    SeguridadCRC(272) = 38
    SeguridadCRC(273) = 73
    SeguridadCRC(274) = 18
    SeguridadCRC(275) = 143
    SeguridadCRC(276) = 62
    SeguridadCRC(277) = 171
    SeguridadCRC(278) = 40
    SeguridadCRC(279) = 216
    SeguridadCRC(280) = 5
    SeguridadCRC(281) = 179
    SeguridadCRC(282) = 57
    SeguridadCRC(283) = 104
    SeguridadCRC(284) = 74
    SeguridadCRC(285) = 67
    SeguridadCRC(286) = 177
    SeguridadCRC(287) = 204
    SeguridadCRC(288) = 250
    SeguridadCRC(289) = 224
    SeguridadCRC(290) = 13
    SeguridadCRC(291) = 93
    SeguridadCRC(292) = 151
    SeguridadCRC(293) = 91
    SeguridadCRC(294) = 237
    SeguridadCRC(295) = 10
    SeguridadCRC(296) = 229
    SeguridadCRC(297) = 176
    SeguridadCRC(298) = 107
    SeguridadCRC(299) = 88
    SeguridadCRC(300) = 231
    SeguridadCRC(301) = 46
    SeguridadCRC(302) = 172
    SeguridadCRC(303) = 166
    SeguridadCRC(304) = 9
    SeguridadCRC(305) = 216
    SeguridadCRC(306) = 180
    SeguridadCRC(307) = 182
    SeguridadCRC(308) = 159
    SeguridadCRC(309) = 12
    SeguridadCRC(310) = 127
    SeguridadCRC(311) = 105
    SeguridadCRC(312) = 142
    SeguridadCRC(313) = 98
    SeguridadCRC(314) = 77
    SeguridadCRC(315) = 202
    SeguridadCRC(316) = 73
    SeguridadCRC(317) = 215
    SeguridadCRC(318) = 61
    SeguridadCRC(319) = 78
    SeguridadCRC(320) = 0
    SeguridadCRC(321) = 43
    SeguridadCRC(322) = 29
    SeguridadCRC(323) = 90
    SeguridadCRC(324) = 19
    SeguridadCRC(325) = 135
    SeguridadCRC(326) = 129
    SeguridadCRC(327) = 6
    SeguridadCRC(328) = 205
    SeguridadCRC(329) = 99
    SeguridadCRC(330) = 18
    SeguridadCRC(331) = 33
    SeguridadCRC(332) = 79
    SeguridadCRC(333) = 167
    SeguridadCRC(334) = 41
    SeguridadCRC(335) = 117
    SeguridadCRC(336) = 202
    SeguridadCRC(337) = 16
    SeguridadCRC(338) = 157
    SeguridadCRC(339) = 76
    SeguridadCRC(340) = 242
    SeguridadCRC(341) = 214
    SeguridadCRC(342) = 216
    SeguridadCRC(343) = 50
    SeguridadCRC(344) = 175
    SeguridadCRC(345) = 140
    SeguridadCRC(346) = 49
    SeguridadCRC(347) = 253
    SeguridadCRC(348) = 21
    SeguridadCRC(349) = 71
    SeguridadCRC(350) = 117
    SeguridadCRC(351) = 11
    SeguridadCRC(352) = 150
    SeguridadCRC(353) = 2
    SeguridadCRC(354) = 199
    SeguridadCRC(355) = 203
    SeguridadCRC(356) = 118
    SeguridadCRC(357) = 65
    SeguridadCRC(358) = 171
    SeguridadCRC(359) = 127
    SeguridadCRC(360) = 128
    SeguridadCRC(361) = 245
    SeguridadCRC(362) = 93
    SeguridadCRC(363) = 64
    SeguridadCRC(364) = 248
    SeguridadCRC(365) = 160
    SeguridadCRC(366) = 103
    SeguridadCRC(367) = 66
    SeguridadCRC(368) = 208
    SeguridadCRC(369) = 185
    SeguridadCRC(370) = 114
    SeguridadCRC(371) = 89
    SeguridadCRC(372) = 30
    SeguridadCRC(373) = 82
    SeguridadCRC(374) = 93
    SeguridadCRC(375) = 188
    SeguridadCRC(376) = 206
    SeguridadCRC(377) = 248
    SeguridadCRC(378) = 140
    SeguridadCRC(379) = 9
    SeguridadCRC(380) = 148
    SeguridadCRC(381) = 219
    SeguridadCRC(382) = 131
    SeguridadCRC(383) = 138
    SeguridadCRC(384) = 37
    SeguridadCRC(385) = 46
    SeguridadCRC(386) = 179
    SeguridadCRC(387) = 183
    SeguridadCRC(388) = 167
    SeguridadCRC(389) = 209
    SeguridadCRC(390) = 147
    SeguridadCRC(391) = 252
    SeguridadCRC(392) = 102
    SeguridadCRC(393) = 46
    SeguridadCRC(394) = 243
    SeguridadCRC(395) = 188
    SeguridadCRC(396) = 200
    SeguridadCRC(397) = 96
    SeguridadCRC(398) = 141
    SeguridadCRC(399) = 149
    SeguridadCRC(400) = 131
    SeguridadCRC(401) = 155
    SeguridadCRC(402) = 222
    SeguridadCRC(403) = 230
    SeguridadCRC(404) = 13
    SeguridadCRC(405) = 200
    SeguridadCRC(406) = 52
    SeguridadCRC(407) = 142
    SeguridadCRC(408) = 84
    SeguridadCRC(409) = 111
    SeguridadCRC(410) = 7
    SeguridadCRC(411) = 247
    SeguridadCRC(412) = 176
    SeguridadCRC(413) = 218
    SeguridadCRC(414) = 140
    SeguridadCRC(415) = 83
    SeguridadCRC(416) = 22
    SeguridadCRC(417) = 120
    SeguridadCRC(418) = 136
    SeguridadCRC(419) = 38
    SeguridadCRC(420) = 142
    SeguridadCRC(421) = 127
    SeguridadCRC(422) = 98
    SeguridadCRC(423) = 5
    SeguridadCRC(424) = 231
    SeguridadCRC(425) = 213
    SeguridadCRC(426) = 125
    SeguridadCRC(427) = 157
    SeguridadCRC(428) = 169
    SeguridadCRC(429) = 49
    SeguridadCRC(430) = 196
    SeguridadCRC(431) = 246
    SeguridadCRC(432) = 75
    SeguridadCRC(433) = 125
    SeguridadCRC(434) = 135
    SeguridadCRC(435) = 249
    SeguridadCRC(436) = 166
    SeguridadCRC(437) = 127
    SeguridadCRC(438) = 133
    SeguridadCRC(439) = 49
    SeguridadCRC(440) = 170
    SeguridadCRC(441) = 185
    SeguridadCRC(442) = 74
    SeguridadCRC(443) = 206
    SeguridadCRC(444) = 80
    SeguridadCRC(445) = 142
    SeguridadCRC(446) = 187
    SeguridadCRC(447) = 239
    SeguridadCRC(448) = 207
    SeguridadCRC(449) = 165
    SeguridadCRC(450) = 239
    SeguridadCRC(451) = 33
    SeguridadCRC(452) = 19
    SeguridadCRC(453) = 147
    SeguridadCRC(454) = 64
    SeguridadCRC(455) = 34
    SeguridadCRC(456) = 107
    SeguridadCRC(457) = 180
    SeguridadCRC(458) = 162
    SeguridadCRC(459) = 235
    SeguridadCRC(460) = 130
    SeguridadCRC(461) = 89
    SeguridadCRC(462) = 52
    SeguridadCRC(463) = 238
    SeguridadCRC(464) = 144
    SeguridadCRC(465) = 41
    SeguridadCRC(466) = 21
    SeguridadCRC(467) = 157
    SeguridadCRC(468) = 209
    SeguridadCRC(469) = 193
    SeguridadCRC(470) = 121
    SeguridadCRC(471) = 43
    SeguridadCRC(472) = 54
    SeguridadCRC(473) = 158
    SeguridadCRC(474) = 252
    SeguridadCRC(475) = 150
    SeguridadCRC(476) = 91
    SeguridadCRC(477) = 61
    SeguridadCRC(478) = 53
    SeguridadCRC(479) = 229
    SeguridadCRC(480) = 186
    SeguridadCRC(481) = 128
    SeguridadCRC(482) = 143
    SeguridadCRC(483) = 174
    SeguridadCRC(484) = 30
    SeguridadCRC(485) = 84
    SeguridadCRC(486) = 84
    SeguridadCRC(487) = 220
    SeguridadCRC(488) = 90
    SeguridadCRC(489) = 145
    SeguridadCRC(490) = 11
    SeguridadCRC(491) = 175
    SeguridadCRC(492) = 58
    SeguridadCRC(493) = 33
    SeguridadCRC(494) = 4
    SeguridadCRC(495) = 4
    SeguridadCRC(496) = 186
    SeguridadCRC(497) = 101
    SeguridadCRC(498) = 49
    SeguridadCRC(499) = 215
    SeguridadCRC(500) = 118
End Sub

Public Function ConvertirFlush(ByVal data As String) As String
Dim i As Integer

For i = 1 To Len(data)
    ConvertirFlush = ConvertirFlush & Chr(Asc(mid(data, i, 1)) Xor SeguridadCRC(CRC)) ' + 29))
Next i
End Function

Public Sub SumarFlush()
CRC = CRC + 1
If CRC = 501 Then CRC = 0
End Sub
