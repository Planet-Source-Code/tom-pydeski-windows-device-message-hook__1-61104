Attribute VB_Name = "WM_Constants"
'Submitted by Tom Pydeski
'code for displaying the description of windows messaging constants
'
'# Several messages for testing ...
' The following constants are the Win32 Message constants.
' WM_Activate Values
Public Const WM_NULL = &H0 '=0
Public Const WM_CREATE = &H1 '=1
Public Const WM_DESTROY = &H2 '=2
Public Const WM_MOVE = &H3 '=3
Public Const WM_SIZE = &H5 '=5
Public Const WM_ACTIVATE = &H6 '=6
Public Const WM_SETFOCUS = &H7 '=7
Public Const WM_KILLFOCUS = &H8 '=8
Public Const WM_ENABLE = &HA '=10
Public Const WM_SETREDRAW = &HB '=11
Public Const WM_SETTEXT = &HC '=12
Public Const WM_GETTEXT = &HD '=13
Public Const WM_GETTEXTLENGTH = &HE '=14
Public Const WM_PAINT = &HF '=15
Public Const WM_CLOSE = &H10 '=16
Public Const WM_QUERYENDSESSION = &H11 '=17
Public Const WM_QUIT = &H12 '=18
Public Const WM_QUERYOPEN = &H13 '=19
Public Const WM_ERASEBKGND = &H14 '=20
Public Const WM_SYSCOLORCHANGE = &H15 '=21
Public Const WM_ENDSESSION = &H16 '=22
Public Const WM_SHOWWINDOW = &H18 '=24
'Public Const WM_ = &H19 '=25
Public Const WM_WinIniChange = &H1A      '=26
Public Const WM_DEVMODECHANGE = &H1B '=27
Public Const WM_ACTIVATEAPP = &H1C '=28
Public Const WM_FONTCHANGE = &H1D '=29
Public Const WM_TIMECHANGE = &H1E '=30
Public Const WM_CANCELMODE = &H1F '=31
Public Const WM_SETCURSOR = &H20 '=32
Public Const WM_MOUSEACTIVATE = &H21 '=33
Public Const WM_CHILDACTIVATE = &H22 '=34
Public Const WM_QUEUESYNC = &H23 '=35
Public Const WM_GETMINMAXINFO = &H24 '=36
Public Const WM_PAINTICON = &H26 '=38
Public Const WM_ICONERASEBKGND = &H27 '=39
Public Const WM_NEXTDLGCTL = &H28 '=40
Public Const WM_DrawItem = &H2B  '=43
Public Const WM_MeasureItem = &H2C       '=44
Public Const WM_DeleteItem = &H2D        '=45
Public Const WM_VkeytoItem = &H2E        '=46
Public Const WM_ChartoItem = &H2F        '=47
Public Const WM_SETFONT = &H30 '=48
Public Const WM_GETFONT = &H31 '=49
Public Const WM_SETHOTKEY = &H32 '=50
Public Const WM_GETHOTKEY = &H33 '=51
Public Const WM_QUERYDRAGICON = &H37 '=55
Public Const WM_COMPAREITEM = &H39 '=57
Public Const WM_GETOBJECT = &H3D '=61
Public Const WM_COMPACTING = &H41 '=65
Public Const WM_OTHERWINDOWCREATED = &H42 '=66
Public Const WM_OTHERWINDOWDESTROYED = &H43 '=67
Public Const WM_COMMNOTIFY = &H44 '=68
'Public Const WM_ = &H45 '=69
Public Const WM_WINDOWPOSCHANGING = &H46 '=70
Public Const WM_WINDOWPOSCHANGED = &H47 '=71
Public Const WM_POWER = &H48 '=72
Public Const WM_CopyData = &H4A  '=74
Public Const WM_NOTIFY = &H4E '=78
Public Const WM_INPUTLANGCHANGEREQUEST = &H50 '=80
Public Const WM_INPUTLANGCHANGE = &H51 '=81
Public Const WM_TCARD = &H52 '=82
Public Const WM_HELP = &H53 '=83
Public Const WM_USERCHANGED = &H54 '=84
Public Const WM_NOTIFYFORMAT = &H55 '=85
Public Const WM_CONTEXTMENU = &H7B '=123
Public Const WM_STYLECHANGING = &H7C '=124
Public Const WM_STYLECHANGED = &H7D '=125
Public Const WM_DISPLAYCHANGE = &H7E '=126
Public Const WM_GETICON = &H7F '=127
Public Const WM_SETICON = &H80 '=128
Public Const WM_NCCREATE = &H81 '=129
Public Const WM_NCDESTROY = &H82 '=130
Public Const WM_NCCALCSIZE = &H83 '=131
Public Const WM_NCHITTEST = &H84 '=132
Public Const WM_NCPAINT = &H85 '=133
Public Const WM_NCACTIVATE = &H86 '=134
Public Const WM_GETDLGCODE = &H87 '=135
Public Const WM_NCMOUSEMOVE = &HA0 '=160
Public Const WM_NCLBUTTONDOWN = &HA1 '=161
Public Const WM_NCLBUTTONUP = &HA2 '=162
Public Const WM_NCLBUTTONDBLCLK = &HA3 '=163
Public Const WM_NCRBUTTONDOWN = &HA4 '=164
Public Const WM_NCRBUTTONUP = &HA5 '=165
Public Const WM_NCRBUTTONDBLCLK = &HA6 '=166
Public Const WM_NCMBUTTONDOWN = &HA7 '=167
Public Const WM_NCMBUTTONUP = &HA8 '=168
Public Const WM_NCMBUTTONDBLCLK = &HA9 '=169
'
Public Const WM_KEYDOWN = &H100 '=256
Public Const WM_KEYUP = &H101 '=257
Public Const WM_CHAR = &H102 '=258
Public Const WM_DEADCHAR = &H103 '=259
Public Const WM_SYSKEYDOWN = &H104 '=260
Public Const WM_SYSKEYUP = &H105 '=261
Public Const WM_SYSCHAR = &H106 '=262
Public Const WM_SYSDEADCHAR = &H107 '=263
Public Const WM_KEYLAST = &H108 '=264
Public Const WM_IM_INFO = &H10C  '=268
Public Const WM_IME_STARTCOMPOSITION = &H10D '=269
Public Const WM_IME_ENDCOMPOSITION = &H10E '=270
Public Const WM_IME_COMPOSITION = &H10F '=271
Public Const WM_IME_KEYLAST = &H10F '=271
Public Const WM_INITDIALOG = &H110 '=272
Public Const WM_COMMAND = &H111 '=273
Public Const WM_SYSCOMMAND = &H112 '=274
Public Const WM_TIMER = &H113 '=275
Public Const WM_HSCROLL = &H114 '=276
Public Const WM_VSCROLL = &H115 '=277
Public Const WM_INITMENU = &H116 '=278
Public Const WM_INITMENUPOPUP = &H117 '=279
Public Const WM_MENUSELECT = &H11F '=287
Public Const WM_MENUCHAR = &H120 '=288
Public Const WM_ENTERIDLE = &H121 '=289
Public Const WM_MENURBUTTONUP = &H122 '=290
Public Const WM_MENUDRAG = &H123 '=291
Public Const WM_MENUGETOBJECT = &H124 '=292
Public Const WM_UNINITMENUPOPUP = &H125 '=293
Public Const WM_MENUCOMMAND = &H126 '=294
'
Public Const WM_CTLCOLORMSGBOX = &H132 '=306
Public Const WM_CTLCOLOREDIT = &H133 '=307
Public Const WM_CTLCOLORLISTBOX = &H134 '=308
Public Const WM_CTLCOLORBTN = &H135 '=309
Public Const WM_CTLCOLORDLG = &H136 '=310
Public Const WM_CTLCOLORSCROLLBAR = &H137 '=311
Public Const WM_CTLCOLORSTATIC = &H138 '=312
Public Const WM_MOUSEFIRST = &H200 '=512
Public Const WM_MOUSEMOVE = &H200 '=512
Public Const WM_LBUTTONDOWN = &H201 '=513
Public Const WM_LBUTTONUP = &H202 '=514
Public Const WM_LBUTTONDBLCLK = &H203 '=515
Public Const WM_RBUTTONDOWN = &H204 '=516
Public Const WM_RBUTTONUP = &H205 '=517
Public Const WM_RBUTTONDBLCLK = &H206 '=518
Public Const WM_MBUTTONDOWN = &H207 '=519
Public Const WM_MBUTTONUP = &H208 '=520
Public Const WM_MBUTTONDBLCLK = &H209 '=521
Public Const WM_MOUSELAST = &H20A '=522
Public Const WM_MOUSEWHEEL = &H20A '=522
Public Const WM_PARENTNOTIFY = &H210 '=528
Public Const WM_ENTERMENULOOP = &H211 '=529
Public Const WM_EXITMENULOOP = &H212 '=530
Public Const WM_NEXTMENU = &H213 '=531
Public Const WM_SIZING = &H214 '=532
Public Const WM_CAPTURECHANGED = &H215 '=533
Public Const WM_MOVING = &H216 '=534
Public Const WM_POWERBROADCAST = &H218 '=536
Public Const WM_DEVICECHANGE = &H219 '=537
Public Const WM_MDICREATE = &H220 '=544
Public Const WM_MDIDESTROY = &H221 '=545
Public Const WM_MDIACTIVATE = &H222 '=546
Public Const WM_MDIRESTORE = &H223 '=547
Public Const WM_MDINEXT = &H224 '=548
Public Const WM_MDIMAXIMIZE = &H225 '=549
Public Const WM_MDITILE = &H226 '=550
Public Const WM_MDICASCADE = &H227 '=551
Public Const WM_MDIICONARRANGE = &H228 '=552
Public Const WM_MDIGETACTIVE = &H229 '=553
Public Const WM_MDISETMENU = &H230 '=560
Public Const WM_ENTERSIZEMOVE = &H231 '=561
Public Const WM_EXITSIZEMOVE = &H232 '=562
Public Const WM_DROPFILES = &H233 '=563
Public Const WM_MDIREFRESHMENU = &H234 '=564
Public Const WM_IME_SETCONTEXT = &H281 '=641
Public Const WM_IME_NOTIFY = &H282 '=642
Public Const WM_IME_CONTROL = &H283 '=643
Public Const WM_IME_COMPOSITIONFULL = &H284 '=644
Public Const WM_IME_SELECT = &H285 '=645
Public Const WM_IME_CHAR = &H286 '=646
Public Const WM_IME_KEYDOWN = &H290 '=656
Public Const WM_IME_KEYUP = &H291 '=657
Public Const WM_MOUSEHOVER = &H2A1 '=673
Public Const WM_MOUSELEAVE = &H2A3 '=675
'
Public Const WM_CUT = &H300 '=768
Public Const WM_COPY = &H301 '=769
Public Const WM_PASTE = &H302 '=770
Public Const WM_CLEAR = &H303 '=771
Public Const WM_UNDO = &H304 '=772
Public Const WM_RENDERFORMAT = &H305 '=773
Public Const WM_RENDERALLFORMATS = &H306 '=774
Public Const WM_DESTROYCLIPBOARD = &H307 '=775
Public Const WM_DRAWCLIPBOARD = &H308 '=776
Public Const WM_PAINTCLIPBOARD = &H309 '=777
Public Const WM_VSCROLLCLIPBOARD = &H30A '=778
Public Const WM_SIZECLIPBOARD = &H30B '=779
Public Const WM_ASKCBFORMATNAME = &H30C '=780
Public Const WM_CHANGECBCHAIN = &H30D '=781
Public Const WM_HSCROLLCLIPBOARD = &H30E '=782
Public Const WM_QUERYNEWPALETTE = &H30F '=783
Public Const WM_PALETTEISCHANGING = &H310 '=784
Public Const WM_PALETTECHANGED = &H311 '=785
Public Const WM_HOTKEY = &H312 '=786
Public Const WM_PRINT = &H317 '=791
Public Const WM_PRINTCLIENT = &H318 '=792
Public Const WM_HANDHELDFIRST = &H358 '=856
Public Const WM_HANDHELDLAST = &H35F '=863
Public Const WM_AFXFIRST = &H360 '=864
Public Const WM_AFXLAST = &H37F '=895
Public Const WM_PENWINFIRST = &H380 '=896
Public Const WM_PENWINLAST = &H38F '=911
Public Const WM_USER = &H400 '=1024
Public Const WM_APP = &H8000 '=32768
'
Public Enum enPowerBroadcastType
    PBT_APMQUERYSUSPEND = &H0
    PBT_APMQUERYSTANDBY = &H1
    PBT_APMQUERYSUSPENDFAILED = &H2
    PBT_APMQUERYSTANDBYFAILED = &H3
    PBT_APMSUSPEND = &H4
    PBT_APMSTANDBY = &H5
    PBT_APMRESUMECRITICAL = &H6
    PBT_APMRESUMESUSPEND = &H7
    PBT_APMRESUMESTANDBY = &H8
End Enum
Public Const PWR_SUSPENDREQUEST = 1
Public Const PWR_FAIL = (-1)
Public Const BROADCAST_QUERY_DENY = &H424D5144
'
Global WinMess() As String * 22
'
'Screensaver launching sends the following
'536 Powerbroadcast wparam=10     lparam=0
'274 Syscommand     wparam=61760  lparam=0
'

Sub SetWinMess()
ReDim WinMess(1025) As String * 22
'For i = 0 To UBound(WinMess) - 1
'    WinMess(i) = Space(25)
'Next i
WinMess(0) = "Null"
WinMess(1) = "Create"
WinMess(2) = "Destroy"
WinMess(3) = "Move"
WinMess(5) = "Size"
WinMess(6) = "Activate"
WinMess(7) = "SetFocus"
WinMess(8) = "KillFocus"
WinMess(10) = "Enable"
WinMess(11) = "SetRedraw"
WinMess(12) = "SetText"
WinMess(13) = "GetText"
WinMess(14) = "GetTextLength"
WinMess(15) = "Paint"
WinMess(16) = "Close"
WinMess(17) = "QueryEndSession"
WinMess(18) = "Quit"
WinMess(19) = "QueryOpen"
WinMess(20) = "EraseBkgnd"
WinMess(21) = "SysColorChange"
WinMess(22) = "EndSession"
WinMess(24) = "ShowWindow"
WinMess(26) = "WinIniChange"
WinMess(27) = "DevModeChange"
WinMess(28) = "ActivateApp"
WinMess(29) = "FontChange"
WinMess(30) = "TimeChange"
WinMess(31) = "CancelMode"
WinMess(32) = "SetCursor"
WinMess(33) = "MouseActivate"
WinMess(34) = "ChildActivate"
WinMess(35) = "QueueSync"
WinMess(36) = "GetMinMaxInfo"
WinMess(38) = "PaintIcon"
WinMess(39) = "IconEraseBkgnd"
WinMess(40) = "Nextdlgctl"
WinMess(43) = "DrawItem"
WinMess(44) = "MeasureItem"
WinMess(45) = "DeleteItem"
WinMess(46) = "VkeytoItem"
WinMess(47) = "ChartoItem"
WinMess(48) = "SetFont"
WinMess(49) = "GetFont"
WinMess(50) = "SetHotKey"
WinMess(51) = "GetHotKey"
WinMess(55) = "QueryDragIcon"
WinMess(57) = "CompareItem"
WinMess(61) = "GetObject"
WinMess(65) = "Compacting"
WinMess(66) = "OtherWindowCreated"
WinMess(67) = "OtherWindowDestroyed"
WinMess(68) = "CommNotify"
WinMess(70) = "WindowPoschanging"
WinMess(71) = "WindowPosChanged"
WinMess(72) = "Power"
WinMess(74) = "CopyData"
WinMess(78) = "Notify"
WinMess(80) = "InputLangChangerequest"
WinMess(81) = "InputLangChange"
WinMess(82) = "Tcard"
WinMess(83) = "Help"
WinMess(84) = "UserChanged"
WinMess(85) = "Notifyformat"
WinMess(123) = "ConTextMenu"
WinMess(124) = "Stylechanging"
WinMess(125) = "StyleChanged"
WinMess(126) = "DisplayChange"
WinMess(127) = "GetIcon"
WinMess(128) = "SetIcon"
WinMess(129) = "NcCreate"
WinMess(130) = "NcDestroy"
WinMess(131) = "NcCalcsize"
WinMess(132) = "NcHittest"
WinMess(133) = "NcPaint"
WinMess(134) = "NcActivate"
WinMess(135) = "GetDlgCode"
WinMess(160) = "NcMouseMove"
WinMess(161) = "NclButtonDown"
WinMess(162) = "NclButtonUp"
WinMess(163) = "NclButtonDBlclk"
WinMess(164) = "NcrButtonDown"
WinMess(165) = "NcrButtonUp"
WinMess(166) = "NcrButtonDBlclk"
WinMess(167) = "NcmButtonDown"
WinMess(168) = "NcmButtonUp"
WinMess(169) = "NcmButtonDBlclk"
WinMess(256) = "KeyFirst"
WinMess(256) = "KeyDown"
WinMess(257) = "KeyUp"
WinMess(258) = "Char"
WinMess(259) = "Deadchar"
WinMess(260) = "SysKeyDown"
WinMess(261) = "SysKeyUp"
WinMess(262) = "Syschar"
WinMess(263) = "Sysdeadchar"
WinMess(264) = "Keylast"
WinMess(268) = "IM_Info"
WinMess(269) = "Ime_startcomPosition"
WinMess(270) = "Ime_endcomPosition"
WinMess(271) = "Ime_comPosition"
WinMess(271) = "Ime_Keylast"
WinMess(272) = "Initdialog"
WinMess(273) = "Command"
WinMess(274) = "Syscommand"
WinMess(275) = "Timer"
WinMess(276) = "Hscroll"
WinMess(277) = "Vscroll"
WinMess(278) = "InitMenu"
WinMess(279) = "InitMenuPopUp"
WinMess(287) = "Menuselect"
WinMess(288) = "MenuChar"
WinMess(289) = "EnterIdle"
WinMess(293) = "WindowsMessage 293"
WinMess(306) = "CtlColorMsgbox"
WinMess(307) = "CtlColorEdit"
WinMess(308) = "CtlColorListbox"
WinMess(309) = "CtlColorBtn"
WinMess(310) = "CtlColorDlg"
WinMess(311) = "CtlColorScrollbar"
WinMess(312) = "CtlColorStatic"
WinMess(512) = "MouseFirst"
WinMess(512) = "MouseMove"
WinMess(513) = "LButtonDown"
WinMess(514) = "LButtonUp"
WinMess(515) = "LButtonDBlclk"
WinMess(516) = "RButtonDown"
WinMess(517) = "RButtonUp"
WinMess(518) = "RButtonDBlclk"
WinMess(519) = "MButtonDown"
WinMess(520) = "MButtonUp"
WinMess(521) = "MButtonDBlclk"
WinMess(522) = "Mouselast"
WinMess(522) = "Mousewheel"
WinMess(528) = "Parentnotify"
WinMess(529) = "EnterMenuloop"
WinMess(530) = "ExitMenuloop"
WinMess(531) = "NextMenu"
WinMess(532) = "Sizing"
WinMess(533) = "CaptureChanged"
WinMess(534) = "Moving"
WinMess(536) = "Powerbroadcast"
WinMess(537) = "DeviceChange"
WinMess(544) = "MDICreate"
WinMess(545) = "MDIDestroy"
WinMess(546) = "MDIActivate"
WinMess(547) = "MDIRestore"
WinMess(548) = "MDINext"
WinMess(549) = "MDIMaximize"
WinMess(550) = "MDItile"
WinMess(551) = "MDIcascade"
WinMess(552) = "MDIiconarrange"
WinMess(553) = "MDIgetActive"
WinMess(560) = "MDIsetMenu"
WinMess(561) = "Entersizemove"
WinMess(562) = "Exitsizemove"
WinMess(563) = "Dropfiles"
WinMess(564) = "MDIrefreshMenu"
WinMess(641) = "Ime_setconText"
WinMess(642) = "Ime_notify"
WinMess(643) = "Ime_control"
WinMess(644) = "Ime_comPositionfull"
WinMess(645) = "Ime_select"
WinMess(646) = "Ime_char"
WinMess(656) = "Ime_KeyDown"
WinMess(657) = "Ime_KeyUp"
WinMess(673) = "MouseHover"
WinMess(674) = "MouseEnter??"
WinMess(675) = "MouseLeave"
WinMess(768) = "Cut"
WinMess(769) = "Copy"
WinMess(770) = "Paste"
WinMess(771) = "Clear"
WinMess(772) = "Undo"
WinMess(773) = "Renderformat"
WinMess(774) = "Renderallformats"
WinMess(775) = "DestroyClipBoard"
WinMess(776) = "DrawClipBoard"
WinMess(777) = "PaintClipBoard"
WinMess(778) = "VscrollClipBoard"
WinMess(779) = "SizeClipBoard"
WinMess(780) = "AskCBFormatName"
WinMess(781) = "ChangeCBChain"
WinMess(782) = "HscrollClipBoard"
WinMess(783) = "QueryNewPalette"
WinMess(784) = "PaletteIsChanging"
WinMess(785) = "PaletteChanged"
WinMess(786) = "HotKey"
WinMess(791) = "Print"
WinMess(792) = "Printclient"
WinMess(856) = "Handheldfirst"
WinMess(863) = "Handheldlast"
WinMess(864) = "Afxfirst"
WinMess(895) = "Afxlast"
WinMess(896) = "PenWinfirst"
WinMess(911) = "PenWinlast"
WinMess(1024) = "User"
'WinMess(32768) = "App"
End Sub
