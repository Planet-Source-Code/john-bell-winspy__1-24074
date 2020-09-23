Attribute VB_Name = "danAPI"
Option Explicit
'these are structures for handling ttf info
Type OffsetTable
    sfntVersionA As Integer
    sfntVersionB As Integer
    numTables As Integer
    searchRange As Integer
    entrySelector As Integer
    rangeShift As Integer
End Type

Type TableDirectoryEntry
    tag(3) As Byte
    checksum As Long
    offset As Long
    Length As Long
End Type

Type NamingTable
    FormatSelector As Integer
    NameRecords As Integer
    offsStrings As Integer
End Type

Type NameRecord
    PlatformID As Integer
    PlatformSecifics As Integer
    LanguageID As Integer
    NameID As Integer
    StringLength As Integer
    StringOffset As Integer
End Type

'this structure handles system memory info
Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
'windows structures
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type POINTAPI
    x As Long
    Y As Long
End Type
'gets system memory availible
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'kernel32 lower level memory functions
'user32 relating to windows management(cursors, carets, messages)
'gdi32 graphics device interface library
'comdlg32,lz32,version32 additional capabilities, including support for file compression, common dialogs, ect.
'mapi32 lets any application work with electronic mail
'netapi32 access and control networks
'odbc32 lets you work with multiple types of databases
'winmm lets you access multimedia
'chapter two most significant thing, bitfie
Public PointMode%

Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long 'loads a system cursor based on a passed constant
Public Declare Function SetClassWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long 'used to set the current cursor(and other stuff)
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long 'what it says stupid
Public Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long 'what it says stupid
Public Declare Function GetCursor Lib "user32" () As Long 'gets the current cursor
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Public Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function UnionRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function SubtractRect Lib "user32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Public Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ptx As Long, pty As Long) As Long
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachto As Long, ByVal fAttach As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam&) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long
' We create a special SendMessage alias that accepts a long value by reference
Public Declare Function SendMessageLongByRef Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' Note the use of two longs to transfer a POINTAPI structure.
Public Declare Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const GCW_HCURSOR = (-12)
Public Const IDC_SIZEALL = 32646&

' Straight port to Win32. The old GWW and GCW constants are gone
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const GWL_ID = (-12)

Public Const GCL_MENUNAME = (-8)
Public Const GCL_HBRBACKGROUND = (-10)
Public Const GCL_HCURSOR = (-12)
Public Const GCL_HICON = (-14)
Public Const GCL_HMODULE = (-16)
Public Const GCL_CBWNDEXTRA = (-18)
Public Const GCL_CBCLSEXTRA = (-20)
Public Const GCL_WNDPROC = (-24)
Public Const GCL_STYLE = (-26)
Public Const GCW_ATOM = (-32)


' Style constants remain the same
' The previous version used the "Global" keyword. Replacing
' it with "Public" is optional.
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000

Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000

Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_DBLCLKS = &H8
Public Const CS_OWNDC = &H20
Public Const CS_CLASSDC = &H40
Public Const CS_PARENTDC = &H80
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOCLOSE = &H200
Public Const CS_SAVEBITS = &H800
Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_PUBLICCLASS = &H4000

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5

Public Const ES_LEFT = &H0&
Public Const ES_CENTER = &H1&
Public Const ES_RIGHT = &H2&
Public Const ES_MULTILINE = &H4&
Public Const ES_UPPERCASE = &H8&
Public Const ES_LOWERCASE = &H10&
Public Const ES_PASSWORD = &H20&
Public Const ES_AUTOVSCROLL = &H40&
Public Const ES_AUTOHSCROLL = &H80&
Public Const ES_NOHIDESEL = &H100&
Public Const ES_OEMCONVERT = &H400&
Public Const ES_READONLY = &H800&
Public Const ES_WANTRETURN = &H1000&

Public Const BS_PUSHBUTTON = &H0&
Public Const BS_DEFPUSHBUTTON = &H1&
Public Const BS_CHECKBOX = &H2&
Public Const BS_AUTOCHECKBOX = &H3&
Public Const BS_RADIOBUTTON = &H4&
Public Const BS_3STATE = &H5&
Public Const BS_AUTO3STATE = &H6&
Public Const BS_GROUPBOX = &H7&
Public Const BS_USERBUTTON = &H8&
Public Const BS_AUTORADIOBUTTON = &H9&
Public Const BS_OWNERDRAW = &HB&
Public Const BS_LEFTTEXT = &H20&
' New button styles for Windows 95
Public Const BS_TEXT = 0&
Public Const BS_ICON = &H40&
Public Const BS_BITMAP = &H80&
Public Const BS_LEFT = &H100&
Public Const BS_RIGHT = &H200&
Public Const BS_CENTER = &H300&
Public Const BS_TOP = &H400&
Public Const BS_BOTTOM = &H800&
Public Const BS_VCENTER = &HC00&
Public Const BS_PUSHLIKE = &H1000&
Public Const BS_MULTILINE = &H2000&
Public Const BS_NOTIFY = &H4000&
Public Const BS_FLAT = &H8000&
Public Const BS_RIGHTBUTTON = &H20&

Public Const SS_LEFT = &H0&
Public Const SS_CENTER = &H1&
Public Const SS_RIGHT = &H2&
Public Const SS_ICON = &H3&
Public Const SS_BLACKRECT = &H4&
Public Const SS_GRAYRECT = &H5&
Public Const SS_WHITERECT = &H6&
Public Const SS_BLACKFRAME = &H7&
Public Const SS_GRAYFRAME = &H8&
Public Const SS_WHITEFRAME = &H9&
Public Const SS_USERITEM = &HA&
Public Const SS_SIMPLE = &HB&
Public Const SS_LEFTNOWORDWRAP = &HC&
Public Const SS_NOPREFIX = &H80           '  Don't do "&" character translation

Public Const DS_ABSALIGN = &H1&
Public Const DS_SYSMODAL = &H2&
Public Const DS_LOCALEDIT = &H20
Public Const DS_SETFONT = &H40
Public Const DS_MODALFRAME = &H80
Public Const DS_NOIDLEMSG = &H100
Public Const DS_SETFOREGROUND = &H200


Global Const WM_USER = &H400

' Watch out here - control message numbers have changed!
Public Const LB_RESETCONTENT = &H184
Public Const LB_SETTABSTOPS = &H192

Public Const LBS_NOTIFY = &H1&
Public Const LBS_SORT = &H2&
Public Const LBS_NOREDRAW = &H4&
Public Const LBS_MULTIPLESEL = &H8&
Public Const LBS_OWNERDRAWFIXED = &H10&
Public Const LBS_OWNERDRAWVARIABLE = &H20&
Public Const LBS_HASSTRINGS = &H40&
Public Const LBS_USETABSTOPS = &H80&
Public Const LBS_NOINTEGRALHEIGHT = &H100&
Public Const LBS_MULTICOLUMN = &H200&
Public Const LBS_WANTKEYBOARDINPUT = &H400&
Public Const LBS_EXTENDEDSEL = &H800&
Public Const LBS_DISABLENOSCROLL = &H1000&
Public Const LBS_NODATA = &H2000&

Public Const CBS_SIMPLE = &H1&
Public Const CBS_DROPDOWN = &H2&
Public Const CBS_DROPDOWNLIST = &H3&
Public Const CBS_OWNERDRAWFIXED = &H10&
Public Const CBS_OWNERDRAWVARIABLE = &H20&
Public Const CBS_AUTOHSCROLL = &H40&
Public Const CBS_OEMCONVERT = &H80&
Public Const CBS_SORT = &H100&
Public Const CBS_HASSTRINGS = &H200&
Public Const CBS_NOINTEGRALHEIGHT = &H400&
Public Const CBS_DISABLENOSCROLL = &H800&

Public Const SBS_HORZ = &H0&
Public Const SBS_VERT = &H1&
Public Const SBS_TOPALIGN = &H2&
Public Const SBS_LEFTALIGN = &H2&
Public Const SBS_BOTTOMALIGN = &H4&
Public Const SBS_RIGHTALIGN = &H4&
Public Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
Public Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Public Const SBS_SIZEBOX = &H8&


'this section deals with assigning and killing a new cursor
Public Function CursorSwitch(ByVal Hwend As Long, ByVal Location As String)
'Hwend is the handle of the object for which to set a new cursor
'location is the location of the cursor file
Dim syscurshandle As Long, Curs1Handle As Long
    'Load a cursor from a file
    Curs1Handle = LoadCursorFromFile(Location)
    CursorSwitch = SetClassWord(Hwend, GCW_HCURSOR, Curs1Handle)
End Function

Public Sub KillCursor(Hwend As Long, ByVal SysCurs As Long, ByVal CurCursor As Long)
'resets default cursor,hwend is the hanld eof the object to reset the cursor on, syscurs is the original cursor, curcursor is the current cursor
Dim syscurshandle As Long
syscurshandle = SetClassWord(Hwend, GCW_HCURSOR, SysCurs)
Call DestroyCursor(CurCursor)
End Sub
'the next section of subs deal with true type font info
Public Sub TTF(filename$, List1 As ListBox, mode%)
Dim fileid%
Dim OT As OffsetTable
Dim TD As TableDirectoryEntry
Dim NT As NamingTable
Dim NR As NameRecord
Dim n&, ntStart&, CurrentLoc&, nrnum&, nPlatformID&, stroffset&, z&
Dim nameInfo() As Byte

List1.Clear
fileid% = FreeFile
Open filename$ For Binary Access Read As #fileid%
Get #fileid%, , OT
List1.AddItem "Version " & SwapInteger(OT.sfntVersionA) & "." & SwapInteger(OT.sfntVersionB)
List1.AddItem "Tables " & SwapInteger(OT.numTables)
Select Case mode
    Case 0
        For n& = 1 To SwapInteger(OT.numTables)
            Get #fileid%, , TD
            List1.AddItem TagToString(TD)
        Next n
    Case 1
        For n& = 1 To SwapInteger(OT.numTables)
            Get #fileid%, , TD
            If TagToString(TD) = "name" Then
                ntStart& = SwapLong(TD.offset) + 1
                Seek #fileid%, ntStart&
                Get #fileid, , NT
                List1.AddItem "NumRecords" & SwapInteger(NT.NameRecords)
                    For nrnum& = 1 To SwapInteger(NT.NameRecords)
                    Get #fileid%, , NR
                    List1.AddItem "Namerecord # " & nrnum&
                    List1.AddItem " PlatformID: " & GetPlatformID(SwapInteger(NR.PlatformID))
                    List1.AddItem " NameID: " & GetNameForID(SwapInteger(NR.NameID))
                    nPlatformID = SwapInteger(NR.PlatformID)
                    If (nPlatformID = 1 Or nPlatformID = 3) And NR.StringLength <> 0 Then
                        CurrentLoc = Seek(fileid%)
                        ReDim nameInfo(SwapInteger(NR.StringLength) - 1)
                        stroffset = 6 + SwapInteger(NT.NameRecords) * 12
                        Get #fileid%, ntStart + stroffset + SwapInteger(NR.StringOffset), nameInfo
                        Select Case nPlatformID
                            Case 1
                                List1.AddItem " " & StrConv(nameInfo, vbUnicode)
                            Case 3
                                SwapArray nameInfo()
                                List1.AddItem " " & CStr(nameInfo)
                        End Select
                        Seek #fileid%, CurrentLoc
                    Else
                        List1.AddItem "-Get a better system that this prog supports"
                    End If
                    Next nrnum&
                    
                    Exit For
            End If
        Next n&
End Select

Close #fileid%
End Sub

Public Function SwapInteger(ByVal i As Long) As Long
    SwapInteger = ((i \ &H100) And &HFF) Or ((i And &HFF) * &H100&)
End Function

Public Function SwapLong(ByVal L As Long) As Long
Dim addbit%
Dim newlow&, newhigh&

newlow& = L \ &H10000
newlow& = SwapInteger(newlow& And &HFFFF&)

newhigh& = SwapInteger(L And &HFFFF&)
If newhigh& And &H8000& Then
    newhigh& = newhigh And &H7FFF
    addbit% = True
End If

newhigh& = (newhigh& * &H10000) Or newlow&
If addbit% Then newhigh = newhigh Or &H80000000
SwapLong = newhigh&

End Function

Public Sub SwapArray(namearray() As Byte)
Dim u%, p%
Dim b As Byte
u% = UBound(namearray)
For p = 0 To u - 1 Step 2
    b = namearray(p) 'create placeholder
    namearray(p) = namearray(p + 1) 'make p = to the next one up
    namearray(p + 1) = b 'make the next one up = to p's original value, thus they are swapped
Next p
End Sub

Public Function TagToString(TD As TableDirectoryEntry)
Dim tagstr As String * 4
Dim x%
For x% = 1 To 4
    Mid(tagstr, x%, 1) = Chr$(TD.tag(x% - 1))
Next x%
TagToString = tagstr
End Function

Public Function GetPlatformID(ByVal id As Long) As String
Dim s$
Select Case id
    Case 0
    s$ = "Apple Unicode"
    Case 1
    s$ = "Macintosh"
    Case 2
    s$ = "ISO"
    Case 3
    s$ = "Microsoft"
End Select
GetPlatformID = s$
End Function

Public Function GetNameForID(ByVal id As Long) As String
Dim s$
Select Case id
    Case 0
    s$ = "CopyRight"
    Case 1
    s$ = "Font Family"
    Case 2
    s$ = "Font Subfamily"
    Case 3
    s$ = "Font Identifier"
    Case 4
    s$ = "Full Font Name"
    Case 5
    s$ = "Version"
    Case 6
    s$ = "Postscript Version"
    Case 7
    s$ = "Trademark"
    Case Else
    s$ = "UnKnown"
End Select
GetNameForID = s$
End Function

Public Function Callback1_EnumWindows(ByVal hwnd As Long, ByVal lpdata As Long, List1 As ListBox) As Long
    If GetParent(hwnd) = lpdata Then
        List1.AddItem GetWindowDesc$(hwnd)
    End If
Callback1_EnumWindows = 1
End Function

Public Function GetWindowDesc(hwnd As Long) As String
Dim desc$
Dim tbuf$
Dim inst&
Dim dl&
Dim hwndProcess&
desc$ = Str$(hwnd&) + Chr$(9) 'get string equiv of hwnd
tbuf$ = String$(256, 0) 'create buffer
dl& = GetWindowThreadProcessId(hwnd, hwndProcess)
    If hwndProcess = GetCurrentProcessId() Then
        inst& = GetWindowLong(hwnd&, GWL_HINSTANCE)
        dl& = GetModuleFileName(inst&, tbuf$, 255)
        tbuf$ = GetBaseName(tbuf$)
        If InStr(tbuf$, Chr$(0)) Then tbuf$ = Left$(tbuf$, InStr(tbuf$, Chr$(0)) - 1)
    Else
        tbuf$ = "Foreign Window"
    End If
desc$ = desc$ + tbuf$ + Chr$(9)
tbuf$ = String$(256, 0)
dl& = GetClassName(hwnd&, tbuf$, 255)
If InStr(tbuf$, Chr$(0)) Then tbuf$ = Left$(tbuf$, InStr(tbuf$, Chr$(0)) - 1)
desc$ = desc$ + tbuf$
GetWindowDesc$ = desc$
End Function
 
Public Function GetBaseName(ByVal Source As String) As String
Do While InStr(Source$, "\") <> 0
    Source$ = Mid$(Source$, InStr(Source$, "\") + 1) 'trim off path to the window
Loop
If InStr(Source$, ":") <> 0 Then
    Source$ = Mid$(Source$, InStr(Source$, ":") + 1) '?
End If
GetBaseName$ = Source$
End Function

Public Function GetHwnd(title$) As String
    Dim p%
    p% = InStr(title$, Chr$(9))
    If p% > 0 Then GetHwnd$ = Val(Left$(title$, p% - 1))
End Function
