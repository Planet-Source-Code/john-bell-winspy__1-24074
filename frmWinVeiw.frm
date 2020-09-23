VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWinVeiw 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WinSpy 2001"
   ClientHeight    =   3390
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtDisplay 
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      TextRTF         =   $"frmWinVeiw.frx":0000
   End
   Begin MSComctlLib.TreeView trevHwnd 
      Height          =   2175
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3836
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   58
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Parent"
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Flash"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "WinStyles"
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "ClassInfo"
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Size"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Position"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label lblAction 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   6735
   End
   Begin VB.Menu mnuLoad 
      Caption         =   "Load List"
      Begin VB.Menu mnuLoadList 
         Caption         =   "TopLevel"
         Index           =   0
      End
      Begin VB.Menu mnuLoadList 
         Caption         =   "Children"
         Index           =   1
      End
      Begin VB.Menu mnuLoadList 
         Caption         =   "Siblings"
         Index           =   2
      End
      Begin VB.Menu mnuLoadList 
         Caption         =   "Parent"
         Index           =   3
      End
      Begin VB.Menu mnuLoadList 
         Caption         =   "SameClassSameLevel"
         Index           =   5
      End
      Begin VB.Menu mnuLoadList 
         Caption         =   "Clear"
         Index           =   6
      End
   End
End
Attribute VB_Name = "frmWinVeiw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'John Bell, 06/14/01
'WinSpy ver. 1.0,
'Goto my site at www.byllbo.com to see how my
'coding progresses, to get ver. 2.0, ect.
'I wrote this for some practice while reading chapter 5 of
'dan appleman's Guide to the win32 api...lol
Private prev As Integer
Option Explicit

Private Sub cmdOp_Click(Index As Integer)
Dim titlestring$
Dim usehwnd&
Dim dl&
Dim style&
Dim crlf$
Dim outstring$
Dim newhwnd&
Dim windowdesc$
Dim hwnd&
Dim WindowRect As RECT
Select Case Index
    Case 0
        crlf$ = vbCrLf
        If List1.ListIndex < 0 Then Exit Sub
        titlestring$ = List1.Text
        usehwnd& = GetHwnd(List1.Text)
        dl& = GetWindowRect(usehwnd&, WindowRect)
        If IsIconic&(usehwnd&) Then
            outstring$ = "Is Iconic" & crlf$
        End If
        If IsZoomed&(usehwnd&) Then
            outstring = outstring$ & "Is Zoomed" & crlf$
        End If
        If IsWindowEnabled&(usehwnd&) Then
            outstring$ = outstring$ & "Is Enabled" & crlf$
        Else
            outstring$ = outstring$ & "Is Disabled" & crlf$
        End If
        If IsWindowVisible&(usehwnd&) Then
             outstring$ = outstring$ & "Is Visible" & crlf$
        Else
            outstring$ = outstring$ & "Is Visible" & crlf$
        End If
        outstring$ = outstring$ & "Rect: " & Str$(WindowRect.Left) & "," & Str$(WindowRect.Top) & "," & Str$(WindowRect.Right) & "," & Str$(WindowRect.Bottom)
        Do While InStr(titlestring$, Chr(9)) <> 0 'trim's off the class name
            titlestring$ = Mid$(titlestring$, InStr(titlestring$, Chr$(9)) + 1)
        Loop
        txtDisplay.Text = outstring$ & crlf$ & "Class: " & titlestring$
    Case 1 'size
        crlf$ = vbCrLf 'carriage return code
        If List1.ListIndex < 0 Then Exit Sub
        titlestring$ = List1.Text 'get class info from list
        usehwnd = GetHwnd(titlestring$) 'get hwnd from info
        dl& = GetClientRect(usehwnd&, WindowRect) 'get client(visible) dimensions
        outstring$ = "Horiz Pixels: " & (WindowRect.Right) & " " & crlf$ & "Vert Pixels: " & (WindowRect.Bottom)  'trim off what we want
        Do While InStr(titlestring$, Chr(9)) <> 0 'trim's off the class name
            titlestring$ = Mid$(titlestring$, InStr(titlestring$, Chr$(9)) + 1)
        Loop
        txtDisplay.Text = outstring$ & crlf$ & "Class: " & titlestring$ 'show user
    Case 2 'classinfo
    Dim clsExtra&, wndExtra&
        crlf$ = vbCrLf
        If List1.ListIndex < 0 Then Exit Sub
        titlestring$ = List1.Text
        usehwnd& = GetHwnd(titlestring$)
        'get class info
        clsExtra& = GetClassLong(usehwnd&, GCL_CBCLSEXTRA)
        wndExtra& = GetClassLong(usehwnd&, GCL_CBWNDEXTRA)
        style& = GetClassLong(usehwnd&, GCL_STYLE) 'style being the style bits....
        outstring$ = "Class & Word Extra: " & Str$(clsExtra) & Str$(wndExtra) & crlf$
        If style& And CS_BYTEALIGNCLIENT Then
        outstring$ = outstring$ & "CS_BYTEALIGNCLIENT" & crlf$
        End If
        If style& And CS_BYTEALIGNWINDOW Then
        outstring$ = outstring$ & "CS_BYTEALIGNWINDOW" & crlf$
        End If
        If style& And CS_CLASSDC Then
        outstring$ = outstring$ & "CS_CLASSDC" & crlf$
        End If
        If style& And CS_DBLCLKS Then
            outstring$ = outstring$ & "CS_DBLCLKS" & crlf$
        End If
        ' Was CS_GLOBALCLASS (has same value)
        If style& And CS_PUBLICCLASS Then
        outstring$ = outstring$ & "CS_GLOBALCLASS" & crlf$
        End If
        If style& And CS_HREDRAW Then
        outstring$ = outstring$ & "CS_HREDRAW" & crlf$
        End If
        If style& And CS_NOCLOSE Then
        outstring$ = outstring$ & "CS_NOCLOSE" & crlf$
        End If
        If style& And CS_OWNDC Then
        outstring$ = outstring$ & "CS_OWNDC" & crlf$
        End If
        If style& And CS_PARENTDC Then
        outstring$ = outstring$ & "CS_PARENTDC" & crlf$
        End If
        If style& And CS_SAVEBITS Then
        outstring$ = outstring$ & "CS_SAVEBITS" & crlf$
        End If
        If style& And CS_VREDRAW Then
        outstring$ = outstring$ & "CS_VREDRAW" & crlf$
        End If
        Do While InStr(titlestring$, Chr(9)) <> 0 'trim's off the class name
            titlestring$ = Mid$(titlestring$, InStr(titlestring$, Chr$(9)) + 1)
        Loop
        txtDisplay.Text = outstring$ & crlf$ & "Class: " & titlestring$
    Case 3
        crlf$ = vbCrLf
        If List1.ListIndex < 0 Then Exit Sub
        titlestring$ = List1.Text
        usehwnd& = GetHwnd(titlestring$)
        style& = GetWindowLong(usehwnd, GWL_STYLE)
        If style& And WS_BORDER Then
        outstring$ = outstring$ + "WS_BORDER" + crlf$
        End If
        If style& And WS_CAPTION Then
        outstring$ = outstring$ + "WS_CAPTION" + crlf$
        End If
        If style& And WS_CHILD Then
        outstring$ = outstring$ + "WS_CHILD" + crlf$
        End If
        If style& And WS_CLIPCHILDREN Then
        outstring$ = outstring$ + "WS_CLIPCHILDREN" + crlf$
        End If
        If style& And WS_CLIPSIBLINGS Then
        outstring$ = outstring$ + "WS_CLIPSIBLINGS" + crlf$
        End If
        If style& And WS_DISABLED Then
        outstring$ = outstring$ + "WS_DISABLED" + crlf$
        End If
        If style& And WS_DLGFRAME Then
        outstring$ = outstring$ + "WS_DLGFRAME" + crlf$
        End If
        If style& And WS_GROUP Then
        outstring$ = outstring$ + "WS_GROUP" + crlf$
        End If
        If style& And WS_HSCROLL Then
        outstring$ = outstring$ + "WS_HSCROLL" + crlf$
        End If
        If style& And WS_MAXIMIZE Then
        outstring$ = outstring$ + "WS_MAXIMIZE" + crlf$
        End If
        If style& And WS_MAXIMIZEBOX Then
        outstring$ = outstring$ + "WS_MAXIMIZEBOX" + crlf$
        End If
        If style& And WS_MINIMIZE Then
        outstring$ = outstring$ + "WS_MINIMIZE" + crlf$
        End If
        If style& And WS_MINIMIZEBOX Then
        outstring$ = outstring$ + "WS_MINIMIZEBOX" + crlf$
        End If
        If style& And WS_POPUP Then
        outstring$ = outstring$ + "WS_POPUP" + crlf$
        End If
        If style& And WS_SYSMENU Then
        outstring$ = outstring$ + "WS_SYSMENU" + crlf$
        End If
        If style& And WS_TABSTOP Then
        outstring$ = outstring$ + "WS_TABSTOP" + crlf$
        End If
        If style& And WS_THICKFRAME Then
        outstring$ = outstring$ + "WS_THICKFRAME" + crlf$
        End If
        If style& And WS_VISIBLE Then
        outstring$ = outstring$ + "WS_VISIBLE" + crlf$
        End If
        If style& And WS_VSCROLL Then
        outstring$ = outstring$ + "WS_VSCROLL" + crlf$
        End If
     
        ' Note: We could tap the style& variable for class
        ' styles as well (especially since it is easy to
        ' determine the class for a window), but that is
        ' beyond the scope of this sample program.
        Do While InStr(titlestring$, Chr(9)) <> 0 'trim's off the class name
            titlestring$ = Mid$(titlestring$, InStr(titlestring$, Chr$(9)) + 1)
        Loop
        txtDisplay = outstring$ & crlf$ & "Class: " & titlestring$
    Case 4
        If List1.ListIndex < 0 Then Exit Sub
        titlestring$ = List1.Text
        usehwnd& = GetHwnd(titlestring$)
        dl& = FlashWindow(usehwnd, -1)
        Sleep 500 'my own additions so you actually notice the flash
        dl& = FlashWindow(usehwnd, 0) 'its kinda neat
    Case 5
        If List1.ListIndex < 0 Then Exit Sub
        titlestring$ = List1.Text
        hwnd& = GetHwnd(List1.Text)
        newhwnd& = GetParent(hwnd&)
        If newhwnd& = 0 Then
            lblAction.Caption = "No Prarent"
            Exit Sub
        End If
        DoEvents
        windowdesc$ = GetWindowDesc$(newhwnd&)
        Do While InStr(titlestring$, Chr(9)) <> 0 'trim's off the class name
            titlestring$ = Mid$(titlestring$, InStr(titlestring$, Chr$(9)) + 1)
        Loop
        txtDisplay = windowdesc$ & crlf$ & "Parent of Class " & titlestring$ & " member " & "&H" & Hex$(hwnd) & " is"
    End Select
End Sub

Private Sub cmdSearch_Click()
End Sub

Private Sub Form_Activate()
SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Private Sub Form_Load()
Dim tabsets&(2)
Dim dl&
    tabsets(0) = 45
    tabsets(1) = 110
    dl& = SendMessageLongByRef&(List1.hwnd, LB_SETTABSTOPS, 2, tabsets(0))
txtDisplay.RightMargin = txtDisplay.Width * 1.5
End Sub

Private Sub List1_Click()
Dim s As Integer
Dim listS As String, trevS As String
listS = Val(List1.List(List1.ListIndex))
s = List1.ListIndex
s = s + 2
trevS = Val(trevHwnd.Nodes(s).Text)
If trevS = listS Then
    trevHwnd.Nodes(s).Selected = True
End If
End Sub

Private Sub mnuLoadList_Click(Index As Integer)
Dim hwnd&
Dim dl&
Dim windowdesc$, windowdesc2$
Dim phwnd&
Dim x As Integer
Select Case Index
    Case 0 'toplevel windows
        List1.Clear
        trevHwnd.Nodes.Clear
        x = 0
        hwnd& = GetDesktopWindow() 'find desktop
        trevHwnd.Nodes.Add
        trevHwnd.Nodes(1).Text = Str$(hwnd&)
        trevHwnd.Nodes(1).Key = "node" & Str$(x) 'first index has to be 1, not zero
        hwnd& = GetWindow(hwnd, GW_CHILD) 'first child is toplevel window
        Do: DoEvents
            x = x + 1
            trevHwnd.Nodes.Add "node" & Str$(0), tvwChild
            trevHwnd.Nodes(x + 1).Text = Str$(hwnd&) 'node index's will be two numbers off from the list index's
            trevHwnd.Nodes(x + 1).Key = "node" & Str$(x)
            List1.AddItem GetWindowDesc(hwnd&) 'add to list
            hwnd& = GetWindow(hwnd&, GW_HWNDNEXT) 'get next toplevel window
        Loop While hwnd& <> 0 'loop if we have not reached the end
        lblAction.Caption = "Top Level Windows"
    Case 1 'children  of parent window
       If List1.ListIndex < 0 Then Exit Sub
       windowdesc$ = List1.Text 'get clas info from list
       phwnd = GetHwnd(windowdesc$) 'code this
        x = 0
        If phwnd = 0 Then Exit Sub 'make sure hwnd not 0
        hwnd& = GetWindow(phwnd&, GW_CHILD)
        If hwnd = 0 Then Exit Sub 'make sure hwnd not 0
        trevHwnd.Nodes.Clear
        List1.Clear 'clear the list
        trevHwnd.Nodes.Add
        trevHwnd.Nodes(1).Text = Str$(phwnd&)
        trevHwnd.Nodes(1).Key = "node" & Str$(x) 'first index has to be 1, not zero
        Do: DoEvents
             x = x + 1
            trevHwnd.Nodes.Add "node" & Str$(0), tvwChild
            trevHwnd.Nodes(x + 1).Text = Str$(hwnd&) 'node index's will be two numbers off from the list index's
            trevHwnd.Nodes(x + 1).Key = "node" & Str$(x)
            List1.AddItem GetWindowDesc$(hwnd&) 'add first child
            hwnd& = GetWindow(hwnd&, GW_HWNDNEXT) 'get next child
        Loop While hwnd <> 0 'if there is one then loop
        Do While InStr(windowdesc$, Chr(9)) <> 0 'trim's off the class name
            windowdesc$ = Mid$(windowdesc$, InStr(windowdesc$, Chr$(9)) + 1)
        Loop
        lblAction.Caption = "Children of: " & windowdesc$
    Case 2 'siblings, yup it works, mine too
        If List1.ListIndex < 0 Then Exit Sub
        x = 0
        windowdesc$ = List1.Text 'get info from list
        hwnd& = GetHwnd(windowdesc$) 'get hwnd from info
        phwnd& = GetParent(hwnd&) 'getparent
        hwnd& = GetWindow(phwnd&, GW_CHILD) 'get first child of parent
        If hwnd& = 0 Then Exit Sub
        trevHwnd.Nodes.Clear
        List1.Clear 'clear the list
        trevHwnd.Nodes.Add
        trevHwnd.Nodes(1).Text = Str$(phwnd&)
        trevHwnd.Nodes(1).Key = "node" & Str$(x) 'first index has to be 1, not zero
        Do: DoEvents
            x = x + 1
            trevHwnd.Nodes.Add "node" & Str$(0), tvwChild
            trevHwnd.Nodes(x + 1).Text = Str$(hwnd&) 'node index's will be two numbers off from the list index's
            trevHwnd.Nodes(x + 1).Key = "node" & Str$(x)
            List1.AddItem GetWindowDesc(hwnd&) 'add to list
            hwnd& = GetWindow(hwnd&, GW_HWNDNEXT) 'get all children of class
        Loop While hwnd <> 0
        Do While InStr(windowdesc$, Chr(9)) <> 0 'trim's off the class name
            windowdesc$ = Mid$(windowdesc$, InStr(windowdesc$, Chr$(9)) + 1)
        Loop
        lblAction.Caption = "Siblings of: " & windowdesc$
    Case 3 'parent
        If List1.ListIndex < 0 Then Exit Sub
        x = 0
        windowdesc$ = List1.Text 'get info from list
        hwnd& = GetHwnd(windowdesc$) 'get hwnd from info
        hwnd& = GetParent(hwnd&) 'get parent of hwnd'
        phwnd& = GetParent(hwnd&) 'get the parents parent for the seek of the treeveiw
        If phwnd = 0 Then phwnd = GetDesktopWindow() 'if phwnd = 0 then we are at a toplevel so get the desktop as parent
        If hwnd = 0 Then Exit Sub 'if it has one otherwise exit sub
        List1.Clear 'clear list
        trevHwnd.Nodes.Clear 'clear nodes
        List1.AddItem GetWindowDesc(hwnd&) 'display parent
        trevHwnd.Nodes.Add 'add parent of parent node
        trevHwnd.Nodes(1).Text = Str$(phwnd&)
        trevHwnd.Nodes(1).Key = "node" & Str$(x)
        x = x + 1 'I added the extra parent to the treeveiw side so that on the list side i could still use the same code to
        trevHwnd.Nodes.Add "node" & Str$(0), tvwChild 'syncronize it to the treeveiw
        trevHwnd.Nodes(x + 1).Text = Str$(hwnd&) 'node index's will be two numbers off from the list index's
        trevHwnd.Nodes(x + 1).Key = "node" & Str$(x)
       DoEvents
        Do While InStr(windowdesc$, Chr(9)) <> 0 'trim's off the class name
            windowdesc$ = Mid$(windowdesc$, InStr(windowdesc$, Chr$(9)) + 1)
        Loop
        lblAction.Caption = "Parent of: " & windowdesc$ 'add caption
    Case 4 'owned windows, this does not work that i can tell
       If List1.ListIndex < 0 Then Exit Sub
       windowdesc$ = List1.Text
       hwnd& = GetHwnd(windowdesc$)
       List1.Clear
       dl& = EnumWindows(AddressOf Callback1_EnumWindows, hwnd) 'aint that some shit??lost me just now
       If List1.ListCount = 0 Then
            lblAction.Caption = "Error during processing"
            Exit Sub
        End If
        lblAction.Caption = "Owned Windows of: " & Mid$(windowdesc$, 22)
    Case 5 'sameclasssamelevel,mine by the way
        If List1.ListIndex < 0 Then Exit Sub
        x = 0
       windowdesc$ = List1.Text 'get class info form list box
       hwnd& = GetHwnd(windowdesc$) 'strip off hwnd form class info list
       phwnd& = GetParent(hwnd&)
       If phwnd& = 0 Then phwnd = GetDesktopWindow()
       Do While InStr(windowdesc$, Chr(9)) <> 0 'trim's off the class name
            windowdesc$ = Mid$(windowdesc$, InStr(windowdesc$, Chr$(9)) + 1)
        Loop
        windowdesc2$ = windowdesc$ 'initialize the search var
        trevHwnd.Nodes.Clear
        List1.Clear 'clear the list
        trevHwnd.Nodes.Add
        trevHwnd.Nodes(1).Text = Str$(phwnd&)
        trevHwnd.Nodes(1).Key = "node" & Str$(x)
        Do: DoEvents 'keep windows from locking up
            If windowdesc2$ = windowdesc$ Then 'if the search var matches the control var then add to list
                List1.AddItem GetWindowDesc(hwnd&)
                x = x + 1
                trevHwnd.Nodes.Add "node" & Str$(0), tvwChild
                trevHwnd.Nodes(x + 1).Text = Str$(hwnd&) 'node index's will be two numbers off from the list index's
                trevHwnd.Nodes(x + 1).Key = "node" & Str$(x)
            End If
            hwnd& = GetWindow(hwnd&, GW_HWNDNEXT) 'get next hwnd at this level
            windowdesc2$ = GetWindowDesc(hwnd&) 'get its class info
            Do While InStr(windowdesc2$, Chr(9)) <> 0 'trim's off its class name
                windowdesc2$ = Mid$(windowdesc2$, InStr(windowdesc2$, Chr$(9)) + 1)
            Loop
        Loop While hwnd& <> 0 'loop and check if its a match
        lblAction.Caption = "All Windows of Class: " & windowdesc$
    Case 6
        List1.Clear
        trevHwnd.Nodes.Clear
        txtDisplay.Text = vbNullString
End Select

End Sub

Private Sub trevHwnd_NodeClick(ByVal Node As MSComctlLib.Node)
Dim trevS As String, listS As String
Dim intIndex As Integer
Dim blnFound As Boolean
intIndex = 0
trevS = Val(Node.Text)
Do While Not blnFound And intIndex < List1.ListCount
    listS = Val(List1.List(intIndex))
    If listS = trevS Then
        List1.ListIndex = intIndex
        blnFound = True
    End If
    intIndex = intIndex + 1
Loop
End Sub

