Attribute VB_Name = "modShowList"
'This Code was written by Dave Andrews
'Feel free to use or modify this module freely
'Special thanks to Joseph Huntley for the skeleton of API forms.

Option Explicit
Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function defWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type


Private Type POINTAPI
    x As Long
    y As Long
End Type


Private Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'Listbox Constants
Const LBS_EXTENDEDSEL = &H800&
Const LBS_SORT = &H2&
Const LBS_NOINTEGRALHEIGHT = &H100&
Const LB_ADDSTRING = &H180
Const LB_GETSELITEMS = &H191
Const LB_GETTEXT = &H189
Const LB_GETTEXTLEN = &H18A
Const LB_GETSELCOUNT = &H190
Const LB_GETSEL = &H187
'------Button Constants
Const BS_USERBUTTON = &H8&
Const BS_CENTER = 768
Const BS_PUSHBUTTON = &H0&
Const BS_AUTORADIOBUTTON = &H9&
Const BS_PUSHLIKE = &H1000&
Const BS_LEFTTEXT = &H20&
Const BM_SETSTATE = &HF3
Const BM_GETSTATE = &HF2
Const BM_SETCHECK = &HF1
Const BM_GETCHECK = &HF0
'-----------Window Style Constants
Const WS_BORDER = &H800000
Const WS_CHILD = &H40000000
Const WS_OVERLAPPED = &H0&
Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
Const WS_SYSMENU = &H80000
Const WS_THICKFRAME = &H40000
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Const WS_VISIBLE = &H10000000
Const WS_POPUP = &H80000000
Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Const WS_VSCROLL = &H200000
Const WS_EX_TOOLWINDOW = &H80
Const WS_EX_TOPMOST = &H8&
Const WS_EX_CLIENTEDGE = &H200&
Const WS_EX_WINDOWEDGE = &H100&
Const WS_SIZEBOX = &H40000
Public Const WS_EX_DLGMODALFRAME = &H1&
'-----------Window Messaging Constants
Const WM_DESTROY = &H2
Const WM_CLOSE = &H10
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_CTLCOLOREDIT = &H133
Const WM_COMMAND = &H111
Const WM_GETTEXT = &HD
Const WM_ENABLE = &HA
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101
Const WM_SETTEXT = &HC
Const WM_VSCROLL = &H115
Const WM_MOVE = &H3
Const WM_SIZE = &H5
'--------Window Heiarchy Constants
Const GWL_WNDPROC = (-4)
Const GW_CHILD = 5
Const GW_OWNER = 4
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const SW_SHOWNORMAL = 1
'----------Misc Constants
Const CS_VREDRAW = &H1
Const CS_HREDRAW = &H2
Const CW_USEDEFAULT = &H80000000
Const COLOR_WINDOW = 5
Const SET_BACKGROUND_COLOR = 4103
Const IDC_ARROW = 32512&
Const IDI_APPLICATION = 32512&
Const MB_OK = &H0&
Const MB_ICONEXCLAMATION = &H30&

Dim MyMousePos As POINTAPI 'for getting the mouse positioning

Const gClassName = "Listbox API"

Dim gAppTitle As String


Dim gHwnd As Long

Dim gListHwnd As Long
Dim gListOldProc As Long

Dim gOKHwnd As Long
Dim gOKOldProc As Long
Dim gCancelHwnd As Long
Dim gCancelOldProc As Long

Dim ListStyle As Long
Dim CurSel() As Variant
Dim inList() As Variant
Dim isSelected As Boolean
Dim wTop As Long
Dim wLeft As Long
Dim wHeight As Long
Dim wWidth As Long
Dim Created As Boolean
Sub CopyArray(Source() As Variant, ByRef Dest() As Variant)
On Error GoTo eTrap:
Dim i As Integer
ReDim Dest(UBound(Source))
For i = 0 To UBound(Source)
    Dest(i) = Source(i)
Next i
eTrap:
End Sub

Sub Main()
    Randomize
    Dim inList() As Variant
    Dim outList() As Variant
    Dim i As Integer
    Dim j As Integer
    ReDim inList(29)
    'Create a list of  "words"
    For i = 0 To 29
        For j = 1 To CInt(Rnd * 20) + 1
            inList(i) = inList(i) & Chr(CInt(Rnd * 26) + 65)
        Next j
    Next i
    'Get our selection
    If ShowList(inList(), outList(), True, True, "List Test", 0, 0, 250, 300) Then
        'output our selection
        For i = 0 To UBound(outList)
            MsgBox outList(i)
        Next i
    End If
    End
End Sub

Function EditClass() As WNDCLASS
EditClass.hbrBackground = vbRed
End Function


Sub MakeSelection()
Dim i As Integer
Dim tLen As Long
Dim tItem As String
Dim sCount As Integer
For i = 0 To UBound(inList)
    If SendMessage(gListHwnd, LB_GETSEL, i, 0&) <> 0 Then
        ReDim Preserve CurSel(sCount)
        tLen = SendMessage(gListHwnd&, LB_GETTEXTLEN, i, 0&)
        tItem = Space(tLen)
        Call SendMessage(gListHwnd&, LB_GETTEXT, i, ByVal tItem)
        CurSel(sCount) = tItem
        sCount = sCount + 1
        isSelected = True
    End If
Next i
End Sub

 Function ShowList(InputList() As Variant, ByRef SelectionList() As Variant, Optional MultiSelect As Boolean, Optional Sorted As Boolean, Optional Title As String, Optional Left As Long, Optional Top As Long, Optional Width As Long = 150, Optional Height As Long = 200) As Boolean
    Call GetCursorPos(MyMousePos)
    If IsMissing(wLeft) Then wLeft = MyMousePos.x
    If IsMissing(wTop) Then wTop = MyMousePos.y
    wWidth = Width
    wHeight = Height
    CopyArray InputList(), inList()
    If Title <> "" Then gAppTitle$ = Title Else gAppTitle$ = "Make A Selection"
    Dim wMsg As Msg
    Dim tSec As String
    If Sorted Then ListStyle = ListStyle Or LBS_SORT
    If MultiSelect Then ListStyle = ListStyle Or LBS_EXTENDEDSEL
    ''Call procedure to register window classname. If false, then exit.
    If RegisterWindowClass = False Then Exit Function
    
      ''Create window
      If CreateWindows() Then
         ''Loop will exit when WM_QUIT is sent to the window.
         Do While GetMessage(wMsg, 0&, 0&, 0&)
            ''TranslateMessage takes keyboard messages and converts
            ''them to WM_CHAR for easier processing.
            Call TranslateMessage(wMsg)
            ''Dispatchmessage calls the default window procedure
            ''to process the window message. (WndProc)
            Call DispatchMessage(wMsg)
            DoEvents
         Loop
      End If
    
    Call UnregisterClass(gClassName$, App.hInstance)
    If isSelected Then CopyArray CurSel(), SelectionList()
    ShowList = isSelected
End Function


 Function RegisterWindowClass() As Boolean

    Dim wc As WNDCLASS
    
    ''Registers our new window with windows so we can use our classname.
    
    wc.style = CS_HREDRAW Or CS_VREDRAW
    wc.lpfnwndproc = GetAddress(AddressOf WndProc) ''Address in memory of default window procedure.
    wc.hInstance = App.hInstance
    wc.hIcon = LoadIcon(0&, IDI_APPLICATION) ''Default application icon
    wc.hCursor = LoadCursor(0&, IDC_ARROW) ''Default arrow
    wc.hbrBackground = COLOR_WINDOW ''Default a color for window.
    wc.lpszClassName = gClassName$

    RegisterWindowClass = RegisterClass(wc) <> 0
    
End Function
 Function CreateWindows() As Boolean
    Dim i As Integer
    Dim tStr As String
    Dim ButtonStyle As Long
    ButtonStyle = WS_CHILD Or WS_VISIBLE Or WS_BORDER
    ListStyle = ListStyle Or WS_CHILD Or WS_VISIBLE Or WS_BORDER Or WS_VSCROLL Or LBS_NOINTEGRALHEIGHT
    'Create form window.
    gHwnd& = CreateWindowEx(WS_EX_TOOLWINDOW Or WS_EX_TOPMOST, gClassName$, gAppTitle$, WS_POPUPWINDOW Or WS_CAPTION Or WS_VISIBLE Or WS_SIZEBOX, wLeft, wTop, wWidth, wHeight, 0&, 0&, App.hInstance, ByVal 0&)
    'Create List Box
    gListHwnd& = CreateWindowEx(0&, "LISTBOX", "", ListStyle, 1, 1, wWidth - 9, wHeight - 47, gHwnd&, 0&, App.hInstance, 0&)
     'Create OK and Cancel Buttons
    gOKHwnd = CreateWindowEx(0&, "BUTTON", "OK", ButtonStyle, 1, wHeight - 44, (wWidth - 9) / 2, 20, gHwnd&, 0&, App.hInstance, 0&)
    gCancelHwnd = CreateWindowEx(0&, "BUTTON", "CANCEL", ButtonStyle, (wWidth - 4) / 2, wHeight - 44, (wWidth - 9) / 2, 20, gHwnd&, 0&, App.hInstance, 0&)
    For i = 0 To UBound(inList)
        tStr$ = CStr(inList(i))
        SendMessage gListHwnd&, LB_ADDSTRING, 0&, ByVal tStr$
    Next i
    
    '-------Hook OK CANCEL-----------
    gOKOldProc& = GetWindowLong(gOKHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gOKHwnd&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
    gCancelOldProc& = GetWindowLong(gCancelHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gCancelHwnd&, GWL_WNDPROC, GetAddress(AddressOf CancelWndProc))
    
    
    CreateWindows = (gHwnd& <> 0)
    Created = True
End Function
Function OKWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_LBUTTONDOWN
            MakeSelection
            Call SendMessage(gHwnd, WM_CLOSE, 0&, 0&)
    End Select
    
  OKWndProc = CallWindowProc(gOKOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function
Function CancelWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_LBUTTONDOWN
            isSelected = False
            Call SendMessage(gHwnd, WM_CLOSE, 0&, 0&)
    End Select
    
  CancelWndProc = CallWindowProc(gCancelOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function

Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ''This our default window procedure for the window. It will handle all
    ''of our incoming window messages and we will write code based on the
    ''window message what the program should do.
    Dim i As Integer
      Select Case uMsg&
         Case WM_DESTROY:
            ''Since DefWindowProc doesn't automatically call
            ''PostQuitMessage (WM_QUIT). We need to do it ourselves.
            ''You can use DestroyWindow to get rid of the window manually.
            'SetDate
            Call PostQuitMessage(0&)
        Case WM_SIZE
            If Not Created Then Exit Function
            Dim wSize As RECT
            GetWindowRect gHwnd&, wSize
            wLeft = wSize.Left
            wTop = wSize.Top
            wWidth = wSize.Right - wSize.Left
            wHeight = wSize.Bottom - wSize.Top
            MoveWindow gListHwnd&, 1, 1, wWidth - 9, wHeight - 47, True
            MoveWindow gOKHwnd&, 1, wHeight - 44, (wWidth - 9) / 2, 20, True
            MoveWindow gCancelHwnd&, (wWidth - 4) / 2, wHeight - 44, (wWidth - 9) / 2, 20, True
            'gListHwnd& = CreateWindowEx(0&, "LISTBOX", "", ListStyle, 1, 1, wWidth - 9, wHeight - 47, gHwnd&, 0&, App.hInstance, 0&)
            'gOKHwnd = CreateWindowEx(0&, "BUTTON", "OK", ButtonStyle, 1, wHeight - 44, (wWidth - 9) / 2, 20, gHwnd&, 0&, App.hInstance, 0&)
            'gCancelHwnd = CreateWindowEx(0&, "BUTTON", "CANCEL", ButtonStyle, (wWidth - 4) / 2, wHeight - 44, (wWidth - 9) / 2, 20, gHwnd&, 0&, App.hInstance, 0&)

      End Select
    ''Let windows call the default window procedure since we're done.
    WndProc = defWindowProc(hwnd&, uMsg&, wParam&, lParam&)

End Function
 

 Function GetAddress(ByVal lngAddr As Long) As Long
    ''Used with AddressOf to return the address in memory of a procedure.

    GetAddress = lngAddr&
    
End Function



