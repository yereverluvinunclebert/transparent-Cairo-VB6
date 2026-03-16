Attribute VB_Name = "modWindowAPI"
'---------------------------------------------------------------------------------------
' Module    : RealWindowMod
' Author    : Andrew Heinlein [Mouse]
' Date      : 11/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

' --- Win32 API Declarations ---

Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' --- Win32 Type Declarations ---
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(32) As Byte
End Type

Private Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

' --- MSG structure ---
Private Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

' --- Constants ---
Private Const CS_HREDRAW = &H2
Private Const CS_VREDRAW = &H1
Private Const CS_PARENTDC = &H80
Private Const WS_OVERLAPPEDWINDOW = &HCF0000
Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const IDC_ARROW = &H7F00
Private Const COLOR_WINDOW = &H5
Private Const SW_SHOW = &H5
Private Const WM_DESTROY = &H2
Private Const WM_PAINT = &HF
Private Const DT_CENTER = &H1
Private Const CW_USEDEFAULT = &H80000000

Private Const WS_CHILD = &H40000000
Private Const SW_NORMAL = 1
Private Const GWL_WNDPROC = (-4)
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const MB_OK = &H0&
Private Const MB_ICONEXCLAMATION = &H30&
  
' vars for the APIs to get/set Window characteristics , opacity &c
'Private Const ULW_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Integer = &H20
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE As Long = -16
'Private Const LWA_COLORKEY = &H1

Private gButOldProc As Long
Private gHwnd As Long
Public hWindowHwnd As Long
Public hVBFormHwnd As Long


'---------------------------------------------------------------------------------------
' Procedure : MainWndProc
' Author    : beededea
' Date      : 11/03/2026
' Purpose   : Custom Window Procedure (callback)
'---------------------------------------------------------------------------------------
'
Private Function MainWndProc(ByVal hwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim udtRect As RECT
    Dim hDC As Long
    Dim ps As PAINTSTRUCT
    Dim thisText As String
        
    On Error GoTo MainWndProc_Error

    If message = WM_PAINT Then

        ' Retrieves the coordinates of the client area
        GetClientRect hwnd, udtRect
        
        ' Get a GDI handle to a device context
        hDC = BeginPaint(hwnd, ps) ' required for WM_PAINT handling
        
        ' set the position and size of the user defined text rectangle
        udtRect.Left = 150
        udtRect.Top = 105
        udtRect.Bottom = 130
        udtRect.Right = 410
        
        thisText = "Hello from subclassed WM_PAINT"
        
        ' draw the text at the location, centered
        DrawText hDC, thisText, Len(thisText), udtRect, DT_CENTER
        EndPaint hwnd, ps
        
        ' now we paint the image using Cairo, Cairo HAS to load from file as the process to get Cairo to load from a collection is rather tricky using VB6 (Cairo requires a callback as input)
        Call drawAlphaPngCairo(thisHDC, hWindowHwnd, App.Path & "\tardis.png", 50, 350)
        
        ' now we paint the image using GDI+ extracting the image from a previously loaded dictionary, in this case Christian Buse's VBA dictionary replacement
        Call drawAlphaPngGDIP(250, 250, 150, 150)
        
        ' since we have handled this message, return 0 and exit the function to prevent the DefWindowProc from handling it.
        MainWndProc = 0
        Exit Function
    End If

    ' watch for WM_DESTROY message, if it is sent, then let the GetMessage loop in createAPIWindow know so it breaks out of the GetMessage loop
    If message = WM_DESTROY Then
        PostQuitMessage 0
        MainWndProc = 0
        Call thisForm_Unload
        Exit Function
    End If
        
     ' Default handling - return value and ensure that every message is processed
     MainWndProc = DefWindowProc(hwnd, message, wParam, lParam)

    On Error GoTo 0
    Exit Function

MainWndProc_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MainWndProc of Module RealWindowMod"
End Function

'---------------------------------------------------------------------------------------
' Procedure : createAPIWindow
' Author    : beededea
' Date      : 11/03/2026
' Purpose   : Entry Point: Create and run the window
'---------------------------------------------------------------------------------------
'
Private Function createAPIWindow(ByVal MyWndProc As Long, ByVal szWindowClass As String, ByVal szWindowTitle As String, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long) As Long
    
    Dim wcex As WNDCLASSEX

    Dim myMsg As Msg
    Dim gButtonHwnd As Long
    
    Dim windowLngReturn As Long
    
    On Error GoTo createAPIWindow_Error

    ' define a WNDCLASSEX class properties
    wcex.cbSize = LenB(wcex)
    wcex.style = CS_HREDRAW Or CS_VREDRAW Or CS_PARENTDC
    wcex.lpfnWndProc = MyWndProc ' Long Pointer to the Windows Procedure function that will be called
    wcex.cbClsExtra = 0
    wcex.cbWndExtra = 0
    wcex.hInstance = App.hInstance
    wcex.hIcon = 0
    wcex.hCursor = LoadCursor(0, IDC_ARROW)
    wcex.hbrBackground = COLOR_WINDOW + 1
    wcex.lpszMenuName = vbNullString
    wcex.lpszClassName = szWindowClass
    wcex.hIconSm = 0
    
    ' registers a window class for subsequent use in calls to the CreateWindowEx API
    If RegisterClassEx(wcex) = 0 Then
        MsgBox "Failed to register window!"
        createAPIWindow = -1
        Exit Function
    End If
    
    ' create the window of the class above
    hWindowHwnd = CreateWindowEx(WS_EX_APPWINDOW Or WS_EX_WINDOWEDGE, _
                              szWindowClass, _
                              szWindowTitle, _
                              WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or WS_OVERLAPPEDWINDOW, _
                              x, y, cx, cy, 0, 0, App.hInstance, 0)
                              
    If hWindowHwnd = 0 Then
        MsgBox "Failed to create the window!"
        UnregisterClass szWindowClass, App.hInstance
        createAPIWindow = -1
        Exit Function
    End If
    
    ' remove all titlebar and associated controls
    ' SetWindowLong hWindowHwnd, GWL_STYLE, 0
    
    'set the transparency of the underlying form with full click through, makes the form completely transparent, the created button will not be clickable as it will not be visible as per Steamydock with GDI+
    'SetWindowLong hWindowHwnd, GWL_EXSTYLE, GetWindowLong(hWindowHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    
    ' the addition of "Or WS_EX_TRANSPARENT" to SetWindowLong will make the transparent form fully click-through but ALL controls will be unresponsive, even the titlebar controls.
    'SetWindowLong hWindowHwnd, GWL_EXSTYLE, GetWindowLong(hWindowHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
    
    ' Subclass the window (optional, here we're already using our own WndProc)
'    lpPrevWndProc = GetWindowLong(hWindowHwnd, GWL_WNDPROC)
'    SetWindowLong hWindowHwnd, GWL_WNDPROC, AddressOf WndProc
     
    ' this brings the form back again and uses the colour key cyan to make the form and any other similar items appear transparent
    'Me.BackColor = vbCyan ' sets the VB6 form to the transparent key colour
    'SetLayeredWindowAttributes hWindowHwnd, vbCyan, 0&, LWA_COLORKEY
    
    
'             PAINTSTRUCT ps;
'         RECT rc;
'         HDC hdc = BeginPaint(hwnd, &ps);
'         GetClientRect(hwnd, &rc);
'         SetBkColor(hdc, 0x000000ff); // red
'         ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &rc, 0, 0, 0);
'         EndPaint(hwnd, &ps);

    'Create a  button
    gButtonHwnd = CreateWindowEx(ByVal 0&, "BUTTON", "OK", WS_CHILD, 58, 90, 100, 50, hWindowHwnd, ByVal 0&, App.hInstance, ByVal 0&)
    
    ' The function sends a WM_PAINT message directly to the window procedure of the specified window, bypassing the application queue.
    UpdateWindow hWindowHwnd
    
    ' Activates the window and displays it in its current size and position.
    ShowWindow hWindowHwnd, SW_SHOW
     
    'Show our button
    ShowWindow gButtonHwnd, SW_NORMAL
    
    ' Get the memory address of the default window
    ' procedure for the button and store it in gButOldProc
    ' This will be used in ButtonWndProc to call the original
    ' window procedure for processing.
    gButOldProc = GetWindowLong(gButtonHwnd, GWL_WNDPROC)
   
    ' Set default window procedure of button to ButtonWndProc. We are using GWL_WNDPROC
    ' to set the address of the window procedure.
    Call SetWindowLong(gButtonHwnd, GWL_WNDPROC, GetAddress(AddressOf ButtonWndProc))
    
    ' Get a device context compatible with the window, this allows placement of the Cairo image on the user created window hDC
    thisHDC = GetDC(hWindowHwnd)  ' <- writing to the desktop dc fully transparent - overwritten shortly after
    
    Call setUpGDIP
   
    ' message loop to process window messages
    While GetMessage(myMsg, 0, 0, 0) <> 0 ' waiting for PostQuitMessage to be called to break out
        ' TranslateMessage takes keyboard messages and converts them to WM_CHAR for easier processing
        TranslateMessage myMsg
        
        ' Dispatchmessage calls the default window procedure to process the window message (MyWndProc)
        DispatchMessage myMsg
    Wend
    
    ' done with window, now clean up what we created
    DestroyWindow hWindowHwnd
    UnregisterClass szWindowClass, App.hInstance
    
    ' return exit code
    createAPIWindow = myMsg.wParam

    On Error GoTo 0
    Exit Function

createAPIWindow_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createAPIWindow of Module RealWindowMod"
End Function

'---------------------------------------------------------------------------------------
' Procedure : initiateAPIWindow
' Author    : beededea
' Date      : 11/03/2026
' Purpose   : determine screen details, initiate the window, pass address of MainWndProc to subclass, so VB6 can intercept messages such as WM_PAINT
'---------------------------------------------------------------------------------------
'
Public Function initiateAPIWindow(ByVal szWindowTitle As String, ByVal szWindowClass As String, Optional ByVal x As Long = CW_USEDEFAULT, Optional ByVal y As Long = CW_USEDEFAULT, Optional ByVal cx As Long = CW_USEDEFAULT, Optional ByVal cy As Long = CW_USEDEFAULT) As Long
    On Error GoTo initiateAPIWindow_Error
    
    Call configWindowParams

    'initiateAPIWindow = createAPIWindow(AddressOf MainWndProc, szWindowClass, szWindowTitle, x, y, windowSize.x, windowSize.y)
    initiateAPIWindow = createAPIWindow(AddressOf MainWndProc, szWindowClass, szWindowTitle, x, y, 500, 500)

    On Error GoTo 0
    Exit Function

initiateAPIWindow_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initiateAPIWindow of Module RealWindowMod"
End Function





'---------------------------------------------------------------------------------------
' Procedure : ButtonWndProc
' Author    : beededea
' Date      : 11/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function ButtonWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo ButtonWndProc_Error

    Select Case uMsg&
       Case WM_LBUTTONUP:
          ''Left mouse button went up (user clicked the button)
          ''You can use WM_LBUTTONDOWN for the MouseDown event.
         
          ''We use the MessageBox API call because the built in
          ''function 'MsgBox' stops thread processes, which causes
          ''flickering.
          Call MessageBox(gHwnd, "You clicked the button! Now unloading the program. ", App.Title, MB_OK Or MB_ICONEXCLAMATION)
          
          Call thisForm_Unload
          
    End Select
   
    'Since in MyCreateWindow we made the default window proc
    'this procedure, we have to call the old one using CallWindowProc
    ButtonWndProc = CallWindowProc(gButOldProc&, hwnd&, uMsg&, wParam&, lParam&)

    On Error GoTo 0
    Exit Function

ButtonWndProc_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ButtonWndProc of Module RealWindowMod"
   
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAddress
' Author    : beededea
' Date      : 11/03/2026
' Purpose   : Used with AddressOf to return the address in memory of a procedure.
'---------------------------------------------------------------------------------------
'
Public Function GetAddress(ByVal lngAddr As Long) As Long

    On Error GoTo GetAddress_Error

    GetAddress = lngAddr&

    On Error GoTo 0
    Exit Function

GetAddress_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetAddress of Module RealWindowMod"
   
End Function
'---------------------------------------------------------------------------------------
' Procedure : thisForm_Unload
' Author    : beededea
' Date      : 18/08/2022
' Purpose   : the standard form unload routine called from several places
'---------------------------------------------------------------------------------------
'
Private Sub thisForm_Unload() ' name follows VB6 standard naming convention
    On Error GoTo Form_Unload_Error
    
    Call SelectObject(dcMemory, hOldBmp) ' releases memory used by any open GDI handles
    Call DeleteObject(hBmpMemory)
    Call DeleteDC(dcMemory)
    Call ReleaseDC(hVBFormHwnd, dcMemory)

    End
    
    Call unloadAllForms(True)

    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Class Module module1"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : unloadAllForms
' Author    : beededea
' Date      : 28/06/2023
' Purpose   : unload all VB6 and RichClient5 forms
'---------------------------------------------------------------------------------------
'
Private Sub unloadAllForms(ByVal endItAll As Boolean)
    
   On Error GoTo unloadAllForms_Error
   
    ' stop all VB6 timers in the timer form
    
    ' unload the native VB6 forms
    
    'Unload frmMessage

    ' remove all variable references to each VB6 form in turn
    
    'Set widgetPrefs = Nothing

    On Error Resume Next
    
    If endItAll = True Then End

   On Error GoTo 0
   Exit Sub

unloadAllForms_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure unloadAllForms of Module Module1"
End Sub
