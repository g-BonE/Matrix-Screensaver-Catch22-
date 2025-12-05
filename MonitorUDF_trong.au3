#include-once
#include <WinAPIGdi.au3>
#include <WinAPISysWin.au3>
#include <WindowsConstants.au3>

; ===============================================================================================================================
; Title .........: Monitor UDF
; Description ...: Provides advanced monitor management and multi-monitor utilities.
; Author ........: Dao Van Trong - TRONG.PRO
; Version .......: 2.2
; Compatibility .: Windows XP SP2 and later (most functions)
;                  Windows 8.1+ required for GetDpiForMonitor API (automatic fallback on older systems)
; ===============================================================================================================================
; IMPORTANT NOTES:
; - Some functions require Administrator privileges (UAC elevation):
;   * _Monitor_Enable()       - Requires admin to attach monitor
;   * _Monitor_Disable()      - Requires admin to detach monitor
;   * _Monitor_SetResolution() - Requires admin to change display settings
;   * _Monitor_SetPrimary()   - Requires admin to set primary monitor
;   * _Monitor_SetDisplayMode() - Requires admin to change display mode
; - These functions may not work with proprietary graphics drivers (NVIDIA, AMD control panels)
; - Always test resolution changes before applying permanently
; ===============================================================================================================================
; FUNCTIONS SUMMARY
; ===============================================================================================================================
;
; === CORE MONITOR INFORMATION FUNCTIONS ===
;   _Monitor_GetList()                              - Enumerate all connected monitors and initialize global list
;   _Monitor_GetCount()                             - Get total number of connected monitors
;   _Monitor_GetPrimary()                           - Get index of the primary monitor
;   _Monitor_GetInfo($iMonitor)                     - Get detailed information about a specific monitor
;   _Monitor_GetBounds($iMonitor, ByRef ...)        - Get full monitor rectangle (including taskbar)
;   _Monitor_GetWorkArea($iMonitor, ByRef ...)      - Get working area (excluding taskbar)
;   _Monitor_GetDisplaySettings($iMonitor [, $iMode]) - Get current display mode settings
;   _Monitor_GetResolution($iMonitor)               - Get monitor resolution (width x height)
;   _Monitor_GetVirtualBounds()                     - Get bounding rectangle of entire virtual screen
;   _Monitor_Refresh()                              - Refresh monitor list (reload from system)
;   _Monitor_IsConnected($iMonitor)                 - Check if monitor is still connected
;   _Monitor_ShowInfo([$bShowMsgBox, $iTimeout])    - Display all monitor information
;
; === MONITOR LOCATION & DETECTION FUNCTIONS ===
;   _Monitor_GetFromPoint([$x, $y])                 - Get monitor from screen coordinates or mouse position
;   _Monitor_GetFromWindow($hWnd [, $iFlag])        - Get monitor containing a specific window
;   _Monitor_GetFromRect($L, $T, $R, $B [, $iFlag]) - Get monitor overlapping a rectangle
;
; === COORDINATE CONVERSION FUNCTIONS ===
;   _Monitor_ToVirtual($iMonitor, $x, $y)           - Convert local monitor coords to virtual coords
;   _Monitor_FromVirtual($iMonitor, $x, $y)         - Convert virtual coords to local monitor coords
;
; === WINDOW MANAGEMENT FUNCTIONS ===
;   _Monitor_IsVisibleWindow($hWnd)                 - Check if window is visible and top-level
;   _Monitor_MoveWindowToScreen($vTitle, ...)       - Move window to specific monitor (centered)
;   _Monitor_MoveWindowToAll($vTitle, ...)          - Move window across all monitors sequentially
;
; === DISPLAY MODE ENUMERATION FUNCTIONS ===
;   _Monitor_EnumAllDisplayModes($iMonitor)         - Enumerate all available display modes
;   _Monitor_GetDPI($iMonitor)                      - Get DPI scaling information (Win8.1+)
;   _Monitor_GetOrientation($iMonitor)              - Get display orientation (0/90/180/270°)
;   _Monitor_GetDisplayMode()                       - Get current display mode (duplicate/extend)
;
; === LAYOUT MANAGEMENT FUNCTIONS ===
;   _Monitor_GetLayout()                            - Get current display layout configuration
;   _Monitor_GetLayoutDescription()                 - Get text description of current layout
;   _Monitor_SaveLayout($sFilePath)                 - Save current layout to INI file
;   _Monitor_LoadLayout($sFilePath)                 - Load layout from INI file (does not apply)
;
; === DISPLAY SETTINGS FUNCTIONS (*** REQUIRE ADMINISTRATOR PRIVILEGES ***) ===
;   _Monitor_Enable($iMonitor)                      - Enable/attach a disabled monitor [ADMIN]
;   _Monitor_Disable($iMonitor)                     - Disable/detach a monitor [ADMIN]
;   _Monitor_SetResolution($iMonitor, $W, $H, ...)  - Set monitor resolution and refresh rate [ADMIN]
;   _Monitor_SetPrimary($iMonitor)                  - Set a monitor as primary [ADMIN]
;   _Monitor_SetDisplayMode($iMode)                 - Set display mode (duplicate/extend) [ADMIN]
;
; === INTERNAL/PRIVATE FUNCTIONS (DO NOT CALL DIRECTLY) ===
;   __Monitor_IsWindowsVersionOrGreater($iMajor, $iMinor) - Check OS version (internal)
;   __Monitor_FallbackOSVersionCheck($iMajor, $iMinor)    - Fallback OS check (internal)
;   __Monitor_IsWindows8_1OrGreater()                     - Check if Win8.1+ (internal)
;
; ===============================================================================================================================
; USAGE EXAMPLES:
; ===============================================================================================================================
; Example 1: Get basic monitor information
;   Local $iMonCount = _Monitor_GetCount()
;   Local $iPrimary = _Monitor_GetPrimary()
;   ConsoleWrite("Total monitors: " & $iMonCount & ", Primary: " & $iPrimary & @CRLF)
;
; Example 2: Get resolution of monitor 1
;   Local $aRes = _Monitor_GetResolution(1)
;   If Not @error Then ConsoleWrite("Resolution: " & $aRes[0] & "x" & $aRes[1] & @CRLF)
;
; Example 3: Move window to monitor 2 (centered)
;   _Monitor_MoveWindowToScreen("Notepad", "", 2)
;
; Example 4: Get monitor at mouse position
;   Local $iMon = _Monitor_GetFromPoint()
;   ConsoleWrite("Mouse is on monitor: " & $iMon & @CRLF)
;
; Example 5: Change resolution (requires admin)
;   Local $iResult = _Monitor_SetResolution(1, 1920, 1080, 32, 60)
;   If @error Then ConsoleWrite("Failed to set resolution. Error: " & @error & @CRLF)
;
; Example 6: Save and load layout
;   _Monitor_SaveLayout(@ScriptDir & "\my_layout.ini")
;   Local $aLayout = _Monitor_LoadLayout(@ScriptDir & "\my_layout.ini")
;
; Example 7: Show all monitor information
;   _Monitor_ShowInfo(True, 15) ; Show in MsgBox with 15 second timeout
;
; ===============================================================================================================================

#Region --- Constants Definition ---
;~ ; System Metrics Constants (Windows XP+)
;~ Global Const $SM_CMONITORS = 80
;~ Global Const $SM_XVIRTUALSCREEN = 76
;~ Global Const $SM_YVIRTUALSCREEN = 77
;~ Global Const $SM_CXVIRTUALSCREEN = 78
;~ Global Const $SM_CYVIRTUALSCREEN = 79

;~ ; Monitor Flags (Windows 2000+)
;~ Global Const $MONITOR_DEFAULTTONULL = 0
;~ Global Const $MONITOR_DEFAULTTOPRIMARY = 1
;~ Global Const $MONITOR_DEFAULTTONEAREST = 2

;~ ; EnumDisplaySettings Mode (Windows 95+)
;~ Global Const $ENUM_CURRENT_SETTINGS = -1
;~ Global Const $ENUM_REGISTRY_SETTINGS = -2

;~ ; Window Style Constants (Windows 95+)
;~ Global Const $GWL_STYLE = -16
;~ Global Const $GWL_EXSTYLE = -20
;~ Global Const $WS_VISIBLE = 0x10000000
;~ Global Const $WS_CHILD = 0x40000000
;~ Global Const $WS_EX_TOOLWINDOW = 0x00000080

; ChangeDisplaySettings Flags (Windows 95+)
Global Const $CDS_UPDATEREGISTRY = 0x00000001
Global Const $CDS_TEST = 0x00000002
Global Const $CDS_FULLSCREEN = 0x00000004
Global Const $CDS_GLOBAL = 0x00000008
Global Const $CDS_SET_PRIMARY = 0x00000010
Global Const $CDS_NORESET = 0x10000000
Global Const $CDS_RESET = 0x40000000

; ChangeDisplaySettings Return Values
Global Const $DISP_CHANGE_SUCCESSFUL = 0
Global Const $DISP_CHANGE_RESTART = 1
Global Const $DISP_CHANGE_FAILED = -1
Global Const $DISP_CHANGE_BADMODE = -2
Global Const $DISP_CHANGE_NOTUPDATED = -3
Global Const $DISP_CHANGE_BADFLAGS = -4
Global Const $DISP_CHANGE_BADPARAM = -5
Global Const $DISP_CHANGE_BADDUALVIEW = -6

;~ ; DEVMODE Field Flags (Windows 95+)
;~ Global Const $DM_ORIENTATION = 0x00000001
;~ Global Const $DM_PAPERSIZE = 0x00000002
;~ Global Const $DM_PAPERLENGTH = 0x00000004
;~ Global Const $DM_PAPERWIDTH = 0x00000008
;~ Global Const $DM_POSITION = 0x00000020
;~ Global Const $DM_DISPLAYORIENTATION = 0x00000080
;~ Global Const $DM_DISPLAYFIXEDOUTPUT = 0x20000000
;~ Global Const $DM_BITSPERPEL = 0x00040000
;~ Global Const $DM_PELSWIDTH = 0x00080000
;~ Global Const $DM_PELSHEIGHT = 0x00100000
;~ Global Const $DM_DISPLAYFLAGS = 0x00200000
;~ Global Const $DM_DISPLAYFREQUENCY = 0x00400000
#EndRegion --- Constants Definition ---

#Region --- Global Variables ---
; ===============================================================================================================================
; Global Monitor Information Array
; ===============================================================================================================================
; $__g_aMonitorList[][] structure:
;
;   [0][0] = Number of monitors detected
;   [0][1] = Virtual desktop Left coordinate (combined area)
;   [0][2] = Virtual desktop Top coordinate
;   [0][3] = Virtual desktop Right coordinate
;   [0][4] = Virtual desktop Bottom coordinate
;   [0][5] = Virtual desktop Width
;   [0][6] = Virtual desktop Height
;
; For each monitor index i (1..$__g_aMonitorList[0][0]):
;   [i][0] = Monitor handle (HMONITOR)
;   [i][1] = Left coordinate of monitor
;   [i][2] = Top coordinate of monitor
;   [i][3] = Right coordinate of monitor
;   [i][4] = Bottom coordinate of monitor
;   [i][5] = IsPrimary (1 if primary, 0 otherwise)
;   [i][6] = Device name string (e.g. "\\.\DISPLAY1")
;
; ===============================================================================================================================
Global $__g_aMonitorList[1][7] = [[0, 0, 0, 0, 0, 0, ""]]
#EndRegion --- Global Variables ---

#Region --- OS Compatibility Functions ---
; #FUNCTION# ====================================================================================================================
; Name...........: __Monitor_IsWindowsVersionOrGreater
; Description....: Check if running on Windows version or greater (internal function)
; Syntax.........: __Monitor_IsWindowsVersionOrGreater($iMajor, $iMinor = 0)
; Parameters.....: $iMajor      - Major version number (e.g., 6 for Vista, 10 for Windows 10)
;                  $iMinor       - [optional] Minor version number. Default is 0
; Return values..: True if OS version is equal or greater, False otherwise
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Internal function for OS compatibility checking. Uses GetVersionEx API directly for accurate
;                  version detection, works on all AutoIt versions (even old ones where @OSVersion may be inaccurate).
;                  Windows XP = 5.1, Vista = 6.0, Win7 = 6.1, Win8 = 6.2, Win8.1 = 6.3, Win10 = 10.0
; ================================================================================================================================
Func __Monitor_IsWindowsVersionOrGreater($iMajor, $iMinor = 0)
    ; Use GetVersionEx API directly (works on all Windows versions and all AutoIt versions)
    ; This is more reliable than @OSVersion which may not be accurate on old AutoIt versions

    ; Define OSVERSIONINFOEX structure
    Local $tOSVI = DllStructCreate("dword;dword;dword;dword;dword;wchar[128];ushort;ushort;ushort;byte;byte")
    If @error Then Return __Monitor_FallbackOSVersionCheck($iMajor, $iMinor)

    ; Set structure size (first field)
    DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))

    ; Call GetVersionExW (Unicode version, available from Windows 2000+)
    Local $aRet = DllCall("kernel32.dll", "bool", "GetVersionExW", "struct*", $tOSVI)
    If @error Or Not IsArray($aRet) Or Not $aRet[0] Then
        ; Fallback: Try ANSI version GetVersionExA (available from Windows 95+)
        Local $tOSVIA = DllStructCreate("dword;dword;dword;dword;dword;char[128];ushort;ushort;ushort;byte;byte")
        If @error Then Return __Monitor_FallbackOSVersionCheck($iMajor, $iMinor)

        DllStructSetData($tOSVIA, 1, DllStructGetSize($tOSVIA))
        $aRet = DllCall("kernel32.dll", "bool", "GetVersionExA", "struct*", $tOSVIA)
        If @error Or Not IsArray($aRet) Or Not $aRet[0] Then
            Return __Monitor_FallbackOSVersionCheck($iMajor, $iMinor)
        EndIf

        ; Use ANSI version data
        Local $iOSMajor = DllStructGetData($tOSVIA, 2)
        Local $iOSMinor = DllStructGetData($tOSVIA, 3)

        ; Compare versions
        If $iOSMajor > $iMajor Then Return True
        If $iOSMajor = $iMajor And $iOSMinor >= $iMinor Then Return True
        Return False
    EndIf

    ; Get version from structure (Unicode version)
    Local $iOSMajor = DllStructGetData($tOSVI, 2)
    Local $iOSMinor = DllStructGetData($tOSVI, 3)

    ; Handle Windows 10/11: GetVersionEx may return 6.3 for compatibility
    ; Check build number to distinguish Windows 10/11 from 8.1
    Local $iBuildNumber = DllStructGetData($tOSVI, 4)
    If $iOSMajor = 6 And $iOSMinor = 3 And $iBuildNumber >= 10000 Then
        ; Windows 10/11 (build number >= 10000)
        $iOSMajor = 10
        $iOSMinor = 0
    EndIf

    ; Compare versions
    If $iOSMajor > $iMajor Then Return True
    If $iOSMajor = $iMajor And $iOSMinor >= $iMinor Then Return True
    Return False
EndFunc   ;==>__Monitor_IsWindowsVersionOrGreater

; #FUNCTION# ====================================================================================================================
; Name...........: __Monitor_FallbackOSVersionCheck
; Description....: Fallback OS version check using @OSVersion macro (internal function)
; Syntax.........: __Monitor_FallbackOSVersionCheck($iMajor, $iMinor)
; Parameters.....: $iMajor      - Major version number
;                  $iMinor       - Minor version number
; Return values..: True if OS version is equal or greater, False otherwise
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Used only when GetVersionEx API fails. Less reliable than API call.
; ================================================================================================================================
Func __Monitor_FallbackOSVersionCheck($iMajor, $iMinor)
    ; Fallback to @OSVersion (may be inaccurate on old AutoIt versions, but better than nothing)
    Local $sOSVersion = @OSVersion
    Local $iOSMajor = 5, $iOSMinor = 1 ; Default to XP

    If StringInStr($sOSVersion, "WIN_11") Then
        $iOSMajor = 10
        $iOSMinor = 0
    ElseIf StringInStr($sOSVersion, "WIN_10") Then
        $iOSMajor = 10
        $iOSMinor = 0
    ElseIf StringInStr($sOSVersion, "WIN_8") Then
        ; Use @OSBuild to distinguish 8.0 vs 8.1 (8.1 has build >= 9600)
        If @OSBuild >= 9600 Then
            $iOSMajor = 6
            $iOSMinor = 3 ; Windows 8.1
        Else
            $iOSMajor = 6
            $iOSMinor = 2 ; Windows 8
        EndIf
    ElseIf StringInStr($sOSVersion, "WIN_7") Then
        $iOSMajor = 6
        $iOSMinor = 1
    ElseIf StringInStr($sOSVersion, "WIN_VISTA") Then
        $iOSMajor = 6
        $iOSMinor = 0
    ElseIf StringInStr($sOSVersion, "WIN_XP") Then
        $iOSMajor = 5
        $iOSMinor = 1
    ElseIf StringInStr($sOSVersion, "WIN_2003") Then
        $iOSMajor = 5
        $iOSMinor = 2
    EndIf

    ; Compare versions
    If $iOSMajor > $iMajor Then Return True
    If $iOSMajor = $iMajor And $iOSMinor >= $iMinor Then Return True
    Return False
EndFunc   ;==>__Monitor_FallbackOSVersionCheck

; #FUNCTION# ====================================================================================================================
; Name...........: __Monitor_IsWindows8_1OrGreater
; Description....: Check if running on Windows 8.1 or greater (internal function)
; Syntax.........: __Monitor_IsWindows8_1OrGreater()
; Parameters.....: None
; Return values..: True if Windows 8.1 or greater, False otherwise
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Internal function for checking if GetDpiForMonitor API is available
; ================================================================================================================================
Func __Monitor_IsWindows8_1OrGreater()
    ; FIXED: Improved logic to correctly detect Windows 8.1+
    Local $tOSVI = DllStructCreate("dword;dword;dword;dword;dword;wchar[128];ushort;ushort;ushort;byte;byte")
    If @error Then Return __Monitor_IsWindowsVersionOrGreater(6, 3)

    DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
    Local $aRet = DllCall("kernel32.dll", "bool", "GetVersionExW", "struct*", $tOSVI)
    If @error Or Not IsArray($aRet) Or Not $aRet[0] Then
        Return __Monitor_IsWindowsVersionOrGreater(6, 3)
    EndIf

    Local $iOSMajor = DllStructGetData($tOSVI, 2)
    Local $iOSMinor = DllStructGetData($tOSVI, 3)
    Local $iBuildNumber = DllStructGetData($tOSVI, 4)

    ; FIXED: Correct logic for Windows 8.1+ detection
    ; Windows 10+ has major >= 10
    If $iOSMajor > 6 Then Return True
    ; Windows 8.1+ has major=6 AND minor>=3
    If $iOSMajor = 6 And $iOSMinor > 3 Then Return True
    ; Windows 8.1 exactly has major=6 AND minor=3 AND build>=9600
    If $iOSMajor = 6 And $iOSMinor = 3 And $iBuildNumber >= 9600 Then Return True

    Return False
EndFunc   ;==>__Monitor_IsWindows8_1OrGreater
#EndRegion --- OS Compatibility Functions ---

#Region --- Core Monitor Functions ---
; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetFromPoint
; Description....: Get the monitor index from a screen coordinate or current mouse position
; Syntax.........: _Monitor_GetFromPoint([$iX = -1 [, $iY = -1]])
; Parameters.....: $iX        - [optional] X coordinate in virtual screen coordinates. Default is -1 (use mouse position)
;                  $iY        - [optional] Y coordinate in virtual screen coordinates. Default is -1 (use mouse position)
; Return values..: Success    - Monitor index (1..N)
;                  Failure    - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid parameters or MouseGetPos failed
;                  |@error = 2 - Monitor not found at specified coordinates
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: If both $iX and $iY are -1 (default), function uses current mouse position.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetList, _Monitor_GetFromWindow, _Monitor_GetFromRect
; ================================================================================================================================
Func _Monitor_GetFromPoint($iX = -1, $iY = -1)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()

    ; Use WinAPI function if available
    If $iX = -1 Or $iY = -1 Then
        Local $aMouse = MouseGetPos()
        If @error Then Return SetError(1, 0, 0)
        $iX = $aMouse[0]
        $iY = $aMouse[1]
    EndIf

    Local $tPoint = DllStructCreate($tagPOINT)
    If @error Then Return SetError(1, 0, 0)
    DllStructSetData($tPoint, "X", $iX)
    DllStructSetData($tPoint, "Y", $iY)

    Local $hMonitor = _WinAPI_MonitorFromPoint($tPoint, $MONITOR_DEFAULTTONEAREST)
    If @error Then
        $hMonitor = 0
    EndIf

    ; Find index in our list
    For $i = 1 To $__g_aMonitorList[0][0]
        If $__g_aMonitorList[$i][0] = $hMonitor Then Return $i
    Next

    ; Fallback to coordinate checking
    For $i = 1 To $__g_aMonitorList[0][0]
        If $iX >= $__g_aMonitorList[$i][1] _
                And $iX < $__g_aMonitorList[$i][3] _
                And $iY >= $__g_aMonitorList[$i][2] _
                And $iY < $__g_aMonitorList[$i][4] Then
            Return $i
        EndIf
    Next
    Return SetError(2, 0, 0)
EndFunc   ;==>_Monitor_GetFromPoint

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetFromWindow
; Description....: Get the monitor index that contains the specified window
; Syntax.........: _Monitor_GetFromWindow($hWnd [, $iFlag = $MONITOR_DEFAULTTONEAREST])
; Parameters.....: $hWnd      - Window handle or title string. Can be HWND or window title
;                  $iFlag     - [optional] Monitor flag. Default is $MONITOR_DEFAULTTONEAREST
;                              Can be: $MONITOR_DEFAULTTONULL, $MONITOR_DEFAULTTOPRIMARY, $MONITOR_DEFAULTTONEAREST
; Return values..: Success    - Monitor index (1..N)
;                  Failure    - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid window handle or window not found
;                  |@error = 2 - WinAPI MonitorFromWindow call failed
;                  |@error = 3 - Monitor handle not found in internal list
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Function accepts both window handles and window titles. Automatically converts title to handle.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetList, _Monitor_GetFromPoint, _Monitor_GetFromRect
; ================================================================================================================================
Func _Monitor_GetFromWindow($hWnd, $iFlag = $MONITOR_DEFAULTTONEAREST)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If Not IsHWnd($hWnd) Then $hWnd = WinGetHandle($hWnd)
    If Not $hWnd Then Return SetError(1, 0, 0)

    Local $hMonitor = _WinAPI_MonitorFromWindow($hWnd, $iFlag)
    If @error Or Not $hMonitor Then Return SetError(2, 0, 0)

    For $i = 1 To $__g_aMonitorList[0][0]
        If $__g_aMonitorList[$i][0] = $hMonitor Then Return $i
    Next
    Return SetError(3, 0, 0)
EndFunc   ;==>_Monitor_GetFromWindow

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetFromRect
; Description....: Get the monitor index that has the largest intersection with the specified rectangle
; Syntax.........: _Monitor_GetFromRect($iLeft, $iTop, $iRight, $iBottom [, $iFlag = $MONITOR_DEFAULTTONEAREST])
; Parameters.....: $iLeft     - Left coordinate of the rectangle in virtual screen coordinates
;                  $iTop      - Top coordinate of the rectangle in virtual screen coordinates
;                  $iRight    - Right coordinate of the rectangle in virtual screen coordinates
;                  $iBottom   - Bottom coordinate of the rectangle in virtual screen coordinates
;                  $iFlag     - [optional] Monitor flag. Default is $MONITOR_DEFAULTTONEAREST
;                              Can be: $MONITOR_DEFAULTTONULL, $MONITOR_DEFAULTTOPRIMARY, $MONITOR_DEFAULTTONEAREST
; Return values..: Success    - Monitor index (1..N)
;                  Failure    - 0, sets @error to non-zero:
;                  |@error = 1 - DllStructCreate failed or WinAPI MonitorFromRect call failed
;                  |@error = 2 - Monitor not found in internal list
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Coordinates should be in virtual screen coordinate system (can span multiple monitors).
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetList, _Monitor_GetFromPoint, _Monitor_GetFromWindow
; ================================================================================================================================
Func _Monitor_GetFromRect($iLeft, $iTop, $iRight, $iBottom, $iFlag = $MONITOR_DEFAULTTONEAREST)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()

    Local $tRect = DllStructCreate($tagRECT)
    If @error Then Return SetError(1, 0, 0)
    DllStructSetData($tRect, "Left", $iLeft)
    DllStructSetData($tRect, "Top", $iTop)
    DllStructSetData($tRect, "Right", $iRight)
    DllStructSetData($tRect, "Bottom", $iBottom)

    Local $hMonitor = _WinAPI_MonitorFromRect($tRect, $iFlag)
    If @error Or Not $hMonitor Then Return SetError(1, 0, 0)

    For $i = 1 To $__g_aMonitorList[0][0]
        If $__g_aMonitorList[$i][0] = $hMonitor Then Return $i
    Next
    Return SetError(2, 0, 0)
EndFunc   ;==>_Monitor_GetFromRect

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetWorkArea
; Description....: Get working area of a specific monitor (excluding taskbar and system bars)
; Syntax.........: _Monitor_GetWorkArea($iMonitor, ByRef $left, ByRef $top, ByRef $right, ByRef $bottom)
; Parameters.....: $iMonitor   - Monitor index (1..N)
;                  $left       - [out] Left coordinate of work area (virtual screen coordinates)
;                  $top        - [out] Top coordinate of work area (virtual screen coordinates)
;                  $right      - [out] Right coordinate of work area (virtual screen coordinates)
;                  $bottom     - [out] Bottom coordinate of work area (virtual screen coordinates)
; Return values..: Success     - 1
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - WinAPI GetMonitorInfo call failed
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Work area excludes taskbar and other system bars. Use _Monitor_GetBounds() for full monitor area.
;                  All coordinates are in virtual screen coordinate system.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetBounds, _Monitor_GetInfo, _Monitor_GetList
; ================================================================================================================================
Func _Monitor_GetWorkArea($iMonitor, ByRef $left, ByRef $top, ByRef $right, ByRef $bottom)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    Local $hMonitor = $__g_aMonitorList[$iMonitor][0]
    Local $aInfo = _WinAPI_GetMonitorInfo($hMonitor)
    If @error Or Not IsArray($aInfo) Then Return SetError(2, 0, 0)

    Local $tWorkArea = $aInfo[1]
    If Not IsDllStruct($tWorkArea) Then Return SetError(2, 0, 0)
    $left = DllStructGetData($tWorkArea, "Left")
    $top = DllStructGetData($tWorkArea, "Top")
    $right = DllStructGetData($tWorkArea, "Right")
    $bottom = DllStructGetData($tWorkArea, "Bottom")
    Return 1
EndFunc   ;==>_Monitor_GetWorkArea

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetBounds
; Description....: Get full bounds of a specific monitor (including taskbar and all system areas)
; Syntax.........: _Monitor_GetBounds($iMonitor, ByRef $left, ByRef $top, ByRef $right, ByRef $bottom)
; Parameters.....: $iMonitor   - Monitor index (1..N)
;                  $left       - [out] Left coordinate of monitor (virtual screen coordinates)
;                  $top        - [out] Top coordinate of monitor (virtual screen coordinates)
;                  $right      - [out] Right coordinate of monitor (virtual screen coordinates)
;                  $bottom     - [out] Bottom coordinate of monitor (virtual screen coordinates)
; Return values..: Success     - 1
;                  Failure     - 0, sets @error = 1 (Invalid monitor index)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Returns the full physical bounds of the monitor including all system bars (taskbar, etc.).
;                  All coordinates are in virtual screen coordinate system.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
;                  For usable area excluding taskbar, use _Monitor_GetWorkArea() instead.
; Related........: _Monitor_GetWorkArea, _Monitor_GetInfo, _Monitor_GetList
; ================================================================================================================================
Func _Monitor_GetBounds($iMonitor, ByRef $left, ByRef $top, ByRef $right, ByRef $bottom)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    $left = $__g_aMonitorList[$iMonitor][1]
    $top = $__g_aMonitorList[$iMonitor][2]
    $right = $__g_aMonitorList[$iMonitor][3]
    $bottom = $__g_aMonitorList[$iMonitor][4]
    Return 1
EndFunc   ;==>_Monitor_GetBounds

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetInfo
; Description....: Get detailed information about a monitor
; Syntax.........: _Monitor_GetInfo($iMonitor)
; Parameters.....: $iMonitor   - Monitor index (1..N)
; Return values..: Success     - Array with 11 elements:
;                  |[0]  - Monitor handle (HMONITOR)
;                  |[1]  - Left coordinate of monitor bounds (virtual screen coordinates)
;                  |[2]  - Top coordinate of monitor bounds (virtual screen coordinates)
;                  |[3]  - Right coordinate of monitor bounds (virtual screen coordinates)
;                  |[4]  - Bottom coordinate of monitor bounds (virtual screen coordinates)
;                  |[5]  - Left coordinate of work area (virtual screen coordinates)
;                  |[6]  - Top coordinate of work area (virtual screen coordinates)
;                  |[7]  - Right coordinate of work area (virtual screen coordinates)
;                  |[8]  - Bottom coordinate of work area (virtual screen coordinates)
;                  |[9]  - IsPrimary flag (1 = Primary, 0 = Secondary)
;                  |[10] - Device name string (e.g., "\\.\DISPLAY1")
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - WinAPI GetMonitorInfo call failed
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: This is the most comprehensive function to get all monitor information at once.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetList, _Monitor_GetBounds, _Monitor_GetWorkArea, _Monitor_GetPrimary
; ================================================================================================================================
Func _Monitor_GetInfo($iMonitor)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    Local $hMonitor = $__g_aMonitorList[$iMonitor][0]
    Local $aInfo = _WinAPI_GetMonitorInfo($hMonitor)
    If @error Or Not IsArray($aInfo) Then Return SetError(2, 0, 0)

    Local $tMonitorRect = $aInfo[0]
    Local $tWorkRect = $aInfo[1]
    If Not IsDllStruct($tMonitorRect) Or Not IsDllStruct($tWorkRect) Then Return SetError(2, 0, 0)

    Local $aResult[11]
    $aResult[0] = $hMonitor
    $aResult[1] = DllStructGetData($tMonitorRect, "Left")
    $aResult[2] = DllStructGetData($tMonitorRect, "Top")
    $aResult[3] = DllStructGetData($tMonitorRect, "Right")
    $aResult[4] = DllStructGetData($tMonitorRect, "Bottom")
    $aResult[5] = DllStructGetData($tWorkRect, "Left")
    $aResult[6] = DllStructGetData($tWorkRect, "Top")
    $aResult[7] = DllStructGetData($tWorkRect, "Right")
    $aResult[8] = DllStructGetData($tWorkRect, "Bottom")
    $aResult[9] = ($aInfo[2] <> 0) ; IsPrimary
    $aResult[10] = $aInfo[3] ; DeviceName

    Return $aResult
EndFunc   ;==>_Monitor_GetInfo

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetDisplaySettings
; Description....: Get display settings for a monitor (resolution, color depth, refresh rate, etc.)
; Syntax.........: _Monitor_GetDisplaySettings($iMonitor [, $iMode = $ENUM_CURRENT_SETTINGS])
; Parameters.....: $iMonitor   - Monitor index (1..N)
;                  $iMode      - [optional] Display mode index. Default is $ENUM_CURRENT_SETTINGS
;                              Use $ENUM_CURRENT_SETTINGS to get current active settings
;                              Use index number (0..N) to enumerate available modes
; Return values..: Success     - Array with 5 elements:
;                  |[0] - Width (pixels)
;                  |[1] - Height (pixels)
;                  |[2] - Bits per pixel (color depth)
;                  |[3] - Refresh rate (Hz)
;                  |[4] - Display mode flags
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - GetInfo failed (could not get device name)
;                  |@error = 3 - WinAPI EnumDisplaySettings call failed
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Use $ENUM_CURRENT_SETTINGS to get the currently active display mode.
;                  Use _Monitor_EnumAllDisplayModes() to get all available display modes.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetResolution, _Monitor_EnumAllDisplayModes, _Monitor_GetList
; ================================================================================================================================
Func _Monitor_GetDisplaySettings($iMonitor, $iMode = $ENUM_CURRENT_SETTINGS)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    Local $sDevice = $__g_aMonitorList[$iMonitor][6]
    If $sDevice = "" Then
        Local $aInfo = _Monitor_GetInfo($iMonitor)
        If @error Then Return SetError(2, 0, 0)
        $sDevice = $aInfo[10]
    EndIf

    Local $aSettings = _WinAPI_EnumDisplaySettings($sDevice, $iMode)
    If @error Or Not IsArray($aSettings) Then Return SetError(3, 0, 0)

    Return $aSettings
EndFunc   ;==>_Monitor_GetDisplaySettings

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetResolution
; Description....: Get the current resolution (width and height) of a monitor
; Syntax.........: _Monitor_GetResolution($iMonitor)
; Parameters.....: $iMonitor   - Monitor index (1..N)
; Return values..: Success     - Array with 2 elements:
;                  |[0] - Width in pixels
;                  |[1] - Height in pixels
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - GetDisplaySettings failed
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: This is a convenience function that returns only width and height from display settings.
;                  For full display settings (color depth, refresh rate, etc.), use _Monitor_GetDisplaySettings().
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetDisplaySettings, _Monitor_GetInfo, _Monitor_GetList
; ================================================================================================================================
Func _Monitor_GetResolution($iMonitor)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    Local $aSettings = _Monitor_GetDisplaySettings($iMonitor)
    If @error Then Return SetError(2, @error, 0)

    Local $aResult[2] = [$aSettings[0], $aSettings[1]]
    Return $aResult
EndFunc   ;==>_Monitor_GetResolution

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetPrimary
; Description....: Get the index of the primary monitor
; Syntax.........: _Monitor_GetPrimary()
; Parameters.....: None
; Return values..: Success     - Monitor index (1..N) of the primary monitor
;                  Failure     - 0, sets @error = 1 (No primary monitor found)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Primary monitor is the monitor that contains the taskbar by default in Windows.
;                  Function first checks cached IsPrimary flags, then falls back to querying system if needed.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetList, _Monitor_GetInfo, _Monitor_GetCount
; ================================================================================================================================
Func _Monitor_GetPrimary()
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()

    ; Use cached IsPrimary flag if available
    For $i = 1 To $__g_aMonitorList[0][0]
        If $__g_aMonitorList[$i][5] = 1 Then Return $i
    Next

    ; Fallback: query from system
    For $i = 1 To $__g_aMonitorList[0][0]
        Local $aInfo = _Monitor_GetInfo($i)
        If Not @error And $aInfo[9] = 1 Then Return $i
    Next
    Return SetError(1, 0, 0)
EndFunc   ;==>_Monitor_GetPrimary

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetCount
; Description....: Returns the total number of connected monitors
; Syntax.........: _Monitor_GetCount()
; Parameters.....: None
; Return values..: Success     - Number of monitors (>= 1)
;                  Failure     - 0, sets @error = 1 (Enumeration failed)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Function first tries GetSystemMetrics API for fastest response.
;                  Falls back to cached monitor list if API call fails.
;                  Automatically validates and refreshes monitor list if count mismatch detected.
; Related........: _Monitor_GetList, _Monitor_Refresh, _Monitor_GetPrimary
; ================================================================================================================================
Func _Monitor_GetCount()
    ; FIXED: Check @error after each DllCall separately
    Local $aRet = DllCall("user32.dll", "int", "GetSystemMetrics", "int", $SM_CMONITORS)
    If @error Then
        ; Fallback to our cached list
        If $__g_aMonitorList[0][0] = 0 Then
            If _Monitor_GetList() = -1 Then Return SetError(1, 0, 0)
        EndIf
        Return $__g_aMonitorList[0][0]
    EndIf

    If Not IsArray($aRet) Or $aRet[0] < 1 Then
        ; Fallback to our cached list
        If $__g_aMonitorList[0][0] = 0 Then
            If _Monitor_GetList() = -1 Then Return SetError(1, 0, 0)
        EndIf
        Return $__g_aMonitorList[0][0]
    Else
        ; Validate count matches our list (refresh if needed)
        If $__g_aMonitorList[0][0] <> $aRet[0] Then
            _Monitor_GetList()
        EndIf
        Return $aRet[0]
    EndIf
EndFunc   ;==>_Monitor_GetCount

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetVirtualBounds
; Description....: Get bounding rectangle of all monitors combined (the "virtual screen")
; Syntax.........: _Monitor_GetVirtualBounds()
; Parameters.....: None
; Return values..: Success     - Array with 4 elements:
;                  |[0] - Left coordinate of virtual screen
;                  |[1] - Top coordinate of virtual screen
;                  |[2] - Width of virtual screen in pixels
;                  |[3] - Height of virtual screen in pixels
;                  Failure     - 0, sets @error = 1 (DllCall failed)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Virtual screen is the bounding rectangle that encompasses all connected monitors.
;                  Function uses GetSystemMetrics API. Falls back to cached virtual bounds if API fails.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetList, _Monitor_GetBounds, _Monitor_GetCount
; ================================================================================================================================
Func _Monitor_GetVirtualBounds()
    ; FIXED: Check @error after each DllCall separately
    Local $aL = DllCall("user32.dll", "int", "GetSystemMetrics", "int", $SM_XVIRTUALSCREEN)
    Local $bErrorL = @error

    Local $aT = DllCall("user32.dll", "int", "GetSystemMetrics", "int", $SM_YVIRTUALSCREEN)
    Local $bErrorT = @error

    Local $aW = DllCall("user32.dll", "int", "GetSystemMetrics", "int", $SM_CXVIRTUALSCREEN)
    Local $bErrorW = @error

    Local $aH = DllCall("user32.dll", "int", "GetSystemMetrics", "int", $SM_CYVIRTUALSCREEN)
    Local $bErrorH = @error

    ; Validate all calls succeeded
    Local $bError = False
    If $bErrorL Or $bErrorT Or $bErrorW Or $bErrorH Then $bError = True
    If Not IsArray($aL) Or Not IsArray($aT) Or Not IsArray($aW) Or Not IsArray($aH) Then $bError = True

    If $bError Then
        ; Fallback to cached virtual bounds
        If $__g_aMonitorList[0][0] = 0 Then
            If _Monitor_GetList() = -1 Then Return SetError(1, 0, 0)
        EndIf
        Local $aRet[4] = [$__g_aMonitorList[0][1], $__g_aMonitorList[0][2], $__g_aMonitorList[0][5], $__g_aMonitorList[0][6]]
        Return $aRet
    EndIf

    ; Validate returned values are reasonable
    If $aL[0] < -32768 Or $aT[0] < -32768 Or $aW[0] < 1 Or $aH[0] < 1 Then
        ; Invalid values, use fallback
        If $__g_aMonitorList[0][0] = 0 Then
            If _Monitor_GetList() = -1 Then Return SetError(1, 0, 0)
        EndIf
        Local $aRet[4] = [$__g_aMonitorList[0][1], $__g_aMonitorList[0][2], $__g_aMonitorList[0][5], $__g_aMonitorList[0][6]]
        Return $aRet
    EndIf

    Local $a[4] = [$aL[0], $aT[0], $aW[0], $aH[0]]
    Return $a
EndFunc   ;==>_Monitor_GetVirtualBounds

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_ToVirtual
; Description....: Convert local monitor coordinates to virtual screen coordinates
; Syntax.........: _Monitor_ToVirtual($iMonitor, $x, $y)
; Parameters.....: $iMonitor   - Monitor index (1..N)
;                  $x          - X coordinate in local monitor coordinates (0-based from monitor's left edge)
;                  $y          - Y coordinate in local monitor coordinates (0-based from monitor's top edge)
; Return values..: Success     - Array with 2 elements [X, Y] in virtual screen coordinates
;                  Failure     - 0, sets @error = 1 (Invalid monitor index)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Local coordinates are relative to the monitor (0,0 is top-left of that monitor).
;                  Virtual coordinates are absolute in the virtual screen coordinate system.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_FromVirtual, _Monitor_GetBounds, _Monitor_GetList
; ================================================================================================================================
Func _Monitor_ToVirtual($iMonitor, $x, $y)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    Local $aRet[2] = [$__g_aMonitorList[$iMonitor][1] + $x, $__g_aMonitorList[$iMonitor][2] + $y]
    Return $aRet
EndFunc   ;==>_Monitor_ToVirtual

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_FromVirtual
; Description....: Convert virtual screen coordinates back to local monitor coordinates
; Syntax.........: _Monitor_FromVirtual($iMonitor, $x, $y)
; Parameters.....: $iMonitor   - Monitor index (1..N)
;                  $x          - X coordinate in virtual screen coordinates
;                  $y          - Y coordinate in virtual screen coordinates
; Return values..: Success     - Array with 2 elements [X, Y] in local monitor coordinates (0-based)
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - Coordinates are not within the specified monitor's bounds
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Local coordinates are relative to the monitor (0,0 is top-left of that monitor).
;                  Virtual coordinates are absolute in the virtual screen coordinate system.
;                  Function validates that coordinates are actually on the specified monitor.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_ToVirtual, _Monitor_GetBounds, _Monitor_GetList
; ================================================================================================================================
Func _Monitor_FromVirtual($iMonitor, $x, $y)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    ; Validate coordinates are on this monitor
    If $x < $__g_aMonitorList[$iMonitor][1] Or $x >= $__g_aMonitorList[$iMonitor][3] _
            Or $y < $__g_aMonitorList[$iMonitor][2] Or $y >= $__g_aMonitorList[$iMonitor][4] Then
        Return SetError(2, 0, 0)
    EndIf

    Local $aRet[2] = [$x - $__g_aMonitorList[$iMonitor][1], $y - $__g_aMonitorList[$iMonitor][2]]
    Return $aRet
EndFunc   ;==>_Monitor_FromVirtual
#EndRegion --- Core Monitor Functions ---

#Region --- Window Management Functions ---
; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_IsVisibleWindow
; Description....: Check if a window is visible and is a top-level window (not a child window or tool window)
; Syntax.........: _Monitor_IsVisibleWindow($hWnd)
; Parameters.....: $hWnd       - Window handle or title string. Can be HWND or window title
; Return values..: Success     - True if window is visible and top-level, False otherwise
;                  Failure     - False, sets @error = 1 (Invalid window handle)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Function checks for: WS_VISIBLE flag, not WS_CHILD, and not WS_EX_TOOLWINDOW.
;                  Accepts both window handles and window titles. Automatically converts title to handle.
;                  This is useful for filtering which windows should be moved between monitors.
; Related........: _Monitor_MoveWindowToScreen, _Monitor_GetFromWindow
; ================================================================================================================================
Func _Monitor_IsVisibleWindow($hWnd)
    If Not IsHWnd($hWnd) Then $hWnd = WinGetHandle($hWnd)
    If Not $hWnd Or Not WinExists($hWnd) Then Return SetError(1, 0, False)

    Local $style = _WinAPI_GetWindowLong($hWnd, $GWL_STYLE)
    If @error Then Return SetError(1, 0, False)
    If BitAND($style, $WS_VISIBLE) = 0 Then Return False
    If BitAND($style, $WS_CHILD) <> 0 Then Return False

    Local $ex = _WinAPI_GetWindowLong($hWnd, $GWL_EXSTYLE)
    If @error Then Return False
    If BitAND($ex, $WS_EX_TOOLWINDOW) <> 0 Then Return False
    Return True
EndFunc   ;==>_Monitor_IsVisibleWindow

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_MoveWindowToScreen
; Description....: Move a visible window to a specific monitor (centered if coordinates not specified)
; Syntax.........: _Monitor_MoveWindowToScreen($vTitle [, $sText = "" [, $iMonitor = -1 [, $x = -1 [, $y = -1 [, $bUseWorkArea = True]]]]])
; Parameters.....: $vTitle     - Window title or handle. Can be HWND, title string, or class string
;                  $sText      - [optional] Window text (for matching with title). Default is ""
;                  $iMonitor   - [optional] Target monitor index (1..N). Default is -1 (uses monitor 1)
;                  $x          - [optional] X position on monitor. Default is -1 (centers horizontally)
;                  $y          - [optional] Y position on monitor. Default is -1 (centers vertically)
;                  $bUseWorkArea - [optional] Use work area instead of full bounds. Default is True
; Return values..: Success     - 1
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Window not visible or not top-level
;                  |@error = 2 - Invalid monitor index or GetWorkArea/GetBounds failed
;                  |@error = 3 - WinGetPos failed (could not get window position/size)
;                  |@error = 4 - Window too large to fit on monitor
;                  |@error = 5 - WinMove failed
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: If both $x and $y are -1, window is centered on the monitor.
;                  If $bUseWorkArea is True, positioning is relative to work area (excludes taskbar).
;                  Function ensures window stays within monitor bounds (adjusts if necessary).
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_MoveWindowToAll, _Monitor_IsVisibleWindow, _Monitor_GetWorkArea, _Monitor_GetBounds
; ================================================================================================================================
Func _Monitor_MoveWindowToScreen($vTitle, $sText = "", $iMonitor = -1, $x = -1, $y = -1, $bUseWorkArea = True)
    Local $hWnd = IsHWnd($vTitle) ? $vTitle : WinGetHandle($vTitle, $sText)
    If Not _Monitor_IsVisibleWindow($hWnd) Then Return SetError(1, 0, 0)

    If $iMonitor = -1 Then $iMonitor = 1
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(2, 0, 0)

    Local $aWinPos = WinGetPos($hWnd)
    If @error Or Not IsArray($aWinPos) Then Return SetError(3, 0, 0)

    Local $iLeft, $iTop, $iRight, $iBottom
    If $bUseWorkArea Then
        If Not _Monitor_GetWorkArea($iMonitor, $iLeft, $iTop, $iRight, $iBottom) Then Return SetError(2, @error, 0)
    Else
        If Not _Monitor_GetBounds($iMonitor, $iLeft, $iTop, $iRight, $iBottom) Then Return SetError(2, @error, 0)
    EndIf

    Local $iWidth = $iRight - $iLeft
    Local $iHeight = $iBottom - $iTop

    ; Check if window fits on monitor
    If $aWinPos[2] > $iWidth Or $aWinPos[3] > $iHeight Then
        Return SetError(4, 0, 0)
    EndIf

    If $x = -1 Or $y = -1 Then
        $x = $iLeft + ($iWidth - $aWinPos[2]) / 2
        $y = $iTop + ($iHeight - $aWinPos[3]) / 2
    Else
        $x += $iLeft
        $y += $iTop
        ; Ensure window stays within bounds
        If $x + $aWinPos[2] > $iRight Then $x = $iRight - $aWinPos[2]
        If $y + $aWinPos[3] > $iBottom Then $y = $iBottom - $aWinPos[3]
        If $x < $iLeft Then $x = $iLeft
        If $y < $iTop Then $y = $iTop
    EndIf

    WinMove($hWnd, "", $x, $y)
    If @error Then Return SetError(5, 0, 0)
    Return 1
EndFunc   ;==>_Monitor_MoveWindowToScreen

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_MoveWindowToAll
; Description....: Move a visible window sequentially across all monitors with a delay between moves
; Syntax.........: _Monitor_MoveWindowToAll($vTitle [, $sText = "" [, $bCenter = True [, $iDelay = 1000]]])
; Parameters.....: $vTitle     - Window title or handle. Can be HWND, title string, or class string
;                  $sText      - [optional] Window text (for matching with title). Default is ""
;                  $bCenter    - [optional] Center window on each monitor. Default is True
;                              If False, window is positioned at (50, 50) on each monitor
;                  $iDelay     - [optional] Delay in milliseconds between moves. Default is 1000
; Return values..: Success     - 1
;                  Failure     - 0, sets @error = 1 (Window not visible or not top-level)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: This is a demonstration function that moves a window to each monitor in sequence.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
;                  Useful for testing multi-monitor setups or demonstrating window movement.
; Related........: _Monitor_MoveWindowToScreen, _Monitor_IsVisibleWindow, _Monitor_GetCount
; ================================================================================================================================
Func _Monitor_MoveWindowToAll($vTitle, $sText = "", $bCenter = True, $iDelay = 1000)
    Local $hWnd = IsHWnd($vTitle) ? $vTitle : WinGetHandle($vTitle, $sText)
    If Not _Monitor_IsVisibleWindow($hWnd) Then Return SetError(1, 0, 0)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()

    For $i = 1 To $__g_aMonitorList[0][0]
        If $bCenter Then
            _Monitor_MoveWindowToScreen($hWnd, "", $i)
        Else
            _Monitor_MoveWindowToScreen($hWnd, "", $i, 50, 50)
        EndIf
        Sleep($iDelay)
    Next
    Return 1
EndFunc   ;==>_Monitor_MoveWindowToAll
#EndRegion --- Window Management Functions ---

#Region --- Display Mode Functions ---
; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_EnumAllDisplayModes
; Description....: Enumerate all available display modes for a monitor
; Syntax.........: _Monitor_EnumAllDisplayModes($iMonitor)
; Parameters.....: $iMonitor   - Monitor index (1..N)
; Return values..: Success     - 2D array with display modes:
;                  |[0][0] - Number of modes found
;                  |[n][0] - Width in pixels for mode n
;                  |[n][1] - Height in pixels for mode n
;                  |[n][2] - Bits per pixel (color depth) for mode n
;                  |[n][3] - Refresh rate in Hz for mode n
;                  |[n][4] - Display mode flags for mode n
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - GetInfo failed (could not get device name)
;                  |@error = 3 - No display modes found
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Returns all supported display modes for the specified monitor.
;                  Use _Monitor_GetDisplaySettings() to get only the current active mode.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetDisplaySettings, _Monitor_GetResolution, _Monitor_GetList
; ================================================================================================================================
Func _Monitor_EnumAllDisplayModes($iMonitor)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    Local $aInfo = _Monitor_GetInfo($iMonitor)
    If @error Then Return SetError(2, 0, 0)
    Local $sDevice = $aInfo[10]

    Local $aModes[1][5]
    $aModes[0][0] = 0
    Local $iIndex = 0

    While True
        Local $aMode = _WinAPI_EnumDisplaySettings($sDevice, $iIndex)
        If @error Then ExitLoop

        ReDim $aModes[$aModes[0][0] + 2][5]
        $aModes[0][0] += 1
        $aModes[$aModes[0][0]][0] = $aMode[0]
        $aModes[$aModes[0][0]][1] = $aMode[1]
        $aModes[$aModes[0][0]][2] = $aMode[2]
        $aModes[$aModes[0][0]][3] = $aMode[3]
        $aModes[$aModes[0][0]][4] = $aMode[4]

        $iIndex += 1
    WEnd

    If $aModes[0][0] = 0 Then Return SetError(3, 0, 0)
    Return $aModes
EndFunc   ;==>_Monitor_EnumAllDisplayModes

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetList
; Description....: Enumerate all connected monitors and fill the global monitor list with their information
; Syntax.........: _Monitor_GetList()
; Parameters.....: None
; Return values..: Success     - Number of monitors detected (>= 1)
;                  Failure     - -1, sets @error = 1 (WinAPI EnumDisplayMonitors call failed)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: This is the core function that initializes the monitor list. Most other functions call this
;                  automatically if the list is not initialized (when $__g_aMonitorList[0][0] = 0).
;                  Populates the global array $__g_aMonitorList with monitor handles, coordinates, and device names.
;                  Also stores virtual desktop bounds and IsPrimary flags for each monitor.
; Related........: _Monitor_Refresh, _Monitor_GetCount, _Monitor_GetInfo
; ================================================================================================================================
Func _Monitor_GetList()
    Local $aMonitors = _WinAPI_EnumDisplayMonitors()
    If @error Or Not IsArray($aMonitors) Or $aMonitors[0][0] = 0 Then
        Return SetError(1, 0, -1)
    EndIf

    ReDim $__g_aMonitorList[$aMonitors[0][0] + 1][7]
    $__g_aMonitorList[0][0] = $aMonitors[0][0]

    Local $l_aVirtual = _Monitor_GetVirtualBounds()
    If @error Then
        ; Fallback calculation
        Local $l_aVirtual[4] = [0, 0, @DesktopWidth, @DesktopHeight]
    EndIf

    Local $l_vRight = $l_aVirtual[0] + $l_aVirtual[2]
    Local $l_vBottom = $l_aVirtual[1] + $l_aVirtual[3]
    $__g_aMonitorList[0][1] = $l_aVirtual[0]
    $__g_aMonitorList[0][2] = $l_aVirtual[1]
    $__g_aMonitorList[0][3] = $l_vRight
    $__g_aMonitorList[0][4] = $l_vBottom
    $__g_aMonitorList[0][5] = $l_aVirtual[2]
    $__g_aMonitorList[0][6] = $l_aVirtual[3]

    For $i = 1 To $aMonitors[0][0]
        Local $hMonitor = $aMonitors[$i][0]
        Local $tRect = $aMonitors[$i][1]

        $__g_aMonitorList[$i][0] = $hMonitor
        $__g_aMonitorList[$i][1] = DllStructGetData($tRect, "Left")
        $__g_aMonitorList[$i][2] = DllStructGetData($tRect, "Top")
        $__g_aMonitorList[$i][3] = DllStructGetData($tRect, "Right")
        $__g_aMonitorList[$i][4] = DllStructGetData($tRect, "Bottom")

        ; Get additional info - Store IsPrimary flag correctly
        Local $aInfo = _WinAPI_GetMonitorInfo($hMonitor)
        If Not @error And IsArray($aInfo) Then
            ; Store IsPrimary flag (0 or 1) instead of pointer
            $__g_aMonitorList[$i][5] = ($aInfo[2] <> 0) ? 1 : 0  ; IsPrimary flag
            $__g_aMonitorList[$i][6] = $aInfo[3] ; Device name
        Else
            ; Safe fallback if GetMonitorInfo fails
            $__g_aMonitorList[$i][5] = 0
            $__g_aMonitorList[$i][6] = ""
        EndIf
    Next

    Return $__g_aMonitorList[0][0]
EndFunc   ;==>_Monitor_GetList

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_ShowInfo
; Description....: Display monitor coordinates and detailed information in a message box and console
; Syntax.........: _Monitor_ShowInfo([$bShowMsgBox = 1 [, $iTimeout = 10]])
; Parameters.....: $bShowMsgBox - [optional] Show message box. Default is 1 (True)
;                  $iTimeout    - [optional] Message box timeout in seconds. Default is 10
; Return values..: Success      - String containing formatted monitor information
;                  Failure      - Empty string "", sets @error = 1 (Enumeration failed)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Displays comprehensive information about all monitors including bounds, work areas,
;                  resolutions, refresh rates, and device names. Information is written to console and
;                  optionally shown in a message box.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
;                  Useful for debugging and displaying system monitor configuration.
; Related........: _Monitor_GetList, _Monitor_GetInfo, _Monitor_GetDisplaySettings
; ================================================================================================================================
Func _Monitor_ShowInfo($bShowMsgBox = 1, $iTimeout = 10)
    If $__g_aMonitorList[0][0] = 0 Then
        If _Monitor_GetList() = -1 Then Return SetError(1, 0, "")
    EndIf
    Local $sMsg = "> Total Monitors: " & $__g_aMonitorList[0][0] & @CRLF & @CRLF
    $sMsg &= StringFormat("+ Virtual Desktop: " & @CRLF & "Left=%d, Top=%d, Right=%d, Bottom=%d, Width=%d, Height=%d", $__g_aMonitorList[0][1], $__g_aMonitorList[0][2], $__g_aMonitorList[0][3], $__g_aMonitorList[0][4], $__g_aMonitorList[0][5], $__g_aMonitorList[0][6]) & @CRLF & @CRLF

    For $i = 1 To $__g_aMonitorList[0][0]
        Local $aInfo = _Monitor_GetInfo($i)
        If @error Then ContinueLoop

        Local $aSettings = _Monitor_GetDisplaySettings($i)
        Local $sResolution = @error ? "N/A" : $aSettings[0] & "x" & $aSettings[1] & " @" & $aSettings[3] & "Hz"

        $sMsg &= StringFormat("+ Monitor %d: %s%s\n", $i, $aInfo[9] ? "(Primary) " : "", $aInfo[10])
        $sMsg &= StringFormat("  Bounds: L=%d, T=%d, R=%d, B=%d (%dx%d)\n", _
                $aInfo[1], $aInfo[2], $aInfo[3], $aInfo[4], _
                $aInfo[3] - $aInfo[1], $aInfo[4] - $aInfo[2])
        $sMsg &= StringFormat("  Work Area: L=%d, T=%d, R=%d, B=%d (%dx%d)\n", _
                $aInfo[5], $aInfo[6], $aInfo[7], $aInfo[8], _
                $aInfo[7] - $aInfo[5], $aInfo[8] - $aInfo[6])
        $sMsg &= "  Resolution: " & $sResolution & @CRLF & @CRLF
    Next
    ConsoleWrite($sMsg)
    If $bShowMsgBox Then MsgBox(64 + 262144, "Monitor Information", $sMsg, $iTimeout)
    Return $sMsg
EndFunc   ;==>_Monitor_ShowInfo

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_Refresh
; Description....: Refresh the monitor list by reloading information from the system
; Syntax.........: _Monitor_Refresh()
; Parameters.....: None
; Return values..: Success     - Number of monitors detected (>= 1)
;                  Failure     - -1, sets @error = 1 (Refresh failed, _Monitor_GetList failed)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Resets the monitor list and re-enumerates all monitors from the system.
;                  Useful when monitors are hot-plugged or display configuration changes.
;                  This forces a complete refresh of all monitor information.
; Related........: _Monitor_GetList, _Monitor_GetCount, _Monitor_IsConnected
; ================================================================================================================================
Func _Monitor_Refresh()
    ; Reset the list
    $__g_aMonitorList[0][0] = 0
    Local $iResult = _Monitor_GetList()
    If $iResult = -1 Then Return SetError(1, 0, -1)
    Return $iResult
EndFunc   ;==>_Monitor_Refresh

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_IsConnected
; Description....: Check if a monitor is still connected and its handle is still valid
; Syntax.........: _Monitor_IsConnected($iMonitor)
; Parameters.....: $iMonitor   - Monitor index (1..N)
; Return values..: Success     - True if monitor is connected and valid, False if disconnected
;                  Failure     - False, sets @error = 1 (Invalid monitor index)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Verifies that the monitor handle is still valid by querying GetMonitorInfo.
;                  Useful for detecting when a monitor has been unplugged or disconnected.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_Refresh, _Monitor_GetList, _Monitor_GetCount
; ================================================================================================================================
Func _Monitor_IsConnected($iMonitor)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, False)

    ; Check if monitor handle is still valid
    Local $hMonitor = $__g_aMonitorList[$iMonitor][0]
    Local $aInfo = _WinAPI_GetMonitorInfo($hMonitor)
    Return (Not @error And IsArray($aInfo))
EndFunc   ;==>_Monitor_IsConnected

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetDPI
; Description....: Get DPI (Dots Per Inch) scaling information for a monitor
; Syntax.........: _Monitor_GetDPI($iMonitor)
; Parameters.....: $iMonitor   - Monitor index (1..N)
; Return values..: Success     - Array with 3 elements:
;                  |[0] - X DPI value
;                  |[1] - Y DPI value
;                  |[2] - Scaling percentage (typically 100, 125, 150, 175, 200, etc.)
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - DPI query failed (fallback uses default 96 DPI)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Tries to use GetDpiForMonitor API (Windows 8.1+) for accurate DPI values.
;                  Falls back to GetDeviceCaps (Windows XP compatible) if GetDpiForMonitor is not available.
;                  On Windows XP/Vista/7/8, function uses GetDeviceCaps which works reliably.
;                  Scaling percentage is calculated as (DPI / 96) * 100.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
;                  Compatible with Windows XP SP2 and later.
; Related........: _Monitor_GetInfo, _Monitor_GetDisplaySettings, _Monitor_GetList
; ================================================================================================================================
Func _Monitor_GetDPI($iMonitor)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    Local $hMonitor = $__g_aMonitorList[$iMonitor][0]
    Local $iDPI_X = 96, $iDPI_Y = 96

    ; Try to get DPI using GetDpiForMonitor (Windows 8.1+ only)
    ; Check OS version first to avoid loading shcore.dll on older systems
    If __Monitor_IsWindows8_1OrGreater() Then
        ; FIXED: Check if shcore.dll exists before calling and close handle properly
        Local $hShCore = DllOpen("shcore.dll")
        If $hShCore <> -1 Then
            Local $aRet = DllCall($hShCore, "long", "GetDpiForMonitor", "handle", $hMonitor, "int", 0, "uint*", 0, "uint*", 0)
            Local $bError = @error
            DllClose($hShCore) ; FIXED: Always close DLL handle to prevent memory leak

            If Not $bError And IsArray($aRet) And $aRet[0] = 0 Then
                $iDPI_X = $aRet[3]
                $iDPI_Y = $aRet[4]
                Local $iScaling = Round(($iDPI_X / 96) * 100)
                Local $aResult[3] = [$iDPI_X, $iDPI_Y, $iScaling]
                Return $aResult
            EndIf
        EndIf
    EndIf

    ; Fallback: Use GetDeviceCaps (compatible with Windows XP and later)
    Local $hDC = DllCall("user32.dll", "handle", "GetDC", "hwnd", 0)
    If @error Or Not IsArray($hDC) Or Not $hDC[0] Then
        ; Return default 96 DPI if GetDC fails
        Local $aResult[3] = [96, 96, 100]
        Return $aResult
    EndIf

    Local $aDPI_X = DllCall("gdi32.dll", "int", "GetDeviceCaps", "handle", $hDC[0], "int", 88) ; LOGPIXELSX
    Local $bErrorX = @error
    Local $aDPI_Y = DllCall("gdi32.dll", "int", "GetDeviceCaps", "handle", $hDC[0], "int", 90) ; LOGPIXELSY
    Local $bErrorY = @error

    ; Release DC handle
    DllCall("user32.dll", "bool", "ReleaseDC", "hwnd", 0, "handle", $hDC[0])

    If Not $bErrorX And IsArray($aDPI_X) And $aDPI_X[0] > 0 Then $iDPI_X = $aDPI_X[0]
    If Not $bErrorY And IsArray($aDPI_Y) And $aDPI_Y[0] > 0 Then $iDPI_Y = $aDPI_Y[0]

    Local $iScaling = Round(($iDPI_X / 96) * 100)
    Local $aResult[3] = [$iDPI_X, $iDPI_Y, $iScaling]
    Return $aResult
EndFunc   ;==>_Monitor_GetDPI

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetOrientation
; Description....: Get the display orientation (rotation angle) for a monitor
; Syntax.........: _Monitor_GetOrientation($iMonitor)
; Parameters.....: $iMonitor   - Monitor index (1..N)
; Return values..: Success     - Orientation angle in degrees:
;                  |0   - Landscape (normal)
;                  |90  - Portrait (rotated 90° clockwise)
;                  |180 - Landscape flipped (rotated 180°)
;                  |270 - Portrait flipped (rotated 270° clockwise / 90° counter-clockwise)
;                  Failure     - -1, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - GetDisplaySettings failed (could not query display mode)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Orientation is extracted from display mode flags.
;                  Most monitors typically return 0 (landscape) unless rotated in Windows display settings.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetDisplaySettings, _Monitor_GetInfo, _Monitor_GetList
; ================================================================================================================================
Func _Monitor_GetOrientation($iMonitor)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, -1)

    Local $aSettings = _Monitor_GetDisplaySettings($iMonitor)
    If @error Then Return SetError(2, 0, -1)

    ; DisplayMode field contains orientation information
    ; DM_DISPLAYORIENTATION values: 0=0°, 1=90°, 2=180°, 3=270°
    Local $iOrientation = 0
    If IsArray($aSettings) And UBound($aSettings) > 4 Then
        ; Check display mode flags for orientation
        Local $iDisplayMode = $aSettings[4]
        ; Orientation is stored in bits 8-9 of display mode
        $iOrientation = BitAND(BitShift($iDisplayMode, 8), 3) * 90
    EndIf

    Return $iOrientation
EndFunc   ;==>_Monitor_GetOrientation
#EndRegion --- Display Mode Functions ---

#Region --- Display Settings Functions (Require Admin) ---
; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_Enable
; Description....: Enable (attach) a monitor that has been disabled
; Syntax.........: _Monitor_Enable($iMonitor)
; Parameters.....: $iMonitor   - Monitor index (1..N)
; Return values..: Success     - 1
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - GetInfo failed (could not get device name)
;                  |@error = 3 - ChangeDisplaySettingsEx failed
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: *** REQUIRES ADMINISTRATOR PRIVILEGES (UAC ELEVATION) ***
;                  Enables a monitor by setting CDS_UPDATEREGISTRY and CDS_NORESET flags.
;                  Changes are applied immediately.
;                  May not work with some proprietary graphics drivers (NVIDIA, AMD control panels).
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_Disable, _Monitor_SetResolution, _Monitor_Refresh
; ================================================================================================================================
Func _Monitor_Enable($iMonitor)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    Local $aInfo = _Monitor_GetInfo($iMonitor)
    If @error Then Return SetError(2, 0, 0)
    Local $sDevice = $aInfo[10]

    ; Get current settings
    Local $aSettings = _Monitor_GetDisplaySettings($iMonitor, $ENUM_REGISTRY_SETTINGS)
    If @error Then $aSettings = _Monitor_GetDisplaySettings($iMonitor)

    ; Create DEVMODE structure
    Local $tDEVMODE = DllStructCreate( _
        "wchar DeviceName[32];" & _
        "ushort SpecVersion;" & _
        "ushort DriverVersion;" & _
        "ushort Size;" & _
        "ushort DriverExtra;" & _
        "dword Fields;" & _
        "short Orientation;" & _
        "short PaperSize;" & _
        "short PaperLength;" & _
        "short PaperWidth;" & _
        "short Scale;" & _
        "short Copies;" & _
        "short DefaultSource;" & _
        "short PrintQuality;" & _
        "short Color;" & _
        "short Duplex;" & _
        "short YResolution;" & _
        "short TTOption;" & _
        "short Collate;" & _
        "wchar FormName[32];" & _
        "ushort LogPixels;" & _
        "dword BitsPerPel;" & _
        "dword PelsWidth;" & _
        "dword PelsHeight;" & _
        "dword DisplayFlags;" & _
        "dword DisplayFrequency")

    If @error Then Return SetError(3, 0, 0)

    DllStructSetData($tDEVMODE, "Size", DllStructGetSize($tDEVMODE))
    DllStructSetData($tDEVMODE, "Fields", BitOR($DM_PELSWIDTH, $DM_PELSHEIGHT, $DM_POSITION))

    If IsArray($aSettings) Then
        DllStructSetData($tDEVMODE, "PelsWidth", $aSettings[0])
        DllStructSetData($tDEVMODE, "PelsHeight", $aSettings[1])
    EndIf

    ; CDS_UPDATEREGISTRY | CDS_NORESET
    Local $aRet = DllCall("user32.dll", "long", "ChangeDisplaySettingsExW", "wstr", $sDevice, "struct*", $tDEVMODE, "hwnd", 0, "dword", BitOR($CDS_UPDATEREGISTRY, $CDS_NORESET), "ptr", 0)
    If @error Or Not IsArray($aRet) Or $aRet[0] <> $DISP_CHANGE_SUCCESSFUL Then Return SetError(3, 0, 0)

    ; Apply changes
    DllCall("user32.dll", "long", "ChangeDisplaySettingsExW", "ptr", 0, "ptr", 0, "hwnd", 0, "dword", 0, "ptr", 0)

    Sleep(500)
    _Monitor_Refresh()
    Return 1
EndFunc   ;==>_Monitor_Enable

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_Disable
; Description....: Disable (detach) a monitor
; Syntax.........: _Monitor_Disable($iMonitor)
; Parameters.....: $iMonitor   - Monitor index (1..N)
; Return values..: Success     - 1
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - Cannot disable primary monitor
;                  |@error = 3 - GetInfo failed (could not get device name)
;                  |@error = 4 - ChangeDisplaySettingsEx failed
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: *** REQUIRES ADMINISTRATOR PRIVILEGES (UAC ELEVATION) ***
;                  Disables a monitor by setting it to 0x0 resolution.
;                  Cannot disable the primary monitor (returns @error = 2).
;                  May not work with some proprietary graphics drivers.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_Enable, _Monitor_SetPrimary, _Monitor_Refresh
; ================================================================================================================================
Func _Monitor_Disable($iMonitor)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    ; Cannot disable primary monitor
    If $__g_aMonitorList[$iMonitor][5] = 1 Then Return SetError(2, 0, 0)

    Local $aInfo = _Monitor_GetInfo($iMonitor)
    If @error Then Return SetError(3, 0, 0)
    Local $sDevice = $aInfo[10]

    ; Create DEVMODE structure with 0x0 resolution (disable)
    Local $tDEVMODE = DllStructCreate( _
        "wchar DeviceName[32];" & _
        "ushort SpecVersion;" & _
        "ushort DriverVersion;" & _
        "ushort Size;" & _
        "ushort DriverExtra;" & _
        "dword Fields;" & _
        "short Orientation;" & _
        "short PaperSize;" & _
        "short PaperLength;" & _
        "short PaperWidth;" & _
        "short Scale;" & _
        "short Copies;" & _
        "short DefaultSource;" & _
        "short PrintQuality;" & _
        "short Color;" & _
        "short Duplex;" & _
        "short YResolution;" & _
        "short TTOption;" & _
        "short Collate;" & _
        "wchar FormName[32];" & _
        "ushort LogPixels;" & _
        "dword BitsPerPel;" & _
        "dword PelsWidth;" & _
        "dword PelsHeight;" & _
        "dword DisplayFlags;" & _
        "dword DisplayFrequency")

    If @error Then Return SetError(4, 0, 0)

    DllStructSetData($tDEVMODE, "Size", DllStructGetSize($tDEVMODE))
    DllStructSetData($tDEVMODE, "Fields", BitOR($DM_PELSWIDTH, $DM_PELSHEIGHT, $DM_POSITION))
    DllStructSetData($tDEVMODE, "PelsWidth", 0)
    DllStructSetData($tDEVMODE, "PelsHeight", 0)

    ; CDS_UPDATEREGISTRY | CDS_NORESET
    Local $aRet = DllCall("user32.dll", "long", "ChangeDisplaySettingsExW", "wstr", $sDevice, "struct*", $tDEVMODE, "hwnd", 0, "dword", BitOR($CDS_UPDATEREGISTRY, $CDS_NORESET), "ptr", 0)
    If @error Or Not IsArray($aRet) Or $aRet[0] <> $DISP_CHANGE_SUCCESSFUL Then Return SetError(4, 0, 0)

    ; Apply changes
    DllCall("user32.dll", "long", "ChangeDisplaySettingsExW", "ptr", 0, "ptr", 0, "hwnd", 0, "dword", 0, "ptr", 0)

    Sleep(500)
    _Monitor_Refresh()
    Return 1
EndFunc   ;==>_Monitor_Disable

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_SetResolution
; Description....: Set monitor resolution and refresh rate
; Syntax.........: _Monitor_SetResolution($iMonitor, $iWidth, $iHeight[, $iBitsPerPixel = 32[, $iFrequency = 0]])
; Parameters.....: $iMonitor      - Monitor index (1..N)
;                  $iWidth        - Width in pixels
;                  $iHeight       - Height in pixels
;                  $iBitsPerPixel - [optional] Color depth (bits per pixel). Default is 32
;                                  Valid values: 16, 24, 32
;                  $iFrequency    - [optional] Refresh rate in Hz. Default is 0 (use current/default)
; Return values..: Success        - 1
;                  Failure        - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - GetInfo failed (could not get device name)
;                  |@error = 3 - Invalid resolution parameters
;                  |@error = 4 - Test mode failed (resolution not supported), @extended = Windows error code
;                  |@error = 5 - Apply mode failed, @extended = Windows error code
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: *** REQUIRES ADMINISTRATOR PRIVILEGES (UAC ELEVATION) ***
;                  Changes display resolution for specified monitor.
;                  Function tests the resolution first before applying (CDS_TEST flag).
;                  Use _Monitor_EnumAllDisplayModes() to check supported resolutions.
;                  May not work with some proprietary graphics drivers.
;                  @extended returns Windows error codes:
;                  |DISP_CHANGE_SUCCESSFUL = 0
;                  |DISP_CHANGE_RESTART = 1 (requires restart)
;                  |DISP_CHANGE_FAILED = -1
;                  |DISP_CHANGE_BADMODE = -2 (mode not supported)
;                  |DISP_CHANGE_NOTUPDATED = -3
;                  |DISP_CHANGE_BADFLAGS = -4
;                  |DISP_CHANGE_BADPARAM = -5
;                  |DISP_CHANGE_BADDUALVIEW = -6
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetResolution, _Monitor_EnumAllDisplayModes, _Monitor_GetDisplaySettings
; ================================================================================================================================
Func _Monitor_SetResolution($iMonitor, $iWidth, $iHeight, $iBitsPerPixel = 32, $iFrequency = 0)
    ; FIXED: Added parameter validation
    If Not IsInt($iMonitor) Or Not IsInt($iWidth) Or Not IsInt($iHeight) Then Return SetError(3, 0, 0)
    If $iBitsPerPixel <> 16 And $iBitsPerPixel <> 24 And $iBitsPerPixel <> 32 Then Return SetError(3, 0, 0)

    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)
    If $iWidth < 1 Or $iHeight < 1 Then Return SetError(3, 0, 0)

    Local $aInfo = _Monitor_GetInfo($iMonitor)
    If @error Then Return SetError(2, 0, 0)
    Local $sDevice = $aInfo[10]

    ; FIXED: Corrected DEVMODE structure definition
    Local $tDEVMODE = DllStructCreate( _
        "wchar DeviceName[32];" & _
        "ushort SpecVersion;" & _
        "ushort DriverVersion;" & _
        "ushort Size;" & _
        "ushort DriverExtra;" & _
        "dword Fields;" & _
        "short Orientation;" & _
        "short PaperSize;" & _
        "short PaperLength;" & _
        "short PaperWidth;" & _
        "short Scale;" & _
        "short Copies;" & _
        "short DefaultSource;" & _
        "short PrintQuality;" & _
        "short Color;" & _
        "short Duplex;" & _
        "short YResolution;" & _
        "short TTOption;" & _
        "short Collate;" & _
        "wchar FormName[32];" & _
        "ushort LogPixels;" & _
        "dword BitsPerPel;" & _
        "dword PelsWidth;" & _
        "dword PelsHeight;" & _
        "dword DisplayFlags;" & _
        "dword DisplayFrequency")

    If @error Then Return SetError(4, 0, 0)

    ; Initialize structure
    DllStructSetData($tDEVMODE, "Size", DllStructGetSize($tDEVMODE))

    ; Set fields flag
    Local $iFields = BitOR($DM_PELSWIDTH, $DM_PELSHEIGHT, $DM_BITSPERPEL)
    If $iFrequency > 0 Then $iFields = BitOR($iFields, $DM_DISPLAYFREQUENCY)

    DllStructSetData($tDEVMODE, "Fields", $iFields)
    DllStructSetData($tDEVMODE, "PelsWidth", $iWidth)
    DllStructSetData($tDEVMODE, "PelsHeight", $iHeight)
    DllStructSetData($tDEVMODE, "BitsPerPel", $iBitsPerPixel)
    If $iFrequency > 0 Then DllStructSetData($tDEVMODE, "DisplayFrequency", $iFrequency)

    ; Test first (CDS_TEST)
    Local $aRet = DllCall("user32.dll", "long", "ChangeDisplaySettingsExW", _
        "wstr", $sDevice, _
        "struct*", $tDEVMODE, _
        "hwnd", 0, _
        "dword", $CDS_TEST, _
        "ptr", 0)

    If @error Then Return SetError(4, @error, 0)
    If Not IsArray($aRet) Then Return SetError(4, 0, 0)

    Local $iTestResult = $aRet[0]
    If $iTestResult <> $DISP_CHANGE_SUCCESSFUL Then
        Return SetError(4, $iTestResult, 0)
    EndIf

    ; Apply changes (CDS_UPDATEREGISTRY)
    $aRet = DllCall("user32.dll", "long", "ChangeDisplaySettingsExW", _
        "wstr", $sDevice, _
        "struct*", $tDEVMODE, _
        "hwnd", 0, _
        "dword", $CDS_UPDATEREGISTRY, _
        "ptr", 0)

    If @error Then Return SetError(5, @error, 0)
    If Not IsArray($aRet) Then Return SetError(5, 0, 0)
    If $aRet[0] <> $DISP_CHANGE_SUCCESSFUL Then Return SetError(5, $aRet[0], 0)

    Sleep(500) ; Wait for display to settle
    _Monitor_Refresh()
    Return 1
EndFunc   ;==>_Monitor_SetResolution

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_SetPrimary
; Description....: Set a monitor as the primary monitor
; Syntax.........: _Monitor_SetPrimary($iMonitor)
; Parameters.....: $iMonitor   - Monitor index (1..N)
; Return values..: Success     - 1
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid monitor index
;                  |@error = 2 - GetInfo failed (could not get device name)
;                  |@error = 3 - ChangeDisplaySettingsEx failed
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: *** REQUIRES ADMINISTRATOR PRIVILEGES (UAC ELEVATION) ***
;                  Sets specified monitor as primary display (where taskbar appears).
;                  This also sets the monitor position to (0,0) in virtual coordinates.
;                  May not work with some proprietary graphics drivers.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetPrimary, _Monitor_GetInfo, _Monitor_Refresh
; ================================================================================================================================
Func _Monitor_SetPrimary($iMonitor)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMonitor < 1 Or $iMonitor > $__g_aMonitorList[0][0] Then Return SetError(1, 0, 0)

    Local $aInfo = _Monitor_GetInfo($iMonitor)
    If @error Then Return SetError(2, 0, 0)
    Local $sDevice = $aInfo[10]

    ; Get current settings
    Local $aSettings = _Monitor_GetDisplaySettings($iMonitor)
    If @error Then Return SetError(2, 0, 0)

    ; Create DEVMODE structure
    Local $tDEVMODE = DllStructCreate( _
        "wchar DeviceName[32];" & _
        "ushort SpecVersion;" & _
        "ushort DriverVersion;" & _
        "ushort Size;" & _
        "ushort DriverExtra;" & _
        "dword Fields;" & _
        "short Orientation;" & _
        "short PaperSize;" & _
        "short PaperLength;" & _
        "short PaperWidth;" & _
        "short Scale;" & _
        "short Copies;" & _
        "short DefaultSource;" & _
        "short PrintQuality;" & _
        "short Color;" & _
        "short Duplex;" & _
        "short YResolution;" & _
        "short TTOption;" & _
        "short Collate;" & _
        "wchar FormName[32];" & _
        "ushort LogPixels;" & _
        "dword BitsPerPel;" & _
        "dword PelsWidth;" & _
        "dword PelsHeight;" & _
        "dword DisplayFlags;" & _
        "dword DisplayFrequency")

    If @error Then Return SetError(3, 0, 0)

    DllStructSetData($tDEVMODE, "Size", DllStructGetSize($tDEVMODE))
    DllStructSetData($tDEVMODE, "Fields", BitOR($DM_PELSWIDTH, $DM_PELSHEIGHT, $DM_POSITION))
    DllStructSetData($tDEVMODE, "PelsWidth", $aSettings[0])
    DllStructSetData($tDEVMODE, "PelsHeight", $aSettings[1])
    DllStructSetData($tDEVMODE, "Orientation", 0) ; Position X = 0
    DllStructSetData($tDEVMODE, "PaperSize", 0)   ; Position Y = 0

    ; CDS_UPDATEREGISTRY | CDS_SET_PRIMARY
    Local $aRet = DllCall("user32.dll", "long", "ChangeDisplaySettingsExW", "wstr", $sDevice, "struct*", $tDEVMODE, "hwnd", 0, "dword", BitOR($CDS_UPDATEREGISTRY, $CDS_SET_PRIMARY), "ptr", 0)
    If @error Or Not IsArray($aRet) Or $aRet[0] <> $DISP_CHANGE_SUCCESSFUL Then Return SetError(3, 0, 0)

    Sleep(500)
    _Monitor_Refresh()
    Return 1
EndFunc   ;==>_Monitor_SetPrimary

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetDisplayMode
; Description....: Get current display mode (duplicate, extend, internal only, external only)
; Syntax.........: _Monitor_GetDisplayMode()
; Parameters.....: None
; Return values..: Success     - Integer representing display mode:
;                  |1 - Internal only (laptop screen only)
;                  |2 - Duplicate/Clone (same content on all displays)
;                  |3 - Extend (extended desktop across displays)
;                  |4 - External only (external monitor only)
;                  Failure     - 0, sets @error = 1 (Could not determine mode)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Analyzes monitor positions to determine display mode.
;                  Duplicate mode: All monitors at position (0,0)
;                  Extend mode: Monitors at different positions
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_SetDisplayMode, _Monitor_GetCount, _Monitor_GetBounds
; ================================================================================================================================
Func _Monitor_GetDisplayMode()
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()

    Local $iCount = $__g_aMonitorList[0][0]
    If $iCount = 0 Then Return SetError(1, 0, 0)
    If $iCount = 1 Then Return 1 ; Single monitor (internal only)

    ; Check if all monitors at same position (duplicate mode)
    Local $bAllSamePos = True
    For $i = 2 To $iCount
        If $__g_aMonitorList[$i][1] <> $__g_aMonitorList[1][1] Or $__g_aMonitorList[$i][2] <> $__g_aMonitorList[1][2] Then
            $bAllSamePos = False
            ExitLoop
        EndIf
    Next

    If $bAllSamePos Then Return 2 ; Duplicate/Clone mode
    Return 3 ; Extend mode
EndFunc   ;==>_Monitor_GetDisplayMode

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_SetDisplayMode
; Description....: Set display mode (duplicate or extend)
; Syntax.........: _Monitor_SetDisplayMode($iMode)
; Parameters.....: $iMode      - Display mode:
;                  |2 - Duplicate/Clone (same content on all displays)
;                  |3 - Extend (extended desktop across displays)
; Return values..: Success     - 1
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - Invalid mode parameter
;                  |@error = 2 - GetInfo failed
;                  |@error = 3 - ChangeDisplaySettingsEx failed
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: *** REQUIRES ADMINISTRATOR PRIVILEGES (UAC ELEVATION) ***
;                  Duplicate mode sets all monitors to position (0,0).
;                  Extend mode arranges monitors horizontally from left to right.
;                  May not work with some proprietary graphics drivers.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetDisplayMode, _Monitor_ApplyLayoutHorizontal, _Monitor_Refresh
; ================================================================================================================================
Func _Monitor_SetDisplayMode($iMode)
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $iMode < 2 Or $iMode > 3 Then Return SetError(1, 0, 0)

    If $iMode = 2 Then
        ; Duplicate mode - set all monitors to (0,0)
        For $i = 1 To $__g_aMonitorList[0][0]
            Local $aInfo = _Monitor_GetInfo($i)
            If @error Then Return SetError(2, 0, 0)
            Local $sDevice = $aInfo[10]

            Local $aSettings = _Monitor_GetDisplaySettings($i)
            If @error Then ContinueLoop

            Local $tDEVMODE = DllStructCreate( _
                "wchar DeviceName[32];" & _
                "ushort SpecVersion;" & _
                "ushort DriverVersion;" & _
                "ushort Size;" & _
                "ushort DriverExtra;" & _
                "dword Fields;" & _
                "short Orientation;" & _
                "short PaperSize;" & _
                "short PaperLength;" & _
                "short PaperWidth;" & _
                "short Scale;" & _
                "short Copies;" & _
                "short DefaultSource;" & _
                "short PrintQuality;" & _
                "short Color;" & _
                "short Duplex;" & _
                "short YResolution;" & _
                "short TTOption;" & _
                "short Collate;" & _
                "wchar FormName[32];" & _
                "ushort LogPixels;" & _
                "dword BitsPerPel;" & _
                "dword PelsWidth;" & _
                "dword PelsHeight;" & _
                "dword DisplayFlags;" & _
                "dword DisplayFrequency")

            If @error Then ContinueLoop

            DllStructSetData($tDEVMODE, "Size", DllStructGetSize($tDEVMODE))
            DllStructSetData($tDEVMODE, "Fields", BitOR($DM_PELSWIDTH, $DM_PELSHEIGHT, $DM_POSITION))
            DllStructSetData($tDEVMODE, "PelsWidth", $aSettings[0])
            DllStructSetData($tDEVMODE, "PelsHeight", $aSettings[1])
            DllStructSetData($tDEVMODE, "Orientation", 0) ; Position X = 0
            DllStructSetData($tDEVMODE, "PaperSize", 0)   ; Position Y = 0

            ; CDS_UPDATEREGISTRY | CDS_NORESET
            DllCall("user32.dll", "long", "ChangeDisplaySettingsExW", "wstr", $sDevice, "struct*", $tDEVMODE, "hwnd", 0, "dword", BitOR($CDS_UPDATEREGISTRY, $CDS_NORESET), "ptr", 0)
        Next

        ; Apply all changes
        DllCall("user32.dll", "long", "ChangeDisplaySettingsExW", "ptr", 0, "ptr", 0, "hwnd", 0, "dword", 0, "ptr", 0)

    ElseIf $iMode = 3 Then
        ; Extend mode - arrange horizontally
        Local $iCurrentX = 0
        For $i = 1 To $__g_aMonitorList[0][0]
            Local $aInfo = _Monitor_GetInfo($i)
            If @error Then Return SetError(2, 0, 0)
            Local $sDevice = $aInfo[10]

            Local $aSettings = _Monitor_GetDisplaySettings($i)
            If @error Then ContinueLoop

            Local $tDEVMODE = DllStructCreate( _
                "wchar DeviceName[32];" & _
                "ushort SpecVersion;" & _
                "ushort DriverVersion;" & _
                "ushort Size;" & _
                "ushort DriverExtra;" & _
                "dword Fields;" & _
                "short Orientation;" & _
                "short PaperSize;" & _
                "short PaperLength;" & _
                "short PaperWidth;" & _
                "short Scale;" & _
                "short Copies;" & _
                "short DefaultSource;" & _
                "short PrintQuality;" & _
                "short Color;" & _
                "short Duplex;" & _
                "short YResolution;" & _
                "short TTOption;" & _
                "short Collate;" & _
                "wchar FormName[32];" & _
                "ushort LogPixels;" & _
                "dword BitsPerPel;" & _
                "dword PelsWidth;" & _
                "dword PelsHeight;" & _
                "dword DisplayFlags;" & _
                "dword DisplayFrequency")

            If @error Then ContinueLoop

            DllStructSetData($tDEVMODE, "Size", DllStructGetSize($tDEVMODE))
            DllStructSetData($tDEVMODE, "Fields", BitOR($DM_PELSWIDTH, $DM_PELSHEIGHT, $DM_POSITION))
            DllStructSetData($tDEVMODE, "PelsWidth", $aSettings[0])
            DllStructSetData($tDEVMODE, "PelsHeight", $aSettings[1])
            DllStructSetData($tDEVMODE, "Orientation", $iCurrentX) ; Position X
            DllStructSetData($tDEVMODE, "PaperSize", 0)            ; Position Y = 0

            ; CDS_UPDATEREGISTRY | CDS_NORESET
            DllCall("user32.dll", "long", "ChangeDisplaySettingsExW", "wstr", $sDevice, "struct*", $tDEVMODE, "hwnd", 0, "dword", BitOR($CDS_UPDATEREGISTRY, $CDS_NORESET), "ptr", 0)

            $iCurrentX += $aSettings[0]
        Next

        ; Apply all changes
        DllCall("user32.dll", "long", "ChangeDisplaySettingsExW", "ptr", 0, "ptr", 0, "hwnd", 0, "dword", 0, "ptr", 0)
    EndIf

    Sleep(500)
    _Monitor_Refresh()
    Return 1
EndFunc   ;==>_Monitor_SetDisplayMode
#EndRegion --- Display Settings Functions (Require Admin) ---

#Region --- Layout Management Functions ---
; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetLayout
; Description....: Get current display layout configuration
; Syntax.........: _Monitor_GetLayout()
; Parameters.....: None
; Return values..: Success     - 2D array with layout information:
;                  |[0][0] - Number of monitors
;                  |[n][0] - Monitor index
;                  |[n][1] - Left coordinate
;                  |[n][2] - Top coordinate
;                  |[n][3] - Width
;                  |[n][4] - Height
;                  |[n][5] - IsPrimary flag (1 or 0)
;                  Failure     - 0, sets @error = 1 (No monitors found)
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Returns the current physical layout of all monitors.
;                  Useful for saving and restoring monitor configurations.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_SaveLayout, _Monitor_LoadLayout, _Monitor_GetLayoutDescription
; ================================================================================================================================
Func _Monitor_GetLayout()
    If $__g_aMonitorList[0][0] = 0 Then _Monitor_GetList()
    If $__g_aMonitorList[0][0] = 0 Then Return SetError(1, 0, 0)

    Local $aLayout[$__g_aMonitorList[0][0] + 1][6]
    $aLayout[0][0] = $__g_aMonitorList[0][0]

    For $i = 1 To $__g_aMonitorList[0][0]
        $aLayout[$i][0] = $i
        $aLayout[$i][1] = $__g_aMonitorList[$i][1]
        $aLayout[$i][2] = $__g_aMonitorList[$i][2]
        $aLayout[$i][3] = $__g_aMonitorList[$i][3] - $__g_aMonitorList[$i][1]
        $aLayout[$i][4] = $__g_aMonitorList[$i][4] - $__g_aMonitorList[$i][2]
        $aLayout[$i][5] = $__g_aMonitorList[$i][5]
    Next

    Return $aLayout
EndFunc   ;==>_Monitor_GetLayout

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_GetLayoutDescription
; Description....: Get text description of current layout
; Syntax.........: _Monitor_GetLayoutDescription()
; Parameters.....: None
; Return values..: Success     - String describing the layout
;                  Failure     - Empty string "", sets @error = 1
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Returns a human-readable description of monitor layout.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_GetLayout, _Monitor_ShowInfo
; ================================================================================================================================
Func _Monitor_GetLayoutDescription()
    Local $aLayout = _Monitor_GetLayout()
    If @error Then Return SetError(1, 0, "")

    Local $sDesc = "Layout: " & $aLayout[0][0] & " monitor(s)" & @CRLF
    For $i = 1 To $aLayout[0][0]
        $sDesc &= StringFormat("Monitor %d: Pos(%d,%d) Size(%dx%d) %s" & @CRLF, _
            $aLayout[$i][0], $aLayout[$i][1], $aLayout[$i][2], _
            $aLayout[$i][3], $aLayout[$i][4], _
            $aLayout[$i][5] ? "[Primary]" : "")
    Next
    Return $sDesc
EndFunc   ;==>_Monitor_GetLayoutDescription

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_SaveLayout
; Description....: Save current layout to file
; Syntax.........: _Monitor_SaveLayout($sFilePath)
; Parameters.....: $sFilePath  - Full path to save layout file
; Return values..: Success     - 1
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - GetLayout failed
;                  |@error = 2 - File write failed
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: Saves layout in INI format for easy editing.
;                  Function automatically calls _Monitor_GetList() if monitor list is not initialized.
; Related........: _Monitor_LoadLayout, _Monitor_GetLayout
; ================================================================================================================================
Func _Monitor_SaveLayout($sFilePath)
    Local $aLayout = _Monitor_GetLayout()
    If @error Then Return SetError(1, 0, 0)

    Local $hFile = FileOpen($sFilePath, 2) ; Overwrite mode
    If $hFile = -1 Then Return SetError(2, 0, 0)

    FileWriteLine($hFile, "[MonitorLayout]")
    FileWriteLine($hFile, "Count=" & $aLayout[0][0])
    FileWriteLine($hFile, "")

    For $i = 1 To $aLayout[0][0]
        FileWriteLine($hFile, "[Monitor" & $i & "]")
        FileWriteLine($hFile, "Left=" & $aLayout[$i][1])
        FileWriteLine($hFile, "Top=" & $aLayout[$i][2])
        FileWriteLine($hFile, "Width=" & $aLayout[$i][3])
        FileWriteLine($hFile, "Height=" & $aLayout[$i][4])
        FileWriteLine($hFile, "IsPrimary=" & $aLayout[$i][5])
        FileWriteLine($hFile, "")
    Next

    FileClose($hFile)
    Return 1
EndFunc   ;==>_Monitor_SaveLayout

; #FUNCTION# ====================================================================================================================
; Name...........: _Monitor_LoadLayout
; Description....: Load layout from file (NOTE: Does not apply layout, only loads data)
; Syntax.........: _Monitor_LoadLayout($sFilePath)
; Parameters.....: $sFilePath  - Full path to layout file
; Return values..: Success     - 2D array with layout information (same format as _Monitor_GetLayout)
;                  Failure     - 0, sets @error to non-zero:
;                  |@error = 1 - File not found or cannot be read
;                  |@error = 2 - Invalid layout file format
; Author.........: Dao Van Trong - TRONG.PRO
; Remarks........: *** NOTE: This function only LOADS layout data, it does NOT apply it ***
;                  To apply loaded layout, you need to use _Monitor_SetResolution() and
;                  _Monitor_SetPrimary() functions manually for each monitor.
;                  This is because applying layout requires Administrator privileges.
; Related........: _Monitor_SaveLayout, _Monitor_GetLayout
; ================================================================================================================================
Func _Monitor_LoadLayout($sFilePath)
    If Not FileExists($sFilePath) Then Return SetError(1, 0, 0)

    Local $iCount = IniRead($sFilePath, "MonitorLayout", "Count", 0)
    If $iCount < 1 Then Return SetError(2, 0, 0)

    Local $aLayout[$iCount + 1][6]
    $aLayout[0][0] = $iCount

    For $i = 1 To $iCount
        $aLayout[$i][0] = $i
        $aLayout[$i][1] = IniRead($sFilePath, "Monitor" & $i, "Left", 0)
        $aLayout[$i][2] = IniRead($sFilePath, "Monitor" & $i, "Top", 0)
        $aLayout[$i][3] = IniRead($sFilePath, "Monitor" & $i, "Width", 0)
        $aLayout[$i][4] = IniRead($sFilePath, "Monitor" & $i, "Height", 0)
        $aLayout[$i][5] = IniRead($sFilePath, "Monitor" & $i, "IsPrimary", 0)
    Next

    Return $aLayout
EndFunc   ;==>_Monitor_LoadLayout
#EndRegion --- Layout Management Functions ---