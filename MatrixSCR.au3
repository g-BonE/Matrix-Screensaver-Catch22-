#include "include\monitorudf_trong.au3"   ; Trong's multi-monitor UDF

; ===========================================================================
; Catch22 Matrix Screen Saver wrapper (TheMatrixWRP_C22.au3)
; - Runs Matrix.scr windowed (/w) and shows it only on a chosen monitor.
; - Uses monitorudf_trong.au3 + WinMove to cover the monitor and hide title/borders
;   by oversizing/offsetting the window.
; - Controls only numeric/font settings; Messages remain as configured via /c.
; ===========================================================================

Global Const $g_sRegBase   = "HKEY_CURRENT_USER\Software\Catch22\Matrix Screen Saver"
Global Const $g_sSaverPath = @ScriptDir & "\Matrix_C22\Matrix.scr"  ; <-- adjust to your .scr path

; ----------------- PROFILE SELECTION -----------------
; 1 = monitor 1
; 2 = monitor 2
; 3 = monitor 3
Global $g_iProfile = 2 ; default to monitor 2

If $CmdLine[0] >= 1 Then
    Local $iArg = Number($CmdLine[1])
    If $iArg = 1 Or $iArg = 2 Or $iArg = 3 Then $g_iProfile = $iArg
EndIf

; Per-profile settings
Global $g_iTargetMonitor
Global $g_iMessageSpeed   ; 50..500 (0x32..0x1F4), default 302 (0x12E)
Global $g_iMatrixSpeed    ; 1..10   (0x1..0xA),    default 6   (0x6)
Global $g_iDensity        ; 5..48   (0x5..0x30),   default 30  (0x1E)
Global $g_iFontSize       ; 8..30   (0x8..0x1E),   default 12  (0x0C)
Global $g_bRandomize      ; 0/1, default 0
Global $g_bFontBold       ; 0/1
Global $g_sFontName       ; installed font name

Switch $g_iProfile
    Case 1 ; monitor 1
        $g_iTargetMonitor = 1
        $g_iMessageSpeed  = 300
        $g_iMatrixSpeed   = 5
        $g_iDensity       = 30
        $g_iFontSize      = 11
        $g_bRandomize     = True
        $g_bFontBold      = True
        $g_sFontName      = "MS Sans Serif"

    Case 2 ; monitor 2
        $g_iTargetMonitor = 2
        $g_iMessageSpeed  = 300
        $g_iMatrixSpeed   = 5
        $g_iDensity       = 30
        $g_iFontSize      = 11
        $g_bRandomize     = True
        $g_bFontBold      = True
        $g_sFontName      = "MS Sans Serif"

    Case 3 ; monitor 3
        $g_iTargetMonitor = 3
        $g_iMessageSpeed  = 300
        $g_iMatrixSpeed   = 5
        $g_iDensity       = 30
        $g_iFontSize      = 11
        $g_bRandomize     = True
        $g_bFontBold      = True
        $g_sFontName      = "MS Sans Serif"
EndSwitch

; ===========================================================================
; MAIN
; ===========================================================================

; Get target monitor bounds via Trong UDF
Local $iTL, $iTT, $iTR, $iTB, $iTHeight
If Not _GetMonitorRegion($g_iTargetMonitor, $iTL, $iTT, $iTR, $iTB, $iTHeight) Then Exit

; Write numeric/font settings (Messages left untouched)
_WriteCatch22Settings_NoMessages( _
    $g_iMessageSpeed, _
    $g_iMatrixSpeed, _
    $g_iDensity, _
    $g_iFontSize, _
    $g_bRandomize, _
    $g_bFontBold, _
    $g_sFontName)

; Start saver in windowed mode
If Not FileExists($g_sSaverPath) Then
    MsgBox(16, "Error", "Matrix.scr not found at:" & @CRLF & $g_sSaverPath)
    Exit
EndIf

Local $iPID = Run('"' & $g_sSaverPath & '" /w', "", @SW_SHOW)
If @error Or $iPID = 0 Then Exit

; Wait for saver window to appear
Local $hWnd = 0, $tStart = TimerInit()
While TimerDiff($tStart) < 5000
    $hWnd = _FindSaverWindow($iPID)
    If $hWnd <> 0 Then ExitLoop
    Sleep(100)
WEnd
If $hWnd = 0 Then Exit

; Make it "fullscreen": oversize to hide title/borders off-screen
_PositionWindowCoverMonitor($hWnd, $iTL, $iTT, $iTR, $iTB)

; Optionally wait until saver exits
; ProcessWaitClose($iPID)
Exit


; ===========================================================================
; FUNCTIONS
; ===========================================================================

Func _GetMonitorRegion($iMon, ByRef $L, ByRef $T, ByRef $R, ByRef $B, ByRef $H)
    If _Monitor_GetList() = -1 Then
        MsgBox(16, "Error", "Monitor UDF: _Monitor_GetList() failed. @error=" & @error)
        Return False
    EndIf

    Local $iCount = _Monitor_GetCount()
    If @error Or $iCount = 0 Then
        MsgBox(16, "Error", "Monitor UDF: _Monitor_GetCount() failed or no monitors.")
        Return False
    EndIf

    If $iMon < 1 Or $iMon > $iCount Then
        MsgBox(16, "Error", "Target monitor " & $iMon & " out of range 1.." & $iCount)
        Return False
    EndIf

    If Not _Monitor_GetBounds($iMon, $L, $T, $R, $B) Then
        MsgBox(16, "Error", "Monitor UDF: _Monitor_GetBounds(" & $iMon & ") failed. @error=" & @error)
        Return False
    EndIf

    $H = $B - $T
    Return True
EndFunc   ;==>_GetMonitorRegion


Func _FindSaverWindow($iPID)
    Local $aList = WinList()
    For $i = 1 To $aList[0][0]
        If $aList[$i][0] = "" Then ContinueLoop
        If WinGetProcess($aList[$i][1]) = $iPID Then
            Return $aList[$i][1]
        EndIf
    Next
    Return 0
EndFunc   ;==>_FindSaverWindow


Func _Clamp($iVal, $iMin, $iMax)
    If $iVal < $iMin Then Return $iMin
    If $iVal > $iMax Then Return $iMax
    Return $iVal
EndFunc


; Oversize/offset window so caption/borders are off-screen, but monitor is filled
Func _PositionWindowCoverMonitor($hWnd, $iL, $iT, $iR, $iB)
    ; Basic monitor dimensions
    Local $iMonW = $iR - $iL
    Local $iMonH = $iB - $iT

    ; Overscan margin to hide title/borders
    Local $iMarginLeft   = 10
    Local $iMarginTop    = 40    ; a bit more for title bar
    Local $iMarginRight  = 10
    Local $iMarginBottom = 10

    Local $iX = $iL - $iMarginLeft
    Local $iY = $iT - $iMarginTop
    Local $iW = $iMonW + $iMarginLeft + $iMarginRight
    Local $iH = $iMonH + $iMarginTop + $iMarginBottom

    WinMove($hWnd, "", $iX, $iY, $iW, $iH)
EndFunc   ;==>_PositionWindowCoverMonitor


Func _WriteCatch22Settings_NoMessages( _
        $iMsgSpeed, $iMatrixSpeed, $iDensity, $iFontSize, _
        $bRandomize, $bFontBold, $sFontName)

    ; Clamp to documented ranges
    $iMsgSpeed    = _Clamp($iMsgSpeed,    50, 500)
    $iMatrixSpeed = _Clamp($iMatrixSpeed, 1,  10)
    $iDensity     = _Clamp($iDensity,     5,  48)
    $iFontSize    = _Clamp($iFontSize,    8,  30)

    RegWrite($g_sRegBase, "MessageSpeed", "REG_DWORD", $iMsgSpeed)
    RegWrite($g_sRegBase, "MatrixSpeed",  "REG_DWORD", $iMatrixSpeed)
    RegWrite($g_sRegBase, "Density",      "REG_DWORD", $iDensity)
    RegWrite($g_sRegBase, "FontSize",     "REG_DWORD", $iFontSize)

    RegWrite($g_sRegBase, "Randomize", "REG_DWORD", $bRandomize ? 1 : 0)
    RegWrite($g_sRegBase, "FontBold",  "REG_DWORD", $bFontBold  ? 1 : 0)
    RegWrite($g_sRegBase, "FontName",  "REG_SZ",    $sFontName)
    RegWrite($g_sRegBase, "Previews",  "REG_DWORD", 1)

    ; Messages left unchanged – configure once via /c dialog.
EndFunc   ;==>_WriteCatch22Settings_NoMessages
