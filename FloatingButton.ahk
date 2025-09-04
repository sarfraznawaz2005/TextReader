; Floating Desktop Button

global g_FloatingGui := ""
global g_ButtonX := 0
global g_ButtonY := 0
global g_IsDragging := false
global g_Transparency := 128 ; Default to 50% opaque

; New function to place the button at the bottom-right
SetButtonXY(buttonWidth, buttonHeight, rightMargin := 20, bottomMargin := 20) {
    MonitorGetWorkArea(MonitorGetPrimary(), &L, &T, &R, &B)
    global g_ButtonX, g_ButtonY
    g_ButtonX := R - buttonWidth - rightMargin
    g_ButtonY := B - buttonHeight - bottomMargin
}

; Windows API functions
SetCapture := DllCall.Bind("user32.dll", "SetCapture", "Ptr")
ReleaseCapture := DllCall.Bind("user32.dll", "ReleaseCapture")

CreateFloatingButton() {
    global g_FloatingGui, g_ButtonX, g_ButtonY

    ; Get desktop window handle for parenting
    desktopHwnd := GetDesktopWindowHandle()
    if (!desktopHwnd) {
        LogError("Could not find desktop window handle for parenting the floating button.")
        return ; Can't create the button without a parent
    }

    LoadButtonPosition()
    
    ; Load button position from settings, or place bottom-right if not found/invalid
    iniFile := A_ScriptDir . "\TextReader.ini"
    if !FileExist(iniFile) {    
        SaveButtonPosition()
    }
    
    if (g_FloatingGui)
		g_FloatingGui.Destroy()

    ; Create floating button GUI, parented to the desktop
    g_FloatingGui := Gui("-Caption -Border +ToolWindow +LastFound +E0x80000 +Parent" . desktopHwnd, "")
	g_FloatingGui.Opt("+E0x08000000")
    g_FloatingGui.SetFont("s10 Bold", "Segoe UI")
	
	DllCall("user32\SetParent", "ptr", g_FloatingGui.Hwnd, "ptr", desktopHwnd)

	; auto-reload in case explorer is re-started
	static __last:=0
	OnMessage(DllCall("RegisterWindowMessage","str","TaskbarCreated","uint"), (*) => (A_TickCount-__last<5000?0:(__last:=A_TickCount, SetTimer(() => Reload(), -1000))))

    ; Use one consistent chroma color (name or hex), and remove padding
    CHROMA := "0xDDDDDD"                 ; or use CHROMA := 0xFF00FF
    g_FloatingGui.BackColor := CHROMA
    g_FloatingGui.MarginX := 0
    g_FloatingGui.MarginY := 0

    ; Create button (pin at 0,0 so GUI hugs the control)
    btnFloat := g_FloatingGui.AddButton("x0 y0", "📁 Text Reader")
    btnFloat.SetFont("s10 Bold", "Segoe UI")
    btnFloat.OnEvent("Click", FloatingButton_Click)

    ; Make the GUI draggable
    g_FloatingGui.OnEvent("Close", FloatingGui_Close)

    ; Position and show
    g_FloatingGui.Show("x" . g_ButtonX . " y" . g_ButtonY . " AutoSize")

    WinSetTransColor(CHROMA, g_FloatingGui.Hwnd)
    WinSetTransparent(g_Transparency, g_FloatingGui.Hwnd)
}

ApplyNewSettings(xPos, yPos, g_Transparency) {
    global g_FloatingGui
    
    CHROMA := "0xDDDDDD"                 ; or use CHROMA := 0xFF00FF
    g_FloatingGui.BackColor := CHROMA

    WinSetTransColor(CHROMA, g_FloatingGui.Hwnd)
    WinSetTransparent(g_Transparency, g_FloatingGui.Hwnd)
	
	g_FloatingGui.Show("x" . xPos . " y" . yPos . " AutoSize")
}

FloatingButton_Click(*) {
    global g_MainGui, g_FloatingGui
    g_FloatingGui.Hide()
    
    ; Show or focus main window
    if (WinExist(g_MainGui.Hwnd)) {
        g_MainGui.Show()
        RefreshFileList()
    } else {
        ; Recreate main GUI if it was closed
        ;CreateMainGUI()
        g_MainGui.Show("w1200 h700")
        RefreshFileList()
    }
}

FloatingGui_Close(*) {
    ; Don't actually close, just hide
    g_FloatingGui.Hide()
}

LoadButtonPosition() {
    global g_ButtonX, g_ButtonY, g_Transparency
    
    iniFile := A_ScriptDir . "\TextReader.ini"
    
    ; Read values from INI, default to empty string if not found
    readX := IniRead(iniFile, "FloatingButton", "X", "")
    readY := IniRead(iniFile, "FloatingButton", "Y", "")
    readTransparency := IniRead(iniFile, "FloatingButton", "Transparency", "")
    
    ; Check if values are valid numbers and within screen bounds
    ; Using A_ScreenWidth and A_ScreenHeight for bounds check as they are readily available.
    ; The actual placement will use MonitorGetWorkArea for more precise positioning.
    isValidX := (readX != "" && IsNumber(readX) && Integer(readX) >= 0)
    isValidY := (readY != "" && IsNumber(readY) && Integer(readY) >= 0)
    
    if (isValidX && isValidY) {
        g_ButtonX := Integer(readX)
        g_ButtonY := Integer(readY)
    } else {
        ; If not valid or not found, place at bottom-right
        SetButtonXY(100, 20) ; Button width and height
    }
    
    ; Validate and set transparency
    if (readTransparency != "" && IsNumber(readTransparency) && Integer(readTransparency) >= 0 && Integer(readTransparency) <= 255) {
        g_Transparency := Integer(readTransparency)
    } else {
        g_Transparency := 128 ; Default to 50% if not valid
    }
}

SaveButtonPosition() {
    global g_ButtonX, g_ButtonY, g_Transparency
    
    iniFile := A_ScriptDir . "\TextReader.ini"
    
    IniWrite(g_ButtonX, iniFile, "FloatingButton", "X")
    IniWrite(g_ButtonY, iniFile, "FloatingButton", "Y")
    IniWrite(g_Transparency, iniFile, "FloatingButton", "Transparency")    
}

; Toggle floating button visibility
ToggleFloatingButton() {
    global g_FloatingGui
    
    if (g_FloatingGui.Visible) {
        g_FloatingGui.Hide()
    } else {
        g_FloatingGui.Show("NoActivate")
    }
}
