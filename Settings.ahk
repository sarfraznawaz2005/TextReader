; Settings Management

LoadSettings() {
    global g_WorkingFolder
    
    iniFile := A_ScriptDir . "\TextReader.ini"
    
    g_WorkingFolder := IniRead(iniFile, "Settings", "WorkingFolder", "")
}

SaveSettings() {
    global g_WorkingFolder
    
    iniFile := A_ScriptDir . "\TextReader.ini"
    
    IniWrite(g_WorkingFolder, iniFile, "Settings", "WorkingFolder")
}

ShowSettingsWindow() {
    global g_WorkingFolder, g_MainGui
    
    ; Create settings GUI
    settingsGui := Gui("+Owner" . g_MainGui.Hwnd . " -MinimizeBox -MaximizeBox", "Settings")
    settingsGui.SetFont("s10", "Segoe UI")
    settingsGui.BackColor := "0xF5F5F5"
    
    ; Working folder setting
    settingsGui.AddText("x20 y20 w100 h20", "Working Folder:")
    txtFolder := settingsGui.AddEdit("x20 y45 w300 h23 +ReadOnly", g_WorkingFolder)
    btnBrowse := settingsGui.AddButton("x330 y44 w80 h25", "Browse...")
    btnBrowse.OnEvent("Click", (*) => BrowseFolder(txtFolder))
    
    ; Floating Button Transparency setting
    INI_FILE_PATH := A_ScriptDir . "\TextReader.ini"
    currentTrans255 := IniRead(INI_FILE_PATH, "FloatingButton", "Transparency", "128") ; Default to 128 (50%)
    currentTrans100 := Round(currentTrans255 / 2.55)
    if (currentTrans100 < 0 || currentTrans100 > 100)
      currentTrans100 := 50 ; Fallback if INI value is out of expected range

    settingsGui.AddText("x20 y85 w250 h20", "Floating Button Transparency (0-100):")
    txtTransparency := settingsGui.AddEdit("x20 y110 w100 h23 +Number", currentTrans100)
    settingsGui.AddUpDown("Range0-100", currentTrans100)

    ; Floating Button X Position setting
    currentXPos := IniRead(INI_FILE_PATH, "FloatingButton", "X", "100")
    settingsGui.AddText("x20 y150 w250 h20", "Floating Button X Position:")
    txtXPos := settingsGui.AddEdit("x20 y175 w100 h23 +Number", currentXPos)
    settingsGui.AddUpDown("Range0-25000", Integer(currentXPos))

    ; Floating Button Y Position setting
    currentYPos := IniRead(INI_FILE_PATH, "FloatingButton", "Y", "100")
    settingsGui.AddText("x20 y215 w250 h20", "Floating Button Y Position:")
    txtYPos := settingsGui.AddEdit("x20 y240 w100 h23 +Number", currentYPos)
    settingsGui.AddUpDown("Range0-25000", Integer(currentYPos))
    
    ; Buttons
    btnApply := settingsGui.AddButton("x245 y280 w80 h30", "Save")
    btnCancel := settingsGui.AddButton("x335 y280 w80 h30", "Cancel")
    
    ApplyClick(*) {
        ApplySettings(settingsGui, txtFolder.Text, txtTransparency.Text, txtXPos.Text, txtYPos.Text)
    }
    
    CancelClick(*) {
        settingsGui.Destroy()
    }
    
    btnApply.OnEvent("Click", ApplyClick)
    btnCancel.OnEvent("Click", CancelClick)
    
    ; Center and show
    CenterWindow(settingsGui)
    settingsGui.Show("w430 h320")
}

BrowseFolder(txtControl) {
    folder := DirSelect("*" . txtControl.Text, 3, "Select Working Folder")
    if (folder != "") {
        txtControl.Text := folder
    }
}

ApplySettings(gui, newFolder, newTransparency, newXPos, newYPos) {
    global g_WorkingFolder, rtfContent
    
    ; Validate transparency
    transparency := Integer(newTransparency)
    if (transparency < 0 || transparency > 100) {
        MsgBox("Transparency must be between 0 and 100.", "Invalid Transparency", "48")
        return
    }
    
    newXPos := RegExReplace(newXPos, ",", "")
    newYPos := RegExReplace(newYPos, ",", "")

    ; Validate X Position
    xPos := Integer(newXPos)
    if (!IsNumber(newXPos)) {
        MsgBox("X Position must be a number.", "Invalid X Position", "48")
        return
    }

    ; Validate Y Position
    yPos := Integer(newYPos)
    if (!IsNumber(newYPos)) {
        MsgBox("Y Position must be a number.", "Invalid Y Position", "48")
        return
    }
    
    ; Update settings
    g_WorkingFolder := newFolder
    
    ; Save settings
    SaveSettings()
    
    ; Save Floating Button settings directly
    INI_FILE_PATH := A_ScriptDir . "\TextReader.ini"
    IniWrite(Round(transparency * 2.55), INI_FILE_PATH, "FloatingButton", "Transparency")
    IniWrite(xPos, INI_FILE_PATH, "FloatingButton", "X")
    IniWrite(yPos, INI_FILE_PATH, "FloatingButton", "Y")
    
    RefreshFileList()

    ApplyNewSettings(xPos, yPos, Round(transparency * 2.55))
    
    ; Close settings window
    gui.Destroy()
    
    ; Clear welcome message if folder was set
    if (g_WorkingFolder != "" && g_CurrentFile == "") {
        rtfContent.SetText("")
    }
}