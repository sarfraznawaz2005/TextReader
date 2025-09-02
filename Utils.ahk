; Utility Functions

ShowInputDialog(title, prompt, defaultValue := "") {
    ; Create input dialog
    inputGui := Gui("+Owner" . g_MainGui.Hwnd . " -MinimizeBox -MaximizeBox", title)
    inputGui.SetFont("s10", "Segoe UI")
    inputGui.BackColor := "0xF5F5F5"
    
    inputGui.AddText("x20 y20 w260 h20", prompt)
    txtInput := inputGui.AddEdit("x20 y45 w260 h23", defaultValue)
    
    btnOK := inputGui.AddButton("x130 y85 w70 h30 Default", "OK")
    btnCancel := inputGui.AddButton("x210 y85 w70 h30", "Cancel")
    
    result := ""
    
    btnOK.OnEvent("Click", OKClick)
    btnCancel.OnEvent("Click", CancelClick)
    
    OKClick(*) {
        result := txtInput.Text
        inputGui.Destroy()
    }
    
    CancelClick(*) {
        result := ""
        inputGui.Destroy()
    }
    
    ; Handle Enter key
    txtInput.OnEvent("Change", InputChange)
    
    InputChange(*) {
        if (GetKeyState("Enter", "P")) {
            result := txtInput.Text
            inputGui.Destroy()
        }
    }
    
    ; Center and show modal
    CenterWindow(inputGui)
    inputGui.Show("w300 h135")
    
    ; Wait for dialog to close
    WinWaitClose(inputGui.Hwnd)
    
    return Trim(result)
}

GetDesktopWindowHandle() {
    ; Ensure Explorer's WorkerW layer exists
    progman := WinExist("ahk_class Progman")
    if progman {
        ; 0x052C is the "spawn WorkerW" message Progman listens for
        DllCall("user32\SendMessageTimeout"
            , "ptr",  progman
            , "uint", 0x052C
            , "ptr",  0
            , "ptr",  0
            , "uint", 0x0002    ; SMTO_BLOCK
            , "uint", 1000
            , "ptr*", 0)
    }

    ; Give Explorer a moment to build WorkerW
    WinWait("ahk_class WorkerW",, 1)

    ; We want the WorkerW that *doesn’t* host SHELLDLL_DefView (that one is behind icons),
    ; OR fall back sensibly if not found yet.
    hwndToUse := 0
    for hwnd in WinGetList("ahk_class WorkerW") {
        if !WinExist("ahk_class SHELLDLL_DefView", "ahk_id " hwnd) {
            hwndToUse := hwnd
            break
        }
    }

    ; Fallbacks if needed
    if !hwndToUse
        hwndToUse := WinExist("ahk_class WorkerW")
    if !hwndToUse
        hwndToUse := WinExist("ahk_class Progman")
    if !hwndToUse
        hwndToUse := DllCall("GetShellWindow", "ptr")

    return hwndToUse
}

; Helper function to format text content
FormatContent(content) {
    ; Basic text formatting - AutoHotkey Edit controls have limited formatting
    ; This function can be expanded if RichEdit control is implemented
    return content
}

; Helper function to validate file names
IsValidFileName(fileName) {
    invalidChars := ['<', '>', ':', '"', '|', '?', '*', '/']
    
    for char in invalidChars {
        if (InStr(fileName, char)) {
            return false
        }
    }
    
    return fileName != "" && StrLen(fileName) <= 255
}

; Helper function to get relative path
GetRelativePath(fullPath, basePath) {
    if (InStr(fullPath, basePath) == 1) {
        return SubStr(fullPath, StrLen(basePath) + 2)
    }
    return fullPath
}

; Helper function to backup file before editing
BackupFile(filePath) {
    try {
        backupPath := filePath . ".backup"
        FileCopy(filePath, backupPath, 1)
        return true
    } catch {
        return false
    }
}

; Helper function to restore from backup
RestoreFromBackup(filePath) {
    try {
        backupPath := filePath . ".backup"
        if (FileExist(backupPath)) {
            FileCopy(backupPath, filePath, 1)
            FileDelete(backupPath)
            return true
        }
    } catch {
        return false
    }
    return false
}

; Helper function to check if file is unsaved
HasUnsavedChanges() {
    global g_isDirty
    return g_isDirty
}

; Helper function to confirm exit if there are unsaved changes
ConfirmExit() {
    if (HasUnsavedChanges()) {
        result := MsgBox("You have unsaved changes. Do you want to save before exiting?", "Unsaved Changes", "YesNo")
        
        if (result == "Yes") {
            SaveCurrentFile()
            return true ; Proceed with close
        } else if (result == "No") {
            return true ; Proceed with close
        } else {
            return true ; Cancel close
        }
    }
    
    return true ; No unsaved changes, proceed with close
}
