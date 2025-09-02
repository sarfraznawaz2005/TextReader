; File Management Functions

RefreshFileList() {
    global lvFiles, g_WorkingFolder, rtfContent, g_CurrentFile, btnSave
    LogDebug("RefreshFileList called. g_WorkingFolder: " . g_WorkingFolder)
    
    if (g_WorkingFolder == "" || !DirExist(g_WorkingFolder)) {
        LogDebug("Working folder is empty or does not exist. Returning.")
        ; Clear content and show welcome message if working folder is invalid
        rtfContent.SetText("")
        g_CurrentFile := ""
        btnSave.Visible := false
        ShowWelcomeMessage(true) ; Pass true for no working folder message
        rtfContent.SetEnabled(false)
        return
    }

    ; Clear existing items
    lvFiles.Delete()

    ; Add rtf files
    fileCount := 0
    Loop Files, g_WorkingFolder . "\*.rtf" {
        LogDebug("Adding file to ListView: " . A_LoopFileName)
        lvFiles.Add(, "📄 " . A_LoopFileName)
        fileCount++
    }
    LogDebug("Found " . fileCount . " .rtf files in " . g_WorkingFolder)

    ; Update header with file count
    lvFiles.ModifyCol(1, "Text", "Files (" . fileCount . ")")
    lvFiles.ModifyCol(1, 228) ; Auto-size column

    ; If no files found (but working folder is valid), clear content
    if (fileCount == 0) {
        rtfContent.SetText("")
        g_CurrentFile := ""
        btnSave.Visible := false
        rtfContent.SetEnabled(false)
    }
}

OpenFileInViewer(fileName) {
    global g_IsSearchMode, g_CurrentFile, g_WorkingFolder, rtfContent, btnSave, g_isDirty
    LogDebug("Opening file in viewer: " . fileName)
    
    cleanedFileName := RegExReplace(fileName, "^📄\s", "")
    fullPath := g_WorkingFolder . "\" . cleanedFileName

    if (!FileExist(fullPath)) {
        LogError("File not found: " . fullPath)
        MsgBox("Error: File not found.`r`n`r`n" . fullPath)
        return
    }

    try {
        rtfContent.SetEventMask(["NONE"]) ; Temporarily disable CHANGE event
        rtfContent.LoadFile(fullPath, "Open")
        rtfContent.ClearAllFormatting()
        rtfContent.SetFont({Size: 13, Name: "Calibri"})
        rtfContent.SetEventMask(["SELCHANGE", "LINK", "CHANGE"])
        SetToolbarVisibility(true) ; Show toolbar when a file is opened
        
        rtfContent.SetEnabled(true)
        g_CurrentFile := cleanedFileName
        btnSave.Visible := false
        g_IsSearchMode := false
        g_isDirty := false ; Reset dirty flag
        LogDebug("File opened successfully: " . fileName)
    } catch as e {
        LogError("Failed to read file " . fullPath . ": " . e.Message . ' in ' . e.Line)
        return
    }
}

OpenFileInExternalEditor(fileName) {
    global g_WorkingFolder
    
    ; Remove the file icon prefix if present
    cleanedFileName := RegExReplace(fileName, "^📄\s", "")

    filePath := g_WorkingFolder . "\" . cleanedFileName

    
    if (FileExist(filePath)) {
        Run('wordpad.exe "' . filePath . '"')
    }
}

AddNewFile() {
    global g_WorkingFolder
    
    if (g_WorkingFolder == "") {
        MsgBox("Please set a working folder in Settings first.", "No Working Folder", "48")
        return
    }
    
    fileName := ShowInputDialog("Add New File", "Enter file name (without .rtf extension):")
    
    if (fileName == "") {
        return
    }
    
    ; Remove .rtf if user added it
    fileName := RegExReplace(fileName, "\.rtf$", "", &count)
    fileName .= ".rtf"
    
    filePath := g_WorkingFolder . "\" . fileName
    
    if (FileExist(filePath)) {
        MsgBox("File already exists!", "Error", "16")
        return
    }
    
    try {
        ; Create empty file
        FileAppend("", filePath)
        rtfContent.ClearAllFormatting()
        rtfContent.SetFont({Size: 13, Name: "Calibri"})
        
        ; Refresh list and select new file
        RefreshFileList()
        
        ; Find and select the new file
        Loop lvFiles.GetCount() {
            ; Compare the cleaned ListView text with the new file name
            if (RegExReplace(lvFiles.GetText(A_Index, 1), "^📄\s", "") == fileName) {
                lvFiles.Modify(A_Index, "+Select +Focus")
                OpenFileInViewer(fileName)
                break
            }
        }
        
    } catch Error as e {
        MsgBox("Error creating file: " . e.message, "Error", "16")
    }
}

SaveCurrentFile() {
    global g_CurrentFile, g_WorkingFolder, rtfContent, btnSave, g_isDirty
    
    if (g_CurrentFile == "") {
        return
    }
    
    filePath := g_WorkingFolder . "\" . g_CurrentFile
    
    try {
        rtfContent.SaveFile(filePath)
        btnSave.Visible := false
        g_isDirty := false ; Reset dirty flag
    } catch Error as e {
        MsgBox("Error saving file: " . e.message, "Error", "16")
    }
}

ShowFileContextMenu(fileName, x, y) {
    fileContextMenu := Menu()
    fileContextMenu.Add("Rename", (*) => RenameFile(fileName))
    fileContextMenu.Add("Delete", (*) => DeleteFile(fileName))
    fileContextMenu.Show(x, y)
}

RenameFile(oldName) {
    global g_WorkingFolder, g_CurrentFile
    
    ; Remove the file icon prefix if present
    cleanedOldName := RegExReplace(oldName, "^📄\s", "")

    newName := ShowInputDialog("Rename File", "Enter new name:", RegExReplace(cleanedOldName, "\.rtf$"))
    
    if (newName == "" || newName == RegExReplace(cleanedOldName, "\.rtf$")) {
        return
    }
    
    ; Add .rtf extension if not present
    if (!RegExMatch(newName, "\.rtf$")) {
        newName .= ".rtf"
    }
    
    oldPath := g_WorkingFolder . "\" . cleanedOldName
    newPath := g_WorkingFolder . "\" . newName

    try {
        FileMove(oldPath, newPath)
        RefreshFileList()
        
        ; Update current file if it was renamed
        if (g_CurrentFile == oldName) {
            g_CurrentFile := newName
        }
        
    } catch Error as e {
        MsgBox("Error renaming file: " . e.message, "Error", "16")
    }
}

DeleteFile(fileName) {
    global g_WorkingFolder, g_CurrentFile
    
    ; Remove the file icon prefix if present
    cleanedFileName := RegExReplace(fileName, "^📄\s", "")

    result := MsgBox("Are you sure you want to delete '" . cleanedFileName . "'?", "Confirm Delete", "YesNo Icon!")
    
    if (result == "No") {
        return
    }
    
    filePath := g_WorkingFolder . "\" . cleanedFileName

    
    try {
        FileDelete(filePath)
        RefreshFileList()
        
        ; Clear content if deleted file was currently open
        if (g_CurrentFile == fileName) {
            rtfContent.SetText("")
            g_CurrentFile := ""
            btnSave.Visible := false
        }
        
    } catch Error as e {
        MsgBox("Error deleting file: " . e.message, "Error", "16")
    }
}