#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn

;try FileDelete(A_ScriptDir . "\TextReader.ini")

try DllCall("SetProcessDPIAware")

; --- Tray Menu ---
A_TrayMenu.Delete()
A_TrayMenu.Add("Reload", (*) => Reload())
A_TrayMenu.Add("Exit", (*) => ExitApp())

; ---------- Logging ----------
global LOG_ERRORS := A_ScriptDir . "\error.log"
global LOG_DEBUG  := A_ScriptDir . "\debug.log"

try FileDelete(LOG_ERRORS)
try FileDelete(LOG_DEBUG)

LogDebug(msg) {
  ;FileAppend(Format("[{1}] DEBUG: {2}`r`n", A_Now, msg), LOG_DEBUG)
}

LogError(msg) {
  FileAppend(Format("[{1}] ERROR: {2}`r`n", A_Now, msg), LOG_ERRORS)
}

OnError(LogUnhandled)

LogUnhandled(e, mode) {
  LogError(Format("Unhandled: {1} at {2}:{3}", e.Message, e.File, e.Line))
  return false
}

; Include other modules
#Include "Utils.ahk"
#Include "FloatingButton.ahk"
#Include "RichEdit.ahk"
#Include "RichEditDlgs.ahk"
#Include "AIChat.ahk"
#Include "GUI.ahk"
#Include "FileManager.ahk"
#Include "Settings.ahk"
#Include "SearchManager.ahk"
#Include "AIChat.ahk"

; Global variables
global g_MainGui := ""
global g_CurrentFile := ""
global g_WorkingFolder := ""
global g_SearchResults := []
global g_IsSearchMode := false
global g_FileContent := ""
global g_SearchTimer := ""
global g_isDirty := false

; Main entry point
Main() {
    LoadSettings()
    
    CreateMainGUI()
    
    CreateFloatingButton()
    
    if (g_WorkingFolder != "") {
        RefreshFileList()
    }
    
    g_MainGui.Show("w1300 h700 Hide")
    g_FloatingGui.Show()
}

Main()
