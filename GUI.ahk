
global MainColors, RE1 ; RE1 is a dummy control for UpdateGui in RichEditSample.ahk
global ContextMenu

; ======================================================================================================================
; Sets multi-line tooltips for any Gui control.
; Parameters:
;     GuiCtrl     -  A Gui.Control object
;     TipText     -  The text for the tooltip. If you pass an empty string for a formerly added control,
;                    its tooltip will be removed.
;     UseAhkStyle -  If set to true, the tooltips will be shown using the visual styles of AHK ToolTips.
;                    Otherwise, the current theme settings will be used.
;                    Default: True
;     CenterTip   -  If set to true, the tooltip will be shown centered below/above the control.
;                    Default: False
;  Return values:
;     True on success, otherwise False.
; Remarks: 
;     Text and Picture controls require the SS_NOTIFY (+0x0100) style.
; MSDN:
;     https://learn.microsoft.com/en-us/windows/win32/controls/tooltip-control-reference
; ======================================================================================================================
GuiCtrlSetTip(GuiCtrl, TipText, UseAhkStyle := True, CenterTip := False) {
   Static SizeOfTI := 24 + (A_PtrSize * 6)
   Static Tooltips := Map()
   Local Flags, HGUI, HCTL, HTT, TI
   ; Check the passed GuiCtrl
   If !(GuiCtrl Is Gui.Control)
      Return False
   HGUI := GuiCtrl.Gui.Hwnd
   ; Create the TOOLINFO structure -> msdn.microsoft.com/en-us/library/bb760256(v=vs.85).aspx
   Flags := 0x11 | (CenterTip ? 0x02 : 0x00) ; TTF_SUBCLASS | TTF_IDISHWND [| TTF_CENTERTIP]
   TI := Buffer(SizeOfTI, 0)
   NumPut("UInt", SizeOfTI, "UInt", Flags, "UPtr", HGUI, "UPtr", HGUI, TI) ; cbSize, uFlags, hwnd, uID
   ; Create a tooltip control for this Gui, if needed
   If !ToolTips.Has(HGUI) {
      If !(HTT := DllCall("CreateWindowEx", "UInt", 0, "Str", "tooltips_class32", "Ptr", 0, "UInt", 0x80000003
                                          , "Int", 0x80000000, "Int", 0x80000000, "Int", 0x80000000, "Int", 0x80000000
                                          , "Ptr", HGUI, "Ptr", 0, "Ptr", 0, "Ptr", 0, "UPtr"))
         Return False
      If (UseAhkStyle)
         DllCall("Uxtheme.dll\SetWindowTheme", "Ptr", HTT, "Ptr", 0, "Ptr", 0)
      SendMessage(0x0432, 0, TI.Ptr, HTT) ; TTM_ADDTOOLW
      Tooltips[HGUI] := {HTT: HTT, Ctrls: Map()}
   }
   HTT := Tooltips[HGUI].HTT
   HCTL := GuiCtrl.HWND
   ; Add / remove a tool for this control
   NumPut("UPtr", HCTL, TI, 8 + A_PtrSize) ; uID
   NumPut("UPtr", HCTL, TI, 24 + (A_PtrSize * 4)) ; uID
   If !Tooltips[HGUI].Ctrls.Has(HCTL) { ; add the control
      If (TipText = "")
         Return False
      SendMessage(0x0432, 0, TI.Ptr, HTT) ; TTM_ADDTOOLW
      SendMessage(0x0418, 0, -1, HTT) ; TTM_SETMAXTIPWIDTH
      Tooltips[HGUI].Ctrls[HCTL] := True
   }
   Else If (TipText = "") { ; remove the control
      SendMessage(0x0433, 0, TI.Ptr, HTT) ; TTM_DELTOOLW
      Tooltips[HGUI].Ctrls.Delete(HCTL)
      Return True
   }
   ; Set / Update the tool's text.
   NumPut("UPtr", StrPtr(TipText), TI, 24 + (A_PtrSize * 3))  ; lpszText
	   SendMessage(0x0439, 0, TI.Ptr, HTT) ; TTM_UPDATETIPTEXTW
	Return True
}

WinRedraw(hWnd) {
    DllCall("InvalidateRect", "Ptr", hWnd, "Ptr", 0, "Int", 1) ; Invalidate the entire window, erase background
    DllCall("UpdateWindow", "Ptr", hWnd) ; Force immediate repaint
}

; Define ContextMenu (at the top level of the script)
ContextMenu := Menu()
ContextMenu.Add("Cut", CutFN)
ContextMenu.Add("Copy", CopyFN)
ContextMenu.Add("Paste", PasteFN)
ContextMenu.Add("Clear", ClearFN)
ContextMenu.Add() ; Separator
ContextMenu.Add("Undo", UndoFN)
ContextMenu.Add("Redo", RedoFN)
ContextMenu.Add() ; Separator
ContextMenu.Add("Select All", SelAllFN)
ContextMenu.Add("Deselect", DeselectFN)

#HotIf WinActive(g_MainGui.Hwnd) && IsObject(rtfContent) && ControlGetFocus() == rtfContent.Hwnd
^s:: {
    SaveCurrentFile()
}
#HotIf

#HotIf IsObject(rtfContent) && rtfContent.Focused
^+z:: {
	rtfContent.Redo()
}

^f:: {
    FindText_Click()
}

^h:: {
    ReplaceText_Click()
}
#HotIf

#HotIf IsObject(txtSearch) && txtSearch.Focused
Enter:: {
    SearchPerform()
}
#HotIf

#HotIf WinActive(g_MainGui.Hwnd)
Escape:: {
    global g_MainGui, g_FloatingGui
    g_MainGui.Hide()
    if (IsObject(g_FloatingGui)) {
        g_FloatingGui.Show("NoActivate")
    }
}
#HotIf
	
; GUI Creation and Management

; --- Event Handlers ---

ToggleBold(*) {
    global rtfContent
    rtfContent.ToggleFontStyle("B")
    rtfContent.Focus()
}

ToggleItalic(*) {
    global rtfContent
    rtfContent.ToggleFontStyle("I")
    rtfContent.Focus()
}

ToggleUnderline(*) {
    global rtfContent
    rtfContent.ToggleFontStyle("U")
    rtfContent.Focus()
}

ToggleStrikeout(*) {
    global rtfContent
    rtfContent.ToggleFontStyle("S")
    rtfContent.Focus()
}

ChooseTextColor(*) {
    global rtfContent
    currColor := rtfContent.GetFont().Color
    if (currColor == "Auto")
        currColor := 0 ; Default to black
    NC := RichEditDlgs.ChooseColor(rtfContent, currColor)
    if (NC != "") {
        rtfContent.SetFont({Color: NC})
    }
    rtfContent.Focus()
}

ChooseTextBkColor(*) {
    global rtfContent
    currColor := rtfContent.GetFont().BkColor
    if (currColor == "Auto")
        currColor := 0xFFFFFF ; Default to white
    NC := RichEditDlgs.ChooseColor(rtfContent, currColor)
    if (NC != "") {
        rtfContent.SetFont({BkColor: NC})
    }
    rtfContent.Focus()
}

SetAlignment(align, *) {
    global rtfContent
    alignMap := Map("L", 1, "C", 3, "R", 2, "J", 4)
    rtfContent.AlignText(alignMap[align])
    rtfContent.Focus()
}

ToggleBulletList(*) {
    global rtfContent
    pf := rtfContent.GetParaFormat()
    ; PFN_BULLET = 1
    if (pf.Numbering == 1) { ; If it's already a bullet list, remove numbering
        rtfContent.SetParaNumbering()
    } else {
        rtfContent.SetParaNumbering({Type: "Bullet"})
    }
    rtfContent.Focus()
}

ToggleNumberList(*) {
    global rtfContent
    pf := rtfContent.GetParaFormat()
    ; PFN_ARABIC = 2
    if (pf.Numbering == 2) { ; If it's already an arabic number list, remove numbering
        rtfContent.SetParaNumbering()
    } else {
        rtfContent.SetParaNumbering({Type: "Arabic", Style: "Period"})
    }
    rtfContent.Focus()
}

ChangeIndent(dir, *) {
    global rtfContent
    pf := rtfContent.GetParaFormat()
    step := 360 ; 0.25 inch in twips
    newIndent := pf.StartIndent
    if (dir == "+") {
        newIndent += step
    } else {
        newIndent := Max(0, newIndent - step)
    }
    
    rtfContent.SetParaIndent({Start: newIndent / 1440})
    rtfContent.Focus()
}

IncreaseFontSize(*) {
    global rtfContent
    rtfContent.ChangeFontSize(1)
    rtfContent.Focus()
}

DecreaseFontSize(*) {
    global rtfContent
    rtfContent.ChangeFontSize(-1)
    rtfContent.Focus()
}

ToggleWordWrap(*) {
    global rtfContent, g_WordWrap
    g_WordWrap := !g_WordWrap
    rtfContent.WordWrap(g_WordWrap)
    rtfContent.Focus()
}

ClearFormatting(*) {
    global rtfContent
    rtfContent.SelAll() ; Select all text
    rtfContent.ClearAllFormatting() ; Call the new method
    rtfContent.SetFont({Size: 13, Name: "Calibri"}) ; Apply default font after clearing
    rtfContent.Deselect() ; Deselect the text after applying format
    rtfContent.Focus()
}

InsertHorizontalRule(*) {
    global rtfContent
    local hrText := StrRepeat("-", 120)
    rtfContent.ReplaceSel(hrText)
    rtfContent.Focus()
}

FindText_Click(*) {
    global rtfContent
    RichEditDlgs.FindText(rtfContent)
    rtfContent.Focus()
}

ReplaceText_Click(*) {
    global rtfContent
    RichEditDlgs.ReplaceText(rtfContent)
    rtfContent.Focus()
}


ChooseFont(*) {
    global rtfContent
    RichEditDlgs.ChooseFont(rtfContent)
    rtfContent.Focus()
}

AddNewFile_Click(*) {
    AddNewFile()
}

Settings_Click(*) {
    ShowSettingsWindow()
}

Search_Change(*) {
    global g_SearchTimer
    ; Reset the timer on every change
    SetTimer(SearchPerform, -500) ; Wait 500ms after last key press
}

SearchPerform() {
    global txtSearch, PLACEHOLDER_TEXT, g_IsSearchMode, rtfContent, lvFiles, btnSave, g_CurrentFile, g_WorkingFolder
    ; Ignore if text is placeholder
    if (txtSearch.Text == PLACEHOLDER_TEXT) {
        return
    }
    ; Perform the search logic here
    if (StrLen(txtSearch.Text) > 1) {
        g_IsSearchMode := true
        lvFiles.Modify(0, "-Select") ; Deselect all items
        rtfContent.SetEnabled(false)
        SetToolbarVisibility(false)
        btnSave.Visible := false
        SearchAllFiles(txtSearch.Text)
    } else if (StrLen(txtSearch.Text) == 0) {
        g_IsSearchMode := false
        rtfContent.SetEnabled(true)
        RefreshFileList()
        if (g_CurrentFile != "") {
            OpenFileInViewer(g_CurrentFile)
            SetToolbarVisibility(true)
        } else {
            ShowWelcomeMessage(g_WorkingFolder == "")
            SetToolbarVisibility(false)
        }
    }
}

FileList_Click(lv, rowNumber) {
    global rtfContent, g_CurrentFile, btnSave
    if (rowNumber > 0) {
        selectedFileName := lv.GetText(rowNumber, 1)
        LogDebug("FileList_Click: Selected file name from ListView: " . selectedFileName)
        OpenFileInViewer(selectedFileName)
    }
}

FileList_DoubleClick(lv, rowNumber) {
    if (rowNumber > 0) {
        selectedFileName := lv.GetText(rowNumber, 1)
        LogDebug("FileList_DoubleClick: Selected file name from ListView: " . selectedFileName)
        OpenFileInExternalEditor(selectedFileName)
    }
}

FileList_ContextMenu(lv, item, isRightClick, x, y) {
    if (item > 0) {
        ShowFileContextMenu(lv.GetText(item, 1), x, y)
    }
}

FileList_ItemFocus(lv, rowNumber) {
    ; lv: The ListView control object
    ; rowNumber: The 1-based index of the item that gained focus. 0 if no item has focus.
    global rtfContent, g_CurrentFile, btnSave

    if (rowNumber > 0) {
        selectedFileName := lv.GetText(rowNumber, 1)
        LogDebug("FileList_ItemFocus: Focused file name from ListView: " . selectedFileName)
        OpenFileInViewer(selectedFileName)
    } else {
        ; No item focused, clear content but do NOT show welcome message
        ShowWelcomeMessage(false)
        g_CurrentFile := ""
        btnSave.Visible := false
        rtfContent.SetEnabled(false)
        SetToolbarVisibility(false) ; Hide toolbar when no file is selected
    }
}

Save_Click(*) {
    SaveCurrentFile()
}

Content_Change(*) {
    global g_isDirty, g_IsSearchMode, btnSave, g_CurrentFile
    if (g_IsSearchMode) {
        return
    }
    if (g_CurrentFile != "") {
        btnSave.Visible := true
        g_isDirty := true
    }
}

RtfContent_Link(RE, L) {
	global rtfContent
	
   If (NumGet(L, A_PtrSize * 3, "Int") = 0x0202) { ; WM_LBUTTONUP
      wParam  := NumGet(L, (A_PtrSize * 3) + 4, "UPtr")
      lParam  := NumGet(L, (A_PtrSize * 4) + 4, "UPtr")
      cpMin   := NumGet(L, (A_PtrSize * 5) + 4, "Int")
      cpMax   := NumGet(L, (A_PtrSize * 5) + 8, "Int")
      URLtoOpen := rtfContent.GetTextRange(cpMin, cpMax)
      Run '"' URLtoOpen '"'
   }
}

txtSearch_OnFocus(*) {
    global txtSearch, PLACEHOLDER_TEXT
    if (txtSearch.Text == PLACEHOLDER_TEXT) {
        txtSearch.Text := ""
    }
}

txtSearch_OnLoseFocus(*) {
    global txtSearch, PLACEHOLDER_TEXT, PLACEHOLDER_COLOR
    if (txtSearch.Text == "") {
        txtSearch.Text := PLACEHOLDER_TEXT
    }
}

; --- RichEdit Action Functions for Context Menu ---
CutFN(*) {
    global rtfContent
    rtfContent.Cut()
    rtfContent.Focus()
}

CopyFN(*) {
    global rtfContent
    rtfContent.Copy()
    rtfContent.Focus()
}

PasteFN(*) {
    global rtfContent
    rtfContent.Paste()
    rtfContent.Focus()
}

UndoFN(*) {
    global rtfContent
    rtfContent.Undo()
    rtfContent.Focus()
}

RedoFN(*) {
    global rtfContent
    rtfContent.Redo()
    rtfContent.Focus()
}

SelAllFN(*) {
    global rtfContent
    rtfContent.SelAll()
    rtfContent.Focus()
}

DeselectFN(*) {
    global rtfContent
    rtfContent.Deselect()
    rtfContent.Focus()
}

ClearFN(*) {
    global rtfContent
    rtfContent.Clear()
    rtfContent.Focus()
}

CreateMainGUI() {
    global g_MainGui, g_WorkingFolder, g_FontSize, g_FontName, rtfContent, lvFiles, txtSearch, lblMatchCount, btnSave, g_WordWrap, btnAddFile, btnSettings

    g_WordWrap := true ; Default word wrap to off

    ; Create main window
    g_MainGui := Gui("+Resize +MaximizeBox -MinimizeBox", "Text Reader" . (g_WorkingFolder != "" ? " (" . g_WorkingFolder . ")" : ""))
    g_MainGui.SetFont("s10", "Segoe UI")
    g_MainGui.BackColor := "0xF5F5F5"

    ; Set title bar color
    ; Set title bar color
    titleBarColor := 0xD3D3D3 ; BGR color for #d3d3d3 (slightly darker than #dddddd)
    colorBuffer := Buffer(4) ; Create a buffer of 4 bytes (for a 32-bit integer)
    NumPut("UInt", titleBarColor, colorBuffer) ; Put the color value into the buffer
    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", g_MainGui.Hwnd, "Int", 35, "Ptr", colorBuffer.Ptr, "Int", 4)

    ; Set minimum size
    g_MainGui.OnEvent("Size", MainGui_Size)
    g_MainGui.OnEvent("Close", MainGui_Close)
    g_MainGui.OnEvent("ContextMenu", MainContextMenu)
    
    ; --- Left panel ---
    txtSearch := g_MainGui.AddEdit("x5 y5 w230 h30 +0x200")
    txtSearch.SetFont("s13", "Calibri")
    txtSearch.OnEvent("Change", Search_Change)
    GuiCtrlSetTip(txtSearch, "Search for text in files")

    ; Add placeholder functionality
    global PLACEHOLDER_TEXT := "Search files..."
    global PLACEHOLDER_COLOR := "0x808080" ; Gray

    txtSearch.Text := PLACEHOLDER_TEXT

    txtSearch.OnEvent("Focus", txtSearch_OnFocus)
    txtSearch.OnEvent("LoseFocus", txtSearch_OnLoseFocus)

    lvFiles := g_MainGui.AddListView("x5 y40 w230 h620 +Grid +Sort +Hdr", ["Files (0)"])
    lvFiles.SetFont("s13", "Calibri")
    lvFiles.OnEvent("Click", FileList_Click)
    lvFiles.OnEvent("DoubleClick", FileList_DoubleClick)
    lvFiles.OnEvent("ContextMenu", FileList_ContextMenu)
    lvFiles.OnEvent("ItemFocus", FileList_ItemFocus) ; New: Open file on item focus change

    btnAddFile := g_MainGui.AddButton("x8 y660 w110 h30 +0x1000", "📄 New File")
    btnAddFile.SetFont("s10")
    btnAddFile.OnEvent("Click", AddNewFile_Click)
    GuiCtrlSetTip(btnAddFile, "Add a new file")

    btnSettings := g_MainGui.AddButton("x123 y660 w110 h30 +0x1000", "⚙️ Settings")
    btnSettings.SetFont("s10")
    btnSettings.OnEvent("Click", Settings_Click)
    GuiCtrlSetTip(btnSettings, "Open settings")
    lvFiles.OnEvent("Click", FileList_Click)
    lvFiles.OnEvent("DoubleClick", FileList_DoubleClick)
    lvFiles.OnEvent("ContextMenu", FileList_ContextMenu)

    ; --- Right panel ---
    lblMatchCount := g_MainGui.AddText("x250 y15 w200 h20", "")
    GuiCtrlSetTip(lblMatchCount, "Number of search matches")

    btnSave := g_MainGui.AddButton("x900 y5 w80 h30 +0x1000", "💾 Save")
    btnSave.SetFont("s9")
    btnSave.OnEvent("Click", Save_Click)
    btnSave.Visible := false
    GuiCtrlSetTip(btnSave, "Save current file (Ctrl+S)")

    ; --- Formatting Toolbar ---
    ypos := 5
    xpos := 240
    btnH := 30
    btnW := 35
    iconFont := "Segoe MDL2 Assets"
    iconFontSize := 10
    separatorWidth := 15

    ; Declare toolbar buttons as global
    global btnBold, btnItalic, btnUnderline, btnStrikeout, btnTextColor, btnBgColor, btnDecreaseFont, btnIncreaseFont, btnFontDialog
    global btnAlignLeft, btnAlignCenter, btnAlignRight, btnAlignJustify, btnBulletList, btnNumberList
    global btnDecreaseIndent, btnIncreaseIndent, btnClearFormatting, btnHorizontalRule, btnWordWrap

    ; Font Style Group
    btnBold := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, "B")
    btnBold.SetFont("s12 Bold")
    btnBold.OnEvent("Click", ToggleBold)
    GuiCtrlSetTip(btnBold, "Bold")
    xpos += btnW

    btnItalic := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, "I")
    btnItalic.SetFont("s12 Italic")
    btnItalic.OnEvent("Click", ToggleItalic)
    GuiCtrlSetTip(btnItalic, "Italic")
    xpos += btnW

    btnUnderline := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, "U")
    btnUnderline.SetFont("s12 Underline")
    btnUnderline.OnEvent("Click", ToggleUnderline)
    GuiCtrlSetTip(btnUnderline, "Underline")
    xpos += btnW

    btnStrikeout := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, "S")
    btnStrikeout.SetFont("s12 Strike")
    btnStrikeout.OnEvent("Click", ToggleStrikeout)
    GuiCtrlSetTip(btnStrikeout, "Strikethrough")
    xpos += btnW

    ; Separator
    xpos += separatorWidth

    ; Color Group
    btnTextColor := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, "TC")
    btnTextColor.SetFont("s12")
    btnTextColor.OnEvent("Click", ChooseTextColor)
    GuiCtrlSetTip(btnTextColor, "Text Color")
    xpos += btnW

    btnBgColor := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, "BC")
    btnBgColor.SetFont("s12")
    btnBgColor.OnEvent("Click", ChooseTextBkColor)
    GuiCtrlSetTip(btnBgColor, "Text Background Color")
    xpos += btnW

    ; Separator
    xpos += separatorWidth

    ; Font Size Group
    btnDecreaseFont := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, "-")
    btnDecreaseFont.SetFont("s12")
    btnDecreaseFont.OnEvent("Click", DecreaseFontSize)
    GuiCtrlSetTip(btnDecreaseFont, "Decrease Font Size")
    xpos += btnW

    btnIncreaseFont := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, "+")
    btnIncreaseFont.SetFont("s12")
    btnIncreaseFont.OnEvent("Click", IncreaseFontSize)
    GuiCtrlSetTip(btnIncreaseFont, "Increase Font Size")
    xpos += btnW

    btnFontDialog := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, "F")
    btnFontDialog.SetFont("s12")
    btnFontDialog.OnEvent("Click", ChooseFont)
    GuiCtrlSetTip(btnFontDialog, "Open Font Dialog")
    xpos += btnW

    ; Separator
    xpos += separatorWidth

    ; Alignment Group
    btnAlignLeft := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, Chr(0xE8E4))
    btnAlignLeft.SetFont("s" . iconFontSize, iconFont)
    btnAlignLeft.OnEvent("Click", SetAlignment.Bind("L"))
    GuiCtrlSetTip(btnAlignLeft, "Align Left")
    xpos += btnW

    btnAlignCenter := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, Chr(0xE8E3))
    btnAlignCenter.SetFont("s" . iconFontSize, iconFont)
    btnAlignCenter.OnEvent("Click", SetAlignment.Bind("C"))
    GuiCtrlSetTip(btnAlignCenter, "Align Center")
    xpos += btnW

    btnAlignRight := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, Chr(0xE8E2))
    btnAlignRight.SetFont("s" . iconFontSize, iconFont)
    btnAlignRight.OnEvent("Click", SetAlignment.Bind("R"))
    GuiCtrlSetTip(btnAlignRight, "Align Right")
    xpos += btnW

    btnAlignJustify := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, Chr(0xE8E1))
    btnAlignJustify.SetFont("s" . iconFontSize, iconFont)
    btnAlignJustify.OnEvent("Click", SetAlignment.Bind("J"))
    GuiCtrlSetTip(btnAlignJustify, "Justify")
    xpos += btnW

    ; Separator
    xpos += separatorWidth

    ; List/Indent Group
    btnBulletList := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, Chr(0xE8FD))
    btnBulletList.SetFont("s" . iconFontSize, iconFont)
    btnBulletList.OnEvent("Click", ToggleBulletList)
    GuiCtrlSetTip(btnBulletList, "Bulleted List")
    xpos += btnW

    btnNumberList := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, "123")
    btnNumberList.SetFont("s12")
    btnNumberList.OnEvent("Click", ToggleNumberList)
    GuiCtrlSetTip(btnNumberList, "Numbered List")
    xpos += btnW

    btnDecreaseIndent := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, Chr(0xE72B))
    btnDecreaseIndent.SetFont("s" . iconFontSize, iconFont)
    btnDecreaseIndent.OnEvent("Click", ChangeIndent.Bind("-"))
    GuiCtrlSetTip(btnDecreaseIndent, "Decrease Indent")
    xpos += btnW

    btnIncreaseIndent := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, Chr(0xE72A))
    btnIncreaseIndent.SetFont("s" . iconFontSize, iconFont)
    btnIncreaseIndent.OnEvent("Click", ChangeIndent.Bind("+"))
    GuiCtrlSetTip(btnIncreaseIndent, "Increase Indent")
    xpos += btnW

    ; Separator
    xpos += separatorWidth

    ; Clear Formatting Button
    btnClearFormatting := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, Chr(0xE894))
    btnClearFormatting.SetFont("s" . iconFontSize, iconFont)
    btnClearFormatting.OnEvent("Click", ClearFormatting)
    GuiCtrlSetTip(btnClearFormatting, "Clear Formatting")
    xpos += btnW

    ; Separator
    xpos += separatorWidth

    ; Horizontal Rule Button
    btnHorizontalRule := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, "---")
    btnHorizontalRule.SetFont("s" . iconFontSize, iconFont)
    btnHorizontalRule.OnEvent("Click", InsertHorizontalRule)
    GuiCtrlSetTip(btnHorizontalRule, "Insert Horizontal Rule")
    xpos += btnW

    ; View Group
    btnWordWrap := g_MainGui.AddButton("x" . xpos . " y" . ypos . " w" . btnW . " h" . btnH, Chr(0xE73B))
    btnWordWrap.SetFont("s" . iconFontSize, iconFont)
    btnWordWrap.OnEvent("Click", ToggleWordWrap)
    GuiCtrlSetTip(btnWordWrap, "Toggle Word Wrap")
    xpos += btnW

    ; Separator
    xpos += separatorWidth

    ; Hide all toolbar buttons initially
    toolbarButtons := [btnBold, btnItalic, btnUnderline, btnStrikeout, btnTextColor, btnBgColor
                     , btnAlignLeft, btnAlignCenter, btnAlignRight, btnAlignJustify, btnBulletList, btnNumberList
                     , btnDecreaseIndent, btnIncreaseIndent, btnClearFormatting, btnHorizontalRule, btnWordWrap]
    For Each, btn In toolbarButtons {
        btn.Visible := false
    }
	
    ; --- Main content area ---
    rtfContent := RichEdit(g_MainGui, "x240 y40 w680 h580")
    rtfContent.SetFont({Size: 13, Name: "Calibri"})
	rtfContent.WordWrap(g_WordWrap)
    rtfContent.OnCommand(0x0300, Content_Change) ; EN_CHANGE notification
	rtfContent.OnNotify(0x070B, RtfContent_Link)
    rtfContent.AutoURL(True) ; Ensure Auto-URL detection is always enabled
    OnMessage(0x0205, RtfContent_RButtonUp, rtfContent.Hwnd) ; Add this line

	rtfContent.SetEventMask(["SELCHANGE", "LINK", "CHANGE"])
	
    CenterWindow(g_MainGui)

    ShowWelcomeMessage(g_WorkingFolder == "")
    if (g_WorkingFolder == "") {
        rtfContent.SetEnabled(false)
    }
	lvFiles.Focus()
    SetToolbarVisibility(false) ; Ensure toolbar is hidden initially
}

SetToolbarVisibility(state) {
    global btnBold, btnItalic, btnUnderline, btnStrikeout, btnTextColor, btnBgColor, btnDecreaseFont, btnIncreaseFont, btnFontDialog
    global btnAlignLeft, btnAlignCenter, btnAlignRight, btnAlignJustify, btnBulletList, btnNumberList
    global btnDecreaseIndent, btnIncreaseIndent, btnClearFormatting, btnHorizontalRule, btnWordWrap

    toolbarButtons := [btnBold, btnItalic, btnUnderline, btnStrikeout, btnTextColor, btnBgColor, btnDecreaseFont, btnIncreaseFont, btnFontDialog
                     , btnAlignLeft, btnAlignCenter, btnAlignRight, btnAlignJustify, btnBulletList, btnNumberList
                     , btnDecreaseIndent, btnIncreaseIndent, btnClearFormatting, btnHorizontalRule, btnWordWrap]
    For Each, btn In toolbarButtons {
        btn.Visible := state
    }
}

StrRepeat(stringToRepeat, numberOfTimes) {
    if (numberOfTimes <= 0) {
        return ""
    }
    local result := ""
    Loop numberOfTimes {
        result .= stringToRepeat
    }
    return result
}

CenterWindow(gui) {
    gui.GetPos(,, &width, &height)
    x := (A_ScreenWidth - width) // 2
    y := (A_ScreenHeight - height) // 2
    gui.Move(x, y)
}

MainGui_Size(gui, minMax, width, height) {
    if (minMax == -1) ; minimized
        return

    ; Resize controls
    try {
        ; Left panel controls
        ; txtSearch (already fixed at y0, h30) - no need to move here
        lvFiles.Move(, , 230, height - 80) ; New height calculation
        btnAddFile.Move(, height - 35) ; Anchor to bottom with 5px top margin
        btnSettings.Move(, height - 35) ; Anchor to bottom with 5px top margin

        ; Right panel adjusts
        rightWidth := width - 240
        rightHeight := height - 40 ; New height calculation for rtfContent

        rtfContent.Move(, 40, rightWidth, rightHeight) ; Move rtfContent to y40

        ; Adjust other controls
        lblMatchCount.Move(width - 415, 15) ; Keep its x, but y should be adjusted if it's overlapping.
        btnSave.Move(width - 110, 5) ; Anchor to top right
    }
    WinRedraw(g_MainGui.Hwnd)
}

MainGui_Close(*) {
    if (ConfirmExit()) {
        SaveSettings()
        g_FloatingGui.Show("NoActivate")
        g_MainGui.Minimize()
    }
}

ShowWelcomeMessage(isNoWorkingFolder := false) {
    global rtfContent
    if (isNoWorkingFolder) {
        rtfContent.SetText("Welcome to Text Reader!`r`n`r`n" .
                          "To get started:`r`n" .
                          "1. Click the 'Settings' button`r`n" .
                          "2. Select a working folder containing your RTF files`r`n" .
                          "3. Click 'Apply' to load your files`r`n`r`n" .
                          "Once configured, you'll see all .rtf files in the left panel.")
        rtfContent.SetSel(0, StrLen("Welcome to Text Reader!"))
        rtfContent.ToggleFontStyle("B")
        rtfContent.SetSel(-1, -1)
    } else {
        fullMessage := "Welcome back to Text Reader!`r`n`r`n" .
                       "Your files are loaded and ready. Select a file from the left panel to view its content."
        rtfContent.SetText(fullMessage)

        startOfWelcomeBack := 0
        endOfWelcomeBack := StrLen("Welcome back to Text Reader!")

        rtfContent.SetSel(startOfWelcomeBack, endOfWelcomeBack)
        rtfContent.ToggleFontStyle("B")
        rtfContent.SetSel(-1, -1) ; Move cursor to end
    }
    rtfContent.SetEnabled(false)
}

MainContextMenu(GuiObj, GuiCtrlObj, *) {
    global rtfContent, ContextMenu
    If (GuiCtrlObj = rtfContent)
        ContextMenu.Show()
}

RtfContent_RButtonUp(wParam, lParam, Msg, Hwnd) {
    global rtfContent, ContextMenu
    ; Check if the message is for our rtfContent control
    If (Hwnd = rtfContent.Hwnd) {
        ContextMenu.Show()
    }
}

