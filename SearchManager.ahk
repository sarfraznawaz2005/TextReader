; Search Management Functions

SearchAllFiles(searchText) {
    global g_WorkingFolder, rtfContent, g_SearchResults, g_SearchLinks, g_IsSearchMode, g_MainGui

    if (searchText == "" || g_WorkingFolder == "") {
        g_IsSearchMode := false
        return
    }

    rtfContent.SetText("") ; Clear the content
    rtfContent.AutoURL(false) ; Prevent RichEdit's own URL auto-detect pass from stripping our manual link formatting
    g_SearchResults := []
    g_SearchLinks := []
    totalMatches := 0

    ; Create a temporary hidden GUI and RichEdit control for text extraction
    tempGui := Gui()
    tempRe := RichEdit(tempGui, "w1 h1")

    filesWithMatches := []

    ; Search through all rtf files
    Loop Files, g_WorkingFolder . "\*.rtf" {
        try {
            tempRe.LoadFile(A_LoopFileFullPath, "Open")
            content := tempRe.GetText()
            matches := FindMatches(content, searchText, tempRe)

            if (matches.Length > 0) {
                filesWithMatches.Push({name: A_LoopFileName, matches: matches})
                totalMatches += matches.Length
            }
        }
    }

    tempGui.Destroy() ; Clean up the temporary GUI and control

    if (totalMatches == 0) {
        rtfContent.ReplaceSel("No matches found for '" . searchText . "'.")
    } else {
        ; 1. Add "TOTAL MATCHES: X" and make it bold
        rtfContent.SetFont({Style: "B"})
        rtfContent.ReplaceSel("TOTAL MATCHES: " . totalMatches)
        rtfContent.SetFont({Style: "N"})
        rtfContent.ReplaceSel("`r`n`r`n")

        ; 2. Add "SEARCH RESULTS FOR: 'searchText'" and make it bold
        rtfContent.SetFont({Style: "B"})
        rtfContent.ReplaceSel("SEARCH RESULTS FOR: '" . searchText . "'")
        rtfContent.SetFont({Style: "N"})
        rtfContent.ReplaceSel("`r`n`r`n")

        ; 3. Loop through files and add results
        for _file in filesWithMatches {
            baseName := RegExReplace(_file.name, "\.rtf$", "")
            firstMatchLine := _file.matches[1].lineNum

            ; File name/header - clickable link that opens the file and jumps to its first match
            WriteSearchLink(rtfContent, "📄 " . _file.name . " (" . _file.matches.Length . " matches)", baseName, firstMatchLine, true)
            rtfContent.ReplaceSel("`r`n")

            for match in _file.matches {
                ; Per-match clickable link that opens the file and jumps to that exact line
                rtfContent.ReplaceSel("   ")
                WriteSearchLink(rtfContent, "▸ Line " . match.lineNum, baseName, match.lineNum)
                rtfContent.ReplaceSel("`r`n")

                ; Surrounding context (up to 5 lines above/below, if they exist)
                rtfContent.SetFont({Style: "N", Color: 0x808080, Size: 11})
                for ctxLine in match.contextLines {
                    prefix := "        " . PadLeft(ctxLine.lineNum, 4) . " │ "

                    if (ctxLine.isMatch) {
                        rtfContent.SetFont({Style: "B", Color: 0x000000})
                        rtfContent.ReplaceSel(prefix)

                        ; Highlight the search term within the matched line
                        lineContent := match.text
                        lastPos := 1
                        while (pos := InStr(lineContent, searchText, false, lastPos)) {
                            rtfContent.ReplaceSel(SubStr(lineContent, lastPos, pos - lastPos))

                            rtfContent.SetFont({BkColor: 0x00FFFF}) ; BGR for Yellow
                            rtfContent.ReplaceSel(SubStr(lineContent, pos, StrLen(searchText)))
                            rtfContent.SetFont({BkColor: "Auto"})

                            lastPos := pos + StrLen(searchText)
                        }
                        rtfContent.ReplaceSel(SubStr(lineContent, lastPos))

                        rtfContent.SetFont({Style: "N", Color: 0x808080})
                    } else {
                        rtfContent.ReplaceSel(prefix . ctxLine.text)
                    }
                    rtfContent.ReplaceSel("`r`n")
                }
                rtfContent.SetFont({Style: "N", Color: "Auto", Size: 13})
                rtfContent.ReplaceSel("`r`n")
            }
            rtfContent.ReplaceSel("`r`n")
        }
    }

    g_IsSearchMode := true
}

FindMatches(content, searchText, tempRe) {
    matches := []
    processedLines := Map()
    lastPos := 1
    totalLines := tempRe.GetLineCount()
    while (pos := InStr(content, searchText, false, lastPos)) {
        lineNum := tempRe.GetLineFromChar(pos - 1)

        if (!processedLines.Has(lineNum)) {
            lineText := tempRe.GetLine(lineNum)
            context := Trim(lineText)
            if (StrLen(context) > 100) {
                matchPosInContext := InStr(context, searchText, false)
                start := Max(1, matchPosInContext - 40)
                context := "..." . SubStr(context, start, 80) . "..."
            }

            ; Gather surrounding context lines: up to 5 above and 5 below, if they exist
            contextLines := []
            startLine := Max(1, lineNum - 5)
            endLine := Min(totalLines, lineNum + 5)
            Loop endLine - startLine + 1 {
                curLine := startLine + A_Index - 1
                txt := Trim(tempRe.GetLine(curLine))
                if (StrLen(txt) > 100)
                    txt := SubStr(txt, 1, 100) . "..."
                contextLines.Push({lineNum: curLine, text: txt, isMatch: (curLine = lineNum)})
            }

            matches.Push({lineNum: lineNum, text: context, contextLines: contextLines})
            processedLines[lineNum] := true
        }
        lastPos := pos + StrLen(searchText)
    }
    return matches
}

; Writes clickable "link" text into a RichEdit control and records the character range
; so a later click on it can be resolved back to a file + line number.
WriteSearchLink(RE, text, fileArg, lineArg, bold := false) {
    global g_SearchLinks

    startPos := RE.GetSel().S

    ; Insert the text first (styled to look like a link)
    RE.SetFont({Style: (bold ? "BU" : "U"), Color: 0x1A73E8})
    RE.ReplaceSel(text)
    endPos := RE.GetSel().E

    ; Now select the text we just inserted and mark that actual range as a link.
    ; CFM_LINK must be applied to a real (non-empty) selection - setting it as
    ; "insertion point" formatting before typing does not reliably stick.
    RE.SetSel(startPos, endPos)
    CF2 := RichEdit.CHARFORMAT2()
    CF2.Mask := 0x20    ; CFM_LINK
    CF2.Effects := 0x20 ; CFE_LINK
    RE.SetCharFormat(CF2)

    g_SearchLinks.Push({start: startPos, end: endPos, file: fileArg, line: lineArg})

    ; Collapse the selection back to the end and clear link/style for subsequent text
    RE.SetSel(endPos, endPos)
    CF2Off := RichEdit.CHARFORMAT2()
    CF2Off.Mask := 0x20 ; CFM_LINK
    CF2Off.Effects := 0 ; not a link
    RE.SetCharFormat(CF2Off)
    RE.SetFont({Style: "N", Color: "Auto"})
}

; Opens the given file (base name, no extension) in the viewer and, if a line number
; is supplied, selects and scrolls to that line.
OpenSearchResultLink(fileName, lineNum := 0) {
    global rtfContent, lvFiles, txtSearch

    txtSearch.Text := ""
    RefreshFileList()
    OpenFileInViewer(fileName)
    txtSearch_OnLoseFocus()

    ; Highlight the matching row in the file list for visual consistency
    Loop lvFiles.GetCount() {
        if (lvFiles.GetText(A_Index, 1) = "📄 " . fileName) {
            lvFiles.Modify(A_Index, "Select Focus Vis")
            break
        }
    }

    if (lineNum > 0) {
        charIdx := rtfContent.GetLineIndex(lineNum - 1)
        if (charIdx != -1) {
            lineLen := StrLen(rtfContent.GetLine(lineNum))
            rtfContent.SetSel(charIdx, charIdx + lineLen)
            rtfContent.ScrollCaret()
        }
    }
    rtfContent.Focus()
}

PadLeft(value, len) {
    str := String(value)
    while (StrLen(str) < len)
        str := " " . str
    return str
}
