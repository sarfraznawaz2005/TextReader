; Search Management Functions

SearchAllFiles(searchText) {
    global g_WorkingFolder, rtfContent, g_SearchResults, g_IsSearchMode, g_MainGui

    if (searchText == "" || g_WorkingFolder == "") {
        g_IsSearchMode := false
        return
    }

    rtfContent.SetText("") ; Clear the content
    g_SearchResults := []
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
            ; File name and match count
            rtfContent.SetFont({Style: "B"})
            rtfContent.ReplaceSel("📄 " . _file.name . " (" . _file.matches.Length . " matches)")
            rtfContent.SetFont({Style: "N"})
            rtfContent.ReplaceSel("`r`n")

            for match in _file.matches {
                ; match is "Line X: some context"
                parts := StrSplit(match, [":"],, 2)
                lineNumberPart := parts[1]
                lineContent := Trim(parts[2])

                rtfContent.ReplaceSel("   • ")
                
                rtfContent.SetFont({Style: "B"})
                rtfContent.ReplaceSel(lineNumberPart)
                rtfContent.SetFont({Style: "N"})

                rtfContent.ReplaceSel(": ")

                ; Highlight search term in lineContent
                lastPos := 1
                while (pos := InStr(lineContent, searchText, false, lastPos)) {
                    ; Write text before the match
                    rtfContent.ReplaceSel(SubStr(lineContent, lastPos, pos - lastPos))
                    
                    ; Write the match with yellow background
                    rtfContent.SetFont({BkColor: 0x00FFFF}) ; BGR for Yellow
                    rtfContent.ReplaceSel(SubStr(lineContent, pos, StrLen(searchText)))
                    rtfContent.SetFont({BkColor: "Auto"}) ; Reset background color

                    lastPos := pos + StrLen(searchText)
                }
                ; Write the rest of the line
                rtfContent.ReplaceSel(SubStr(lineContent, lastPos))

                rtfContent.ReplaceSel("`r`n")
            }
            rtfContent.ReplaceSel("`r`n")
        }
    }
    
    g_IsSearchMode := true
    g_MainGui.Title := "Text Reader - Search Results"
}

FindMatches(content, searchText, tempRe) {
    matches := []
    processedLines := Map()
    lastPos := 1
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
            matches.Push("Line " . lineNum . ": " . context)
            processedLines[lineNum] := true
        }
        lastPos := pos + StrLen(searchText)
    }
    return matches
}