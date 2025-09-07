; AI Document Chat window module
#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn

global g_AIChatGui := ""
global chatRe := ""
global chatInput := ""
global AIChat_Messages := []
global AIChat_Selected := -1
global AIChat_IconCtrls := []
global gAssistantStart := 0 ; New global variable to track assistant message start

; delete chat history file as we dont need in this app
global CHAT_HISTORY_PATH := A_ScriptDir . "\\chat_history.json"
try FileDelete(CHAT_HISTORY_PATH)

#HotIf IsObject(g_AIChatGui) && WinActive(g_AIChatGui.Hwnd) && IsObject(chatInput) && chatInput.Focused
Enter:: {
    AIChat_Submit()
}
#HotIf

AIChat_Open() {
    global g_AIChatGui
    try {
        if IsObject(g_AIChatGui) {
            g_AIChatGui.Show("AutoSize Center")
            return
        }
        g_AIChatGui := Gui("-MaximizeBox +MinimizeBox", "AI Document Chat")
        g_AIChatGui.SetFont("s10", "Segoe UI")
        AIChat_InitUI(g_AIChatGui)
        g_AIChatGui.OnEvent("Close", (*) => g_AIChatGui := "")
        g_AIChatGui.Show("AutoSize Center")
    } catch as e {
        LogError("AIChat open error: " . e.Message)
    }
}

AIChat_InitUI(gui) {
    global chatRe, chatInput

    chatRe := RichEdit(gui, "x0 y0 w600 h600", True)
    chatRe.SetReadOnly(true)
    chatRe.WordWrap(true)
	chatRe.ShowScrollBar(0)
	
	gui.OnEvent("Size", (thisGui, minMax, width, height) => chatRe.Move(, , width))
	
	; Set default font here
    chatRe.SetEventMask(["SELCHANGE"]) ; track selection
    chatRe.OnNotify(0x0702, AIChat_SelChange)

    chatInput := gui.AddEdit("x10 y610 w585 h25")
    chatInput.SetFont("s12", "Calibri")
    chatInput.Value := "Ask me anything..."
    chatInput.OnEvent("Focus", (*) => (chatInput.Value == "Ask me anything..." ? chatInput.Value := "" : 0))
    chatInput.OnEvent("LoseFocus", (*) => (chatInput.Value == "" ? chatInput.Value := "Ask me anything..." : 0))
	chatInput.Focus
}

AIChat_AddAssistant(text) {
    global chatRe, AIChat_Messages
    chatRe.SetReadOnly(false)
    start := chatRe.GetTextLen()
    chatRe.SetSel(-1, -1)
    chatRe.AlignText("LEFT")
    chatRe.SetFont({Name: "Calibri", Size: 11, Color: 0x000000})

    chatRe.ReplaceSel(text . "`r`n`r`n")
    chatRe.SetReadOnly(true)
    AIChat_Messages.Push({role: "assistant", text: text, start: start, end: chatRe.GetTextLen()})
    chatRe.ScrollCaret()
    chatRe.Redraw()
}

AIChat_AddUser(text) {
    global chatRe, AIChat_Messages
    chatRe.SetReadOnly(false)
    start := chatRe.GetTextLen()
    chatRe.SetSel(-1, -1)
    chatRe.SetFont({Name: "Calibri", Size: 12, Color: 0x2D6CDF, Style: "B"})

    chatRe.ReplaceSel(text . "`r`n`r`n")
    chatRe.SetReadOnly(true)
    AIChat_Messages.Push({role: "user", text: text, start: start, end: chatRe.GetTextLen()})
    chatRe.ScrollCaret()
    chatRe.Redraw()
}

AIChat_SelChange(*) {
    global chatRe, AIChat_Messages, AIChat_Selected
    sel := chatRe.GetSel()
    idx := 1
    for i, m in AIChat_Messages {
        if (sel.S >= m.start && sel.S <= m.end) {
            idx := i
            break
        }
        idx := i
    }
    if (idx != AIChat_Selected) {
        AIChat_Selected := idx
        
    }
}

AIChat_SelectLast() {
    global AIChat_Messages, AIChat_Selected
    AIChat_Selected := AIChat_Messages.Length
}

AIChat_AddAssistantPrefix() {
    global chatRe, gAssistantStart
    chatRe.SetReadOnly(false)
    gAssistantStart := chatRe.GetTextLen()
    chatRe.SetSel(-1, -1)
    chatRe.AlignText("LEFT")
    chatRe.SetFont({Name: "Calibri", Size: 11, Color: 0x000000})
    chatRe.ReplaceSel("Thinking...")
    chatRe.ScrollCaret()
    chatRe.Redraw()
}

; Function to overwrite the assistant's message from gAssistantStart
AIChat_OverwriteAssistant(text) {
    global chatRe, gAssistantStart
    if (gAssistantStart > 0 && gAssistantStart <= chatRe.GetTextLen()) {
        chatRe.SetReadOnly(false)
        ; Select from gAssistantStart to the end of the current text
        chatRe.SetSel(gAssistantStart, chatRe.GetTextLen())
        chatRe.ReplaceSel(text)
        chatRe.ScrollCaret()
        chatRe.Redraw()
        chatRe.SetReadOnly(true)
    } else {
        AIChat_AddAssistant("Thinking..." . text)
    }
}

AIChat_Submit() {
    global chatInput, chatRe, AIChat_Messages, gAssistantStart

    txt := Trim(chatInput.Value)
    if (txt = "")
        return

    chatInput.Value := ""

    AIChat_AddUser(txt) ; Adds user message and updates AIChat_Messages

    ; --- Start of Assistant's Streaming Response ---
    AIChat_AddAssistantPrefix() 

    ; Create a message object for the assistant's response.
    ; Its 'text' property will be updated as chunks arrive.
    local currentAssistantMessage := {role: "assistant", text: "", start: gAssistantStart, end: gAssistantStart}
    AIChat_Messages.Push(currentAssistantMessage)

    cfg := LoadConfig() ; Load AI configuration from AIAssistant.ahk
    local buf := "" ; Buffer to accumulate streamed deltas

    ; Define onDelta as a nested function
    onDelta(delta) {
        ; chatRe and currentAssistantMessage are accessible from the enclosing scope
        buf .= delta ; Accumulate delta
        AIChat_OverwriteAssistant(buf) ; Overwrite with accumulated text
        currentAssistantMessage.text := buf ; Update internal message object
        currentAssistantMessage.end := chatRe.GetTextLen()
    }

    ; Define onDone as a nested function
    onDone(fullText, status) {
        LogDebug("AIChat.ahk: onDone received fullText: '" fullText "'")
        ; chatRe, AIChat_Messages, currentAssistantMessage are accessible from the enclosing scope
        AIChat_OverwriteAssistant(fullText) ; Final overwrite with full text
        currentAssistantMessage.text := fullText ; Final update to internal message object
        currentAssistantMessage.end := chatRe.GetTextLen()

        chatRe.SetReadOnly(false)
        chatRe.SetSel(-1, -1)
        chatRe.ReplaceSel("`r`n`r`n") ; Add newlines after the assistant's response
        chatRe.SetReadOnly(true)
        chatRe.ScrollCaret()
        chatRe.Redraw()

        AIChat_SelectLast()
        LogDebug("Chat stream finished with status: " status)
        gAssistantStart := 0 ; Reset gAssistantStart
    }

    ; Call the streaming RAG chat function
    ChatWithRAGStreamAsync(txt, cfg, 5, onDelta, onDone)
}