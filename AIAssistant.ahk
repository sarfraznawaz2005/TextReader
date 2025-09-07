#Requires AutoHotkey v2.0+
#SingleInstance Force
#Warn

; =============================================================
; AIAssistant.ahk — Provider‑agnostic AI + RAG library (AHK v2)
; =============================================================
; Overview
; - Single‑file, AutoHotkey v2 library for:
;   - Provider‑agnostic chat (OpenAI‑compatible, Gemini, Ollama).
;   - Streaming and non‑streaming chat with retry and SSE parsing.
;   - RAG: local vector store (JSON), text splitting, ingest flows.
;   - Embeddings via provider or local hashing fallback.
;   - Minimal dependencies (JSON.ahk + MSXML2.XMLHTTP COM).
;
; Quick Start (Code)
;   cfg := LoadConfig()                        ; reads ./config.ini [api]
;   res := RAG_Query("your question", 5, false) ; local-only retrieval
;   ChatWithRAGAsync("your question", cfg, 5, (txt, status) => MsgBox(txt))
;   ChatWithRAGStreamAsync("your question", cfg, 5,
;       (delta) => ToolTip(delta), (full, status) => ToolTip("done:" full))
;
; Quick Start (CLI)
;   "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe" cli.ahk ingest test.txt
;   "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe" cli.ahk chat "Hi"
;   "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe" cli.ahk chatstream "Hi"
;   "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe" cli.ahk ingestpe test.txt  ; provider embeddings
;   "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe" cli.ahk rebuild            ; revectorize with provider
;
; Configuration (./config.ini [api])
;   provider=openai | gemini | openai-compatible | ollama
;   baseUrl=...          ; optional (defaults per provider)
;   apiKey=...           ; provider key (not needed for local ollama)
;   chatModel=...        ; e.g., gpt-4o-mini | gemini-1.5-flash | llama3
;   embedModel=...       ; e.g., text-embedding-3-small | text-embedding-004
;   timeoutMs=60000      ; request timeout
;   streamStallMs=8000   ; stream stall detection
;   ; Gemini extra (optional): taskType, outputDimensionality, safety*
;
; Logging
;   - debug.log: appended with timestamps (deleted at module load).
;   - Never writes secrets. Use for tracing requests, ingest steps, saves.
;
; Vector Store (./vector_store.json)
;   { items: [ {id,path,start,end,text,left,right,vector:Array,hash,...} ],
;     docs:  { "path": { hash, size, mtime, chunkIds:[ids...] } } }
;
; RAG & Embeddings (Key APIs)
;   - LoadConfig() -> Map
;   - BuildContextPrompt(prompt, retrieved:Array) -> String
;   - RAG_Query(prompt, topK:=5, useProvider:=false, cfg?, cb?) -> Array|[]
;       useProvider=true -> async (cb(results)) after embeddings->query.
;   - IngestFile(path, opts?) -> Bool                      ; local vectors
;   - IngestDirectory(dir, opts?)                          ; local vectors
;   - IngestFileProviderEmbeddingsAsync(path, cfg?, opts?, doneCb?)
;   - IngestDirectoryProviderEmbeddingsAsync(dir, cfg?, opts?, doneCb?)
;   - RevectorizeStoreProviderEmbeddingsAsync(cfg?, opts?, doneCb?)
;   - VS_Load()/VS_Save(store)/VS_AddChunks(store,...)/VS_Query(...)
;   - VS_ListDocs(store)/VS_RemoveDoc(store,path)
;   - EmbeddingsAsync(cfg, inputText, cb)                  ; single text
;   - EmbeddingsBatchOpenAIAsync(cfg, texts, cb)           ; batch helper
;   - EmbeddingsBatchGeminiAsync(cfg, texts, cb)           ; batch helper
;   - ParseEmbeddingVector/ParseEmbeddingVectors           ; response helpers
;
; Chat (Key APIs)
;   - ChatWithRAGAsync(userPrompt, cfg?, topK:=5, cb?)
;   - ChatWithRAGStreamAsync(userPrompt, cfg?, topK:=5, onDelta?, onDone?)
;   - ChatCompletionAsync(cfg, messages, cb)               ; provider-agnostic
;   - ChatCompletionStreamAsync(cfg, messages, onDelta, onDone)
;   - ExtractChatText(provider, rawResponseText)           ; normalize text
;
; Strict RAG Policy
;   - Enforced globally via StrictRAGInstruction():
;       "Answer strictly and only from provided context ..."
;   - OpenAI/Ollama: inserted as {role:"system"} message.
;   - Gemini: inserted as systemInstruction in request body.
;   - To disable, change StrictRAGInstruction() to return "" (empty).
;
; Error Handling
;   - try/catch around file I/O, HTTP, JSON parse. Retries for HTTP.
;   - Streaming has stall detection and retries (limited).
;
; Testing
;   - See test/runner.ahk (local RAG, no network)
;   - runner_chatstream.ahk (streams to stdout)
;   - CLI usage (see cli.ahk) for ingest/rag/chat.
;
; Notes for Agents
;   - Keep AHK v2 expression syntax; avoid v1 commands.
;   - Prefer AddSystemToMessages() when crafting custom messages.
;   - For provider embeddings, ensure vectors are preserved (c.vector).
;   - Check debug.log after operations; error.log is populated by runners.
;

; =======================================
; SAMPLE CONFIG
; =======================================

/*
[AIChat]
enabled=1

; provider can be: openai, gemini, openai-compatible, ollama
provider=gemini

; baseUrl defaults per provider, override if self-hosting or using a proxy
; openai: https://api.openai.com
; gemini: https://generativelanguage.googleapis.com
; ollama: http://localhost:11434

baseUrl=https://generativelanguage.googleapis.com

; API key for provider (not needed for local ollama default)
apiKey=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

; default models
;chatModel=gpt-4o-mini
chatModel=gemini-2.0-flash

embedModel=gemini-embedding-001
;embedModel=text-embedding-004
;embedModel=text-embedding-3-small

; numeric dimension if supported by model (e.g., 768)
outputDimensionality=768
;outputDimensionality=1536

; Gemini embeddings options (see https://ai.google.dev/gemini-api/docs/embeddings)
; taskType can be: TASK_TYPE_UNSPECIFIED, RETRIEVAL_QUERY, RETRIEVAL_DOCUMENT, SEMANTIC_SIMILARITY, etc.
taskType=SEMANTIC_SIMILARITY

; Gemini Safety Settings
safetyHarassment=BLOCK_NONE
safetyHateSpeech=BLOCK_NONE
safetySexual=BLOCK_NONE
safetyDangerous=BLOCK_NONE
safetyCivic=BLOCK_NONE

; request timeout in milliseconds (optional)
timeoutMs=60000

; Streaming stall detection (ms): if no new bytes for this long, abort and retry
streamStallMs=8000

; Generation options
; For OpenAI: temperature, max_tokens
; For Gemini: temperature, maxOutputTokens (and optionally topP, topK, candidateCount, stopSequences)
temperature=
max_tokens=
maxOutputTokens=

; Show citations appended to replies (on/off/1/0/true/false)
showCitations=1
citationMaxFiles=3
citationMatchThreshold=0.12

; Text Splitting
chunkSize=1000
overlap=200
pad=50

*/

#Include "JSON.ahk"

ROOT_DIR := __DetectRootDir()
global CONFIG_INI_PATH := ROOT_DIR "\\TextReader.ini"

global CHAT_HISTORY_PATH := ROOT_DIR "\\chat_history.json"
global VECTOR_STORE_PATH := ROOT_DIR "\\vector_store.json"

/*
global LOG_DEBUG  := A_ScriptDir . "\debug.log"

try FileDelete(LOG_DEBUG)

LogDebug(msg) {
  FileAppend(Format("[{1}] DEBUG: {2}`r`n", A_Now, msg), LOG_DEBUG)
}
*/


__DetectRootDir() {
    ; Heuristic: prefer directory containing requirements.txt or .git, else parent of script dir
    cand := NormalizePath(A_ScriptDir)
    p1 := cand "\\requirements.txt"
    if FileExist(p1)
        return cand
    p2 := cand "\\.git"
    if DirExist(p2)
        return cand
    parent := NormalizePath(cand "\\..")
    p3 := parent "\\requirements.txt"
    if FileExist(p3)
        return parent
    return parent ; fallback
}

NormalizePath(path) {
    ; Use WinAPI GetFullPathNameW to normalize paths
    buf := Buffer(32768 * 2, 0) ; wide chars
    len := DllCall("Kernel32.dll\GetFullPathNameW", "WStr", path, "UInt", 32768, "Ptr", buf, "Ptr", 0, "UInt")
    if (len = 0)
        return path
    return StrGet(buf, "UTF-16")
}

; ---------------- Prompt Policy ----------------
StrictRAGInstruction() {
    return "
    (Join`r`n
    You are a RAG assistant. Answer strictly and only from the provided context.
    If the context does not contain the answer, reply: 'Sorry, I don't know based on the provided context.'
    Do not use outside knowledge or make assumptions other than replying to salutations or acknowledgments.

    OUTPUT FORMAT:
    Your answer must always be simple text, no HTML, no markdown, no rich text. Text must be nicely formatted with 
	appropriate new lines and spacing, indentation, special marks such as *, -, etc. so it becomes easy 
	to read for humans.
    )"
}

AddSystemToMessages(messages) {
    sys := StrictRAGInstruction()
    if (sys = "")
        return messages
    arr := []
    arr.Push({ role: "system", content: sys })
    for m in messages
        arr.Push(m)
    return arr
}

; ---------------- Configuration ----------------
LoadConfig() {
    global CONFIG_INI_PATH
    cfg := Map()
    try {
        if FileExist(CONFIG_INI_PATH) {
            ; Read simple INI (flat) keys under [AIChat]
            sec := "AIChat"
            keys := [
                "enabled","provider","baseUrl","apiKey",
                "chatModel","embedModel",
                "timeoutMs","streamStallMs",
                ; generation options
                "temperature","max_tokens","maxOutputTokens","topP","topK","candidateCount","stopSequences",
                ; Gemini embeddings options
                "taskType","outputDimensionality",
                ; Gemini safety
                "safetyHarassment","safetyHateSpeech","safetySexual","safetyDangerous",
                ; Citations toggle
                "showCitations","chunkSize","overlap", "pad"
            ]
            for k in keys {
                val := IniRead(CONFIG_INI_PATH, sec, k, "")
                cfg[k] := val
				
				;if (val)
					;MsgBox k . " = " . val
            }
        }
    } catch as e {
        LogDebug("LoadConfig error")
    }
    return cfg
}

; ---------------- Utility ----------------
ToUTF8(str) {
    ; Ensure UTF-8 strings for HTTP payloads
    return StrReplace(str, "\u", "\\u") ; keep minimal; AHK v2 strings are Unicode
}


HtmlEntityDecode(s) {
    try {
        s := StrReplace(s, "&amp;", "&")
        s := StrReplace(s, "&lt;", "<")
        s := StrReplace(s, "&gt;", ">")
        s := StrReplace(s, "&quot;", '"')
        s := StrReplace(s, "&apos;", "'")
        s := StrReplace(s, "&nbsp;", " ")
        s := StrReplace(s, "&ndash;", "–")
        s := StrReplace(s, "&mdash;", "—")
        s := StrReplace(s, "&lsquo;", "‘")
        s := StrReplace(s, "&rsquo;", "’")
        s := StrReplace(s, "&ldquo;", "“")
        s := StrReplace(s, "&rdquo;", "”")
        s := StrReplace(s, "&hellip;", "…")
        s := StrReplace(s, "&copy;", "©")
        s := StrReplace(s, "&reg;", "®")
        s := StrReplace(s, "&trade;", "™")
        s := RegExReplace(s, "&#(\d+);", (m) => Chr(Integer(m[1])))
        s := RegExReplace(s, "&#x([0-9A-Fa-f]+);", (m) => Chr("0x" m[1]))
        return s
    } catch as e {
        return s
    }
}

__RtfDecodeUnicode(text) {
    out := ""
    i := 1
    while (i <= StrLen(text)) {
        if (SubStr(text, i, 2) = "\u") {
            if RegExMatch(SubStr(text, i), "^\\u(-?\d+)(\??)", &mm) {
                code := Integer(mm[1])
                if (code < 0)
                    code := code + 65536
                out .= Chr(code)
                i += StrLen(mm[0])
                if (mm[2] != "")
                    i += 1
                continue
            }
        }
        out .= SubStr(text, i, 1)
        i += 1
    }
    return out
}

; Robust HTML to text (COM if available, fallback to regex + entity decode)
StripHtml(html) {
    try {
        doc := ComObject("HTMLFile")
        doc.open()
        doc.write(html)
        doc.close()
        txt := ""
        try txt := doc.body.innerText
        if (txt = "")
            try txt := doc.documentElement.innerText
        if (txt != "")
            return Trim(HtmlEntityDecode(txt))
    } catch as e {
        ; fall back
    }
    try {
        t := RegExReplace(html, "is)<script\b.*?</script>", "")
        t := RegExReplace(t, "is)<style\b.*?</style>", "")
        t := RegExReplace(t, "is)<[^>]+>", " ")
        t := RegExReplace(t, "\s+", " ")
        return Trim(HtmlEntityDecode(t))
    } catch as e {
        LogDebug("StripHtml error: " e.Message)
        return html
    }
}

; --- Public API ---
StripRtf(rtfText) {
    try {
        ; Use improved regex-based conversion that preserves hyperlinks
        return RegexRtfToText(rtfText)
    } catch {
        return rtfText
    }
}

; --- Internal: Improved RTF → Text ---
RegexRtfToText(rtfText) {
    text := rtfText

    ; Handle escaped RTF if needed (\\rtf → \rtf)
    if (InStr(text, "\\\\rtf"))
        text := StrReplace(text, "\\\\", "\")

    ; Step 1: Extract hyperlinks BEFORE removing RTF structures
    hyperlinks := []
    pos := 1
    while (pos := RegExMatch(
        text
      , '{\s*\\field\s*{\s*\\\*\s*\\fldinst\s*{\s*HYPERLINK\s+([^\s}]+)\s*}\s*}\s*{\s*\\fldrslt\s*{([^}]*)}\s*}\s*}'
      , &match, pos)) 
    {
        url := Trim(match[1])
        displayText := Trim(match[2])

        ; Clean display text of RTF control words + braces
        displayText := RegExReplace(displayText, "\\[a-zA-Z]+\d*", "")
        displayText := RegExReplace(displayText, "[{}]", "")
        displayText := Trim(displayText)

        finalText := (displayText != "" && displayText != url)
            ? displayText . " (" . url . ")"
            : url

        hyperlinks.Push({pattern: match[0], replacement: finalText})
        pos := match.Pos + match.Len
    }

    ; Step 2: Replace hyperlink structures with preserved text
    for _, link in hyperlinks
        text := StrReplace(text, link.pattern, " " . link.replacement . " ")

    ; Step 3: Remove font tables / headers / metadata
    text := RegExReplace(text, '{\\fonttbl[^{}]*(?:{[^{}]*}[^{}]*)*}', '')
    text := RegExReplace(text, '{\\colortbl[^}]*}', '')
    text := RegExReplace(text, '{\\stylesheet[^}]*}', '')
    text := RegExReplace(text, '{\\info[^}]*}', '')
    text := RegExReplace(text, '{\\\*\s*\\generator[^}]*}', '')
    ; Remove single font lines if any slipped through
    text := RegExReplace(text, '{\\f\d+\\fnil\\fcharset0\s*[^;}]*;}', '')
    ; Remove explicit font names that sometimes linger
    text := RegExReplace(text, 'Cascadia Code ExtraLight[^;]*;', '')
    text := RegExReplace(text, 'Symbol[^;]*;', '')

    ; Step 4: Convert hex-encoded characters (\'B7 → •, etc.)
    pos := 1
    while (pos := RegExMatch(text, "\\'([0-9a-fA-F]{2})", &match, pos)) {
        charCode := "0x" . match[1]
        char := Chr(Integer(charCode))
        text := SubStr(text, 1, match.Pos - 1) . char . SubStr(text, match.Pos + match.Len)
        pos := match.Pos + StrLen(char)
    }

    ; Step 5: Convert RTF line breaks → actual newlines
    ; IMPORTANT: handle \pard BEFORE \par to avoid creating a stray "d"
    text := RegExReplace(text, "\\pard(?![a-zA-Z])", "`n")
    text := RegExReplace(text, "\\par(?![a-zA-Z])",  "`n")
    text := RegExReplace(text, "\\line(?![a-zA-Z])", "`n")
    text := RegExReplace(text, "\\page(?![a-zA-Z])", "`n`n")

    ; Step 6: Remove remaining RTF control words (standalone)
    text := RegExReplace(text, "\\[a-zA-Z]+\d*", " ")
    text := RegExReplace(text, "\\\*", " ")

    ; Step 7: Remove RTF braces and clean escapes
    text := RegExReplace(text, "[{}]", "")
    text := StrReplace(text, "\\~", " ")  ; non-breaking space
    text := StrReplace(text, "\\_", "-")  ; non-breaking hyphen
    ; Escaped literal braces/backslashes
    text := StrReplace(text, "\\{", "{")
    text := StrReplace(text, "\\}", "}")
    text := StrReplace(text, "\\\\", "\")

    ; Step 8: Tidy whitespace / newlines
    text := RegExReplace(text, "[ \t]+", " ")
    text := RegExReplace(text, " *`n *", "`n")
    text := RegExReplace(text, "`n{3,}", "`n`n")

    ; Step 9: Fix common numbered-list quirks
    text := RegExReplace(text, "(\d+)\.\s*\.", "$1.")       ; "1.." → "1."
    text := RegExReplace(text, "(\d+)\.([a-zA-Z])", "$1. $2") ; "1.A" → "1. A"

    ; Step 10: Trim leading non-text garbage (kept from your original)
    text := RegExReplace(text, "^[^A-Za-z]*", "")

    return Trim(text)
}

StripPdf(path) {
    ; Try to use pdftotext.exe if available (tools\pdftotext.exe) or in PATH
    try {
        exe := A_ScriptDir "\\tools\\pdftotext.exe"
        if !FileExist(exe) {
            exe := "pdftotext.exe"
            ; check PATH
            found := DllCall("Kernel32\\SearchPathW", "wstr", "", "wstr", exe, "ptr", 0, "uint", 0, "ptr", 0, "int")
            if (found = 0) {
                LogDebug("StripPdf: pdftotext.exe not found in tools or PATH")
                return ""
            }
        }
        outTxt := A_Temp "\\ahk_pdf_" Djb2Hex(path) ".txt"
        ; -layout preserves reading order better for some docs
        cmd := '"' exe '" -layout -enc UTF-8 ' '"' path '" ' '"' outTxt '"'
        workdir := ""
        if FileExist(A_ScriptDir "\\tools\\pdftotext.exe")
            workdir := A_ScriptDir "\\tools"
        oldPath := EnvGet("PATH")
        if (workdir != "")
            EnvSet("PATH", workdir ";" oldPath)
        try RunWait(cmd, workdir, "Hide")
        finally {
            if (workdir != "")
                EnvSet("PATH", oldPath)
        }
        if FileExist(outTxt) {
            txt := FileRead(outTxt, "UTF-8")
            try FileDelete(outTxt)
            return txt
        }
    } catch as e {
        LogDebug("StripPdf error: " e.Message)
        ; Fallback to pdftohtml -> html -> text
        try {
            exe2 := A_ScriptDir "\\tools\\pdftohtml.exe"
            if !FileExist(exe2)
                exe2 := "pdftohtml.exe"
            htmlOut := A_Temp "\\ahk_pdf_" Djb2Hex(path) ".html"
            cmd2 := '"' exe2 '" -enc UTF-8 -noframes ' '"' path '" ' '"' htmlOut '"'
            workdir2 := ""
            if FileExist(A_ScriptDir "\\tools\\pdftohtml.exe")
                workdir2 := A_ScriptDir "\\tools"
            oldPath2 := EnvGet("PATH")
            if (workdir2 != "")
                EnvSet("PATH", workdir2 ";" oldPath2)
            try RunWait(cmd2, workdir2, "Hide")
            finally {
                if (workdir2 != "")
                    EnvSet("PATH", oldPath2)
            }
            if FileExist(htmlOut) {
                rawHtml := FileRead(htmlOut, "UTF-8")
                try FileDelete(htmlOut)
                return StripHtml(rawHtml)
            }
        } catch as e2 {
            LogDebug("StripPdf fallback error: " e2.Message)
        }
    }
    ; Fallback when no extractor: return placeholder to avoid failure
    return ""
}

ReadTextFile(path) {
    try {
        return FileRead(path, "UTF-8")
    } catch as e {
        LogDebug("ReadTextFile error")
        Throw e
    }
}

LoadDocument(path) {
    full := ResolvePath(path)
    ext := StrLower(Trim(RegExReplace(full, ".*\.([^.]+)$", "$1")))
    text := ""
    if (ext = "pdf") {
        text := StripPdf(full)
    } else if (ext = "html" || ext = "htm") {
        raw := ReadTextFile(full)
        text := StripHtml(raw)
    } else if (ext = "rtf") {
        raw := ReadTextFile(full)
        text := StripRtf(raw)
    } else {
        text := ReadTextFile(full)
    }
    return CleanForRAG(text)
}

ResolvePath(p) {
    if (RegExMatch(p, "^[a-zA-Z]:\\|^\\\\"))
        return p
    ; relative: anchor to ROOT_DIR
    return NormalizePath(ROOT_DIR "\\" p)
}

; ---------------- Text Splitter ----------------
; Splits text into chunks with overlap and optional padding context
SplitText(text, chunkSize := 1200, overlap := 200, pad := 0) {
    if chunkSize <= 0
        Throw Error("chunkSize must be > 0")
    if overlap < 0
        overlap := 0
    if overlap >= chunkSize
        overlap := chunkSize - 1
    parts := []
    len := StrLen(text)
    pos := 1
    while (pos <= len) {
        endPos := Min(len, pos + chunkSize - 1)
        chunk := SubStr(text, pos, endPos - pos + 1)
        ; padding context
        leftStart := Max(1, pos - pad)
        leftCtx := SubStr(text, leftStart, pos - leftStart)
        rightEnd := Min(len, endPos + pad)
        rightCtx := SubStr(text, endPos + 1, rightEnd - endPos)
        parts.Push({
            text: chunk,
            start: pos,
            end: endPos,
            left: leftCtx,
            right: rightCtx
        })
        if endPos = len
            break
        pos := endPos - overlap + 1
    }
    return parts
}

; ---------------- Local Vectorizer (fallback) ----------------
; Fast, hashing-based vectorizer for zero-dependency operation.
LocalHashVector(text, dim := 256) {
    v := []
    Loop dim
        v.Push(0.0)
    for token in Tokenize(text) {
        h := DJB2(token)
        idx := Mod(h, dim) + 1
        v[idx] := v[idx] + 1.0
    }
    ; L2 normalize
    norm := 0.0
    for f in v
        norm += f * f
    if norm > 0 {
        norm := Sqrt(norm)
        for i, f in v
            v[i] := f / norm
    }
    return v
}

Tokenize(text) {
    ; Simple word tokenizer
    cleaned := RegExReplace(StrLower(text), "[^a-z0-9]+", " ")
    arr := []
    for part in StrSplit(Trim(cleaned), " ")
        if (part != "")
            arr.Push(part)
    return arr
}

DJB2(s) {
    h := 5381
    Loop Parse s
        h := ((h << 5) + h) + Ord(A_LoopField)
    if (h < 0)
        h := -h
    return h
}

CosineSim(a, b) {
    if a.Length != b.Length
        Throw Error("Vector length mismatch")
    dot := 0.0, na := 0.0, nb := 0.0
    for i, x in a {
        y := b[i]
        dot += x * y
        na += x * x
        nb += y * y
    }
    if (na = 0 || nb = 0)
        return 0.0
    return dot / (Sqrt(na) * Sqrt(nb))
}

; ---------------- Vector Store ----------------
; structure: { items: [ { id, path, start, end, text, left, right, vector:Array } ] }
VS_Load(path := "") {
    global VECTOR_STORE_PATH
    p := (path = "") ? VECTOR_STORE_PATH : path
    if !FileExist(p)
        return { items: [], docs: {} }
    try {
        s := FileRead(p, "UTF-8")
        o := JSON.parse(s, false, false)
        if !o.HasOwnProp("items")
            o.items := []
        if !o.HasOwnProp("docs")
            o.docs := {}
        return o
    } catch as e {
        LogDebug("VS_Load error")
        return { items: [], docs: {} }
    }
}

Djb2Hex(s) {
    h := DJB2(s)
    ; produce unsigned 32-bit hex
    if (h < 0)
        h := h + 0x100000000
    return Format("{:08X}", h)
}

SplitTextSmart(text, chunkSize := 1200, overlap := 200, pad := 0) {
    sentences := RegExReplace(text, "[\r\n]+", " ")
    sentArr := []
    acc := ""
    for s in StrSplit(sentences, ". ") {
        if (StrLen(acc) + StrLen(s) + 2) <= chunkSize {
            acc := acc (acc = "" ? "" : ". ") s
        } else {
            if acc != ""
                sentArr.Push(acc)
            acc := s
        }
    }
    if acc != ""
        sentArr.Push(acc)
    return SplitText(StrJoin(sentArr, "`n`n"), chunkSize, overlap, pad)
}

VS_Save(store, path := "") {
    global VECTOR_STORE_PATH
    p := (path = "") ? VECTOR_STORE_PATH : path
    try {
        LogDebug("VS_Save: stringify start")
        s := JSON.stringify(store, 0, "  ")
        LogDebug("VS_Save: stringify ok len=" StrLen(s))
        try FileDelete(p)
        LogDebug("VS_Save: delete old ok or not present")
        FileAppend(s, p, "UTF-8")
        LogDebug("VS_Save: append ok")
        return true
    } catch as e {
        LogDebug("VS_Save error")
        return false
    }
}

VS_EnsureIndex(store) {
    if !store.HasOwnProp("items")
        store.items := []
    if !store.HasOwnProp("docs")
        store.docs := {}
}

VS_AddChunks(store, path, chunks, vectorizer?, docMeta := unset) {
    VS_EnsureIndex(store)
    items := store.items
    docs := store.docs
    ; document metadata
    sz := 0
    mt := ""
    try sz := FileGetSize(path)
    try mt := FileGetTime(path, "M")
    ; document hash
    docHash := ""
    if IsSet(docMeta) && docMeta.HasOwnProp("hash")
        docHash := docMeta.hash
    else {
        acc := ""
        for c0 in chunks
            acc .= c0.text
        docHash := Djb2Hex(acc)
    }
    ; if any doc has the same hash, mirror its chunkIds without duplicating
    found := ""
    for k, d in docs.OwnProps() {
        if d.HasOwnProp("hash") && d.hash = docHash {
            found := k
            break
        }
    }
    if (found != "") {
        docs.%path% := { path: path, size: sz, mtime: mt, hash: docHash, chunkIds: docs.%found%.chunkIds }
        return
    }
    ; if this path exists with different hash, prune old items for this doc
    if docs.HasOwnProp(path) {
        old := docs.%path%
        if old.HasOwnProp("hash") && old.hash != docHash {
            newItems := []
            for it in items
                if (it.path != path)
                    newItems.Push(it)
            store.items := newItems
            items := store.items
        } else if old.HasOwnProp("hash") && old.hash = docHash {
            return
        }
    }
    addedIds := []
    for c in chunks {
        vec := c.HasOwnProp("vector") ? c.vector : (IsSet(vectorizer) ? vectorizer(c.text) : LocalHashVector(c.text))
        chash := Djb2Hex(c.text)
        dup := false
        for ex in items {
            if (ex.path = path && ex.start = c.start && ex.end = c.end) {
                dup := true
                break
            }
            if (ex.HasOwnProp("hash") && ex.hash = chash) {
                dup := true
                break
            }
        }
        if dup
            continue
        newId := A_TickCount "-" items.Length + 1
        items.Push({
            id: newId,
            path: path,
            start: c.start,
            end: c.end,
            text: c.text,
            left: c.left,
            right: c.right,
            vector: vec,
            hash: chash,
            size: sz,
            mtime: mt
        })
        addedIds.Push(newId)
    }
    docs.%path% := { path: path, size: sz, mtime: mt, hash: docHash, chunkIds: addedIds }
}

VS_Query(store, queryVec, topK := 5) {
    res := []
    for it in store.items {
        try {
            s := CosineSim(queryVec, it.vector)
            sc := { score: s, item: it }
            ; insertion keep res sorted desc by score
            inserted := false
            idx := 1
            for cur in res {
                if (sc.score > cur.score) {
                    res.InsertAt(idx, sc)
                    inserted := true
                    break
                }
                idx++
            }
            if !inserted
                res.Push(sc)
            if res.Length > topK
                res.RemoveAt(res.Length)
        } catch as e {
            ; ignore malformed items
        }
    }
    return res
}

; ---------------- Async HTTP Core ----------------
; Provides async request with callback(responseText, status, headers)
HttpRequestAsync(method, url, headersObj, bodyText, callback) {
    req := ComObject("MSXML2.XMLHTTP")
    req.open(method, url, true)
    if IsSet(headersObj) {
        for k, v in headersObj {
            if (k != "__timeoutMs")
                req.setRequestHeader(k, v)
        }
    }
    deadline := 0
    try {
        if IsSet(headersObj) && headersObj.HasOwnProp("__timeoutMs")
            deadline := A_TickCount + Integer(headersObj.__timeoutMs)
    } catch {
    }
    try {
        req.send(IsSet(bodyText) ? bodyText : "")
    } catch as e {
        LogDebug("HttpRequestAsync error: " e.Message)
        try callback("", -1, "")
        return
    }
    poll := 0
    poll := () => (
        (deadline && A_TickCount > deadline)
            ? ( SetTimer(poll, 0), callback ? callback("", -1, "") : 0 )
        : ( req.readyState = 4
            ? ( SetTimer(poll, 0), callback ? callback(req.responseText, req.status, req.getAllResponseHeaders()) : 0 )
            : 0 )
    )
    SetTimer(poll, 30)
}

HttpRequestWithRetry(method, url, headersObj, bodyText, callback, retryOpts := unset) {
    ; retryOpts: {retries:int, baseDelayMs:int, backoff:float}
    maxRetries := 2, baseDelay := 400, backoff := 2.0
    if IsSet(retryOpts) {
        if retryOpts.HasOwnProp("retries")
            maxRetries := Integer(retryOpts.retries)
        if retryOpts.HasOwnProp("baseDelayMs")
            baseDelay := Integer(retryOpts.baseDelayMs)
        if retryOpts.HasOwnProp("backoff")
            backoff := retryOpts.backoff
    }
    attempt := 0
    doAttempt := 0
    doAttempt := () => (
        HttpRequestAsync(method, url, headersObj, bodyText, (txt, status, hdrs) => (
            (status = 200)
                ? callback(txt, status, hdrs)
                : (
                    (attempt < maxRetries)
                        ? (
                            delay := Integer(baseDelay * (backoff ** attempt) + Random(0, 100)),
                            attempt := attempt + 1,
                            LogDebug("Retry #" attempt " in " delay "ms status=" status),
                            SetTimer(doAttempt, -delay)
                        )
                        : callback(txt, status, hdrs)
                )
        ))
    )
    doAttempt()
}

; ---------------- Providers ----------------
; Provider-agnostic chat and embeddings

ProviderHeaders(cfg) {
    h := Map()
    prov := StrLower(cfg.Get("provider", ""))
    if (prov = "openai" || prov = "ollama" || prov = "openai-compatible") {
        if cfg.Get("apiKey", "") != ""
            h["Authorization"] := "Bearer " cfg["apiKey"]
        h["Content-Type"] := "application/json"
    } else if (prov = "gemini") {
        h["Content-Type"] := "application/json"
    } else {
        h["Content-Type"] := "application/json"
    }
    to := cfg.Get("timeoutMs", "")
    if (to != "")
        h["__timeoutMs"] := to
    stall := cfg.Get("streamStallMs", "")
    if (stall != "")
        h["__stallMs"] := stall
    return h
}

EmbeddingsAsync(cfg, inputText, callback) {
    prov := StrLower(cfg.Get("provider", "openai"))
    headers := ProviderHeaders(cfg)
    if (prov = "gemini") {
        model := cfg.Get("embedModel", "text-embedding-004")
        key := cfg.Get("apiKey", "")
        base := cfg.Get("baseUrl", "https://generativelanguage.googleapis.com")
        url := base "/v1beta/models/" model ":embedContent?key=" key
        body := { content: { parts: [{text: inputText}] } }
        tt := cfg.Get("taskType", "")
        if (tt != "")
            body.taskType := tt
        dim := cfg.Get("outputDimensionality", "")
        if (dim != "")
            body.outputDimensionality := Integer(dim)
        payload := JSON.stringify(body, 0, "  ")
        HttpRequestWithRetry("POST", url, headers, payload, (txt, status, hdrs) => callback(txt, status, hdrs))
        return
    }
    if (prov = "ollama") {
        model := cfg.Get("embedModel", "nomic-embed-text")
        base := cfg.Get("baseUrl", "http://localhost:11434")
        url := base "/api/embeddings"
        payload := JSON.stringify({ model: model, prompt: inputText }, 0, "  ")
        HttpRequestWithRetry("POST", url, headers, payload, (txt, status, hdrs) => callback(txt, status, hdrs))
        return
    }
    ; OpenAI-compatible embeddings
    model := cfg.Get("embedModel", "text-embedding-3-small")
    base := cfg.Get("baseUrl", "https://api.openai.com")
    url := base "/v1/embeddings"
    payload := JSON.stringify({ model: model, input: inputText }, 0, "  ")
    HttpRequestWithRetry("POST", url, headers, payload, (txt, status, hdrs) => callback(txt, status, hdrs))
}

; Streaming request with provider-specific parsing; delivers onDelta(text) and onDone(fullText, status)
HttpStreamRequest(prov, method, url, headersObj, bodyText, onDelta, onDone) {
    req := ComObject("MSXML2.XMLHTTP")
    req.open(method, url, true)
    if IsSet(headersObj) {
        for k, v in headersObj {
            if (k != "__timeoutMs")
                req.setRequestHeader(k, v)
        }
    }
    deadline := 0
    stallMs := 0
    try {
        if IsSet(headersObj) && headersObj.HasOwnProp("__timeoutMs")
            deadline := A_TickCount + Integer(headersObj.__timeoutMs)
        if IsSet(headersObj) && headersObj.HasOwnProp("__stallMs")
            stallMs := Integer(headersObj.__stallMs)
    } catch {
    }
    try req.send(IsSet(bodyText) ? bodyText : "")
    catch as e {
        LogDebug("HttpStreamRequest send error: " e.Message)
        try onDone("", -1)
        return
    }
    bufStream := ""
    lastLen := 0
    full := ""
    lowerProv := StrLower(prov)
    poll := 0
    poll := () => (
        (deadline && A_TickCount > deadline)
            ? ( SetTimer(poll, 0), onDone ? onDone(full, -1) : 0 )
        : (
            (req.readyState >= 3)
                ? (
                    ( __TryReadText(req, &t) ),
                    add := (IsSet(t) ? SubStr(t, lastLen + 1) : ""),
                    lastLen := (IsSet(t) ? StrLen(t) : lastLen),
                    (add != "") ? ( lastChange := A_TickCount, bufStream .= add, __StreamParse(lowerProv, &bufStream, &full, onDelta) ) : 0
                ) : 0,
            (stallMs > 0 && req.readyState < 4 && (A_TickCount - lastChange) > stallMs)
                ? ( __TryAbort(req), SetTimer(poll, 0), onDone ? onDone(full, -2) : 0 )
                : ((req.readyState = 4) ? ( SetTimer(poll, 0), onDone ? onDone(full, req.status) : 0 ) : 0)
        )
    )
    SetTimer(poll, 30)
}

__SSEPop(&buffer) {
    ; Return next SSE block (string before a blank line). Handles CRLF/CR/LF.
    try {
        if (buffer = "")
            return ""
        pos := RegExMatch(buffer, "\R\R", &m)
        if (pos = 0)
            return ""
        block := SubStr(buffer, 1, pos - 1)
        buffer := SubStr(buffer, pos + m.Len(0))
        return block
    } catch as e {
        return ""
    }
}

__StreamParse(lowerProv, &buffer, &full, onDelta) {
    if (lowerProv = "gemini") {
        loop {
            block := __SSEPop(&buffer)
            if (block = "")
                break
            for line in StrSplit(Trim(block, " `r`n"), "`n") {
                if SubStr(line, 1, 5) = "data:" {
                    payload := Trim(SubStr(line, 6))
                    if (payload = "" || payload = "[DONE]")
                        continue
                    try {
                        o := JSON.parse(payload, false, false)
                        if o.HasOwnProp("candidates") {
                            for cand in o.candidates {
                                ; Preferred: delta parts
                                if cand.HasOwnProp("delta") {
                                    d := cand.delta
                                    if d.HasOwnProp("parts") {
                                        for p in d.parts {
                                            if p.HasOwnProp("text") {
                                                t := p.text
                                                if (t != "") {
                                                    full .= t
                                                    if IsSet(onDelta)
                                                        onDelta(t)
                                                }
                                            }
                                        }
                                    }
                                    if d.HasOwnProp("text") { ; fallback shape
                                        t := d.text
                                        if (t != "") {
                                            full .= t
                                            if IsSet(onDelta)
                                                onDelta(t)
                                        }
                                    }
                                }
                                ; Some events stream full content.parts repeatedly
                                if cand.HasOwnProp("content") && cand.content.HasOwnProp("parts") {
                                    for p in cand.content.parts {
                                        if p.HasOwnProp("text") {
                                            t := p.text
                                            if (t != "") {
                                                full .= t
                                                if IsSet(onDelta)
                                                    onDelta(t)
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    } catch {
                    }
                }
            }
        }
        return
    }
    if (lowerProv = "openai" || lowerProv = "openai-compatible") {
        loop {
            block := __SSEPop(&buffer)
            if (block = "")
                break
            for line in StrSplit(Trim(block, " `r`n"), "`n") {
                if SubStr(line, 1, 5) = "data:" {
                    payload := Trim(SubStr(line, 6))
                    if (payload = "[DONE]")
                        continue
                    try {
                        o := JSON.parse(payload, false, false)
                        if o.HasOwnProp("choices") && o.choices.Length >= 1 {
                            ch := o.choices[1]
                            if ch.HasOwnProp("delta") && ch.delta.HasOwnProp("content") {
                                text := ch.delta.content
                                if (text != "") {
                                    full .= text
                                    if IsSet(onDelta)
                                        onDelta(text)
                                }
                            }
                        }
                    } catch {
                    }
                }
            }
        }
        return
    }
    if (lowerProv = "ollama") {
        loop {
            pos := InStr(buffer, "`n")
            if !pos
                break
            line := SubStr(buffer, 1, pos - 1)
            buffer := SubStr(buffer, pos + 1)
            if (Trim(line) = "")
                continue
            try {
                o := JSON.parse(line, false, false)
                if o.HasOwnProp("message") && o.message.HasOwnProp("content") {
                    text := o.message.content
                    if (text != "") {
                        full .= text
                        if IsSet(onDelta)
                            onDelta(text)
                    }
                }
            } catch {
            }
        }
        return
    }
}

__TryReadText(req, &outText := unset) {
    try {
        outText := req.responseText
    } catch as e {
        ; ignore; not yet available
    }
}

__TryAbort(req) {
    try req.abort()
    catch as e {
    }
}

ChatCompletionStreamAsync(cfg, messages, onDelta, onDone) {
    prov := StrLower(cfg.Get("provider", "openai"))
    headers := ProviderHeaders(cfg)
    headers["Accept"] := "text/event-stream"
    headers["Cache-Control"] := "no-cache"
    ; add system message for non-gemini providers
    if !(prov = "gemini")
        messages := AddSystemToMessages(messages)
    if (prov = "gemini") {
        model := cfg.Get("chatModel", "gemini-1.5-flash")
        key := cfg.Get("apiKey", "")
        base := cfg.Get("baseUrl", "https://generativelanguage.googleapis.com")
        url := base "/v1beta/models/" model ":streamGenerateContent?alt=sse&key=" key
        msgs := messages
        payload := JSON.stringify(BuildGeminiGeneratePayload(msgs, cfg), 0, "  ")
        HttpStreamRequest(prov, "POST", url, headers, payload, onDelta, onDone)
        return
    }
    if (prov = "ollama") {
        model := cfg.Get("chatModel", "llama3")
        base := cfg.Get("baseUrl", "http://localhost:11434")
        url := base "/api/chat"
        obj := { model: model, messages: messages }
        obj.stream := JSON.true
        payload := JSON.stringify(obj, 0, "  ")
        HttpStreamRequest(prov, "POST", url, headers, payload, onDelta, onDone)
        return
    }
    ; OpenAI-compatible streaming via SSE
    model := cfg.Get("chatModel", "gpt-4o-mini")
    base := cfg.Get("baseUrl", "https://api.openai.com")
    url := base "/v1/chat/completions"
    obj := { model: model, messages: messages }
    temp := cfg.Get("temperature", "")
    if (temp != "")
        obj.temperature := (temp + 0)
    mt := cfg.Get("max_tokens", "")
    if (mt != "")
        obj.max_tokens := Integer(mt)
    obj.stream := JSON.true
    payload := JSON.stringify(obj, 0, "  ")
    HttpStreamRequest(prov, "POST", url, headers, payload, onDelta, onDone)
}

; Stream wrapper with retry: retries only if no data was received
ChatCompletionStreamWithRetry(cfg, messages, onDelta, onDone, retryOpts := unset) {
    maxRetries := 1, baseDelay := 500, backoff := 2.0
    if IsSet(retryOpts) {
        if retryOpts.HasOwnProp("retries")
            maxRetries := Integer(retryOpts.retries)
        if retryOpts.HasOwnProp("baseDelayMs")
            baseDelay := Integer(retryOpts.baseDelayMs)
        if retryOpts.HasOwnProp("backoff")
            backoff := retryOpts.backoff
    }
    attempt := 0
    emittedLen := 0
    totalBuf := ""
    startStreamInit := 0
    startStream := () => (
        ChatCompletionStreamAsync(cfg, messages,
            (delta) => (
                totalBuf .= delta,
                newPart := SubStr(totalBuf, emittedLen + 1),
                (newPart != "" && IsSet(onDelta)) ? onDelta(newPart) : 0,
                emittedLen := StrLen(totalBuf)
            ),
            (full, status) => (
                (status = 200)
                    ? ( IsSet(onDone) ? onDone(full, status) : 0 )
                    : (
                        (attempt < maxRetries)
                            ? (
                                delay := Integer(baseDelay * (backoff ** attempt)),
                                attempt := attempt + 1,
                                SetTimer(startStream, -delay)
                              )
                            : ( IsSet(onDone) ? onDone(full, status) : 0 )
                      )
            )
        )
    )
    startStream()
}

; Helper to extract embedding vectors from provider responses.
ParseEmbeddingVector(prov, responseText) {
    try {
        o := JSON.parse(responseText, false, false)
        p := StrLower(prov)
        if (p = "gemini") {
            ; {embedding:{values:[...]}}
            if o.HasOwnProp("embedding") && o.embedding.HasOwnProp("values")
                return o.embedding.values
        } else if (p = "ollama") {
            if o.HasOwnProp("embedding")
                return o.embedding
        } else {
            ; OpenAI-compatible: {data:[{embedding:[...]}]}
            if o.HasOwnProp("data") && o.data.Length >= 1
                return o.data[1].embedding
        }
    } catch as e {
        LogDebug("ParseEmbeddingVector error")
    }
    return []
}

ExtractChatText(prov, responseText) {
    try {
        p := StrLower(prov)
        o := JSON.parse(responseText, false, false)
        if (p = "gemini") {
            if o.HasOwnProp("candidates") && o.candidates.Length >= 1 {
                cand := o.candidates[1]
                if cand.HasOwnProp("content") && cand.content.HasOwnProp("parts") && cand.content.parts.Length >= 1 {
                    part := cand.content.parts[1]
                    return part.HasOwnProp("text") ? part.text : responseText
                }
            }
        } else if (p = "ollama") {
            if o.HasOwnProp("message") && o.message.HasOwnProp("content")
                return o.message.content
        } else {
            if o.HasOwnProp("choices") && o.choices.Length >= 1 {
                ch := o.choices[1]
                if ch.HasOwnProp("message") && ch.message.HasOwnProp("content")
                    return ch.message.content
            }
        }
    } catch as e {
        LogDebug("ExtractChatText parse error: " e.Message)
    }
    return responseText
}

; ---------------- RAG High-level ----------------
; Ingest a file path into vector store using provided embedding backend.
IngestFile(path, opts := unset) {
    try {
        text := LoadDocument(path)
        chunkSize := 1200, overlap := 200, pad := 0
        if IsSet(opts) {
            if opts.HasOwnProp("chunkSize")
                chunkSize := opts.chunkSize
            if opts.HasOwnProp("overlap")
                overlap := opts.overlap
            if opts.HasOwnProp("pad")
                pad := opts.pad
        }
        chunks := SplitText(text, chunkSize, overlap, pad)
        store := VS_Load()
        ; choose vectorizer
        vectorizer := (IsSet(opts) && opts.HasOwnProp("vectorizer")) ? opts.vectorizer : LocalHashVector
        docHash := Djb2Hex(text)
        VS_AddChunks(store, path, chunks, vectorizer, { hash: docHash })
        saveOk := VS_Save(store)
        LogDebug("Ingested " chunks.Length " chunks from: " path)
        Return saveOk
    } catch as e {
        LogDebug("IngestFile error: " e.Message)
        return false
    }
}

EmbeddingsBatchGeminiAsync(cfg, inputTexts, callback) {
    headers := ProviderHeaders(cfg)
    model := cfg.Get("embedModel", "text-embedding-004")
    key := cfg.Get("apiKey", "")
    base := cfg.Get("baseUrl", "https://generativelanguage.googleapis.com")
    url := base "/v1beta/models/" model ":batchEmbedContents?key=" key
    reqs := []
    tt := cfg.Get("taskType", "")
    dim := cfg.Get("outputDimensionality", "")
    for t in inputTexts {
        r := { model: "models/" model, content: { parts: [{ text: t }] } }
        if (tt != "")
            r.taskType := tt
        if (dim != "")
            r.outputDimensionality := Integer(dim)
        reqs.Push(r)
    }
    body := { requests: reqs }
    payload := JSON.stringify(body, 0, "  ")
    HttpRequestWithRetry("POST", url, headers, payload, (txt, status, hdrs) => callback(txt, status, hdrs))
}

EmbeddingsBatchOpenAIAsync(cfg, inputTexts, callback) {
    prov := StrLower(cfg.Get("provider", "openai"))
    if !(prov = "openai" || prov = "openai-compatible") {
        results := []
        pending := inputTexts.Length
        for t in inputTexts {
            EmbeddingsAsync(cfg, t, (txt, status, hdrs) => (
                vec := (status = 200) ? ParseEmbeddingVector("openai", txt) : [],
                results.Push(vec),
                pending := pending - 1,
                (pending = 0) ? callback(JSON.stringify({ data: results }, 0, "  "), 200, "") : 0
            ))
        }
        return
    }
    headers := ProviderHeaders(cfg)
    model := cfg.Get("embedModel", "text-embedding-3-small")
    base := cfg.Get("baseUrl", "https://api.openai.com")
    url := base "/v1/embeddings"
    payload := JSON.stringify({ model: model, input: inputTexts }, 0, "  ")
    HttpRequestWithRetry("POST", url, headers, payload, (txt, status, hdrs) => callback(txt, status, hdrs))
}

ParseEmbeddingVectors(prov, responseText) {
    try {
        p := StrLower(prov)
        o := JSON.parse(responseText, false, false)
        if (p = "gemini") {
            arr := []
            if o.HasOwnProp("embeddings") {
                for it in o.embeddings
                    arr.Push(it.values)
            }
            return arr
        }
        if (p = "openai" || p = "openai-compatible") {
            arr := []
            if o.HasOwnProp("data") {
                for it in o.data
                    arr.Push(it.embedding)
            }
            return arr
        }
    } catch as e {
        LogDebug("ParseEmbeddingVectors error")
    }
    return []
}

VS_RemoveDoc(store, path) {
    VS_EnsureIndex(store)
    if !store.docs.HasOwnProp(path)
        return
    newItems := []
    for it in store.items
        if (it.path != path)
            newItems.Push(it)
    store.items := newItems
    store.docs.Delete(path)
}

VS_ListDocs(store) {
    VS_EnsureIndex(store)
    return store.docs
}

; Query RAG using either provider embeddings or local vectorizer.
RAG_Query(prompt, topK := 5, useProvider := false, cfg := unset, cb := unset) {
    ; cb optional: if provided and useProvider=true, executes async and calls cb(results:Array)
    store := VS_Load()
    if store.items.Length = 0 {
        LogDebug("RAG_Query: empty store")
        return []
    }
    if !useProvider {
        qv := LocalHashVector(CleanForRAG(prompt))
        res := VS_Query(store, qv, topK)
        return res
    }
    if !IsSet(cfg)
        cfg := LoadConfig()
    prov := StrLower(cfg.Get("provider", "openai"))
    EmbeddingsAsync(cfg, CleanForRAG(prompt), (txt, status, hdrs) => (
        vec := (status = 200) ? ParseEmbeddingVector(prov, txt) : [],
        v := (vec.Length > 0) ? vec : LocalHashVector(CleanForRAG(prompt)),
        results := VS_Query(store, v, topK),
        IsSet(cb) ? cb(results) : 0
    ))
    return []
}

; Compose a contextual prompt from retrieved chunks.
BuildContextPrompt(prompt, retrieved) {
    ctx := ""
    i := 1
    for sc in retrieved {
        it := sc.item
        fileName := __GetFileName(it.path)
        span := __MapChunkToLineSpan(it.path, it.start, it.end)
        lineInfo := "lines " span.lineStart "-" span.lineEnd
        
        ctx .= "\n---\n"
        ctx .= "Source " i ": " fileName " (" lineInfo ")\n"
        ctx .= it.text "\n"
        i++
    }
    clean := CleanForRAG(ctx)
    return "Use only the following context to answer.\n" clean "\n\nQuestion: " CleanForRAG(prompt)
}

; ---------------- Citations ----------------

__NormalizeForMatch(s) {
    try {
        s := StrLower(CleanForRAG(s))
        s := RegExReplace(s, "[^a-z0-9]+", " ")
        return Trim(s)
    } catch as e {
        return s
    }
}

__WordSet(s) {
    set := Map()
    for w in StrSplit(__NormalizeForMatch(s), " ") {
        if (w != "")
            set[w] := 1
    }
    return set
}

; Very light overlap score: |words∩| / |words(text)|
__OverlapScore(ansSet, text) {
    tset := __WordSet(text)
    tot := 0, inter := 0
    for word, _ in tset {
        tot++
        if ansSet.Has(word)
            inter++
    }
    return (tot = 0) ? 0.0 : inter / tot
}

__MergeRanges(ranges) {
    if (ranges.Length = 0)
        return ranges
    ; insertion sort by start
    sorted := []
    for r in ranges {
        inserted := false
        i := 1
        while (i <= sorted.Length) {
            if (r[1] < sorted[i][1]) {
                sorted.InsertAt(i, r), inserted := true, i := sorted.Length + 1
            } else i++
        }
        if !inserted
            sorted.Push(r)
    }
    ; merge overlaps/adjacent
    out := []
    curStart := sorted[1][1], curEnd := sorted[1][2]
    i := 2
    while (i <= sorted.Length) {
        s := sorted[i][1], e := sorted[i][2]
        if (s <= curEnd + 1) {
            if (e > curEnd)
                curEnd := e
        } else {
            out.Push([curStart, curEnd])
            curStart := s, curEnd := e
        }
        i++
    }
    out.Push([curStart, curEnd])
    return out
}

__CfgCitationMaxFiles(cfg) {
    try {
        v := Integer(cfg.Get("citationMaxFiles", "3"))
    } catch {
        v := 3
    }
    return (v >= 1) ? v : 3
}

__CfgCitationThreshold(cfg) {
    try {
        f := cfg.Get("citationMatchThreshold", "0.12")
        v := Float(f)
    } catch {
        v := 0.12
    }
    return (v > 0 && v < 1) ? v : 0.12
}

shouldShowCitations(responseText) {
    responseText := StrLower(responseText)
    denialPhrases := [
        "don't know",
        "do not know",
        "cannot answer",
        "can't answer",
        "unable to answer",
        "context does not contain",
        "no information"
    ]

    for phrase in denialPhrases {
        if InStr(responseText, phrase) {
            return false
        }
    }

    salutationPhrases := [
        "hello",
        "hi",
        "greetings",
        "you're welcome",
        "you are welcome",
        "my pleasure",
        "glad to help"
        "glad i could help"
        "doing well",
        "doing pretty well",
		"help you",
		"thank you"
    ]

    if (StrLen(responseText) < 60) {
        for phrase in salutationPhrases {
            if InStr(responseText, phrase) {
                return false
            }
        }
    }

    return true
}

__CfgShowCitations(cfg) {
    v := ""
    try v := StrLower(Trim(cfg.Get("showCitations", "")))
    catch {
    }
    return (v = "1" || v = "on" || v = "true" || v = "yes")
}

__MapChunkToLineSpan(path, startPos, endPos) {
    ; startPos/endPos are 1-based indices in CleanForRAG(text)
    lineStart := 1, lineEnd := 1
    try {
        full := LoadDocument(path) ; already CleanForRAG
        if (StrLen(full) = 0)
            return { lineStart: 1, lineEnd: 1 }
        pre := (startPos > 1) ? SubStr(full, 1, startPos - 1) : ""
        upto := (endPos >= 1) ? SubStr(full, 1, endPos) : pre
        ; Count `n occurrences
        cnt := 0
        if (pre != "") {
            Loop Parse pre
                if (A_LoopField = "`n")
                    cnt := cnt + 1
        }
        lineStart := cnt + 1
        cnt2 := 0
        if (upto != "") {
            Loop Parse upto
                if (A_LoopField = "`n")
                    cnt2 := cnt2 + 1
        }
        lineEnd := cnt2 + 1
    } catch as e {
        ; default lines
    }
    return { lineStart: lineStart, lineEnd: lineEnd }
}

__BuildCitationsBlock(retrieved, answerText := "", cfg := unset) {
    try {
        if !IsSet(retrieved) || retrieved.Length = 0
            return ""
        maxFiles := IsSet(cfg) ? __CfgCitationMaxFiles(cfg) : 3
        thresh   := IsSet(cfg) ? __CfgCitationThreshold(cfg) : 0.12

        ansSet := __WordSet(answerText)
        byFile := Map()
        fileScore := Map()

        ; score chunks against the FINAL answer, keep per-file line ranges
        for sc in retrieved {
            it := sc.item
            fname := __GetFileName(it.path)
            score := (answerText = "") ? 0.0 : __OverlapScore(ansSet, it.text)

            if !byFile.Has(fname)
                byFile[fname] := []
            if !fileScore.Has(fname)
                fileScore[fname] := 0.0

            ; keep spans only if they reasonably appear to contribute to the answer
            if (score >= thresh || answerText = "") {
                span := __MapChunkToLineSpan(it.path, it.start, it.end)
                s := span.HasOwnProp("lineStart") ? span.lineStart : it.start
                e := span.HasOwnProp("lineEnd")   ? span.lineEnd   : it.end
                byFile[fname].Push([s, e])
                fileScore[fname] := fileScore[fname] + score
            } else {
                ; still track score for fallback
                fileScore[fname] := Max(fileScore[fname], score)
            }
        }

        ; fallback: if no file has ranges (e.g. threshold too high), take best single file
        haveAny := false
        for f, arr in byFile {
            if (arr.Length > 0) {
                haveAny := true
                break
            }
        }
        if !haveAny {
            ; choose the file with the highest score and include its top chunk span(s)
            bestFile := ""
            best := -1.0
            for f, s in fileScore {
                if (s > best)
                    best := s, bestFile := f
            }
            if (bestFile != "") {
                ; include all its chunks (they're the only ones we can justify)
                for sc in retrieved {
                    if (__GetFileName(sc.item.path) = bestFile) {
                        span := __MapChunkToLineSpan(sc.item.path, sc.item.start, sc.item.end)
                        s := span.HasOwnProp("lineStart") ? span.lineStart : sc.item.start
                        e := span.HasOwnProp("lineEnd")   ? span.lineEnd   : sc.item.end
                        byFile[bestFile].Push([s, e])
                    }
                }
            }
        }

        ; order files by score (desc) and cap at maxFiles
        files := []
        for f, s in fileScore
            files.Push({ file: f, score: s })  ; property name "file" is OK

        ; sort by score desc (simple)
        i := 1
        while (i <= files.Length) {
            j := i + 1
            while (j <= files.Length) {
                if (files[j].score > files[i].score) {
                    tmp := files[i], files[i] := files[j], files[j] := tmp
                }
                j++
            }
            i++
        }

        out := "Sources:`r`n"
        shown := 0, idx := 1
        for fobj in files {
            f := fobj.file
            if !byFile.Has(f) || byFile[f].Length = 0
                continue
            rng := __MergeRanges(byFile[f])
            parts := []
            for r in rng {
                parts.Push((r[1] = r[2]) ? "line " r[1] : "lines " r[1] "-" r[2])
            }
            out .= Format("{1}) {2} ({3})`r`n", idx, f, StrJoin(parts, ", "))
            idx++, shown++
            if (shown >= maxFiles)
                break
        }
        return out
    } catch as e {
        return ""
    }
}


StrJoin(arr, sep := ", ") {
    s := ""
    for v in arr
        s .= (s = "" ? v : sep v)
    return s
}


__GetFileName(path) {
    try {
        SplitPath(path, &name)
        if (name != "")
            return name
    } catch as e {
    }
    ; Fallback: strip dirs
    p := path
    p := RegExReplace(p, ".*[\\/]", "")
    return p
}

CleanForRAG(text) {
    try {
        s := text
        s := StrReplace(s, "`r`n", "`n")
        s := StrReplace(s, "`r", "`n")
        s := StrReplace(s, "`t", " ")
        s := RegExReplace(s, " {2,}", " ")
        s := RegExReplace(s, "\R{3,}", "`n`n")
        out := ""
        for line in StrSplit(s, "`n") {
            out .= (out = "" ? "" : "`n") Trim(line)
        }
        return out
    } catch as e {
        return text
    }
}

BuildContextDump(retrieved) {
    out := ""
    i := 1
    for sc in retrieved {
        it := sc.item
        out .= "#" i " score=" sc.score " path=" it.path " (" it.start "-" it.end ")\r\n"
        out .= it.text "\r\n\r\n"
        i++
    }
    return out
}

; Chat with provider; non-blocking async only (to avoid UI stalls)
ChatWithRAGAsync(userPrompt, cfg := unset, topK := 5, cb := unset) {
    if !IsSet(cfg)
        cfg := LoadConfig()
    prov := StrLower(cfg.Get("provider", "openai"))
    handler := (retrieved) => __RAG_Retrieved(userPrompt, prov, cfg, cb, retrieved)
    RAG_Query(userPrompt, topK, true, cfg, handler)
}

; Non-streaming chat completion
ChatCompletionAsync(cfg, messages, callback) {
    prov := StrLower(cfg.Get("provider", "openai"))
    headers := ProviderHeaders(cfg)
    if (prov = "gemini") {
        model := cfg.Get("chatModel", "gemini-1.5-flash")
        key := cfg.Get("apiKey", "")
        base := cfg.Get("baseUrl", "https://generativelanguage.googleapis.com")
        url := base "/v1beta/models/" model ":generateContent?key=" key
        payload := JSON.stringify(BuildGeminiGeneratePayload(messages, cfg), 0, "  ")
        HttpRequestWithRetry("POST", url, headers, payload, (txt, status, hdrs) => callback(txt, status, hdrs))
        return
    }
    if (prov = "ollama") {
        model := cfg.Get("chatModel", "llama3")
        base := cfg.Get("baseUrl", "http://localhost:11434")
        url := base "/api/chat"
        obj := { model: model, messages: messages }
        payload := JSON.stringify(obj, 0, "  ")
        HttpRequestWithRetry("POST", url, headers, payload, (txt, status, hdrs) => callback(txt, status, hdrs))
        return
    }
    ; OpenAI-compatible non-streaming chat
    model := cfg.Get("chatModel", "gpt-4o-mini")
    base := cfg.Get("baseUrl", "https://api.openai.com")
    url := base "/v1/chat/completions"
    messages := AddSystemToMessages(messages)
    obj := { model: model, messages: messages }
    temp := cfg.Get("temperature", "")
    if (temp != "")
        obj.temperature := (temp + 0)
    mt := cfg.Get("max_tokens", "")
    if (mt != "")
        obj.max_tokens := Integer(mt)
    payload := JSON.stringify(obj, 0, "  ")
    HttpRequestWithRetry("POST", url, headers, payload, (txt, status, hdrs) => callback(txt, status, hdrs))
}

__RAG_Retrieved(userPrompt, prov, cfg, cb, retrieved) {
    full := BuildContextPrompt(userPrompt, retrieved)
    messages := [ {role: "user", content: full} ]
    doneCb := (txt, status, hdrs := unset) => __ChatDoneCitations(userPrompt, prov, cfg, cb, retrieved, txt, status)
    __CallChat(cfg, messages, doneCb)
}

__ChatDoneCitations(userPrompt, prov, cfg, cb, retrieved, txt, status) {
    reply := ExtractChatText(prov, txt)
    if __CfgShowCitations(cfg) && shouldShowCitations(reply) {
        reply := reply "`r`n`r`n" __BuildCitationsBlock(retrieved, reply, cfg)
        LogDebug("Citations appended to reply")
    }
    SaveChatHistoryEntry(userPrompt, reply, status)
    if IsSet(cb)
        cb(reply, status)
}


__Chat_Done(userPrompt, prov, cb, txt, status, hdrs := unset) {
    LogDebug("ChatWithRAGAsync got status=" status)
    reply := ExtractChatText(prov, txt)
    LogDebug("Saving chat history")
    SaveChatHistoryEntry(userPrompt, reply, status)
    if IsSet(cb)
        cb(reply, status)
}

__CallChat(cfg, messages, callback) {
    ChatCompletionAsync(cfg, messages, callback)
}

ChatWithRAGStreamAsync(userPrompt, cfg := unset, topK := 5, onDelta := unset, onDone := unset) {
    if !IsSet(cfg)
        cfg := LoadConfig()
    prov := StrLower(cfg.Get("provider", "openai"))
    RAG_Query(userPrompt, topK, true, cfg, (retrieved) => (
        full := BuildContextPrompt(userPrompt, retrieved),
        messages := [ {role: "user", content: full} ],
        ChatCompletionStreamWithRetry(cfg, messages,
            (text) => ( IsSet(onDelta) ? onDelta(text) : 0 ),
            (fullText, status) => __StreamDoneCitations(userPrompt, cfg, retrieved, onDone, fullText, status)
        )
    ))
}

__StreamDoneCitations(userPrompt, cfg, retrieved, onDone, fullText, status) {
    finalText := fullText
    if __CfgShowCitations(cfg) && shouldShowCitations(finalText) {
        finalText := finalText "`r`n`r`n" __BuildCitationsBlock(retrieved, finalText, cfg)
        LogDebug("Citations appended to stream reply")
    }
    SaveChatHistoryEntry(userPrompt, finalText, status)
    if IsSet(onDone)
        onDone(finalText, status)
}


; ---------------- Chat History ----------------
LoadChatHistory() {
    global CHAT_HISTORY_PATH
    if !FileExist(CHAT_HISTORY_PATH)
        return []
    try {
        s := FileRead(CHAT_HISTORY_PATH, "UTF-8")
        return JSON.parse(s, false, true)
    } catch as e {
        LogDebug("LoadChatHistory error")
        return []
    }
}

SaveChatHistory(arr := unset) {
    global CHAT_HISTORY_PATH
    try {
        data := IsSet(arr) ? arr : []
        s := JSON.stringify(data, 0, "  ")
        ; ensure directory exists
        dir := SubStr(CHAT_HISTORY_PATH, 1, InStr(CHAT_HISTORY_PATH, "\",, -1) - 1)
        try DirCreate(dir)
        try FileDelete(CHAT_HISTORY_PATH)
        f := FileOpen(CHAT_HISTORY_PATH, "w", "UTF-8")
        if !f
            throw Error("Failed to open chat history for write")
        f.Write(s)
        f.Close()
    } catch as e {
        LogDebug("SaveChatHistory error: " e.Message)
    }
}

; ---------------- Bulk Ingest Queue (non-blocking) ----------------
global gIngestQueue := []
global gIngestActive := false
global gIngestOpts := unset
global gIngestDoneCb := unset

IngestQueueStart(files, opts := unset, doneCb := unset, intervalMs := 15) {
    global gIngestQueue, gIngestActive, gIngestOpts, gIngestDoneCb
    gIngestQueue := []
    for f in files
        gIngestQueue.Push(f)
    gIngestOpts := IsSet(opts) ? opts : unset
    gIngestDoneCb := IsSet(doneCb) ? doneCb : unset
    if gIngestActive {
        SetTimer(__IngestTick, 0)
        gIngestActive := false
    }
    gIngestActive := true
    SetTimer(__IngestTick, -intervalMs) ; negative = run once after delay
    SetTimer(__IngestTick, intervalMs)  ; periodic
    LogDebug("IngestQueue started: " gIngestQueue.Length " files")
}

IngestQueueStop() {
    global gIngestActive
    gIngestActive := false
    SetTimer(__IngestTick, 0)
    LogDebug("IngestQueue stopped")
}

__IngestTick() {
    global gIngestQueue, gIngestActive, gIngestOpts, gIngestDoneCb
    if !gIngestActive {
        SetTimer(__IngestTick, 0)
        return
    }
    if gIngestQueue.Length = 0 {
        SetTimer(__IngestTick, 0)
        gIngestActive := false
        LogDebug("IngestQueue done")
        if IsSet(gIngestDoneCb)
            try gIngestDoneCb()
        return
    }
    nextPath := gIngestQueue.RemoveAt(1)
    LogDebug("IngestQueue tick -> " nextPath)
    ; Ingest one file per tick
    ok := IngestFile(nextPath, IsSet(gIngestOpts) ? gIngestOpts : unset)
    LogDebug("IngestQueue result: " (ok ? "ok" : "fail"))
}

; Convenience: ingest all supported files in a directory (non-blocking queue)
IngestDirectory(dir, opts := unset) {
    files := []
    for ext in ["*.txt", "*.htm", "*.html", "*.rtf"] {
        Loop Files dir "\\" ext
            files.Push(A_LoopFileFullPath)
    }
    IngestQueueStart(files, opts)
}

SaveChatHistoryEntry(user, rawResponseText, status) {
    hist := LoadChatHistory()
    entry := {
        ts: A_Now,
        user: user,
        status: status,
        raw: rawResponseText
    }
    hist.Push(entry)
    SaveChatHistory(hist)
}

; ---------------- Module Init Log ----------------
LogDebug("AIAssistant loaded.")

; ---------------- Provider-Embeddings Ingest (non-blocking) ----------------
global gPE_State := {}

IngestFileProviderEmbeddingsAsync(path, cfg := unset, opts := unset, doneCb := unset) {
    try {
        global gPE_State
        if !IsSet(cfg)
            cfg := LoadConfig()
        text := LoadDocument(path)
        chunkSize := 1200, overlap := 200, pad := 0
        if IsSet(opts) {
            if opts.HasOwnProp("chunkSize")
                chunkSize := opts.chunkSize
            if opts.HasOwnProp("overlap")
                overlap := opts.overlap
            if opts.HasOwnProp("pad")
                pad := opts.pad
        }
        chunks := SplitText(text, chunkSize, overlap, pad)
        prov := StrLower(cfg.Get("provider", "openai"))
        LogDebug("PE: start provider=" prov " path=" path " chunks=" chunks.Length)
        if (prov = "openai" || prov = "openai-compatible") {
            texts := []
            for c in chunks
                texts.Push(c.text)
            LogDebug("PE: OpenAI batch texts=" texts.Length)
            EmbeddingsBatchOpenAIAsync(cfg, texts, (txt, status, hdrs) => (
                __PE_AssignAndSaveBatch(prov, path, text, chunks, txt, status, doneCb)
            ))
            return
        }
        if (prov = "gemini") {
            texts := []
            for c in chunks
                texts.Push(c.text)
            LogDebug("PE: Gemini batch texts=" texts.Length)
            EmbeddingsBatchGeminiAsync(cfg, texts, (txt, status, hdrs) => (
                __PE_AssignAndSaveBatch(prov, path, text, chunks, txt, status, doneCb)
            ))
            return
        }
        gPE_State := {
            path: path,
            cfg: cfg,
            chunks: chunks,
            i: 1,
            store: VS_Load(),
            done: doneCb,
            docHash: Djb2Hex(text)
        }
        SetTimer(__IngestPE_Tick, -10)
        SetTimer(__IngestPE_Tick, 25)
    } catch as e {
        LogDebug("IngestFileProviderEmbeddingsAsync error: " e.Message)
        try doneCb(false)
    }
}

__IngestPE_Tick() {
    global gPE_State
    st := gPE_State
    if !IsSet(st) || !st.HasOwnProp("chunks") {
        SetTimer(__IngestPE_Tick, 0)
        return
    }
    idx := st.i
    if (idx > st.chunks.Length) {
        SetTimer(__IngestPE_Tick, 0)
        try {
            VS_AddChunks(st.store, st.path, st.chunks, LocalHashVector, { hash: st.docHash })
            VS_Save(st.store)
        } catch as e {
            LogDebug("PE finalize save error: " e.Message)
        }
        if st.HasOwnProp("done")
            try st.done(true)
        gPE_State := {}
        return
    }
    c := st.chunks[idx]
    cfg := st.cfg
    EmbeddingsAsync(cfg, c.text, (txt, status, hdrs) => (
        LogDebug("PE tick status=" status)
        vec := (status = 200) ? ParseEmbeddingVector(StrLower(cfg.Get("provider", "openai")), txt) : [],
        c.vector := (vec.Length > 0) ? vec : LocalHashVector(c.text),
        gPE_State.i := idx + 1
    ))
}

__PE_AssignAndSaveBatch(prov, path, fullText, chunks, responseText, status, doneCb := unset) {
    try {
        LogDebug("PE batch status=" status)
        vecs := (status = 200) ? ParseEmbeddingVectors(prov, responseText) : []
        i := 1
        for c in chunks {
            v := (vecs.Length >= i && vecs[i].Length > 0) ? vecs[i] : LocalHashVector(c.text)
            c.vector := v
            i++
        }
        st := VS_Load()
        docHash := Djb2Hex(fullText)
        if st.HasOwnProp("docs") && st.docs.HasOwnProp(path) && st.docs.%path%.HasOwnProp("chunkIds") && st.docs.%path%.HasOwnProp("hash") && st.docs.%path%.hash = docHash {
            ; Update in-place vectors for existing items
            ; Build index for quick lookup
            idxByKey := Map()
            j := 1
            for it in st.items {
                key := it.path ":" it.start ":" it.end
                idxByKey[key] := j
                j++
            }
            for c in chunks {
                key := path ":" c.start ":" c.end
                if idxByKey.Has(key) {
                    k := idxByKey[key]
                    st.items[k].vector := c.vector
                } else {
                    newId := A_TickCount "-" st.items.Length + 1
                    st.items.Push({
                        id: newId,
                        path: path,
                        start: c.start,
                        end: c.end,
                        text: c.text,
                        left: c.left,
                        right: c.right,
                        vector: c.vector,
                        hash: Djb2Hex(c.text)
                    })
                    st.docs.%path%.chunkIds.Push(newId)
                }
            }
            ok := VS_Save(st)
        } else {
            VS_AddChunks(st, path, chunks, LocalHashVector, { hash: docHash })
            ok := VS_Save(st)
        }
        if IsSet(doneCb)
            try doneCb(ok)
    } catch as e {
        LogDebug("__PE_AssignAndSaveBatch error: " e.Message)
        if IsSet(doneCb)
            try doneCb(false)
    }
}

; ---------------- Revectorize Existing Store With Provider Embeddings ----------------
global gREV_State := {}

RevectorizeStoreProviderEmbeddingsAsync(cfg := unset, opts := unset, doneCb := unset) {
    try {
        global gREV_State
        if !IsSet(cfg)
            cfg := LoadConfig()
        st := VS_Load()
        docsPlan := []
        ; Build per-document plan using existing chunkIds order
        if st.HasOwnProp("docs") {
            for path, d in st.docs.OwnProps() {
                if !d.HasOwnProp("chunkIds")
                    continue
                texts := []
                idxs := [] ; indexes into st.items
                ; create fast index from id -> item index
                idToIdx := Map()
                i := 1
                for it in st.items {
                    idToIdx[it.id] := i
                    i++
                }
                for id in d.chunkIds {
                    if idToIdx.Has(id) {
                        idx := idToIdx[id]
                        idxs.Push(idx)
                        texts.Push(st.items[idx].text)
                    }
                }
                if (texts.Length > 0)
                    docsPlan.Push({ path: path, idxs: idxs, texts: texts })
            }
        }
        gREV_State := {
            cfg: cfg,
            store: st,
            docs: docsPlan,
            i: 1,
            done: doneCb
        }
        LogDebug("Revectorize: start docs=" gREV_State.docs.Length)
        __REV_DoNext()
    } catch as e {
        LogDebug("Revectorize init error: " e.Message)
        if IsSet(doneCb)
            try doneCb(false)
    }
}

__REV_DoNext() {
    global gREV_State
    st := gREV_State
    if !IsSet(st) || !st.HasOwnProp("docs") {
        return
    }
    if (st.i > st.docs.Length) {
        ok := false
        try ok := VS_Save(st.store)
        LogDebug("Revectorize: done ok=" (ok?"1":"0"))
        if st.HasOwnProp("done")
            try st.done(ok)
        gREV_State := {}
        return
    }
    prov := StrLower(st.cfg.Get("provider", "openai"))
    doc := st.docs[st.i]
    texts := doc.texts
    LogDebug("Revectorize: doc=" doc.path " chunks=" texts.Length " prov=" prov)
    try {
        if (prov = "openai" || prov = "openai-compatible") {
            EmbeddingsBatchOpenAIAsync(st.cfg, texts, (txt, status, hdrs) => (
                __REV_AssignAndAdvance("openai", txt, status)
            ))
            return
        }
        if (prov = "gemini") {
            EmbeddingsBatchGeminiAsync(st.cfg, texts, (txt, status, hdrs) => (
                __REV_AssignAndAdvance("gemini", txt, status)
            ))
            return
        }
    } catch as e {
        LogDebug("Revectorize dispatch error: " e.Message)
    }
    ; Fallback: local vectors if provider not supported
    __REV_AssignLocalAndAdvance()
}

__REV_AssignAndAdvance(prov, responseText, status) {
    global gREV_State
    st := gREV_State
    doc := st.docs[st.i]
    try {
        vecs := (status = 200) ? ParseEmbeddingVectors(prov, responseText) : []
        j := 1
        for idx in doc.idxs {
            txt := doc.texts[j]
            v := (vecs.Length >= j && vecs[j].Length > 0) ? vecs[j] : LocalHashVector(txt)
            st.store.items[idx].vector := v
            j++
        }
    } catch as e {
        LogDebug("Revectorize assign error: " e.Message)
        __REV_AssignLocal() ; ensure vectors set
    }
    gREV_State.i := st.i + 1
    __REV_DoNext()
}

__REV_AssignLocal() {
    global gREV_State
    st := gREV_State
    doc := st.docs[st.i]
    j := 1
    for idx in doc.idxs {
        txt := doc.texts[j]
        st.store.items[idx].vector := LocalHashVector(txt)
        j++
    }
}

__REV_AssignLocalAndAdvance() {
    __REV_AssignLocal()
    global gREV_State
    st := gREV_State
    gREV_State.i := st.i + 1
    __REV_DoNext()
}

; ---------------- Directory Provider-Embeddings Ingest (non-blocking) ----------------
global gPEDir_State := {}

IngestDirectoryProviderEmbeddingsAsync(dir, cfg := unset, opts := unset, doneCb := unset) {
    try {
        files := []
        for ext in ["*.txt", "*.htm", "*.html", "*.rtf"] {
            Loop Files dir "\\" ext
                files.Push(A_LoopFileFullPath)
        }
        global gPEDir_State
        gPEDir_State := {
            files: files,
            i: 1,
            cfg: IsSet(cfg) ? cfg : LoadConfig(),
            opts: IsSet(opts) ? opts : unset,
            done: doneCb
        }
        LogDebug("IngestPEDir start files=" files.Length)
        __PEDir_Next()
    } catch as e {
        LogDebug("IngestDirectoryProviderEmbeddingsAsync error: " e.Message)
        if IsSet(doneCb)
            try doneCb(false)
    }
}

__PEDir_Next() {
    global gPEDir_State
    st := gPEDir_State
    if !IsSet(st) || !st.HasOwnProp("files")
        return
    if (st.i > st.files.Length) {
        LogDebug("IngestPEDir done")
        if st.HasOwnProp("done")
            try st.done(true)
        gPEDir_State := {}
        return
    }
    p := st.files[st.i]
    LogDebug("IngestPEDir file=" p)
    IngestFileProviderEmbeddingsAsync(p, st.cfg, st.opts, (ok) => (
        gPEDir_State.i := st.i + 1,
        __PEDir_Next()
    ))
}

BuildGeminiGeneratePayload(messages, cfg) {
  parts := []
  for m in messages {
    parts.Push({ role: m.role, parts: [{ text: m.content }] })
  }
  body := { contents: parts }
  sys := StrictRAGInstruction()
  if (sys != "")
    body.systemInstruction := { role: "system", parts: [{ text: sys }] }
  gen := Map()
  for key in ["temperature","topP","topK","candidateCount","maxOutputTokens"] {
    v := cfg.Get(key, "")
    if (v != "")
      gen[key] := (key = "candidateCount" || key = "maxOutputTokens" || key = "topK") ? Integer(v) : (key = "temperature" || key = "topP") ? v+0 : v
  }
  stops := cfg.Get("stopSequences", "")
  if (stops != "") {
    seqs := []
    for s in StrSplit(stops, ",")
      if (Trim(s) != "")
        seqs.Push(Trim(s))
    if (seqs.Length > 0)
      gen["stopSequences"] := seqs
  }
  if (gen.Count > 0)
    body.generationConfig := gen
  ; safety settings
  saf := []
  AddSafe := (catKey, catName) => (
    th := cfg.Get(catKey, ""), (th != "") ? saf.Push({ category: catName, threshold: th }) : 0
  )
  AddSafe("safetyHarassment", "HARM_CATEGORY_HARASSMENT")
  AddSafe("safetyHateSpeech", "HARM_CATEGORY_HATE_SPEECH")
  AddSafe("safetySexual", "HARM_CATEGORY_SEXUALLY_EXPLICIT")
  AddSafe("safetyDangerous", "HARM_CATEGORY_DANGEROUS_CONTENT")
  AddSafe("safetyCivic", "HARM_CATEGORY_CIVIC_INTEGRITY")
  if (saf.Length > 0)
    body.safetySettings := saf
	
  return body
}

