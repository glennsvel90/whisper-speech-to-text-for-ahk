; ======================================================================================================================
; | WhisperAI Audio Transcriber AHK Script                                                                             |
; |                                                                                                                    |
; | This script listens for a hotkey combination to start and stop audio recording.                                    |
; | When recording is stopped, it sends the audio file to the OpenAI Whisper API for transcription.                    |
; | The resulting text is then pasted at the current cursor location and copied to the clipboard.                      |
; ======================================================================================================================

#SingleInstance, Force
#NoEnv
SetWorkingDir, %A_ScriptDir%

; ======================================================================================================================
; | !! IMPORTANT !! USER CONFIGURATION                                                                                 |
; | Paste your OpenAI API Key in the variable below.                                                                   |
; | You can get your API key from: https://platform.openai.com/account/api-keys                                        |
; ======================================================================================================================
Global YOUR_OPENAI_API_KEY := "PASTE_YOUR_OPENAI_API_KEY_HERE_BUT_DO_NOT_REMOVE_THE_QUOTATION_MARKS"

; ======================================================================================================================
; | Script Variables                                                                                                   |
; ======================================================================================================================
Global isRecording := false
Global audioFilePath := A_Temp . "\whisper_recording.wav" ; Using .wav for native recording
Global recordingIndicator := ""
Global q_combo_activated := 0 ; Flag for the custom hotkey logic

; ======================================================================================================================
; | Hotkey Definition                                                                                                  |
; |                                                                                                                    |
; | This new logic allows 'q' to be typed normally upon release, unless it's used in the 'q & z' combo.                |
; | It now correctly handles the Shift key for typing a capital 'Q'.                                                   |
; ======================================================================================================================

; This fires when Z is pressed while Q is held down. It blocks Q's normal function.
q & z::
    ; We set a flag to tell the 'q up' event that the combo fired.
    q_combo_activated := 1
    ToggleRecording()
return

; This hotkey is defined for Q's key-up event.
; The wildcard '*' makes it fire regardless of modifiers like Shift.
*q up::
    ; If the combo was activated, we reset the flag and do nothing else.
    if (q_combo_activated = 1) {
        q_combo_activated := 0
        return
    }
    ; If we reach here, the combo was NOT activated.
    ; Check if Shift was held down to determine case.
    if GetKeyState("Shift") {
        Send, {Q}
    } else {
        Send, {q}
    }
return

; ======================================================================================================================
; | Main Recording Logic                                                                                               |
; ======================================================================================================================
ToggleRecording() {
    Global isRecording, audioFilePath

    If (isRecording) {
        ; --- STOP RECORDING ---
        isRecording := false
        StopAudioRecording()
        HideRecordingIndicator()
        
        Sleep, 500 

        If !FileExist(audioFilePath) {
            TrayTip, Transcription Error, Audio file was not created., 10, 16
            Return
        }

        FileGetSize, size, %audioFilePath%, B
        if (size < 1024) {
            TrayTip, Transcription Error, Recording is too short or mic is not working., 10, 16
            FileDelete, %audioFilePath%
            Return
        }

        TrayTip, Transcribing...,, 10
        
        ; --- SEND TO WHISPER API ---
        transcribedText := TranscribeAudio(audioFilePath, YOUR_OPENAI_API_KEY)
        
        TrayTip

        If (transcribedText != "") {
            ; --- PASTE AND COPY TEXT ---
            SendInput, %transcribedText%
            Clipboard := transcribedText
            TrayTip, Transcription Complete, Text has been pasted and copied., 10, 1
        } 

        If FileExist(audioFilePath) {
            FileDelete, %audioFilePath%
        }

    } Else {
        ; --- START RECORDING ---
        isRecording := true
        
        If FileExist(audioFilePath) {
            FileDelete, %audioFilePath%
        }
        
        StartAudioRecording(audioFilePath)
        ShowRecordingIndicator()
        TrayTip, Recording Started, Press Q & Z again to stop., 10, 1
    }
}

; ======================================================================================================================
; | System Media Control Interface (MCI) Function Definition                                                           |
; ======================================================================================================================
mciSendString(sCommand, ByRef sReturnString="", nReturnLength=0, hCallback=0) {
    return DllCall("winmm\mciSendString" . (A_IsUnicode ? "W" : "A")
        , "str", sCommand
        , "ptr", &sReturnString
        , "uint", nReturnLength
        , "ptr", hCallback)
}

; ======================================================================================================================
; | Audio Recording Functions                                                                                          |
; ======================================================================================================================
StartAudioRecording(filePath) {
    mciSendString("open new type waveaudio alias rec",, 0, 0)
    mciSendString("record rec",, 0, 0)
}

StopAudioRecording() {
    Global audioFilePath
    mciSendString("stop rec",, 0, 0)
    mciSendString("save rec " . audioFilePath, "", 0, 0)
    mciSendString("close rec",, 0, 0)
}

; ======================================================================================================================
; | Whisper API Transcription Function (using ADODB.Stream)                                                            |
; ======================================================================================================================
TranscribeAudio(filePath, apiKey) {
    if (apiKey = "PASTE_YOUR_OPENAI_API_KEY_HERE" or apiKey = "") {
        MsgBox, 16, API Key Error, Please edit the script and replace "PASTE_YOUR_OPENAI_API_KEY_HERE".
        Return ""
    }
    
    CRLF := "`r`n"
    boundary := "ahk-boundary-" . A_TickCount

    ; --- Create the multipart body using a robust ADODB.Stream object ---
    try {
        pDataStream := ComObjCreate("ADODB.Stream")
        pDataStream.Type := 1 ; 1=binary
        pDataStream.Open()

        pTextStream := ComObjCreate("ADODB.Stream")
        pTextStream.Type := 2 ; 2=text
        pTextStream.Charset := "us-ascii"
        pTextStream.Open()

        pTextStream.WriteText("--" . boundary . CRLF . "Content-Disposition: form-data; name=""model""" . CRLF . CRLF . "whisper-1" . CRLF)
        pTextStream.Position := 0
        pTextStream.CopyTo(pDataStream)

        pTextStream.Position := 0
        pTextStream.WriteText("--" . boundary . CRLF . "Content-Disposition: form-data; name=""file""; filename=""audio.wav""" . CRLF . "Content-Type: audio/wav" . CRLF . CRLF)
        pTextStream.Position := 0
        pTextStream.CopyTo(pDataStream)
        
        pTextStream.Close()

        pFilestream := ComObjCreate("ADODB.Stream")
        pFilestream.Type := 1 ; binary
        pFilestream.Open()
        pFilestream.LoadFromFile(filePath)
        pFilestream.CopyTo(pDataStream)
        pFilestream.Close()
        
        pFinalBoundaryStream := ComObjCreate("ADODB.Stream")
        pFinalBoundaryStream.Type := 2
        pFinalBoundaryStream.Charset := "us-ascii"
        pFinalBoundaryStream.Open()
        pFinalBoundaryStream.WriteText(CRLF . "--" . boundary . "--" . CRLF)
        pFinalBoundaryStream.Position := 0
        pFinalBoundaryStream.CopyTo(pDataStream)
        pFinalBoundaryStream.Close()
        
        pDataStream.Position := 0
        postData := pDataStream.Read()
        pDataStream.Close()
    } catch e {
        errorMessage := "Failed to construct the request data.`n" . e.Message
        TrayTip, Script Error, %errorMessage%, 20, 16
        Return ""
    }

    ; --- Send the HTTP request ---
    try {
        oHttp := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        oHttp.Open("POST", "https://api.openai.com/v1/audio/transcriptions", true)
        oHttp.Option(4) := 13056 ; Ignore all SSL certificate errors
        oHttp.SetRequestHeader("Authorization", "Bearer " . apiKey)
        oHttp.SetRequestHeader("Content-Type", "multipart/form-data; boundary=" . boundary)
        
        oHttp.Send(postData)
        oHttp.WaitForResponse()

        response := oHttp.ResponseText
        status := oHttp.Status
        
        if (status != 200) {
            statusText := oHttp.StatusText
            jsonMessage := GetJsonValue(response, "message")
            fullErrorMsg := "HTTP Status: " . status . " " . statusText . "`nResponse: " . jsonMessage
            TrayTip, Transcription Failed, %fullErrorMsg%, 20, 16
            Return ""
        }
        
        transcribedText := GetJsonValue(response, "text")
        if (transcribedText = "") {
            errorText := "API returned success, but no text was found in the response."
            TrayTip, Transcription Failed, %errorText%, 20, 16
            Return ""
        }
        
        Return transcribedText
        
    } catch e {
        errorMessage := e.Message
        fullErrorText := "An error occurred while contacting the API.`n" . errorMessage
        TrayTip, Script Error, %fullErrorText%, 20, 16
        Return ""
    }
}

; A simple helper to extract a value from a JSON string.
GetJsonValue(json, key) {
    try {
        keyWithQuotes := """" . key . """:"""
        startPos := InStr(json, keyWithQuotes)
        if (startPos = 0)
            Return ""
        
        startPos += StrLen(keyWithQuotes)
        endPos := InStr(json, """",, startPos)
        
        length := endPos - startPos
        value := SubStr(json, startPos, length)

        value := StrReplace(value, "\""", """")
        value := StrReplace(value, "\\", "\")
        Return value
    } catch {
        Return ""
    }
}

; ======================================================================================================================
; | GUI and Indicator Functions                                                                                        |
; ======================================================================================================================
ShowRecordingIndicator() {
    Global recordingIndicator
    Gui, recordingIndicator:New, +AlwaysOnTop -Caption +ToolWindow +E0x20, Recording
    Gui, recordingIndicator:Color, CC0000
    Gui, recordingIndicator:Font, s10 cWhite Bold, Verdana
    Gui, recordingIndicator:Add, Text, ,  REC  
    SysGet, MonitorWidth, 76
    SysGet, MonitorHeight, 77
    xPos := MonitorWidth - 100
    yPos := MonitorHeight - 80
    Gui, recordingIndicator:Show, NoActivate, x%xPos% y%yPos%
}

HideRecordingIndicator() {
    Global recordingIndicator
    Gui, recordingIndicator:Destroy
}

; Make the script persistent
#Persistent
