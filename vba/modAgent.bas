' vba/modAgent.bas
Attribute VB_Name = "Agent"
Option Explicit

Private Const AGENT_PORT As Long = 8731
Private Const RECEIVE_TIMEOUT_MS As Long = 3600000

Public Sub Probe(id As Long, values As Variant)
    Dim json As String
    json = "{""probe_id"":" & id & ",""vars"":["

    Dim i As Long
    Dim isFirst As Boolean
    isFirst = True
    i = LBound(values)
    Do While i <= UBound(values)
        If Not isFirst Then json = json & ","
        isFirst = False
        json = json & "{""n"":" & JsonValue(values(i)) & _
            ",""t"":" & JsonValue(TypeName(values(i + 1))) & _
            ",""v"":" & JsonValue(values(i + 1)) & "}"
        i = i + 2
    Loop

    json = json & "]}"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", "http://localhost:" & AGENT_PORT & "/probe/", False
    http.SetTimeouts 0, 0, 0, RECEIVE_TIMEOUT_MS
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send json

    Dim responseText As String
    responseText = http.ResponseText

    If InStr(responseText, """abort""") > 0 Then
        Err.Raise vbObjectError + 3000, "Agent", "Debug session aborted by user."
    End If
End Sub

Private Function JsonValue(v As Variant) As String
    Dim s As String
    If IsArray(v) Then
        s = "<array:" & TypeName(v) & ">"
    ElseIf IsObject(v) Then
        s = "<obj:" & TypeName(v) & ">"
    ElseIf IsNull(v) Then
        s = ""
    Else
        s = CStr(v)
    End If
    JsonValue = """" & JsonEscape(s) & """"
End Function

Private Function JsonEscape(s As String) As String
    Dim result As String
    result = Replace(s, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    JsonEscape = result
End Function
