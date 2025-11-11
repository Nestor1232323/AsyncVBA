VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AsyncVBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents http As WinHttp.WinHttpRequest
Attribute http.VB_VarHelpID = -1

Public Sub SendRequest(url As String, method As String, Optional headers As Variant, Optional body As String)
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    http.Open method, url, True
    
    If Not IsMissing(headers) Then
        Dim h As Variant
        For Each h In headers
            Dim parts() As String
            parts = Split(h, ":", 2)
            If UBound(parts) = 1 Then
                http.setRequestHeader Trim(parts(0)), Trim(parts(1))
            End If
        Next h
    End If

    If Not IsMissing(body) Then
        http.Send body
    Else
        http.Send
    End If
End Sub

Private Sub http_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
    Debug.Print "status:", Status
    Debug.Print "type:", ContentType
End Sub

Private Sub http_OnResponseFinished()
    Debug.Print "Response:"
    Debug.Print StrConv(http.responseBody, vbUnicode)
End Sub

