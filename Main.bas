Attribute VB_Name = "Main"
Public Token As String

Sub EnableTLS(ByRef Http)
    On Error Resume Next
    'Http.SetTimeouts ResolveTimeoutMs, ConnectTimeoutMs, SendTimeoutMs, ReceiveTimeoutMs
    Http.Option(9) = 2048
    Http.Option(6) = True
End Sub

Sub SetFont(frm As Form, Optional font As String = "±¼¸²", Optional fbFont As String = "Gulim")
    On Error Resume Next
    For Each ctrl In frm.Controls
        ctrl.FontName = "Tahoma"
        ctrl.FontName = "Segoe UI"
        ctrl.FontName = fbFont
        ctrl.FontName = font
        ctrl.FontSize = 9
    Next ctrl
End Sub

