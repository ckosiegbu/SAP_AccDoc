Attribute VB_Name = "Tester"
Sub test()
    Dim aSAPZFI_CHECK_F_BKPF_BUK As New SAPZFI_CHECK_F_BKPF_BUK
    Dim aAuth As Integer
    ' test auth
    aAuth = aSAPZFI_CHECK_F_BKPF_BUK.checkAuthority("0200")
End Sub
