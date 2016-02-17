VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPZFI_CHECK_F_BKPF_BUK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsConnected As Integer

Public Function checkAuthority(pBUKRS As String) As Integer
    Dim Ret As Integer
    Dim aSUBRC
    Dim oFunction As Object
    On Error GoTo checkAuthority_Error
    Set oFunction = MySAPCon.SAPFunctions.Add("ZFI_CHECK_F_BKPF_BUK")
    oFunction.Exports("I_BUKRS") = pBUKRS
    Ret = oFunction.call
    checkAuthority = Ret
    Exit Function
checkAuthority_Error:
    checkAuthority = 8
    MySAPErr.MSGProt "SAPZFI_CHECK_F_BKPF_BUK", "checkAuthority", "", Err.Number, Err.Description
    Exit Function
End Function

Private Sub Class_Initialize()
    IsConnected = SAPCheck()
End Sub

