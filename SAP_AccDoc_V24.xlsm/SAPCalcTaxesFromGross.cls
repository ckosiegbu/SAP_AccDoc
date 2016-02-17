VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPCalcTaxesFromGross"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private goTAX_ITEM_IN As Object
Public goTAX_ITEM_OUT As Object
Public goERROR_RETURN As Object
Private goFunction As Object
Public IsConnected As Integer

Public Function InitFunction() As Integer
    On Error GoTo InitFunction_Error
    Set goFunction = MySAPCon.SAPFunctions.Add("CALCULATE_TAXES_GROSS")
    Set goTAX_ITEM_IN = goFunction.Tables("TAX_ITEM_IN")
    Set goTAX_ITEM_OUT = goFunction.Tables("TAX_ITEM_OUT")
    RetText = ""
    InitFunction = True
    Exit Function
InitFunction_Error:
    InitFunction = False
    MySAPErr.MSGProt "SAPCalcTaxesFromGross", "InitFunction", "", Err.Number, Err.Description
    Exit Function
End Function

Public Function getTaxAmount(pBUKRS As String, pMWSKZ As String, pWAERS As String, pBUDAT As String, pWRBTR As Double) As Boolean
    On Error GoTo getTaxAmount_Error
    Dim aTAX_ITEM_IN_Row As Object
    goTAX_ITEM_IN.Rows.RemoveAll
    goTAX_ITEM_OUT.Rows.RemoveAll
    Dim aRet As Integer
    Set aTAX_ITEM_IN_Row = goTAX_ITEM_IN.Rows.Add
    aTAX_ITEM_IN_Row.Value("BUKRS") = pBUKRS
    aTAX_ITEM_IN_Row.Value("MWSKZ") = pMWSKZ
    aTAX_ITEM_IN_Row.Value("WAERS") = pWAERS
    aTAX_ITEM_IN_Row.Value("BUDAT") = pBUDAT
    aTAX_ITEM_IN_Row.Value("WRBTR") = Format$(pWRBTR, "0.00")
    getTaxAmount = goFunction.call()
    Exit Function
getTaxAmount_Error:
    getTaxAmount = False
    MySAPErr.MSGProt "SAPCalcTaxesFromGross", "getTaxAmount", "", Err.Number, Err.Description
    Exit Function
End Function

Public Function RemoveFunction() As Integer
    On Error GoTo RemoveFunction_Error
    MySAPCon.SAPFunctions.Remove ("CALCULATE_TAXES_GROSS")
    RetText = ""
    RemoveFunction = True
    Exit Function
RemoveFunction_Error:
    RemoveFunction = False
    MySAPErr.MSGProt "SAPCalcTaxesFromGross", "RemoveFunction", "", Err.Number, Err.Description
    Exit Function
End Function

Private Sub Class_Initialize()
    IsConnected = SAPCheck()
    If IsConnected Then
        InitFunction
    End If
End Sub


