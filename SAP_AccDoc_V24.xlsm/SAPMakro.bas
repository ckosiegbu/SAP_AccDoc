Attribute VB_Name = "SAPMakro"
Const CP = 40 'column of post indicator
Const CD = 41 'column of first header value
Const CM = 50 'column of return message

Sub sap_AccDoc_check()
    SAP_AccDoc_execute pTest:=True
End Sub

Sub sap_AccDoc_post()
    SAP_AccDoc_execute pTest:=False
End Sub

Sub SAP_AccDoc_execute(pTest As Boolean)
    Dim aSAPAcctngDocument As New SAPAcctngDocument
    Dim aSAPDocItem As New SAPDocItem
    Dim aData As New Collection
    Dim aDateFormatString As New DateFormatString

    Dim i As Integer
    Dim aRetStr As String

    Dim aCURRTYP2 As String
    Dim aWAERS2 As String
    Dim aCURRTYP3 As String
    Dim aWAERS3 As String
    Dim aCURRTYP4 As String
    Dim aWAERS4 As String

    Dim adBLDAT As String
    Dim adBLART As String
    Dim adBUKRS As String
    Dim adBUDAT As String
    Dim adWAERS As String
    Dim adXBLNR As String
    Dim adBKTXT As String
    Dim adFIS_PERIOD As Integer
    Dim adACC_PRINCIPLE As String

    Dim aBLDAT As String
    Dim aBLART As String
    Dim aBUKRS As String
    Dim aCOMP_CODE As String
    Dim aBUDAT As String
    Dim aWAERS As String
    Dim aXBLNR As String
    Dim aBKTXT As String
    Dim aACC_PRINCIPLE As String
    Dim aFIS_PERIOD As Integer

    Dim aGEGKO As String
    Dim aKONTO As String
    Dim aBETRA As Double

    Dim aSGTXT As String
    Dim aMWSKZ As String

    Dim aMATNR As String
    Dim aWERKS As String
    Dim aLIFNR As String
    Dim aKOSTL As String
    Dim aAUFNR As String

    Worksheets("Parameter").Activate
    adBUDAT = Format(Cells(2, 2), aDateFormatString.getString)
    adBLDAT = Format(Cells(3, 2), aDateFormatString.getString)
    adXBLNR = Cells(4, 2)
    adBKTXT = Cells(5, 2)
    adBUKRS = Cells(6, 2)
    adWAERS = Cells(7, 2)
    adBLART = Cells(8, 2)
    adFIS_PERIOD = Cells(9, 2)
    adACC_PRINCIPLE = Cells(10, 2)

    aCURRTYP2 = Cells(11, 2)
    aWAERS2 = Cells(12, 2)
    aCURRTYP3 = Cells(13, 2)
    aWAERS3 = Cells(14, 2)
    aCURRTYP4 = Cells(15, 2)
    aWAERS4 = Cells(16, 2)
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Check Authority
    '  Dim aSAPZFI_CHECK_F_BKPF_BUK As New SAPZFI_CHECK_F_BKPF_BUK
    '  Dim aAuth As Integer
    '  aAuth = aSAPZFI_CHECK_F_BKPF_BUK.checkAuthority(adBUKRS)
    '  If aAuth = False Then
    '    MsgBox "User " & MySAPCon.SAPCon.User & " is not authorized in Company Code " & adBUKRS, vbCritical + vbOKOnly
    '    Exit Sub
    '  End If
    ' Read the Data
    Worksheets("Data").Activate
    i = 2
    Do
        aKONTO = Cells(i, 2).Value
        aMATNR = Cells(i, 3).Value
        aWERKS = Cells(i, 4).Value
        aLIFNR = Cells(i, 5).Value
        aKOSTL = Cells(i, 7).Value
        aAUFNR = Cells(i, 8).Value
        aSGTXT = Cells(i, 31).Value
        aMWSKZ = Cells(i, 32).Value
        aBETRA = Cells(i, 34).Value
        Set aSAPDocItem = New SAPDocItem
        aSAPDocItem.create Cells(i, 1), aKONTO, aBETRA, aMWSKZ, aSGTXT, aAUFNR, aMATNR, aWERKS, aKOSTL, aLIFNR, _
        Cells(i, 12), Cells(i, 13), Cells(i, 14), Cells(i, 15), _
        Cells(i, 16), Cells(i, 17), Cells(i, 19), Cells(i, 28), Cells(i, 29), _
        Cells(i, 33), _
        Cells(i, 35), aCURRTYP2, aWAERS2, _
        Cells(i, 36), aCURRTYP3, aWAERS3, _
        Cells(i, 37), aCURRTYP4, aWAERS4, _
        Cells(i, 9).Value, Cells(i, 21).Value, Cells(i, 6).Value, _
        Cells(i, 23).Value, Cells(i, 24).Value, Cells(i, 38).Value, _
        Cells(i, 10).Value, Cells(i, 11).Value, Cells(i, CD + 4).Value, _
        Cells(i, 22).Value, Cells(i, 20).Value, Cells(i, 25).Value, Cells(i, 26).Value, Cells(i, 27).Value, _
        Cells(i, 18).Value, Cells(i, 39).Value, Cells(i, 30).Value
        aData.Add aSAPDocItem
        If (Cells(i, CP) = "X" Or Cells(i, CP) = "x") Then
            If IsNull(InStr(1, Cells(i, CM), "BKPFF")) Or InStr(1, Cells(i, CM), "BKPFF") = 0 Then
                If Cells(i, CD) <> "" Then
                    aBUDAT = Cells(i, CD)
                Else
                    aBUDAT = adBUDAT
                End If
                If Cells(i, CD + 1) <> "" Then
                    aBLDAT = Cells(i, CD + 1)
                Else
                    aBLDAT = adBLDAT
                End If
                If Cells(i, CD + 2) <> "" Then
                    aXBLNR = Cells(i, CD + 2)
                Else
                    aXBLNR = adXBLNR
                End If
                If Cells(i, CD + 3) <> "" Then
                    aBKTXT = Cells(i, CD + 3)
                Else
                    aBKTXT = adBKTXT
                End If
                If Cells(i, CD + 4) <> "" Then
                    aBUKRS = Cells(i, CD + 4)
                Else
                    aBUKRS = adBUKRS
                End If
                If Cells(i, CD + 5) <> "" Then
                    aWAERS = Cells(i, CD + 5)
                Else
                    aWAERS = adWAERS
                End If
                If Cells(i, CD + 6) <> "" Then
                    aBLART = Cells(i, CD + 6)
                Else
                    aBLART = adBLART
                End If
                If Cells(i, CD + 7) <> "" Then
                    aFIS_PERIOD = Cells(i, CD + 7)
                Else
                    aFIS_PERIOD = adFIS_PERIOD
                End If
                If Cells(i, CD + 8) <> "" Then
                    aACC_PRINCIPLE = Cells(i, CD + 8)
                Else
                    aACC_PRINCIPLE = adACC_PRINCIPLE
                End If
                If IsNull(InStr(1, Cells(i, CM), "BKPFF")) Or InStr(1, Cells(i, CM), "BKPFF") = 0 Then
                    aRetStr = aSAPAcctngDocument.post(aBLDAT, aBLART, aBUKRS, aBUDAT, aWAERS, aXBLNR, aBKTXT, aFIS_PERIOD, aACC_PRINCIPLE, aData, pTest)
                    Cells(i, CM) = aRetStr
                    Cells(i, CM + 1) = ExtractDocNumberFromMessage(aRetStr)
                End If
            End If
            Cells(i, CM + 1) = ExtractDocNumberFromMessage(Cells(i, CM))
            Set aData = New Collection
        End If
        i = i + 1
        Loop While Not IsNull(Cells(i, 1)) And Cells(i, 1) <> ""
    End Sub

    Function ExtractDocNumberFromMessage(Message As String) As String
        Dim aPos As Integer
        Dim aTemp As String
        Dim aLen As Long

        aLen = Len(Message)
        aPos = InStr(1, Message, "BKPFF")
        If Not IsNull(aPos) And aPos <> 0 Then
            ExtractDocNumberFromMessage = Mid(Message, aPos + 6, 18)
        Else
            ExtractDocNumberFromMessage = ""
        End If
    End Function
