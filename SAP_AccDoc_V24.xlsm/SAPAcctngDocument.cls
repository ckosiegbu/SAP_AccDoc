VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPAcctngDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsConnected As Integer

Private Sub Class_Initialize()
    IsConnected = SAPCheck()
End Sub

Public Function post(pBldat As String, pBLART As String, pBUKRS As String, _
    pBUDAT As String, pWAERS As String, pXBLNR As String, _
    pBKTXT As String, pFIS_PERIOD As Integer, pACC_PRINCIPLE As String, pData As Collection, pTest As Boolean) As String
    On Error GoTo post_Error

    Dim lSAPFormat As New SAPFormat
    Dim lSAPWbsElement As New SAPWbsElement
    Dim boAcctngDocument As Object
    Dim oDocumentHeader As Object
    Dim oAccountGl As Object
    Dim oAccountTax As Object
    Dim oAccountPayable As Object
    Dim oAccountReceivable As Object
    Dim oCurrencyAmount As Object
    Dim oCriteria As Object
    Dim oExtension2 As Object
    Dim oRETURN As Object

    Set boAcctngDocument = MySAPCon.SAPBapiControl.GetSAPObject("AcctngDocument")
    Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boAcctngDocument, "Post", "Return")
    Set oDocumentHeader = MySAPCon.SAPBapiControl.DimAs(boAcctngDocument, "Post", "DocumentHeader")
    Set oAccountGl = MySAPCon.SAPBapiControl.DimAs(boAcctngDocument, "Post", "AccountGl")
    Set oAccountTax = MySAPCon.SAPBapiControl.DimAs(boAcctngDocument, "Post", "AccountTax")
    Set oAccountPayable = MySAPCon.SAPBapiControl.DimAs(boAcctngDocument, "Post", "AccountPayable")
    Set oAccountReceivable = MySAPCon.SAPBapiControl.DimAs(boAcctngDocument, "Post", "AccountReceivable")
    Set oCurrencyAmount = MySAPCon.SAPBapiControl.DimAs(boAcctngDocument, "Post", "CurrencyAmount")
    Set oCriteria = MySAPCon.SAPBapiControl.DimAs(boAcctngDocument, "Post", "Criteria")
    Set oExtension2 = MySAPCon.SAPBapiControl.DimAs(boAcctngDocument, "Post", "Extension2")
    oDocumentHeader("BUS_ACT") = "RFBU"
    oDocumentHeader("ACC_PRINCIPLE") = pACC_PRINCIPLE
    oDocumentHeader("COMP_CODE") = pBUKRS
    oDocumentHeader("PSTNG_DATE") = pBUDAT
    oDocumentHeader("FIS_PERIOD") = pFIS_PERIOD '23.01.2012 Buchungsperiode
    oDocumentHeader("DOC_DATE") = pBldat
    oDocumentHeader("USERNAME") = MySAPCon.SAPCon.User
    oDocumentHeader("DOC_TYPE") = pBLART
    oDocumentHeader("REF_DOC_NO") = pXBLNR
    oDocumentHeader("HEADER_TXT") = pBKTXT
    Dim aGLRow As Object
    Dim aARRow As Object
    Dim aAPRow As Object
    Dim aTxRow As Object
    Dim aAmRow As Object
    Dim aCrRow As Object
    Dim aExRow As Object
    Dim lRow As Object
    Dim lCnt As Integer
    lCnt = 0
    For Each lRow In pData
        lCnt = lCnt + 1
        If lRow.ACCTYPE = "S" Or lRow.ACCTYPE = "G" Then
            Set aGLRow = oAccountGl.Rows.Add
            aGLRow.Value("ITEMNO_ACC") = lCnt
            aGLRow.Value("GL_ACCOUNT") = lSAPFormat.unpack(lRow.NEWKO, 10)
            aGLRow.Value("ITEM_TEXT") = lRow.SGTXT
            aGLRow.Value("TAX_CODE") = lRow.MWSKZ
            aGLRow.Value("ALLOC_NMBR") = lRow.ALLOC_NMBR
            aGLRow.Value("REF_KEY_3") = lRow.REF_KEY_3
            If lRow.COMP_CODE <> "" Then
                aGLRow.Value("COMP_CODE") = lRow.COMP_CODE
            Else
                aGLRow.Value("COMP_CODE") = pBUKRS
            End If
            If lRow.PRCTR <> "" Then
                aGLRow.Value("PROFIT_CTR") = lRow.PRCTR
            End If
            If lRow.PART_PRCTR <> "" Then
                aGLRow.Value("PART_PRCTR") = lRow.PART_PRCTR
            End If
            If lRow.SEGMENT <> "" Then
                aGLRow.Value("SEGMENT") = lSAPFormat.unpack(lRow.SEGMENT, 10)
            End If
            If lRow.PARTNER_SEGMENT <> "" Then
                aGLRow.Value("PARTNER_SEGMENT") = lSAPFormat.unpack(lRow.PARTNER_SEGMENT, 10)
            End If
            If lRow.BEWAR <> "" Then
                aGLRow.Value("CS_TRANS_T") = Right(lSAPFormat.unpack(lRow.BEWAR, 10), 3)    '28.11.2011 Bewegungsart hinzugefügt
            End If
            If lRow.FUNC_AREA <> "" Then
                aGLRow.Value("FUNC_AREA") = lSAPFormat.unpack(lRow.FUNC_AREA, 4)
            End If
            If lRow.TRADE_ID <> "" Then
                aGLRow.Value("TRADE_ID") = lSAPFormat.unpack(lRow.TRADE_ID, 6)
            End If
            If lRow.BUS_AREA <> "" Then
                aGLRow.Value("BUS_AREA") = lSAPFormat.unpack(lRow.BUS_AREA, 4)
            End If
            ' Extensions Fields
            If lRow.ZZETXT <> "" Then
                Set aExRow = oExtension2.Rows.Add
                aExRow.Value("STRUCTURE") = "ZFI_EXT2_ZZETXT"
                aExRow.Value("VALUEPART1") = lSAPFormat.unpack(CStr(lCnt), 10)
                aExRow.Value("VALUEPART2") = lRow.ZZETXT
            End If
            ' HFM Sales Country (in ISO 3 Char)
            If lRow.ZZHFMC1 <> "" Then
                Set aExRow = oExtension2.Rows.Add
                aExRow.Value("STRUCTURE") = "ZFI_EXT2_ZZHFMC1"
                aExRow.Value("VALUEPART1") = lSAPFormat.unpack(CStr(lCnt), 10)
                aExRow.Value("VALUEPART2") = lRow.ZZHFMC1
            End If
            ' HFM Customer Group
            If lRow.ZZHFMC3 <> "" Then
                Set aExRow = oExtension2.Rows.Add
                aExRow.Value("STRUCTURE") = "ZFI_EXT2_ZZHFMC3"
                aExRow.Value("VALUEPART1") = lSAPFormat.unpack(CStr(lCnt), 10)
                aExRow.Value("VALUEPART2") = lSAPFormat.unpack(lRow.ZZHFMC3, 3)
            End If
            ' CO-PA charactereistics
            If lRow.PA = "X" Or lRow.PA = "x" Then
                '  BUKRS
                Set aCrRow = oCriteria.Rows.Add
                aCrRow.Value("ITEMNO_ACC") = lCnt
                aCrRow.Value("FIELDNAME") = "BUKRS"
                If lRow.COMP_CODE <> "" Then
                    aCrRow.Value("CHARACTER") = lRow.COMP_CODE
                Else
                    aCrRow.Value("CHARACTER") = pBUKRS
                End If
                '  VKORG
                Set aCrRow = oCriteria.Rows.Add
                aCrRow.Value("ITEMNO_ACC") = lCnt
                aCrRow.Value("FIELDNAME") = "VKORG"
                aCrRow.Value("CHARACTER") = lRow.VKORG
                '  VTWEG
                Set aCrRow = oCriteria.Rows.Add
                aCrRow.Value("ITEMNO_ACC") = lCnt
                aCrRow.Value("FIELDNAME") = "VTWEG"
                aCrRow.Value("CHARACTER") = lSAPFormat.unpack(lRow.VTWEG, 2)
                '  SPART
                Set aCrRow = oCriteria.Rows.Add
                aCrRow.Value("ITEMNO_ACC") = lCnt
                aCrRow.Value("FIELDNAME") = "SPART"
                aCrRow.Value("CHARACTER") = lSAPFormat.unpack(lRow.SPART, 2)
                '  KNDNR
                Set aCrRow = oCriteria.Rows.Add
                aCrRow.Value("ITEMNO_ACC") = lCnt
                aCrRow.Value("FIELDNAME") = "KNDNR"
                aCrRow.Value("CHARACTER") = lSAPFormat.unpack(lRow.KNDNR, 10)
                '  WERKS
                Set aCrRow = oCriteria.Rows.Add
                aCrRow.Value("ITEMNO_ACC") = lCnt
                aCrRow.Value("FIELDNAME") = "WERKS"
                aCrRow.Value("CHARACTER") = lRow.WERKS
                '  ARTNR
                Set aCrRow = oCriteria.Rows.Add
                aCrRow.Value("ITEMNO_ACC") = lCnt
                aCrRow.Value("FIELDNAME") = "ARTNR"
                aCrRow.Value("CHARACTER") = lRow.MATNR
                '  KTGRM
                Set aCrRow = oCriteria.Rows.Add
                aCrRow.Value("ITEMNO_ACC") = lCnt
                aCrRow.Value("FIELDNAME") = "KTGRM"
                aCrRow.Value("CHARACTER") = lSAPFormat.unpack(lRow.KTGRM, 2)
                '  GSBER
                If lRow.BUS_AREA <> "" Then
                    Set aCrRow = oCriteria.Rows.Add
                    aCrRow.Value("ITEMNO_ACC") = lCnt
                    aCrRow.Value("FIELDNAME") = "GSBER"
                    aCrRow.Value("CHARACTER") = lSAPFormat.unpack(lRow.BUS_AREA, 4)
                End If
                '  GSBER
                If lRow.SEGMENT <> "" Then
                    Set aCrRow = oCriteria.Rows.Add
                    aCrRow.Value("ITEMNO_ACC") = lCnt
                    aCrRow.Value("FIELDNAME") = "SEGMENT"
                    aCrRow.Value("CHARACTER") = lSAPFormat.unpack(lRow.SEGMENT, 10)
                End If
                'PARTNER_SEGMENT
                If lRow.PARTNER_SEGMENT <> "" Then
                    Set aCrRow = oCriteria.Rows.Add
                    aCrRow.Value("ITEMNO_ACC") = lCnt
                    aCrRow.Value("FIELDNAME") = "WWPSE"
                    aCrRow.Value("CHARACTER") = lSAPFormat.unpack(lRow.PARTNER_SEGMENT, 10)
                End If
                ' PRCTR
                If lRow.PRCTR <> "" Then
                    Set aCrRow = oCriteria.Rows.Add
                    aCrRow.Value("ITEMNO_ACC") = lCnt
                    aCrRow.Value("FIELDNAME") = "PRCTR"
                    aCrRow.Value("CHARACTER") = lRow.PRCTR
                End If
                ' PART_PRCTR
                If lRow.PART_PRCTR <> "" Then
                    Set aCrRow = oCriteria.Rows.Add
                    aCrRow.Value("ITEMNO_ACC") = lCnt
                    aCrRow.Value("FIELDNAME") = "PPRCTR"
                    aCrRow.Value("CHARACTER") = lRow.PART_PRCTR
                End If
                ' FKBER
                If lRow.FUNC_AREA <> "" Then
                    Set aCrRow = oCriteria.Rows.Add
                    aCrRow.Value("ITEMNO_ACC") = lCnt
                    aCrRow.Value("FIELDNAME") = "FKBER"
                    aCrRow.Value("CHARACTER") = lSAPFormat.unpack(lRow.FUNC_AREA, 4)
                End If
                ' ZZHFMC1
                If lRow.ZZHFMC3 <> "" Then
                    Set aCrRow = oCriteria.Rows.Add
                    aCrRow.Value("ITEMNO_ACC") = lCnt
                    aCrRow.Value("FIELDNAME") = "WWHC3"
                    aCrRow.Value("CHARACTER") = lSAPFormat.unpack(lRow.ZZHFMC3, 3)
                End If
                If lRow.WBS <> "" Then
                    Set aCrRow = oCriteria.Rows.Add
                    aCrRow.Value("ITEMNO_ACC") = lCnt
                    aCrRow.Value("FIELDNAME") = "PSPNR"
                    aCrRow.Value("CHARACTER") = lSAPWbsElement.GetPspnr(lRow.WBS)
                End If
                If lRow.MTART <> "" Then
                    Set aCrRow = oCriteria.Rows.Add
                    aCrRow.Value("ITEMNO_ACC") = lCnt
                    aCrRow.Value("FIELDNAME") = "MTART"
                    aCrRow.Value("CHARACTER") = lRow.MTART
                End If
            Else
                aGLRow.Value("COSTCENTER") = lSAPFormat.unpack(lRow.KOSTL, 10)
                aGLRow.Value("MATERIAL") = lSAPFormat.unpack(lRow.MATNR, 18)
                aGLRow.Value("PLANT") = lRow.WERKS
                aGLRow.Value("VENDOR_NO") = lSAPFormat.unpack(lRow.LIFNR, 10)
                aGLRow.Value("ORDERID") = lSAPFormat.unpack(lRow.AUFNR, 12)
                aGLRow.Value("WBS_ELEMENT") = lRow.WBS
                aGLRow.Value("NETWORK") = lSAPFormat.unpack(lRow.NETWORK, 12)
                aGLRow.Value("ACTIVITY") = lSAPFormat.unpack(lRow.ACTIVITY, 4)
            End If
            Set aAmRow = oCurrencyAmount.Rows.Add
            aAmRow.Value("ITEMNO_ACC") = lCnt
            aAmRow.Value("CURRENCY") = pWAERS
            aAmRow.Value("AMT_DOCCUR") = Format$(lRow.Betrag, "0.00")
            If lRow.MWSKZ <> "" Then
                Dim aSAPCalcTaxesFromGross As New SAPCalcTaxesFromGross
                Dim lTaxSum As Double
                Dim lTaxBase As Double
                Dim lTaxLines As Integer
                Dim lTaxRow As Object
                aRet = aSAPCalcTaxesFromGross.getTaxAmount(pBUKRS, lRow.MWSKZ, pWAERS, pBUDAT, lRow.Betrag)
                lTaxLines = aSAPCalcTaxesFromGross.goTAX_ITEM_OUT.Rows.Count
                ' calculate the taxsum
                lTaxSum = 0
                For Each lTaxRow In aSAPCalcTaxesFromGross.goTAX_ITEM_OUT.Rows
                    lTaxSum = lTaxSum + lTaxRow("FWSTE")
                Next lTaxRow
                lTaxBase = lRow.Betrag - lTaxSum
                ' change the ammount of the row to the net value
                aAmRow.Value("AMT_DOCCUR") = Format$(lTaxBase, "0.00")
                ' add the tax positions
                If lTaxSum <> 0 Or lTaxLines > 1 Then
                    For Each lTaxRow In aSAPCalcTaxesFromGross.goTAX_ITEM_OUT.Rows
                        lCnt = lCnt + 1
                        Set aTxRow = oAccountTax.Rows.Add
                        aTxRow.Value("ITEMNO_ACC") = lCnt
                        aTxRow.Value("COND_KEY") = lTaxRow("KSCHL")
                        aTxRow.Value("TAX_CODE") = lTaxRow("MWSKZ")
                        Set aAmRow = oCurrencyAmount.Rows.Add
                        aAmRow.Value("ITEMNO_ACC") = lCnt
                        aAmRow.Value("CURRENCY") = pWAERS
                        aAmRow.Value("AMT_DOCCUR") = Format$(lTaxRow("FWSTE"), "0.00")
                        aAmRow.Value("AMT_BASE") = Format$(lTaxBase, "0.00")
                    Next lTaxRow
                End If
            End If
            If lRow.BETR2 <> 0 Then
                Set aAmRow = oCurrencyAmount.Rows.Add
                aAmRow.Value("ITEMNO_ACC") = lCnt
                aAmRow.Value("CURR_TYPE") = lRow.CURRTYP2
                aAmRow.Value("CURRENCY") = lRow.WAERS2
                aAmRow.Value("AMT_DOCCUR") = Format$(lRow.BETR2, "0.00")
            End If
            If lRow.BETR3 <> 0 Then
                Set aAmRow = oCurrencyAmount.Rows.Add
                aAmRow.Value("ITEMNO_ACC") = lCnt
                aAmRow.Value("CURR_TYPE") = lRow.CURRTYP3
                aAmRow.Value("CURRENCY") = lRow.WAERS3
                aAmRow.Value("AMT_DOCCUR") = Format$(lRow.BETR3, "0.00")
            End If
            If lRow.BETR4 <> 0 Then
                Set aAmRow = oCurrencyAmount.Rows.Add
                aAmRow.Value("ITEMNO_ACC") = lCnt
                aAmRow.Value("CURR_TYPE") = lRow.CURRTYP4
                aAmRow.Value("CURRENCY") = lRow.WAERS4
                aAmRow.Value("AMT_DOCCUR") = Format$(lRow.BETR4, "0.00")
            End If
        End If
        If lRow.ACCTYPE = "D" Or lRow.ACCTYPE = "C" Then
            Set aARRow = oAccountReceivable.Rows.Add
            aARRow.Value("ITEMNO_ACC") = lCnt
            aARRow.Value("CUSTOMER") = lSAPFormat.unpack(lRow.NEWKO, 10)
            aARRow.Value("ITEM_TEXT") = lRow.SGTXT
            aARRow.Value("TAX_CODE") = lRow.MWSKZ
            aARRow.Value("PMNTTRMS") = lRow.PMNTTRMS
            aARRow.Value("PMNT_BLOCK") = lRow.PMNT_BLOCK
            aARRow.Value("ALLOC_NMBR") = lRow.ALLOC_NMBR
            aARRow.Value("REF_KEY_3") = lRow.REF_KEY_3
            If lRow.BLINE_DATE <> "" Then
                aARRow.Value("BLINE_DATE") = lRow.BLINE_DATE
            End If
            If lRow.BUS_AREA <> "" Then
                aARRow.Value("BUS_AREA") = lSAPFormat.unpack(lRow.BUS_AREA, 4)
            End If
            If lRow.PRCTR <> "" Then
                aARRow.Value("PROFIT_CTR") = lRow.PRCTR
            End If
            If lRow.SP_GL_IND <> "" Then
                aARRow.Value("SP_GL_IND") = lRow.SP_GL_IND
            End If
            Set aAmRow = oCurrencyAmount.Rows.Add
            aAmRow.Value("ITEMNO_ACC") = lCnt
            aAmRow.Value("CURRENCY") = pWAERS
            aAmRow.Value("AMT_DOCCUR") = Format$(lRow.Betrag, "0.00")
            If lRow.BETR2 <> 0 Then
                Set aAmRow = oCurrencyAmount.Rows.Add
                aAmRow.Value("ITEMNO_ACC") = lCnt
                aAmRow.Value("CURR_TYPE") = lRow.CURRTYP2
                aAmRow.Value("CURRENCY") = lRow.WAERS2
                aAmRow.Value("AMT_DOCCUR") = Format$(lRow.BETR2, "0.00")
            End If
            If lRow.BETR3 <> 0 Then
                Set aAmRow = oCurrencyAmount.Rows.Add
                aAmRow.Value("ITEMNO_ACC") = lCnt
                aAmRow.Value("CURR_TYPE") = lRow.CURRTYP3
                aAmRow.Value("CURRENCY") = lRow.WAERS3
                aAmRow.Value("AMT_DOCCUR") = Format$(lRow.BETR3, "0.00")
            End If
            If lRow.BETR4 <> 0 Then
                Set aAmRow = oCurrencyAmount.Rows.Add
                aAmRow.Value("ITEMNO_ACC") = lCnt
                aAmRow.Value("CURR_TYPE") = lRow.CURRTYP4
                aAmRow.Value("CURRENCY") = lRow.WAERS4
                aAmRow.Value("AMT_DOCCUR") = Format$(lRow.BETR4, "0.00")
            End If
        End If
        If lRow.ACCTYPE = "K" Or lRow.ACCTYPE = "V" Then
            Set aAPRow = oAccountPayable.Rows.Add
            aAPRow.Value("ITEMNO_ACC") = lCnt
            aAPRow.Value("VENDOR_NO") = lSAPFormat.unpack(lRow.NEWKO, 10)
            aAPRow.Value("ITEM_TEXT") = lRow.SGTXT
            aAPRow.Value("TAX_CODE") = lRow.MWSKZ
            aAPRow.Value("PMNTTRMS") = lRow.PMNTTRMS
            aAPRow.Value("PMNT_BLOCK") = lRow.PMNT_BLOCK
            aAPRow.Value("ALLOC_NMBR") = lRow.ALLOC_NMBR
            aAPRow.Value("REF_KEY_3") = lRow.REF_KEY_3
            If lRow.BLINE_DATE <> "" Then
                aAPRow.Value("BLINE_DATE") = lRow.BLINE_DATE
            End If
            If lRow.BUS_AREA <> "" Then
                aAPRow.Value("BUS_AREA") = lSAPFormat.unpack(lRow.BUS_AREA, 4)
            End If
            If lRow.PRCTR <> "" Then
                aAPRow.Value("PROFIT_CTR") = lRow.PRCTR
            End If
            If lRow.SP_GL_IND <> "" Then
                aAPRow.Value("SP_GL_IND") = lRow.SP_GL_IND
            End If
            Set aAmRow = oCurrencyAmount.Rows.Add
            aAmRow.Value("ITEMNO_ACC") = lCnt
            aAmRow.Value("CURRENCY") = pWAERS
            aAmRow.Value("AMT_DOCCUR") = Format$(lRow.Betrag, "0.00")
            If lRow.BETR2 <> 0 Then
                Set aAmRow = oCurrencyAmount.Rows.Add
                aAmRow.Value("ITEMNO_ACC") = lCnt
                aAmRow.Value("CURR_TYPE") = lRow.CURRTYP2
                aAmRow.Value("CURRENCY") = lRow.WAERS2
                aAmRow.Value("AMT_DOCCUR") = Format$(lRow.BETR2, "0.00")
            End If
            If lRow.BETR3 <> 0 Then
                Set aAmRow = oCurrencyAmount.Rows.Add
                aAmRow.Value("ITEMNO_ACC") = lCnt
                aAmRow.Value("CURR_TYPE") = lRow.CURRTYP3
                aAmRow.Value("CURRENCY") = lRow.WAERS3
                aAmRow.Value("AMT_DOCCUR") = Format$(lRow.BETR3, "0.00")
            End If
            If lRow.BETR4 <> 0 Then
                Set aAmRow = oCurrencyAmount.Rows.Add
                aAmRow.Value("ITEMNO_ACC") = lCnt
                aAmRow.Value("CURR_TYPE") = lRow.CURRTYP4
                aAmRow.Value("CURRENCY") = lRow.WAERS4
                aAmRow.Value("AMT_DOCCUR") = Format$(lRow.BETR4, "0.00")
            End If
        End If
    Next lRow
    Application.Cursor = xlWait
    If pTest Then
        boAcctngDocument.check DocumentHeader:=oDocumentHeader, _
        AccountGl:=oAccountGl, _
        AccountTax:=oAccountTax, _
        AccountPayable:=oAccountPayable, _
        AccountReceivable:=oAccountReceivable, _
        CurrencyAmount:=oCurrencyAmount, _
        Criteria:=oCriteria, _
        Extension2:=oExtension2, _
        Return:=oRETURN
    Else
        boAcctngDocument.post DocumentHeader:=oDocumentHeader, _
        AccountGl:=oAccountGl, _
        AccountTax:=oAccountTax, _
        AccountPayable:=oAccountPayable, _
        AccountReceivable:=oAccountReceivable, _
        CurrencyAmount:=oCurrencyAmount, _
        Criteria:=oCriteria, _
        Extension2:=oExtension2, _
        Return:=oRETURN
    End If
    Application.Cursor = xlDefault
    Set aRow = oRETURN.Rows(1)
    If aRow("TYPE") = "S" Then
        Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
        If Not pTest Then
            aSAPBapiTranctionCommit.commit
        End If
        post = aRow("MESSAGE")
    Else
        Dim RetRow As Object
        For Each RetRow In oRETURN.Rows
            post = post & ";" & RetRow("MESSAGE")
        Next RetRow
    End If
    Set boAcctngDocument = Nothing
    Exit Function
post_Error:
    Application.Cursor = xlDefault
    post = "Error"
    MySAPErr.MSGProt "SAPAcctngDocument", "post", "", Err.Number, Err.Description
    Exit Function
End Function





