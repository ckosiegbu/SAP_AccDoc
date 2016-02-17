VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPDocItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public ACCTYPE As String
Public NEWKO As String
Public Betrag As Double
Public MWSKZ As String
Public SGTXT As String
Public AUFNR As String
Public MATNR As String
Public WERKS As String
Public KOSTL As String
Public LIFNR As String
Public PA As String
Public VKORG As String
Public VTWEG As String
Public SPART As String
Public KNDNR As String
Public KTGRM As String
Public PRCTR As String
Public PMNTTRMS As String
Public BLINE_DATE As String
Public ALLOC_NMBR As String
Public BETR2 As Double
Public CURRTYP2 As String
Public WAERS2 As String
Public BETR3 As Double
Public CURRTYP3 As String
Public WAERS3 As String
Public BETR4 As Double
Public CURRTYP4 As String
Public WAERS4 As String
Public WBS As String
Public SEGMENT As String
Public FUNC_AREA As String
Public TRADE_ID As String
Public BUS_AREA As String
Public BEWAR As String
Public NETWORK As String
Public ACTIVITY As String
Public COMP_CODE As String
Public PARTNER_SEGMENT As String
Public PART_PRCTR As String
Public ZZETXT As String
Public ZZHFMC1 As String
Public ZZHFMC3 As String
Public MTART As String
Public REF_KEY_3 As String
Public PMNT_BLOCK As String

Public Function create(pACCTYPE As String, pNEWKO As String, pBetrag As Double, pMWSKZ As String, pSGTXT As String, _
    pAUFNR As String, pMATNR As String, pWERKS As String, pKOSTL As String, _
    pLIFNR As String, _
    pPA As String, pVKORG As String, pVTWEG As String, pSPART As String, _
    pKNDNR As String, pKTGRM As String, pPRCTR As String, _
    pPMNTTRMS As String, pBLINE_DATE As String, pALLOC_NMBR As String, _
    pBETR2 As Double, pCURRTYP2 As String, pWAERS2 As String, _
    pBETR3 As Double, pCURRTYP3 As String, pWAERS3 As String, _
    pBETR4 As Double, pCURRTYP4 As String, pWAERS4 As String, _
    pWBS As String, pSEGMENT As String, pFUNC_AREA As String, _
    pTRADE_ID As String, pBUS_AREA As String, pBEWAR As String, _
    pNETWORK As String, pACTIVITY As String, pCOMP_CODE As String, pPARTNER_SEGMENT As String, _
    pPART_PRCTR As String, pZZETXT As String, pZZHFMC1 As String, pZZHFMC3 As String, pMTART As String, _
    pREF_KEY_3 As String, pPMNT_BLOCK As String)
    ACCTYPE = pACCTYPE
    NEWKO = pNEWKO
    Betrag = pBetrag
    MWSKZ = pMWSKZ
    SGTXT = pSGTXT
    AUFNR = pAUFNR
    MATNR = pMATNR
    WERKS = pWERKS
    KOSTL = pKOSTL
    LIFNR = pLIFNR
    PA = pPA
    VKORG = pVKORG
    VTWEG = pVTWEG
    SPART = pSPART
    KNDNR = pKNDNR
    KTGRM = pKTGRM
    PRCTR = pPRCTR
    MWSKZ = pMWSKZ
    PMNTTRMS = pPMNTTRMS
    BLINE_DATE = pBLINE_DATE
    ALLOC_NMBR = pALLOC_NMBR
    BETR2 = pBETR2
    CURRTYP2 = pCURRTYP2
    WAERS2 = pWAERS2
    BETR3 = pBETR3
    CURRTYP3 = pCURRTYP3
    WAERS3 = pWAERS3
    BETR4 = pBETR4
    CURRTYP4 = pCURRTYP4
    WAERS4 = pWAERS4
    WBS = pWBS
    SEGMENT = pSEGMENT
    FUNC_AREA = pFUNC_AREA
    TRADE_ID = pTRADE_ID
    BUS_AREA = pBUS_AREA
    BEWAR = pBEWAR
    NETWORK = pNETWORK
    ACTIVITY = pACTIVITY
    COMP_CODE = pCOMP_CODE
    PARTNER_SEGMENT = pPARTNER_SEGMENT
    PART_PRCTR = pPART_PRCTR
    ZZETXT = pZZETXT
    ZZHFMC1 = pZZHFMC1
    ZZHFMC3 = pZZHFMC3
    MTART = pMTART
    REF_KEY_3 = pREF_KEY_3
    PMNT_BLOCK = pPMNT_BLOCK
End Function

