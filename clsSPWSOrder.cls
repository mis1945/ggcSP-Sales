VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPWSOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Rex S. Adversalo
' XerSys Computing
' Canaoalan, Binmaley, Pangasinan
'
' Copyright 2003-2004 and beyond
' All Rights Reserved
'
'     Spareparts WholeSale Order Object
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-0863      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  iMac [ 02/16/2017 09:42 am ]
'     Start creating this object.
'        A modified version of clsSPWholesale
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME = "clsSPWSOrder"
Private Const pxeSystemCode As String = "SP"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oSCA_Const As New clsCodeGenConst

Private p_sBranchCd As String
Private p_sFullName As String
Private p_sAddressx As String

Private p_xChecks As CheckInfo

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pbBranchxx As Boolean
Private pnCtr As Integer

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property

   ItemCount = p_oBaseClas.ItemCount
End Property

Property Get MasFldSize(ByVal Index As Integer) As Variant
   On Error Resume Next

   If pbInitTran = False Then Exit Property
   
   MasFldSize = p_oBaseClas.MasFldSize(Index)
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      Master = p_sFullName
   Case 4, "xaddressx"
      Master = p_sAddressx
   Case Else
      Master = p_oBaseClas.Master(Index)
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "sCompnyNm"
      getMaster 2, Value, False
   Case 3, "xfullname"
      If pbBranchxx Then
         getCustomer Value, False
      Else
         p_sFullName = Value
      End If
   Case 4, "xaddressx"
      p_sAddressx = Value
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
End Property

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next

   If pbInitTran = False Then Exit Property

   Detail = p_oBaseClas.Detail(Row, Index)
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   On Error Resume Next

   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "sbarrcode"
      getDetail Row, 1, Value, False
   Case 2, "sdescript"
      getDetail Row, 2, Value, False
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
      
      RaiseEvent DetailRetrieved(Index)
   End Select
End Property

Property Get Checks(ByVal Index As Variant) As Variant
   Select Case Index
   Case 0, "CheckNox"
      Checks = p_xChecks.CheckNox
   Case 1, "sAcctNoxx"
      Checks = p_xChecks.AcctNoxx
   Case 2, "sBankIDxx"
      Checks = p_xChecks.BankIDxx
   Case 3, "dCheckDte"
      Checks = p_xChecks.CheckDte
   Case 4, "nAmountxx"
      Checks = p_xChecks.Amountxx
   End Select
End Property

Property Let Checks(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 0, "CheckNox"
      p_xChecks.CheckNox = Value
   Case 1, "sAcctNoxx"
      p_xChecks.AcctNoxx = Value
   Case 2, "sBankIDxx"
      p_xChecks.BankIDxx = Value
   Case 3, "dCheckDte"
      p_xChecks.CheckDte = Value
   Case 4, "nAmountxx"
      p_xChecks.Amountxx = Value
   End Select
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "InitTransaction"
   On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail

   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .DisplayMessage = False
      .MasterTable = "SP_WS_Order_Master"
      .DetailTable = "SP_WS_Order_Detail"
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", b.sCompnyNm" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName) xFullname" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName) xAddressx" & _
                        ", a.sRemarksx" & _
                        ", a.nCredtAmt" & _
                        ", a.nCredtChk" & _
                        ", a.nDiscount" & _
                        ", a.nAddDiscx" & _
                        ", a.nAmtPaidx" & _
                        ", a.nTranTotl" & _
                        ", a.sPostedxx" & _
                        ", a.dPostedxx" & _
                        ", a.cTranStat" & _
                        ", a.cOrdrStat" & _
                        ", a.cPostedxx" & _
                        ", a.dStatChng" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", a.sClientID" & _
                        ", a.nEntryNox" & _
                        ", a.sApproved"
                        
      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                        ", Client_Master b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx"

      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", d.sTypeCode" & _
                        ", e.sModelNme" & _
                        ", a.nUnitPrce" & _
                        ", a.nDiscount" & _
                        ", a.nAddDiscx" & _
                        ", a.nQuantity" & _
                        ", a.sTransNox" & _
                        ", a.nIssuedxx" & _
                        ", a.nCanceled" & _
                        ", a.nAllocatn" & _
                        ", a.dModified" & _
                        ", a.sPartsIDx" & _
                        ", c.nQtyOnHnd" & _
                        ", c.nQtyOnHnd + a.nQuantity xQtyOnHnd" & _
                        ", b.nSMaxDisc" & _
                        ", b.nMMaxDisc" & _
                        ", b.nSelPrice" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo"
      .DetailQuery = .DetailQuery & _
                     " FROM " & .DetailTable & " a" & _
                        ", Spareparts b" & _
                           " LEFT JOIN SP_Model e" & _
                              " ON b.sModelIdx = e.sModelIdx" & _
                        ", SP_Inventory c" & _
                        ", SP_Type d" & _
                     " WHERE a.sPartsIDx = b.sPartsIDx" & _
                        " AND b.sPartsIDx = c.sPartsIDx" & _
                        " AND b.cPartType = d.cPartType" & _
                        " AND LEFT(a.sTransNox,4) = c.sBranchCd " & _
                     " ORDER BY a.nEntryNox"

      .Reference = "sTransNox|a.sTransNox"
      
      InitTransaction = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                           ", CASE a.cTranStat" & _
                              " WHEN '0' THEN 'OPEN'" & _
                              " WHEN '1' THEN 'CLOSED'" & _
                              " WHEN '2' THEN 'POSTED'" & _
                              " WHEN '3' THEN 'CANCELLED'" & _
                              " ELSE 'UNKNOWN'" & _
                              " END cTranStat" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID" & _
                        " ORDER BY b.sCompnyNm" & _
                           ", a.dTransact"
      .BrowseColumn(0) = "sTransNox»sCompnyNm»dTransact»cTranStat"
      .BrowseCriteria(0) = "a.sTransNox»b.sCompnyNm»a.dTransact»a.cTranStat"
      .BrowseTitle(0) = "Transaction No»Company Name»Date»Status"
      .BrowseFormat(0) = "@@-@@@@@@»@»MMMM DD, YYYY»@"

      .BrowseDetailQuery(0) = "SELECT" & _
                                 " DISTINCT a.sBarrCode" & _
                                 ", a.sDescript" & _
                                 ", c.sTypeCode" & _
                                 ", d.sModelNme" & _
                                 ", a.nSelPrice" & _
                                 ", a.sPartsIDx" & _
                                 ", b.nQtyOnHnd" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nReorderx" & _
                                 ", b.nLedgerNo" & _
                                 ", a.nSMaxDisc" & _
                                 ", a.nMMaxDisc" & _
                              " FROM Spareparts a" & _
                                    " LEFT JOIN SP_Model d" & _
                                       " ON a.sModelIDx = d.sModelIDx" & _
                                 ", SP_Inventory b" & _
                                 ", SP_Type c" & _
                              " WHERE a.sPartsIDx = b.sPartsIDx" & _
                                 " AND a.cPartType = c.cPartType" & _
                                 " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                              " ORDER BY a.sDescript"
   
      .BrowseDetailColumn(0) = "sBarrCode»sDescript»sTypeCode»sModelNme»nSelPrice"
      .BrowseDetailTitle(0) = "Bar Code»Description»Type»Model»Unit Price"
      .BrowseDetailFormat(0) = "@»@»@»@»#,##0.00"

      ' Query statement for the Customer field
      .BrowseQuery(11) = "SELECT" & _
                           "  a.sClientID" & _
                           ", a.sCompnyNm" & _
                           ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) xFullName" & _
                           ", CONCAT(a.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                        " FROM Client_Master a" & _
                           ", SP_Customer b" & _
                           ", TownCity c" & _
                           ", Province d" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           " AND a.sTownIDxx = c.sTownIdxx" & _
                           " AND c.sProvIDxx = d.sProvIDxx" & _
                           " AND a.cRecdStat = " & xeRecStateActive & _
                           " AND b.sBranchCd = " & strParm(p_sBranchCd)
                           
      .BrowseColumn(11) = "sClientID»sCompnyNm»xFullName»xAddressx"
      .BrowseTitle(11) = "Client ID»Company Name»Owner»Address"
      .BrowseCriteria(11) = "a.sClientID»a.sCompnyNm"
   End With

   pbInitTran = True
   InitTransaction = True
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

' »»» Rex - 2010-08-03
'  Search customer by name
Private Sub getCustomer(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "getCustomer"
   On Error GoTo errProc
   
   If lsValue <> "" Then
      If StrComp(lsValue, p_sFullName, vbTextCompare) = 0 Then GoTo endProc
      If lbSearch Then
         lsMaster = "CONCAT(a.sLastName, ', ', a.sFrstName) LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsMaster = "CONCAT(a.sLastName, ', ', a.sFrstName) = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If

   lsMaster = p_oBaseClas.getMaster(19, lsMaster)
   If lsMaster = Empty Then
      If lbSearch = False Then
         GoTo endWithClear
      Else
         GoTo endProc
      End If
   End If
   lasMaster = Split(lsMaster, "»")
   With p_oBaseClas
      .Master("sCompnyNm") = lasMaster(1)
      .Master("sClientID") = lasMaster(0)
      .Master("nDiscount") = lasMaster(5)
      
      p_sFullName = lasMaster(2)
      p_sAddressx = lasMaster(3)
'      getTerm lasMaster(4), False
   End With

endProc:
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master("sCompnyNm") = ""
      .Master("sClientID") = ""
      p_sFullName = ""
      p_sAddressx = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Sub


Private Sub getMaster(ByVal lnIndex As Integer, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "GetMaster"
   On Error GoTo errProc
   
   With p_oBaseClas
      If lsValue <> "" Then
         If StrComp(lsValue, .Master("sCompnyNm"), vbTextCompare) = 0 Then GoTo endProc
         
         If lbSearch Then
            lsMaster = "a.sCompnyNm LIKE " & strParm(Trim(lsValue) & "%")
         Else
            lsMaster = "a.sCompnyNm = " & strParm(Trim(lsValue))
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
   
      lsMaster = .getMaster(11, lsMaster)
      If lsMaster = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")
      
      .Master("sCompnyNm") = lasMaster(1)
      .Master("sClientID") = lasMaster(0)
      p_sFullName = lasMaster(2)
      p_sAddressx = lasMaster(3)
   End With

endProc:
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   
   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master("sCompnyNm") = ""
      .Master("sClientID") = ""
      p_sFullName = ""
      p_sAddressx = ""
   End With
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lnIndex _
                       & ", " & lsValue _
                       & ", " & lbSearch & " )"
End Sub

Private Sub getDetail(ByVal lnRow As Integer, _
                      ByVal lnIndex As Integer, _
                      ByVal lsValue As String, _
                      ByVal lbSearch As Boolean)
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsProcName As String

   lsProcName = "GetDetail"
   On Error GoTo errProc
   
   With p_oBaseClas
      If lsValue <> "" Then
         If lnIndex = 1 Then     ' Bar Code
            If lsValue = .Detail(lnRow, "sBarrCode") Then GoTo endProc
            If lbSearch Then
               lsDetail = "a.sBarrcode LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "a.sBarrcode = " & strParm(Trim(lsValue))
            End If
         Else
            If lsValue = .Detail(lnRow, "sDescript") Then GoTo endProc
            If lbSearch Then
               lsDetail = "a.sDescript LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "a.sDescript = " & strParm(Trim(lsValue))
            End If
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
   
      lsDetail = .getDetail(lnRow, lsDetail)
      If lsDetail = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasDetail = Split(lsDetail, "»")
      
'      'mac 2021.01.21
'      '  find the record if already exists on record
'      pnCtr = 0
'      Do While pnCtr <= .ItemCount - 1
'         If .Detail(pnCtr, "sPartsIDx") = lasDetail(5) Then
'            'get the row index
'            lnRow = pnCtr
'            Exit Do
'         End If
'
'         pnCtr = pnCtr + 1
'      Loop
'      'end - 'mac 2021.01.21

      .Detail(lnRow, "sBarrCode") = lasDetail(0)
      .Detail(lnRow, "sDescript") = lasDetail(1)
      .Detail(lnRow, "sTypeCode") = lasDetail(2)
      .Detail(lnRow, "sModelNme") = lasDetail(3)
      .Detail(lnRow, "nUnitPrce") = IIf(.Detail(lnRow, "nQuantity") = 0, lasDetail(4), .Detail(lnRow, "nUnitPrce"))
      .Detail(lnRow, "nDiscount") = 0#
      .Detail(lnRow, "nAddDiscx") = 0#
      .Detail(lnRow, "nQuantity") = 0 '.Detail(lnRow, "nQuantity") + 1
      .Detail(lnRow, "nIssuedxx") = 0
      .Detail(lnRow, "nCanceled") = 0
      .Detail(lnRow, "nAllocatn") = 0
      .Detail(lnRow, "sPartsIDx") = lasDetail(5)
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(6)
      .Detail(lnRow, "nResvOrdr") = lasDetail(7)
      .Detail(lnRow, "nBackOrdr") = lasDetail(8)
      .Detail(lnRow, "nReorderx") = lasDetail(9)
      .Detail(lnRow, "nLedgerNo") = lasDetail(10)
      .Detail(lnRow, "nSelPrice") = lasDetail(4)
      .Detail(lnRow, "nSMaxDisc") = lasDetail(11)
      .Detail(lnRow, "nMMaxDisc") = lasDetail(12)
   End With
   
endProc:
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)
   RaiseEvent DetailRetrieved(7)
   RaiseEvent DetailRetrieved(8)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "sTypeCode") = ""
      .Detail(lnRow, "sModelNme") = ""
      .Detail(lnRow, "nUnitPrce") = 0#
      .Detail(lnRow, "nDiscount") = 0#
      .Detail(lnRow, "nAddDiscx") = 0#
      .Detail(lnRow, "nIssuedxx") = 0
      .Detail(lnRow, "nCanceled") = 0
      .Detail(lnRow, "nAllocatn") = 0
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "sPartsIDx") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nResvOrdr") = 0
      .Detail(lnRow, "nBackOrdr") = 0
      .Detail(lnRow, "nReorderx") = 0
      .Detail(lnRow, "nLedgerNo") = 0
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lnIndex _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Sub

Function AddDetail() As Boolean
   If Trim(p_oBaseClas.Detail(ItemCount - 1, "sPartsIDx")) <> "" Then
      AddDetail = p_oBaseClas.AddDetail

      pnCtr = p_oBaseClas.ItemCount - 1
      With p_oBaseClas
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
         .Detail(pnCtr, "sBarrCode") = ""
         .Detail(pnCtr, "sDescript") = ""
         .Detail(pnCtr, "sTypeCode") = ""
         .Detail(pnCtr, "sModelNme") = ""
         .Detail(pnCtr, "nUnitPrce") = 0#
         .Detail(pnCtr, "nDiscount") = 0#
         .Detail(pnCtr, "nAddDiscx") = 0#
         .Detail(pnCtr, "nIssuedxx") = 0
         .Detail(pnCtr, "nCanceled") = 0
         .Detail(pnCtr, "nAllocatn") = 0
         .Detail(pnCtr, "nQuantity") = 0
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(pnCtr, "sPartsIDx") = ""
         .Detail(pnCtr, "nQtyOnHnd") = 0
         .Detail(pnCtr, "nResvOrdr") = 0
         .Detail(pnCtr, "nBackOrdr") = 0
         .Detail(pnCtr, "nReorderx") = 0
         .Detail(pnCtr, "nLedgerNo") = 0
      End With
   End If

   AddDetail = True
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   DeleteDetail = p_oBaseClas.DeleteDetail(Index)
End Function

Function SearchDetail(ByVal Row As Long, _
                        ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsProcName As String

   lsProcName = "SearchDetail"
   On Error GoTo errProc
   SearchDetail = False

   If pbInitTran = False Then GoTo endProc

   Index = LCase(Index)
   Select Case Index
   Case 1, "sbarrcode"
      getDetail Row, 1, Value, True
   Case 2, "sdescript"
      getDetail Row, 2, Value, True
   End Select
   SearchDetail = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & Row _
                       & ", " & Index _
                       & ", " & Value & " )"
End Function

Function SearchMaster(ByVal Index As Variant, Optional Value As Variant = "") As Boolean
   If pbInitTran = False Then Exit Function
   
   Index = LCase(Index)
   Select Case Index
   Case 2, "scompnynm"
      getMaster Index, Value, True
   End Select
   SearchMaster = True
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "CancelTransaction"
   On Error GoTo errProc
      
   p_oBaseClas.Master("dStatChng") = p_oAppDrivr.ServerDate
   
   If pbInitTran Then CancelTransaction = p_oBaseClas.CancelTransaction()

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "CloseTransaction"
   On Error GoTo errProc
   
   p_oBaseClas.Master("dStatChng") = p_oAppDrivr.ServerDate
   
   If pbInitTran Then CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "NewTransaction"
   On Error GoTo errProc
   
   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "SaveTransaction"
   On Error GoTo errProc
      
   If pbInitTran Then SaveTransaction = p_oBaseClas.SaveTransaction

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "OpenTransaction"
   On Error GoTo errProc
      
   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "DeleteTransaction"
   On Error GoTo errProc
      
   If pbInitTran Then DeleteTransaction = p_oBaseClas.DeleteTransaction

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lsRecord As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String

   lsProcName = "SearchTransaction"
   On Error GoTo errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sTransNox"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "b.sCompnyNm"
            lsDescript = "b.sCompnyNm"
         End If
      End If
   End If

   SearchTransaction = p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript)

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sSearch _
                       & ", " & bByCode & " )"
End Function

Function PostTransaction(ByVal sTransNox As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "PostTransaction"
   On Error GoTo errProc

   With p_oBaseClas
      .Master("sPostedxx") = p_oAppDrivr.UserID
      .Master("dPostedxx") = p_oAppDrivr.ServerDate
      .Master("cPostedxx") = xeYes
      .Master("dStatChng") = p_oAppDrivr.ServerDate
      
      If pbInitTran Then PostTransaction = .PostTransaction(sTransNox)
   End With

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNox & " )"
End Function

Function UpdateTransaction() As Boolean
   p_oBaseClas.EditMode = xeModeUpdate
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_InitMaster"
   On Error GoTo errProc
   
   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
                                 
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sCompnyNm") = ""
      .Master("nAmtPaidx") = 0#
      .Master("sRemarksx") = ""
      .Master("nCredtAmt") = 0#
      .Master("nCredtChk") = 0#
      .Master("nDiscount") = 0#
      .Master("nAddDiscx") = 0#
      .Master("nAmtPaidx") = 0#
      .Master("nTranTotl") = 0#
      .Master("sPostedxx") = ""
      .Master("dPostedxx") = Null
      .Master("cTranStat") = xeStateOpen
      .Master("cOrdrStat") = xeNo
      .Master("cPostedxx") = xeNo
      .Master("dStatChng") = p_oAppDrivr.ServerDate
      .Master("sClientID") = ""
      .Master("nEntryNox") = 0
      .Master("sRemarksx") = ""
      
      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "sTypeCode") = ""
      .Detail(0, "sModelNme") = ""
      .Detail(0, "nUnitPrce") = 0#
      .Detail(0, "nDiscount") = 0#
      .Detail(0, "nAddDiscx") = 0#
      .Detail(0, "nQuantity") = 0
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "nIssuedxx") = 0
      .Detail(0, "nCanceled") = 0
      .Detail(0, "nAllocatn") = 0
      .Detail(0, "sPartsIDx") = ""
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nResvOrdr") = 0
      .Detail(0, "nBackOrdr") = 0
      .Detail(0, "nReorderx") = 0
      .Detail(0, "nLedgerNo") = 0
   End With
   
   clearReceipt
   p_sFullName = ""
   p_sAddressx = ""

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub clearReceipt()
   p_xChecks.TransNox = ""
   p_xChecks.AcctNoxx = ""
   p_xChecks.BankIDxx = ""
   p_xChecks.CheckNox = ""
   p_xChecks.CheckDte = Date
   p_xChecks.Amountxx = 0
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      p_sFullName = .Master("xFullName")
      p_sAddressx = .Master("xAddressx")
   End With
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = Not pbModified
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_WillSave"
   On Error GoTo errProc
   Cancel = True

   pbModified = True
   With p_oBaseClas
      If .isMasterModify = False And .isDetailModify = False Then
         pbModified = False
         Cancel = False
         GoTo endProc
      End If
      
      ' »»» Rex - 2010-08-03
      '  Discount is based on diff field
      If pbBranchxx Then
         If isDealerDiscOK() = False Then GoTo endProc
      Else
         If isTotalDiscOK() = False Then GoTo endProc
      End If
   
      'MS-2007.12.18
      If .EditMode = xeModeAddNew Then
         .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
      End If
   
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, 1) = "" Then Exit For
   
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      Next
   End With
   
   Cancel = False
   
endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub Class_Terminate()
   Set p_oBaseClas = Nothing
End Sub

' »»» Rex - 2010-08-03
'  Discount of branch wholesaler is based from Dealer Discount field
Private Function isDealerDiscOK() As Boolean
   Dim lnItemPrice As Double
   Dim lnDiscPrice As Double, lnDiscount As Double
   Dim lnMDiscount As Double
   Dim lsAppvID As String, lsAppvName As String, lsAppvCD As String
   Dim lnAppvRights As Integer
   Dim lbGetApproval As Boolean
   Dim lsSQL As String
                                                                                                          
   isDealerDiscOK = False
                                                                                                          
   lnAppvRights = p_oAppDrivr.UserLevel
   lsAppvID = p_oAppDrivr.UserID
   
   With p_oBaseClas
      lnMDiscount = 0
      lnDiscount = 0
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sPartsIDx") = "" Then Exit For
         lnItemPrice = .Detail(pnCtr, "nSelPrice") * .Detail(pnCtr, "nQuantity")
         lnMDiscount = lnMDiscount + _
                  Round(lnItemPrice * .Detail(pnCtr, "nDelrDisc") / 100, 2)
         lnDiscPrice = .Detail(pnCtr, "nSelPrice") * .Detail(pnCtr, "nQuantity")
         lnDiscPrice = (lnItemPrice - lnDiscPrice) + Round(lnDiscPrice * .Detail(pnCtr, "nDiscount") / 100, 2)
         lnDiscPrice = lnDiscPrice - .Detail(pnCtr, "nAddDiscx")
         lnDiscount = lnDiscount + lnDiscPrice
      Next
      
      lnDiscount = lnDiscount + _
                  Round(.Master("nTranTotl") * .Master("nDiscount") / 100, 2) + _
                  .Master("nAddDiscx")
      
      If lnDiscount = 0# Then
         isDealerDiscOK = True
         GoTo endProc
      End If
      
      lbGetApproval = False
      If p_oAppDrivr.UserLevel = xeEncoder Or p_oAppDrivr.UserLevel = xeSupervisor Then
         lbGetApproval = True
      ElseIf p_oAppDrivr.UserLevel = xeManager And lnMDiscount < lnDiscount Then
         lbGetApproval = True
      End If
      
      'TODO: incorporate APPROVAL OF discount here...
      If lbGetApproval Then
         If GetApproval(p_oAppDrivr, lnAppvRights, lsAppvID, lsAppvName, p_oAppDrivr.MenuName) = False Then GoTo endProc
         
         If lnAppvRights = xeSupervisor Then
            MsgBox "Discount is not within the range of the approving officer!!!" & vbCrLf & _
               "Request can not be granted!!!", vbCritical, "Warning"
            GoTo endProc
         ElseIf lnAppvRights = xeManager Then
            If lnMDiscount < lnDiscount Then
               MsgBox "Discount is not within the range of the approving officer!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         ElseIf Not (lnAppvRights = xeSysOwner Or lnAppvRights = xeEngineer) Then
            'kalyptus - 2014.04.12 04:17pm
            'if discount is greater than supposed discount then asked for approval code
            If Not GetCodeApproval(p_oAppDrivr, lsAppvCD, lsAppvID, lsAppvName) Then
               MsgBox "Discount given is greater than approved discount!!!" & vbCrLf & _
                     "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            Else
               If isValidApproveCode( _
                  p_oSCA_Const.pxeWholeSaleDiscount, _
                  p_oAppDrivr.BranchCode, _
                  Mid(lsAppvCD, 4, 1), _
                  .Master("dTransact"), _
                  p_sFullName, _
                  lsAppvCD) Then
                  
                  lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                             p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
                  lsSQL = "INSERT INTO xxxSCA_Usage" & _
                         " SET sTransNox = " & strParm(lsSQL) & _
                            ", sApprCode = " & strParm(lsAppvCD) & _
                            ", sApproved = " & strParm(lsAppvID) & _
                            ", sSystemCD = " & strParm(p_oSCA_Const.pxeWholeSaleDiscount) & _
                            ", sSourceNo = " & strParm(.Master("sTransNox")) & _
                            ", sSourceCD = " & strParm(pxeSPWholesale) & _
                            ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                            ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
                  
                  p_oAppDrivr.Execute lsSQL, "xxxSCA_Usage"
               
               Else
                  MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                     "Verify entry then try again!", vbCritical, "Warning"
                  GoTo endProc
               End If
            End If
         End If
      End If
      .Master("sApproved") = lsAppvID
   End With
   isDealerDiscOK = True

endProc:
   Exit Function
End Function

Private Function isTotalDiscOK() As Boolean
   Dim lnItemPrice As Double
   Dim lnDiscPrice As Double, lnDiscount As Double
   Dim lnMDiscount As Double, lnSDiscount As Double
   Dim lsAppvID As String, lsAppvName As String, lsAppvCD As String
   Dim lnAppvRights As Integer
   Dim lbGetApproval As Boolean
   Dim lsSQL As String
   
                                                                                                          
   isTotalDiscOK = False
                                                                                                          
   lnAppvRights = p_oAppDrivr.UserLevel
   lsAppvID = p_oAppDrivr.UserID
   
   With p_oBaseClas
      lnSDiscount = 0
      lnMDiscount = 0
      lnDiscount = 0
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sPartsIDx") = "" Then Exit For
         Debug.Print .Detail(pnCtr, "sPartsIDx")
         lnItemPrice = .Detail(pnCtr, "nUnitPrce") * .Detail(pnCtr, "nQuantity")
         lnMDiscount = lnMDiscount + _
                  Round(lnItemPrice * .Detail(pnCtr, "nMMaxDisc") / 100, 2)
         lnSDiscount = lnSDiscount + _
                  Round(lnItemPrice * .Detail(pnCtr, "nSMaxDisc") / 100, 2)
         lnDiscPrice = .Detail(pnCtr, "nUnitPrce") * .Detail(pnCtr, "nQuantity")
         lnDiscPrice = (lnItemPrice - lnDiscPrice) + Round(lnDiscPrice * .Detail(pnCtr, "nDiscount") / 100, 2)
         lnDiscPrice = lnDiscPrice - .Detail(pnCtr, "nAddDiscx")
         lnDiscount = lnDiscount + lnDiscPrice
      Next
      
      lnDiscount = lnDiscount + _
                  Round(.Master("nTranTotl") * .Master("nDiscount") / 100, 2) + _
                  .Master("nAddDiscx")
      
      If lnDiscount = 0# Then
         isTotalDiscOK = True
         GoTo endProc
      End If
      
      lbGetApproval = False
      If p_oAppDrivr.UserLevel = xeEncoder Then
         lbGetApproval = True
      ElseIf p_oAppDrivr.UserLevel = xeSupervisor And lnSDiscount < lnDiscount Then
         lbGetApproval = True
      ElseIf p_oAppDrivr.UserLevel = xeManager And lnMDiscount < lnDiscount Then
         lbGetApproval = True
      End If
      
      'TODO: incorporate APPROVAL OF discount here...
      If lbGetApproval Then
         If GetApproval(p_oAppDrivr, lnAppvRights, lsAppvID, lsAppvName, p_oAppDrivr.MenuName) = False Then GoTo endProc
         
         If lnAppvRights = xeSupervisor Then
            If lnSDiscount < lnDiscount Then
               MsgBox "Discount is not within the range of the approving officer!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         ElseIf lnAppvRights = xeManager Then
            If lnMDiscount < lnDiscount Then
               MsgBox "Discount is not within the range of the approving officer!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         ElseIf Not (lnAppvRights = xeSysOwner Or lnAppvRights = xeEngineer) Then
            'kalyptus - 2014.04.12 05:04pm
            'if discount is greater than supposed discount then asked for approval code
            If Not GetCodeApproval(p_oAppDrivr, lsAppvCD, lsAppvID, lsAppvName) Then
               MsgBox "Discount given is greater than approved discount!!!" & vbCrLf & _
                     "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            Else
               If isValidApproveCode( _
                  p_oSCA_Const.pxeWholeSaleDiscount, _
                  p_oAppDrivr.BranchCode, _
                  Mid(lsAppvCD, 4, 1), _
                  .Master("dTransact"), _
                  p_sFullName, _
                  lsAppvCD) Then
                  
                  lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                             p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
                  lsSQL = "INSERT INTO xxxSCA_Usage" & _
                         " SET sTransNox = " & strParm(lsSQL) & _
                            ", sApprCode = " & strParm(lsAppvCD) & _
                            ", sApproved = " & strParm(lsAppvID) & _
                            ", sSystemCD = " & strParm(p_oSCA_Const.pxeWholeSaleDiscount) & _
                            ", sSourceNo = " & strParm(.Master("sTransNox")) & _
                            ", sSourceCD = " & strParm(pxeSPWholesale) & _
                            ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                            ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
                  
                  p_oAppDrivr.Execute lsSQL, "xxxSCA_Usage"
               
               Else
                  MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                     "Verify entry then try again!", vbCritical, "Warning"
                  GoTo endProc
               End If
            End If
         End If
      End If
      .Master("sApproved") = lsAppvID
   End With
   isTotalDiscOK = True

endProc:
   Exit Function
End Function
