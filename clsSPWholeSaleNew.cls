VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPWholeSaleNew"
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
'     Spareparts Sales Object
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
'  XerSys [ 10/29/2004 10:00 am ]
'     Start creating this object.
'  XerSys [ 12/06/2004 10:10 am ]
'     Finalized coding this object.
'  XerSys [ 12/15/2004 03:52 pm ]
'     Oh boy, till now, i hadn't finalized this object yet. Nevertheless lets proccee
'  XerSys [ 10/26/2005 04:52 pm ]
'     Add the issuance of ordered parts...
'  Kalyptus [ 12/11/2006 10:08 am]
'     Implements the BugTracker
'  Kalyptus [ 04/02/2007 10:30 am]
'     Updated the p_oBase_Clas to conform it with the clsMasterDetail specification
'  XerSys [ 08/03/2007 10:21 am ]
'     Update this object to adopt the new implementation of SP Inventory Transaction
'  XerSys [ 08/24/2007 09:31 am ]
'     Payment of this transaction should be credited to main office, thus create a
'        transaction that will hold the payment made by this.
'  XerSys [ 08/02/2010 10:33 am
'     Incorporate this to branch level with the following consideration:
'        - Sales Invoice must be issued
'        - Payment Mode
'           - Cash. Status of Wholesale must be Closed
'           - Check. Status of Wholesale must be Open. Upon clearing, invoice must be
'              issued. Status will be changed to Posted.
'        - Dealer Discount must be used as discount rate
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeSystemCode As String = "SP"
Private Const pxeSourceCode As String = "SPWl"
Private Const pxeMODULENAME = "clsSPWholeSale"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oSCA_Const As New clsCodeGenConst
Private p_xChecks As CheckInfo

Private p_sSalesInv As String
Private p_sBranchCd As String
Private p_sFullName As String
Private p_sAddressx As String
Private p_nPrevTotl As Currency
Private p_sPrevIDxx As String

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pbChecksOK As Boolean
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
   End Select
   
   RaiseEvent DetailRetrieved(Index)
End Property

'kalyptus - 2022.10.12 10:03am
'Added way on viewing the credit limit and current balance of the customer
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
   Case 1, "dtransact"
'      p_oBaseClas.Master(Index) = getValidTranDate(p_oAppDrivr.SysDate, Value)
       p_oBaseClas.Master(Index) = getValidDate(p_oAppDrivr, Value)
   Case 2, "scompnynm"
      getCompany Value, False
   Case 3, "xfullname"
      ' »»» Rex - 2010-08-03
      '  allow search of customer name for branch level used
      If pbBranchxx Then
         getCustomer Value, False
      Else
         p_sFullName = Value
      End If
   Case 4, "xaddressx"
      p_sAddressx = Value
   Case 5, "stermidxx"
      getTerm Value, False
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
End Property

Property Get SalesInv() As String
   SalesInv = p_sSalesInv
End Property

Property Let SalesInv(ByVal Value As String)
   p_sSalesInv = Value
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

Function InitTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   pbBranchxx = Not p_oAppDrivr.IsWarehouse
   
   Set p_oBaseClas = New clsMasterDetail

   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "SP_WSO_Master"
      .DetailTable = "SP_WSO_Detail"
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", b.sCompnyNm" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xFullName" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName) xAddressx" & _
                        ", e.sTermName" & _
                        ", a.nGrossAmt" & _
                        ", a.nDiscRate" & _
                        ", a.dDueDatex" & _
                        ", a.nAddDiscx" & _
                        ", a.nFreightx" & _
                        ", a.sRemarksx" & _
                        ", a.sApproved" & _
                        ", a.nAmtPaidx" & _
                        ", a.nTranTotl" & _
                        ", a.cTranStat" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", a.sClientID" & _
                        ", a.sTermIDxx" & _
                        ", f.nDiscount" & _
                        ", a.sCheckTNo" & _
                        ", a.nEntryNox"
                        
      .MasterQuery = .MasterQuery & _
                        ", a.cWithOrdr" & _
                        ", a.nCheckAmt" & _
                        ", f.nCredLimt" & _
                        ", f.nABalance" & _
                     " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN Term e" & _
                              " ON a.sTermIDxx = e.sTermIDxx" & _
                        ", Client_Master b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                        ", AR_Client_Master f" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND a.sClientId = f.sClientID" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx"
                        
'                        " AND f.sBranchCd = " & strParm(p_sBranchCd)
                  
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", d.sTypeCode" & _
                        ", e.sModelNme" & _
                        ", c.nQtyOnHnd" & _
                        ", a.nUnitPrce" & _
                        ", a.nQuantity" & _
                        ", a.nDiscount" & _
                        ", a.sTransNox" & _
                        ", a.dModified" & _
                        ", c.nQtyOnHnd + a.nQuantity xQtyOnHnd" & _
                        ", b.nSMaxDisc" & _
                        ", b.nMMaxDisc" & _
                        ", b.nSelPrice" & _
                        ", a.sPartsIDx"
      .DetailQuery = .DetailQuery & _
                        ", b.nDelrDisc" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nLedgerNo" & _
                     " FROM " & .DetailTable & " a" & _
                        ", Spareparts b" & _
                              " LEFT JOIN SP_Model e" & _
                                 " ON b.sModelIDx = e.sModelIDx" & _
                        ", SP_Inventory c" & _
                        ", SP_Type d" & _
                     " WHERE a.sPartsIDx = b.sPartsIDx" & _
                        " AND a.sPartsIDx = c.sPartsIDx" & _
                        " AND b.cPartType = d.cPartType" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                     " ORDER BY a.nEntryNox"
      .Reference = "a.sTransNox|a.sTransNox"

      InitTransaction = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                           " AND IFNULL(a.sOrderNox, '') = ''" & _
                        " ORDER BY b.sCompnyNm" & _
                           ", a.dTransact"
      .BrowseColumn(0) = "sTransNox»sCompnyNm»dTransact"
      .BrowseTitle(0) = "Transaction No»Company Name»Date"
      .BrowseCriteria(0) = "a.sTransNox»b.sCompnyNm»a.dTransact"
      .BrowseFormat(0) = "@@@@-@@@@@@»@»MMMM DD, YYYY"

      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sBarrCode" & _
                                 ", a.sDescript" & _
                                 ", c.sTypeCode" & _
                                 ", d.sModelNme" & _
                                 ", b.nQtyOnHnd" & _
                                 ", a.nSelPrice" & _
                                 ", a.nSMaxDisc" & _
                                 ", a.nMMaxDisc" & _
                                 ", a.sPartsIDx" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nLedgerNo" & _
                                 ", a.nDelrDisc" & _
                              " FROM Spareparts a" & _
                                    " LEFT JOIN SP_Model d" & _
                                       " ON a.sModelIDx = d.sModelIDx" & _
                                 ", SP_Inventory b" & _
                                 ", SP_Type c" & _
                              " WHERE a.sPartsIDx = b.sPartsIDx" & _
                                 " AND a.cPartType = c.cPartType" & _
                                 " AND b.sBranchCd = " & strParm(p_sBranchCd)
      .BrowseDetailColumn(0) = "sBarrCode»sDescript»sTypeCode»sModelNme»nQtyOnHnd»nSelPrice"
      .BrowseDetailTitle(0) = "Bar Code»Description»Type»Model»QtyOnHand»Price"
      .BrowseDetailFormat(0) = "@»@»@»@»#0»#,##0.00"
      .BrowseDetailCriteria(0) = "a.sBarrCode»a.sDescript"

      ' Query statement for the Customer field
      'kalyptus - 2022.10.12 09:57am
      'Added AR_Client_Master table in our left join
      'Added AR_Client_Master->nCredLimt and AR_Client_Master->nABalance in our query
      .BrowseQuery(19) = "SELECT DISTINCT" & _
                           "  a.sClientID" & _
                           ", a.sCompnyNm" & _
                           ", CONCAT(a.sLastName, ', ', a.sFrstName) as xFullName" & _
                           ", CONCAT(a.sAddressx, ', ', c.sTownName, ', ', d.sProvName) as xAddressx" & _
                           ", e.sTermName" & _
                           ", IFNULL(f.nDiscount, b.nDiscount) nDiscount" & _
                           ", IFNULL(f.nCredLimt, b.nCredLimt) nCredLimt" & _
                           ", IFNULL(f.nABalance, b.nABalance) nABalance" & _
                        " FROM Client_Master a" & _
                                 " LEFT JOIN AR_Client_Master f ON a.sClientID = f.sClientID" & _
                           ", SP_Customer b" & _
                                 " LEFT JOIN Term e ON b.sTermIDxx = e.sTermIDxx" & _
                           ", TownCity c" & _
                           ", Province d" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           " AND a.sTownIDxx = c.sTownIdxx" & _
                           " AND c.sProvIDxx = d.sProvIDxx" & _
                           " AND b.cRecdStat = " & strParm(xeRecStateActive)
      
      ' »»» Rex - 2010-08-03
      '  search by customer name in branch level implementation
      If pbBranchxx Then
         .BrowseColumn(19) = "sClientID»xFullName»xAddressx"
         .BrowseTitle(19) = "Client ID»Owner»Address"
         .BrowseCriteria(19) = "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName)"
      Else
         .BrowseColumn(19) = "sClientID»sCompnyNm»xFullName»xAddressx"
         .BrowseTitle(19) = "Client ID»Company Name»Owner»Address"
         .BrowseCriteria(19) = "a.sClientID»a.sCompnyNm"
      End If

      ' Query statement for the term field
      .BrowseQuery(20) = "SELECT" & _
                        "  sTermIDxx" & _
                        ", sTermName" & _
                        ", nTermDays" & _
                        ", nDiscDays" & _
                        ", nDiscount" & _
                     " FROM Term" & _
                     " WHERE cRecdStat = " & xeRecStateActive
      .BrowseColumn(20) = "sTermName»nTermDays»nDiscDays»nDiscount"
      .BrowseTitle(20) = "Term Name»Term»Disc Days»Disc"
      .BrowseFormat(20) = "@»0 Day/s»0 Day/s»0.00"
   End With

   pbInitTran = InitTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "NewTransaction"
   'On Error GoTo errProc
      
   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "SaveTransaction"
   'On Error GoTo errProc
      
   If pbInitTran Then SaveTransaction = p_oBaseClas.SaveTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "OpenTransaction"
   'On Error GoTo errProc
      
   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "DeleteTransaction"
   'On Error GoTo errProc
      
   If pbInitTran Then DeleteTransaction = p_oBaseClas.DeleteTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, _
                           Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lsRecord As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String

   lsProcName = "SearchTransaction"
   'On Error GoTo errProc
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

Function PostTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "PostTransaction"
   'On Error GoTo errProc

   If pbInitTran Then PostTransaction = p_oBaseClas.PostTransaction(sTransNo)
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "CancelTransaction"
   'On Error GoTo errProc
   
   If pbInitTran Then CancelTransaction = p_oBaseClas.CancelTransaction()
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "CloseTransaction"
   'On Error GoTo errProc
   
   If pbInitTran Then CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function UpdateTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "UpdateTransaction"
   'On Error GoTo errProc
   UpdateTransaction = False
   
'   If pbChecksOK = False Then
'      MsgBox "Check payment of this transaction has been deposited!!!" & _
'               vbCrLf & "Transaction is Not Allowed to be Modify!!!", vbCritical, "Warning"
'      GoTo endProc
'   End If
   p_oBaseClas.EditMode = xeModeUpdate
   UpdateTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function AddDetail() As Boolean
   With p_oBaseClas
      If .Detail(.ItemCount - 1, "sPartsIDx") <> "" Then
         .AddDetail

         pnCtr = .ItemCount - 1
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
         .Detail(pnCtr, "sBarrCode") = ""
         .Detail(pnCtr, "sDescript") = ""
         .Detail(pnCtr, "sTypeCode") = ""
         .Detail(pnCtr, "sModelNme") = ""
         .Detail(pnCtr, "nQtyOnHnd") = 0
         .Detail(pnCtr, "nUnitPrce") = 0#
         .Detail(pnCtr, "nQuantity") = 0
         .Detail(pnCtr, "nDiscount") = 0
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(pnCtr, "sPartsIDx") = ""
         .Detail(pnCtr, "nResvOrdr") = 0
         .Detail(pnCtr, "nBackOrdr") = 0
         .Detail(pnCtr, "nLedgerNo") = 0
      End If
      
      'AddDetail = .AddDetail
      AddDetail = True
   End With
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   If p_oBaseClas.DeleteDetail(Index) Then
      If p_oBaseClas.Detail(p_oBaseClas.ItemCount - 1, "sPartsIDx") <> "" Then
         DeleteDetail = AddDetail
      Else
         DeleteDetail = True
      End If
   End If

   'DeleteDetail = p_oBaseClas.DeleteDetail(Index)
End Function

Function SearchDetail(ByVal Row As Long, _
                        ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsProcName As String

   lsProcName = "SearchDetail"
   'On Error GoTo errProc
   SearchDetail = False

   If pbInitTran = False Then GoTo endProc

   Index = LCase(Index)
   Select Case Index
   Case 1, "sbarrcode"
      SearchDetail = getDetail(Row, 1, Value, True)
   Case 2, "sdescript"
      SearchDetail = getDetail(Row, 2, Value, True)
   End Select

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & Row _
                       & ", " & Index _
                       & ", " & Value & " )"
End Function

Function SearchMaster(ByVal Index As Variant, Optional Value As Variant = "") As Boolean
   If pbInitTran = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "scompnynm"
      getCompany Value, True
   Case 3, "sfullname"
      ' »»» Rex - 2010-08-03
      '  allow search of customer name for branch level used
      If pbBranchxx Then
         getCustomer Value, True
      End If
   Case 5, "stermidxx"
      getTerm Value, True
   Case Else
      SearchMaster = p_oBaseClas.SearchMaster(Index)
   End Select
   SearchMaster = True
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

Private Sub clearReceipt()
   p_xChecks.TransNox = ""
   p_xChecks.AcctNoxx = ""
   p_xChecks.BankIDxx = ""
   p_xChecks.CheckNox = ""
   p_xChecks.CheckDte = Date
   p_xChecks.Amountxx = 0
End Sub

Private Function getDetail(ByVal lnRow As Integer, _
                        ByVal lnIndex As Integer, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsProcName As String

   lsProcName = "GetDetail"
   'On Error GoTo errProc
   getDetail = False

   With p_oBaseClas
      If lsValue <> "" Then
         If lnIndex = 1 Then     ' Bar Code
            If lsValue = .Detail(lnRow, "sBarrCode") Then GoTo endProc
            If lbSearch Then
               lsDetail = "a.sBarrCode LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "a.sBarrcode = " & strParm(Trim(lsValue))
            End If
         Else                    ' Description
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
      
      .Detail(lnRow, "sBarrCode") = lasDetail(0)
      .Detail(lnRow, "sDescript") = lasDetail(1)
      .Detail(lnRow, "sTypeCode") = lasDetail(2)
      .Detail(lnRow, "sModelNme") = lasDetail(3)
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(4)
      .Detail(lnRow, "nUnitPrce") = lasDetail(5)
      .Detail(lnRow, "nSelPrice") = lasDetail(5)
      .Detail(lnRow, "nSMaxDisc") = lasDetail(6)
      .Detail(lnRow, "nMMaxDisc") = lasDetail(7)
      .Detail(lnRow, "sPartsIDx") = lasDetail(8)
      .Detail(lnRow, "nResvOrdr") = lasDetail(9)
      .Detail(lnRow, "nBackOrdr") = lasDetail(10)
      .Detail(lnRow, "nLedgerNo") = lasDetail(11)
      .Detail(lnRow, "nDelrDisc") = lasDetail(12)
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "nDiscount") = 0
   End With
   
   getDetail = True
   
endProc:
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)
   RaiseEvent DetailRetrieved(7)

   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "sTypeCode") = ""
      .Detail(lnRow, "sModelNme") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nUnitPrce") = 0#
      .Detail(lnRow, "nSelPrice") = 0#
      .Detail(lnRow, "nSMaxDisc") = 0
      .Detail(lnRow, "nMMaxDisc") = 0
      .Detail(lnRow, "sPartsIDx") = ""
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "nDiscount") = 0
      .Detail(lnRow, "nResvOrdr") = 0
      .Detail(lnRow, "nBackOrdr") = 0
      .Detail(lnRow, "nLedgerNo") = 0
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lnIndex _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Function

'Private Function getDetail(ByVal lnRow As Integer, _
'                        ByVal lnIndex As Integer, _
'                        ByVal lsValue As String, _
'                        ByVal lbSearch As Boolean) As Boolean
'   Dim lsDetail As String
'   Dim lasDetail() As String
'   Dim lsProcName As String
'
'   lsProcName = "GetDetail"
'   'On Error GoTo errProc
'   getDetail = False
'
'   With p_oBaseClas
'      If lsValue <> "" Then
'         If lnIndex = 1 Then     ' Bar Code
'            If lsValue = .Detail(lnRow, "sBarrCode") Then GoTo endProc
'            If lbSearch Then
'               'she 2021-08-13
'               'replace "-" to "" for barcode reader use
''               lsDetail = "a.sBarrCode LIKE " & strParm(Trim(lsValue) & "%")
'               lsDetail = "(a.sBarrCode LIKE " & strParm(Trim(lsValue) + "%") & _
'                           " OR REPLACE(a.sBarrcode, '-', '') LIKE " & strParm(Trim(lsValue) + "%") & ")"
'            Else
''               lsDetail = "a.sBarrcode = " & strParm(Trim(lsValue))
'               lsDetail = "(a.sBarrCode = " & strParm(Trim(lsValue)) & _
'                           " OR REPLACE(a.sBarrcode, '-', '') = " & strParm(Trim(lsValue)) & ")"
'            End If
'         Else                    ' Description
'            If lsValue = .Detail(lnRow, "sDescript") Then GoTo endProc
'            If lbSearch Then
'               lsDetail = "a.sDescript LIKE " & strParm(Trim(lsValue) & "%")
'            Else
'               lsDetail = "a.sDescript = " & strParm(Trim(lsValue))
'            End If
'         End If
'      ElseIf lbSearch = False Then
'         GoTo endWithClear
'      End If
'
'      lsDetail = .getDetail(lnRow, lsDetail)
'      If lsDetail = Empty Then
'         If lbSearch = False Then
'            GoTo endWithClear
'         Else
'            GoTo endProc
'         End If
'      End If
'      lasDetail = Split(lsDetail, "»")
'
'      'mac 2020.11.04
'      '  find the record if already exists on record
'      pnCtr = 0
'      Do While pnCtr <= .ItemCount - 1
'         If .Detail(pnCtr, "sPartsIDx") = lasDetail(8) Then
'            'get the row index
'            lnRow = pnCtr
'            Exit Do
'         End If
'
'         pnCtr = pnCtr + 1
'      Loop
'      'end - 'mac 2020.11.04
'
'      .Detail(lnRow, "sBarrCode") = lasDetail(0)
'      .Detail(lnRow, "sDescript") = lasDetail(1)
'      .Detail(lnRow, "sTypeCode") = lasDetail(2)
'      .Detail(lnRow, "sModelNme") = lasDetail(3)
'      .Detail(lnRow, "nQtyOnHnd") = lasDetail(4)
'      .Detail(lnRow, "nUnitPrce") = IIf(.Detail(lnRow, "nQuantity") = 0, lasDetail(5), .Detail(lnRow, "nUnitPrce"))
'      .Detail(lnRow, "nSelPrice") = lasDetail(5)
'      .Detail(lnRow, "nSMaxDisc") = lasDetail(6)
'      .Detail(lnRow, "nMMaxDisc") = lasDetail(7)
'      .Detail(lnRow, "sPartsIDx") = lasDetail(8)
'      .Detail(lnRow, "nResvOrdr") = lasDetail(9)
'      .Detail(lnRow, "nBackOrdr") = lasDetail(10)
'      .Detail(lnRow, "nLedgerNo") = lasDetail(11)
'      .Detail(lnRow, "nDelrDisc") = lasDetail(12)
'      .Detail(lnRow, "nQuantity") = .Detail(lnRow, "nQuantity") + 1
'
'      If .Detail(lnRow, "nQuantity") > .Detail(lnRow, "nQtyOnHnd") Then .Detail(lnRow, "nQuantity") = .Detail(lnRow, "nQtyOnHnd")
'
''      .Detail(lnRow, "nUnitPrce") = lasDetail(5)
''      .Detail(lnRow, "nQuantity") = 0
''      .Detail(lnRow, "nDiscount") = 0
'   End With
'
'   getDetail = True
'
'endProc:
'   RaiseEvent DetailRetrieved(1)
'   RaiseEvent DetailRetrieved(2)
'   RaiseEvent DetailRetrieved(3)
'   RaiseEvent DetailRetrieved(4)
'   RaiseEvent DetailRetrieved(5)
'   RaiseEvent DetailRetrieved(6)
'   RaiseEvent DetailRetrieved(7)
'
'   Exit Function
'endWithClear:
'   With p_oBaseClas
'      .Detail(lnRow, "sBarrCode") = ""
'      .Detail(lnRow, "sDescript") = ""
'      .Detail(lnRow, "sTypeCode") = ""
'      .Detail(lnRow, "sModelNme") = ""
'      .Detail(lnRow, "nQtyOnHnd") = 0
'      .Detail(lnRow, "nUnitPrce") = 0#
'      .Detail(lnRow, "nSelPrice") = 0#
'      .Detail(lnRow, "nSMaxDisc") = 0
'      .Detail(lnRow, "nMMaxDisc") = 0
'      .Detail(lnRow, "sPartsIDx") = ""
'      .Detail(lnRow, "nQuantity") = 0
'      .Detail(lnRow, "nDiscount") = 0
'      .Detail(lnRow, "nResvOrdr") = 0
'      .Detail(lnRow, "nBackOrdr") = 0
'      .Detail(lnRow, "nLedgerNo") = 0
'   End With
'   GoTo endProc
'errProc:
'    ShowError lsProcName & "( " & lnRow _
'                        & ", " & lnIndex _
'                        & ", " & lsValue _
'                        & ", " & lbSearch & " )"
'End Function

Private Sub getCompany(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "getCompany"
   'On Error GoTo errProc
   
   If lsValue <> "" Then
      If StrComp(lsValue, p_oBaseClas.Master("sCompnyNm"), vbTextCompare) = 0 Then GoTo endProc
      If lbSearch Then
         lsMaster = "sCompnyNm LIKE " & strParm(Trim(lsValue) & "%") '(previous code of this line) lsMaster = "a.sCompnyNm LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsMaster = "sCompnyNm = " & strParm(Trim(lsValue)) '(previous code of this line) lsMaster = "a.sCompnyNm = " & strParm(Trim(lsValue))
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
      'kalyptus - 2022.10.12 09:57am
      'Added credit limit and balance
      .Master("nCredLimt") = lasMaster(6)
      .Master("nABalance") = lasMaster(7)
      
      p_sFullName = lasMaster(2)
      p_sAddressx = lasMaster(3)
      getTerm lasMaster(4), False
   End With

endProc:
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master("sCompnyNm") = ""
      .Master("sClientId") = ""
      .Master("nDiscount") = 0#
      p_sFullName = ""
      p_sAddressx = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Sub

' »»» Rex - 2010-08-03
'  Search customer by name
Private Sub getCustomer(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "getCustomer"
   'On Error GoTo errProc
   
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
      'kalyptus - 2022.10.12 09:57am
      'Added credit limit and balance
      .Master("nCredLimt") = lasMaster(6)
      .Master("nABalance") = lasMaster(7)
      
      p_sFullName = lasMaster(2)
      p_sAddressx = lasMaster(3)
      
      
      getTerm lasMaster(4), False
   End With

endProc:
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   
   'kalyptus - 2022.10.12 09:57am
   'Added credit limit and balance
   RaiseEvent MasterRetrieved(6)
   RaiseEvent MasterRetrieved(7)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master("sCompnyNm") = ""
      .Master("sClientId") = ""
      .Master("nDiscount") = 0#
      
      'kalyptus - 2022.10.12 09:57am
      'Added credit limit and balance
      .Master("nCredLimt") = 0#
      .Master("nABalance") = 0#
      
      p_sFullName = ""
      p_sAddressx = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Sub

Private Sub getTerm(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "GetTerm"
   'On Error GoTo errProc
   
   With p_oBaseClas
      If lsValue <> "" Then
         If StrComp(lsValue, .Master("sTermName"), vbTextCompare) = 0 Then GoTo endProc
         
         If lbSearch Then
            lsMaster = "sTermName LIKE " & strParm(Trim(lsValue) & "%")
         Else
            lsMaster = "sTermName = " & strParm(Trim(lsValue))
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
   
      lsMaster = .getMaster(20, lsMaster)
      If lsMaster = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")
      
      .Master("sTermIDxx") = lasMaster(0)
      .Master("sTermName") = lasMaster(1)
      .Master("dDueDatex") = DateAdd("d", lasMaster(2), .Master("dTransact"))
   End With

endProc:
   RaiseEvent MasterRetrieved(5)
   RaiseEvent MasterRetrieved(8)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master("sTermIDxx") = ""
      .Master("sTermName") = ""
      .Master("dDueDatex") = .Master("dTransact")
   End With
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lsValue _
                       & ", " & lbSearch & " )"
End Sub

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
         lnItemPrice = .Detail(pnCtr, "nSelPrice") * .Detail(pnCtr, "nQuantity")
         lnMDiscount = lnMDiscount + _
                  Round(lnItemPrice * .Detail(pnCtr, "nMMaxDisc") / 100, 2)
         lnSDiscount = lnSDiscount + _
                  Round(lnItemPrice * .Detail(pnCtr, "nSMaxDisc") / 100, 2)
         lnDiscPrice = .Detail(pnCtr, "nUnitPrce") * .Detail(pnCtr, "nQuantity")
         lnDiscPrice = (lnItemPrice - lnDiscPrice) + Round(lnDiscPrice * .Detail(pnCtr, "nDiscount") / 100, 2)
         lnDiscount = lnDiscount + lnDiscPrice
      Next
      
      lnDiscount = lnDiscount + _
                  Round(.Master("nTranTotl") * .Master("nDiscRate") / 100, 2) + _
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
         If getApproval(p_oAppDrivr, lnAppvRights, lsAppvID, lsAppvName, p_oAppDrivr.MenuName) = False Then GoTo endProc
         
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
         lnDiscPrice = .Detail(pnCtr, "nUnitPrce") * .Detail(pnCtr, "nQuantity")
         lnDiscPrice = (lnItemPrice - lnDiscPrice) + Round(lnDiscPrice * .Detail(pnCtr, "nDiscount") / 100, 2)
         lnDiscount = lnDiscount + lnDiscPrice
      Next
      
      lnDiscount = lnDiscount + _
                  Round(.Master("nTranTotl") * .Master("nDiscRate") / 100, 2) + _
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
         If getApproval(p_oAppDrivr, lnAppvRights, lsAppvID, lsAppvName, p_oAppDrivr.MenuName) = False Then GoTo endProc
         
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

Private Function saveIssuedOrder(ByVal lnCtr As Integer) As Boolean
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnOrderQty As Long, lnIssuedxx As Long
   Dim lnRow As Long
   
   lsProcName = "SaveIssuedOrder"
   'On Error GoTo errProc
   saveIssuedOrder = False
   
   With p_oBaseClas
      lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", a.nEntryNox" & _
                  ", a.sPartsIDx" & _
                  ", a.nQuantity - a.nIssuedxx xOrderQty" & _
               " FROM SP_Customer_Order_Detail a" & _
                  ", SP_Customer_Order_Master b" & _
               " WHERE a.sTransNox = b.sTransNox" & _
                  " AND a.cCanceled = " & strParm(xeNo) & _
                  " AND b.cTranStat = " & strParm(xeStatePosted) & _
                  " AND b.sClientID = " & strParm(.Master("sClientID")) & _
                  " AND a.sPartsIDx = " & strParm(.Detail(lnCtr, "sPartsIDx")) & _
               " ORDER BY b.dTransact"
      
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If loRS.EOF Then
         MsgBox lsSQL & vbCrLf & "Unable to Retrieve Spareparts Order Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lnOrderQty = .Detail(lnCtr, "nQuantity")
      Do
         lnIssuedxx = loRS("xOrderQty")
         
         If lnOrderQty < lnIssuedxx Then lnIssuedxx = lnOrderQty
         
         lnOrderQty = lnOrderQty - lnIssuedxx
         lsSQL = "UPDATE SP_Customer_Order_Detail SET" & _
                     " nIssuedxx = nIssuedxx + " & lnIssuedxx & _
                  " WHERE sTransNox = " & strParm(loRS("sTransNox")) & _
                     " AND sPartsIDx = " & strParm(loRS("sPartsIDx")) & _
                     " AND nEntryNox = " & loRS("nEntryNox")
                     
         lnRow = p_oAppDrivr.Execute(lsSQL, "SP_Customer_Order_Detail", p_sBranchCd)
         If lnRow <= 0 Then
            MsgBox lsSQL & vbCrLf & "Unable to Update Spareparts Pending Branch Order Info!!!", vbCritical, "Warning"
            GoTo endProc
         End If
         
         If lnOrderQty = 0 Then Exit Do
         loRS.MoveNext
      Loop Until loRS.EOF
   End With
   
   saveIssuedOrder = True
   
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & lnCtr & " )"
End Function

Private Function saveReceipt() As Boolean
   Dim loReceipt As clsARReceipt
   Dim lsProcName As String

   lsProcName = "saveReceipt"
   'On Error GoTo errProc

   Set loReceipt = New clsARReceipt
   With loReceipt
      Set .AppDriver = p_oAppDrivr
      .Parent = pxeMODULENAME
      .SystemCode = pxeSystemCode
      .Branch = p_sBranchCd
      .DisplayConfirmation = False
      If .InitTransaction() = False Then GoTo endProc

      If p_oBaseClas.EditMode = xeModeAddNew Then
         Call .NewTransaction
      Else
         If .SearchTransaction(p_oBaseClas.Master("sTransNox"), True) = False Then
            Call .NewTransaction

'            MsgBox "Unable to Load Receipt Transaction!", vbCritical, "Warning"
'            GoTo endProc
         End If
      End If

      .Master("dTransact") = p_oBaseClas.Master("dTransact")
      .Master("sClientID") = p_oBaseClas.Master("sClientID")
      .Master("nCheckAmt") = p_xChecks.Amountxx
      .Master("nCashAmtx") = p_oBaseClas.Master("nTranTotl") - p_xChecks.Amountxx
      .Master("nTranTotl") = p_oBaseClas.Master("nTranTotl")
      .Master("sReferNox") = p_oBaseClas.Master("sTransNox")

      .Checks("CheckNox") = p_xChecks.CheckNox
      .Checks("AcctNoxx") = p_xChecks.AcctNoxx
      .Checks("BankIDxx") = p_xChecks.BankIDxx
      .Checks("CheckDte") = p_xChecks.CheckDte
      .Checks("Amountxx") = p_xChecks.Amountxx

      If .SaveTransaction() = False Then GoTo endProc
   End With
   
   saveReceipt = True
   
endProc:
'   Set loReceipt = Nothing
   
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function unSaveReceipt() As Boolean
   Dim loReceipt As clsARReceipt
   Dim lsProcName As String
   
   lsProcName = "unSaveReceipt"
   'On Error GoTo errProc
   
   Set loReceipt = New clsARReceipt
   With loReceipt
      Set .AppDriver = p_oAppDrivr
      .Parent = pxeMODULENAME
      .SystemCode = pxeSystemCode
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc
     
      If .LoadByCondition("sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                           " AND sSystemCd = " & strParm(pxeSystemCode)) = False Then
         MsgBox "Unable to Load Receipt Transaction!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      ' after locating the transaction, delete it
      If .DeleteTransaction() = False Then GoTo endProc
   End With
   
   unSaveReceipt = True
   
endProc:
   Set loReceipt = Nothing
   
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

' »»» Rex - 2010-08-03
'  Branch wholesale issues a sales invoice
Function saveInvoice() As Boolean
   Dim loSPSales As clsSPPOSBranch
   Dim lsOldProc As String
  
   lsOldProc = "saveInvoice"
   'On Error GoTo errProc
   
   If Not pbInitTran Then Exit Function
   
   Set loSPSales = New clsSPPOSBranch
   With loSPSales
      Set .AppDriver = p_oAppDrivr
      .CreateLedger = False
      .DisplayConfirmation = False
      If .InitTransaction() = False Then GoTo endProc
            
      If p_oBaseClas.EditMode = xeModeUpdate Then
         If .LoadByCondition("sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                              " AND sSourceCd = " & strParm(pxeSourceCode)) = False Then
            MsgBox "Unable to Load Invoice!" & vbCrLf & _
                     "Verify your entry then Try Again!", vbCritical, "Warning"
            GoTo endProc
         End If
         Call .UpdateTransaction
      Else
         If .NewTransaction() = False Then GoTo endProc
      End If
      
      .Master("sClientID") = p_oBaseClas.Master("sClientID")
      .Master("dTransact") = p_oBaseClas.Master("dTransact")
      .Master("sSalesInv") = p_sSalesInv
      .Master("nTranTotl") = p_oBaseClas.Master("nTranTotl")
      .Master("nAmtPaidx") = p_oBaseClas.Master("nAmtPaidx")
      .Master("nGiftCpnx") = 0
      .Master("sReferNox") = p_oBaseClas.Master("sTransNox")
      .Master("sSourceCd") = pxeSourceCode
      .Master("sAutoTran") = xeNo
      .Master("nAdvPaymx") = 0
      .Master("sAdvRefer") = ""
  
'      .Checks("sCheckNox") = Checks("sCheckNox")
'      .Checks("sAcctNoxx") = Checks("sAcctNoxx")
'      .Checks("sBankIDxx") = Checks("sBankIDxx")
'      .Checks("dCheckDte") = Checks("dCheckDte")
'      .Checks("nAmountxx") = Checks("nAmountxx")
'      .Checks("sPRNoxxxx") = Checks("sPRNoxxxx")
      
      For pnCtr = 0 To p_oBaseClas.ItemCount - 1
         If p_oBaseClas.Detail(pnCtr, "sPartsIDx") = "" Then Exit For
         
         If pnCtr > .ItemCount - 1 Then .AddDetail
         .Detail(pnCtr, "sPartsIDx") = p_oBaseClas.Detail(pnCtr, "sPartsIDx")
         .Detail(pnCtr, "nQuantity") = p_oBaseClas.Detail(pnCtr, "nQuantity")
         .Detail(pnCtr, "nUnitPrce") = p_oBaseClas.Detail(pnCtr, "nUnitPrce")
         .Detail(pnCtr, "nDiscount") = p_oBaseClas.Detail(pnCtr, "nDiscount")
         .Detail(pnCtr, "cLaborxxx") = xeNo
      Next
      If .SaveTransaction() = False Then GoTo endProc
   End With
   Debug.Print "Invoice Save"
   saveInvoice = True
   
endProc:
   Set loSPSales = Nothing

   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

' »»» Rex - 2010-08-03
'  Remove issued sales invoice
Function unSaveInvoice() As Boolean
   Dim loSPSales As clsSPPOSBranch
   Dim lsOldProc As String
   
   lsOldProc = "unSaveInvoice"
   'On Error GoTo errProc
   
   If Not pbInitTran Then Exit Function
   
   Set loSPSales = New clsSPPOSBranch
   With loSPSales
      Set .AppDriver = p_oAppDrivr
      .CreateLedger = False
      .DisplayConfirmation = False
      If .InitTransaction() = False Then GoTo endProc
   
      If .LoadByCondition("sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                           " AND sSourceCd = " & strParm(pxeSourceCode)) = False Then GoTo endProc
     
      If .DeleteTransaction = False Then GoTo endProc
   End With
   
   unSaveInvoice = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Function unSaveIssuedOrder(ByVal lnQuantity As Integer, ByVal lsPartsIDx As String) As Boolean
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnIssuedxx As Long, lnRow As Long
   
   lsProcName = "UnSaveIssuedOrder"
   'On Error GoTo errProc
   
   With p_oBaseClas
      lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", a.nEntryNox" & _
                  ", a.sPartsIDx" & _
                  ", a.nIssuedxx" & _
               " FROM SP_Customer_Order_Detail a" & _
                  ", SP_Customer_Order_Master b" & _
               " WHERE a.sTransNox = b.sTransNox" & _
                  " AND a.cCanceled = " & strParm(xeNo) & _
                  " AND b.cTranStat = " & strParm(xeStatePosted) & _
                  " AND b.sClientID = " & strParm(.MasterOriginal("sClientID")) & _
                  " AND a.sPartsIDx = " & strParm(lsPartsIDx) & _
                  " AND a.nIssuedxx > 0" & _
               " ORDER BY b.dTransact DESC"
      
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If loRS.EOF Then
         MsgBox "Unable to Retrieve Spareparts Order Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      Do
         lnIssuedxx = loRS("nIssuedxx")
         
         If lnQuantity < lnIssuedxx Then lnIssuedxx = lnQuantity
         
         lnQuantity = lnQuantity - lnIssuedxx
         lsSQL = "UPDATE SP_Customer_Order_Detail SET" & _
                     " nIssuedxx = nIssuedxx - " & lnIssuedxx & _
                  " WHERE sTransNox = " & strParm(loRS("sTransNox")) & _
                     " AND sPartsIDx = " & strParm(loRS("sPartsIDx")) & _
                     " AND nEntryNox = " & loRS("nEntryNox")
                     
         lnRow = p_oAppDrivr.Execute(lsSQL, "SP_Customer_Order_Detail", p_sBranchCd)
         If lnRow <= 0 Then
            MsgBox "Unable to Update Spareparts Pending Branch Order Info!!!", vbCritical, "Warning"
            GoTo endProc
         End If
         
         If lnQuantity = 0 Then Exit Do
         loRS.MoveNext
      Loop Until loRS.EOF
  End With
   
   unSaveIssuedOrder = True
   
endProc:
   Set loRS = Nothing

   Exit Function
errProc:
   ShowError lsProcName & "( " & lnQuantity _
                       & ", " & lsPartsIDx & " )"
End Function

Private Function saveSPTransaction() As Boolean
`Dim loSPTrans As clsSPInventoryTrans
   Dim loCltTrans As clsClientTrans
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lbReorder As Boolean

   lsProcName = "p_oBaseClas_SaveOthers"
   'On Error GoTo errProc
   Cancel = True

   With p_oBaseClas
      Set loSPTrans = New clsSPInventoryTrans
      With loSPTrans
         Set .AppDriver = p_oAppDrivr
         .Branch = p_sBranchCd
         If .InitTransaction() = False Then GoTo endProc
      End With
         
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, 1) = "" Then Exit For
   
'         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
'         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
   
         loSPTrans.Detail(pnCtr, "sPartsIDx") = .Detail(pnCtr, "sPartsIDx")
         loSPTrans.Detail(pnCtr, "nQuantity") = .Detail(pnCtr, "nQuantity")
         loSPTrans.Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
         loSPTrans.Detail(pnCtr, "nResvOrdr") = .Detail(pnCtr, "nResvOrdr")
         loSPTrans.Detail(pnCtr, "nBackOrdr") = .Detail(pnCtr, "nBackOrdr")
         loSPTrans.Detail(pnCtr, "nLedgerNo") = .Detail(pnCtr, "nLedgerNo")
         
         ' add this to update the pending customer order...
         If .Master("cWithOrdr") = xeYes Then
            If saveIssuedOrder(pnCtr) = False Then GoTo endProc
         End If
      Next
      
      ' if a spareparts is not successfully updated, cancel saving.
      If loSPTrans.WholeSale(.Master("sTransNox"), _
            .Master("dTransact"), _
            .EditMode) = False Then
         GoTo endProc
      End If
      
      If .Master("nAmtPaidx") > 0 Then
         ' »»» Rex - 2010-08-03
         '  Branch wholesale issues a sales invoice
         If pbBranchxx Then
            If saveInvoice = False Then GoTo endProc
         Else
            If saveReceipt = False Then GoTo endProc
         End If
      End If
   
      'kalyptus - 2022.10.12 10:36am
      'Save the transaction to AR_Client instead of SP_Customer
'      If SaveSPCustomerTrans(.EditMode) = False Then GoTo endProc
      If Not saveClientTransaction(.EditMode) Then GoTo endProc
      
      Debug.Print "hola"
      ' check if reorder level of some parts has reached
      If loSPTrans.ReOrderParts Then
         MsgBox "Some Parts Included in the Transaction has Reached it Reorder Level Limit!!!" & _
                  vbCrLf & vbCrLf & "Please Inform the Spareparts Manager on this Alert!!!", _
                  vbInformation, "Notice"
      End If
   End With
   
   Cancel = False
   
endProc:
   Set loSPTrans = Nothing
   Set loCltTrans = Nothing

   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Function

Private Function delSPTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim loCltTrans As clsClientTrans
   Dim lsProcName As String
   
   lsProcName = "delSPTransaction"
   'On Error GoTo errProc
   
   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then
         MsgBox "hola"
         GoTo endProc
      End If
      
      For pnCtr = 0 To .ItemCount - 1
         .Detail(pnCtr, "sPartsIDx") = p_oBaseClas.Detail(pnCtr, "sPartsIDx")
         .Detail(pnCtr, "nQtyOnHnd") = p_oBaseClas.Detail(pnCtr, "nQtyOnHnd")
         .Detail(pnCtr, "nResvOrdr") = p_oBaseClas.Detail(pnCtr, "nResvOrdr")
         .Detail(pnCtr, "nBackOrdr") = p_oBaseClas.Detail(pnCtr, "nBackOrdr")
         .Detail(pnCtr, "nLedgerNo") = p_oBaseClas.Detail(pnCtr, "nLedgerNo")
      Next
      
      ' if a serial is not successfully updated, cancel saving.
      If .WholeSale(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("dTransact"), _
            xeModeDelete) = False Then
         MsgBox "wholesale"
         GoTo endProc
      End If
   End With
   
   With p_oBaseClas
      'kalyptus - 2022.10.12 10:36am
      'Save the transaction to AR_Client instead of SP_Customer
      'If SaveSPCustomerTrans(xeModeDelete) = False Then GoTo endProc
      If saveClientTransaction(xeModeDelete) = False Then GoTo endProc

      ' check if payment was made
      If .Master("nAmtPaidx") > 0 Then
         If pbBranchxx Then
            ' »»» Rex - 2010-08-03
            '  Branch wholesale issues a sales invoice, so remove the issued invoice
            If unSaveInvoice() = False Then GoTo endProc
         Else
            If unSaveReceipt() = False Then GoTo endProc
         End If
      End If
   End With
   
   delSPTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub Class_Terminate()
   Set p_oBaseClas = Nothing
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_InitMaster"
   'On Error GoTo errProc
   
   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sCompnyNm") = ""
      .Master("sTermIDxx") = ""
      .Master("nGrossAmt") = 0#
      .Master("nDiscRate") = 0#
      .Master("dDueDatex") = 0#
      .Master("nAddDiscx") = 0#
      .Master("nFreightx") = 0#
      .Master("nTranTotl") = 0#
      .Master("nAmtPaidx") = 0#
      .Master("sRemarksx") = ""
      .Master("cTranStat") = xeStateOpen
      .Master("sApproved") = ""
      .Master("sClientID") = ""
      .Master("cWithOrdr") = 0
      .Master("sCheckTNo") = ""
      .Master("nEntryNox") = 0
      
      'kalyptus - 2022.10.12 10:00am
      'Added previous total for our computation of credit limit
      .Master("nCredLimt") = 0#
      .Master("nABalance") = 0#
      
      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "sTypeCode") = ""
      .Detail(0, "sModelNme") = ""
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nUnitPrce") = 0#
      .Detail(0, "nQuantity") = 0
      .Detail(0, "nDiscount") = 0
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sPartsIDx") = ""
      .Detail(0, "nSelPrice") = 0#
      .Detail(0, "nSMaxDisc") = 0
      .Detail(0, "nMMaxDisc") = 0
      .Detail(0, "nLedgerNo") = 0
   End With
   
   clearReceipt
   p_sFullName = ""
   p_sAddressx = ""
   
   'kalyptus - 2022.10.12 10:00am
   'Added previous customer id for possible update of customer
   'Added previous total for our computation of credit limit
   p_nPrevTotl = 0
   p_sPrevIDxx = ""
   
endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   Dim loRS As Recordset
   Dim loReceipt As clsARReceipt
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "p_oBaseClas_LoadOthers"
   'On Error GoTo errProc

   With p_oBaseClas
      For pnCtr = 0 To .ItemCount - 1
         .Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "xQtyOnHnd")
      Next
      'p_sFullName = .Master("xFullName") '2017-10-04 she
      p_sFullName = .Master("sCompnyNm")
      p_sAddressx = .Master("xAddressx")

      p_nPrevTotl = .Master("nTranTotl") - .Master("nAmtPaidx")
      p_sPrevIDxx = .Master("sClientID")
      
   End With
   
   If p_oBaseClas.Master("nAmtPaidx") > 0 Then
      If p_oAppDrivr.IsWarehouse Then
         Set loReceipt = New clsARReceipt
         With loReceipt
            Set .AppDriver = p_oAppDrivr
            .Parent = pxeMODULENAME
            .SystemCode = pxeSystemCode
            .Branch = p_sBranchCd
            If .InitTransaction() = False Then GoTo endProc
      
            If .SearchTransaction(p_oBaseClas.Master("sTransNox"), True) = False Then
               MsgBox "Unable to Load Receipt Transaction!", vbCritical, "Warning"
               GoTo endProc
            End If
      
            p_xChecks.CheckNox = .Checks("CheckNox")
            p_xChecks.AcctNoxx = .Checks("AcctNoxx")
            p_xChecks.BankIDxx = .Checks("BankIDxx")
            p_xChecks.CheckDte = .Checks("CheckDte")
            p_xChecks.Amountxx = .Checks("Amountxx")
            p_xChecks.TransNox = .Master("sTransNox")
         End With
      Else
         lsSQL = "SELECT sSalesInv" & _
                  " FROM SP_SO_Master" & _
                  " WHERE sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                     " AND sSourceCd = " & strParm(pxeSPWholesale)
                     
         Set loRS = New Recordset
         loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
         
         If loRS.EOF = False Then
            p_sSalesInv = loRS("sSalesInv")
         End If
      End If
   End If
   
endProc:
   Set loReceipt = Nothing

   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = Not pbModified
End Sub

Private Sub p_oBaseClas_WillCancel(Cancel As Boolean)
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_WillCancel"
   'On Error GoTo errProc
      
   Cancel = Not delSPTransaction

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillDelete(Cancel As Boolean)
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Long

   lsProcName = "p_oBaseClas_WillDelete"
   'On Error GoTo errProc

   Cancel = Not delSPTransaction

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_WillSave"
   'On Error GoTo errProc
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
      
      If p_sPrevIDxx = "" Then
         If Not saveARClient Then
            MsgBox "Unable to save SP Client to AR Client!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
      
      'kalyptus - check credit limit
      If .Master("nCredLimt") < ((.Master("nABalance") - p_nPrevTotl) + (.Master("nTranTotl") - .Master("nAmtPaidx"))) Then
            MsgBox "Unable to save transaction!" & vbCrLf & _
                   "The previous balance and the amount of this transaction will excess the Wholesaler's CREDIT LIMIT", vbCritical, "Warning"
            GoTo endProc
      End If
      
      'MS-2007.12.18
      If .EditMode = xeModeAddNew Then
         'kalyptus - 2022.10.12 10:18am
         'Save the SP_Customer record to AR_Client_Master
         .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
      End If
   
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, 1) = "" Then Exit For
   
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      Next
      
      'mac 2020.11.04
      '  delete last detail row if not parts is selected
      If .Detail(.ItemCount - 1, "sPartsIDx") = "" Then .DeleteDetail (.ItemCount - 1)
   End With
   
   Cancel = False
   
endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Sub

'Private Function SaveSPCustomerTrans(fnEditMode As xeEditMode) As Boolean
'   Dim loCltTrans As clsClientTrans
'   Dim lsOldProc As String
'
'   lsOldProc = "SaveSPCustomerTrans"
'   'On Error GoTo errProc
'
'   Set loCltTrans = New clsClientTrans
'   With loCltTrans
'      Set .AppDriver = p_oAppDrivr
'      .Branch = p_sBranchCd
'      .ClientType = xeClientCustomer
'      .SourceSystem = pxeSystemCode
'      Call .InitTransaction
'
'      ' now save the transaction
'      If .Sales(p_oBaseClas.Master("sTransNox"), _
'            p_oBaseClas.Master("sClientID"), _
'            p_oBaseClas.Master("dTransact"), _
'            p_oBaseClas.Master("nTranTotl"), _
'            p_oBaseClas.Master("nAmtPaidx"), _
'            fnEditMode) = False Then
'         GoTo endProc
'      End If
'   End With
'
'   SaveSPCustomerTrans = True
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsOldProc & "(" & fnEditMode & ")"
'   GoTo endProc
'End Function

Private Function saveARClient() As Boolean
   Dim loClient As clsARClient
   Dim loRS As Recordset
   Dim lsProcName As String
   
   lsProcName = "saveARClient"
   Debug.Print pxeMODULENAME & "." & lsProcName
   '''On Error GoTo errProc
   
   Set loRS = New Recordset

   loRS.Open "SELECT a.*, b.sClientID xClientID" & _
               " FROM SP_Customer a" & _
                  " LEFT JOIN AR_Client_Master b" & _
                     " ON a.sClientID = b.sClientID" & _
               " WHERE a.sClientID = " & strParm(Master("sClientID")) _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If Not IsNull(loRS("xClientID")) Then
      saveARClient = True
      GoTo endProc
   End If
   
   Set loClient = New clsARClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .InitRecord
      .NewRecord
            
      .Master("sClientID") = loRS("sClientID")
      .Master("sTelNoxxx") = loRS("sTelNoxxx")
      .Master("sFaxNoxxx") = loRS("sFaxNoxxx")
      .Master("sCPerson1") = loRS("sCPerson1")
      .Master("sCPPosit1") = loRS("sCPPosit1")
      .Master("sTermIDxx") = loRS("sTermIDxx")
      .Master("sRemarksx") = loRS("sRemarksx")
      .Master("nCredLimt") = loRS("nCredLimt")
      .Master("nABalance") = loRS("nABalance")
      .Master("dCltSince") = loRS("dCltSince")
      .Master("nLedgerNo") = loRS("nLedgerNo")
      .Master("cHoldAcct") = xeNo
      .Master("cAutoHold") = xeNo
      .Master("cRecdStat") = xeYes
         
      saveARClient = .SaveRecord
   End With
   
endProc:
   Set loClient = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

'kalyptus - 2022.10.12 09:46am
Private Function saveClientTransaction(fnEditMode As xeEditMode) As Boolean
   Dim loClientsTrans As clsARClientTrans
   Dim lsProcName As String
   
   lsProcName = "saveClientTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   '''On Error GoTo errProc
   
   Set loClientsTrans = New clsARClientTrans
   With loClientsTrans
      Set .AppDriver = p_oAppDrivr
      .ClientType = xeClientCustomer
      .SourceSystem = pxeSystemCode
      .Branch = p_sBranchCd
      Call .InitTransaction
   End With
   
   With p_oBaseClas
      If Not loClientsTrans.WholeSale(.Master("sTransNox") _
                                          , .Master("sClientID") _
                                          , p_oAppDrivr.ServerDate _
                                          , .Master("nTranTotl") _
                                          , 0# _
                                          , fnEditMode) Then
         GoTo endProc
      End If
   End With
   saveClientTransaction = True


endProc:
   Set loClientsTrans = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Function ViewEncoder()
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsDate As Date
   Dim lsEncoder As String
   Dim lsApprved As String
   
   lsSQL = "SELECT sModified, dModified from xxxReplicationLog" & _
            " WHERE sTransNox LIKE " & strParm(Left(p_oBaseClas.Master("sTransNox"), 6) & "%") & _
               " AND sStatemnt LIKE '%INSERT INTO%sTransNox = %" & p_oBaseClas.Master("sTransNox") & "%'" & _
               " AND sTableNme = 'SP_WSO_Master'"
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   If loRS.EOF Then
      MsgBox "Transaction has no history of encoding.", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lsDate = CDate(loRS("dModified"))
   lsSQL = "SELECT sUserName" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & strParm(loRS("sModified"))

   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If loRS.EOF Then
      MsgBox "User does not Exist in the User Table!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lsEncoder = loRS("sUserName")
   
   lsSQL = "SELECT sUserName" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & strParm(p_oBaseClas.Master("sApproved"))
   
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If loRS.EOF Then
      lsApprved = "NONE"
   Else
      lsApprved = loRS("sUserName")
   End If
   
   MsgBox "    ««« WHOLESALE TRANSACTION »»»" & vbCrLf & vbCrLf & _
                "Encoded By   : " & Decrypt(lsEncoder, p_oAppDrivr.Machinex) & vbCrLf & _
                "Date Encoded : " & Format(lsDate, "MMMM DD, YYYY HH:NN:SS") & vbCrLf & _
                "Approved By  : " & Decrypt(lsApprved, p_oAppDrivr.Machinex) _
                , vbInformation, "Notice"
endProc:
    Exit Function
End Function

