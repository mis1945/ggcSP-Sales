VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPPOSBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Rex S. Adversalo
' XerSys Computing
' Canaoalan, Binmaley, Pangasinan
'
' Copyright 2003 and beyond
' All Rights Reserved
'
'     Spareparts Point of Sale Object (Branch Version)
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
'  XerSys [ 06/10/2006 12:40 pm ]
'     Start creating this object. Adopt SPSales Object
'  XerSys [ 07/24/2006 11:52 am ]
'     Implement the constant value retrieval for special information.
'  XerSys [ 09/21/2006 12:15 pm ]
'     SP Sales must accept Customer Name
'     Allow the client to pay more than the tran total if the payment is thru gift coupon.
'  Kalyptus [ 12/11/2006 10:08 am]
'     Implements the BugTracker
'  Kalyptus [ 04/02/2007 10:23 am]
'     Updated the p_oBase_Clas to conform it with the clsMasterDetail specification
'  XerSys [ 08/14/2007 09:51 am ]
'     Update this object to adopt the new implementation of SP Inventory Transaction
'  XerSys [ 08/28/2007 11:51 am ]
'     Update this object to allow Job Order save it's parts included in the transaction
'        without creating inventory ledger
'  XerSys [ 09/25/2007 08:51 am ]
'     As per approval by Bos Jo/Guan labor must be included in Sales Invoice
'  XerSys [ 09/26/2007 10:59 am ]
'     Add LoadByCondition function to allow user of this object to open it using its
'        own condition.
'  Jheff [ 10/10/2007 12:00 pm ]
'     Add Client for searching and updating.
'  XerSys [ 03/29/2008 01:20 pm ]
'     Implement advance payment for spareparts.
'  XerSys [ 09/16/2008 11:38 am ]
'     Remove the assignment of the quantity on hand for loaded transaction
'  XerSys [ 01/12/2009 09:27 am ]
'     Remove the dealer discount for Parts POS
'  Kalyptus [ 2010/09/22 04:25 pm ]
'     Implemented GCard discounting. cCardType = 17
'  XerSys [ 08/03/2011 09:08 am ]
'     Incorporate order issuance
'  XerSys [ 08/05/2011 12:05 pm ]
'     Remove update capability of this object
'  XerSys [ 10/03/2011 03:08 pm ]
'     Allow manager to approve discount without seeking approval from head office.
'  XerSys [ 01/18/2012 03:32 pm ]
'     Compare total advance payment from total transaction amount to avoid over crediting
'  XerSys [ 07/23/2012 03:02 pm ]
'     Advance payment can be clicked even though no amount is available for crediting.
'  kalyptus [ 08/23/2012 12:02 pm ]
'     Added a capability to reset the status of Provisionary Receipt Master if SI is cancelled
'     or deleted...
'  XerSys [ 10/16/2013 11:13 am ]
'     Allow unissue of order with out affected parts
'  XerSys [ 04/07/2014 10:46 am ]
'     Incorporate crediting of advance payment
'  XerSys [ 12/17/2014 02:52 pm ]
'     Incorporate discount card of other partners
'  XerSys [ 06/08/2015 10:48 ]
'     Incorporate the minimum amount of parts for discounting
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsSPPOSBranch"
Private Const pxeSourceCd As String = "SPSl"
Private Const pxeDiscount As Double = 0.05
Private Const xeGiftCpn As String = "GiftCpnx"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oSCA_Const As New clsCodeGenConst
Private p_oClient As clsClient
Private p_oAdvance As Recordset
Private p_xChecks As CheckInfo

Private p_sBranchCd As String
Private p_nGCpnAmtx As Double
Private p_bNoLedger As Boolean

Private p_sFullName As String
Private p_sAddressx As String
Private p_sAdvRefer As String
Private p_nAdvPaymx As Double
Private p_sParent As String
Private p_bHsParent As Boolean

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pbWHSearch As Boolean
Private pbShowMsg As Boolean
Private pnCtr As Integer

Private p_oReturn As clsSPSalesReturn
Private p_oRSIssd As Recordset

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)
Public Event WillSave(Cancel As Boolean)
Public Event WillDelete(Cancel As Boolean)
Public Event SaveOthers(Cancel As Boolean)
Public Event CancelOthers(Cancel As Boolean)
Public Event DeleteOthers(Cancel As Boolean)

Private p_oDiscount As Recordset
Private p_sDiscCard As String
Private p_sDiscSQL As String


Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(Value As String)
   p_sBranchCd = Value
End Property

'Solution for the supersession of joborder
Property Get IssuedOrder() As Recordset
   Set IssuedOrder = p_oRSIssd
End Property

Property Get Checks(ByVal Index As Variant) As Variant
   Select Case Index
   Case 0, "sCheckNox"
      Checks = p_xChecks.CheckNox
   Case 1, "sAcctNoxx"
      Checks = p_xChecks.AcctNoxx
   Case 2, "sBankIDxx"
      Checks = p_xChecks.BankIDxx
   Case 3, "dCheckDte"
      Checks = p_xChecks.CheckDte
   Case 4, "nAmountxx"
      Checks = p_xChecks.Amountxx
   Case 5, "sPRNoxxxx"
      Checks = p_xChecks.PRNoxxxx
   End Select
End Property

Property Let Checks(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 0, "sCheckNox"
      p_xChecks.CheckNox = Value
   Case 1, "sAcctNoxx"
      p_xChecks.AcctNoxx = Value
   Case 2, "sBankIDxx"
      p_xChecks.BankIDxx = Value
   Case 3, "dCheckDte"
      p_xChecks.CheckDte = Value
   Case 4, "nAmountxx"
      p_xChecks.Amountxx = Value
   Case 5, "sPRNoxxxx"
      p_xChecks.PRNoxxxx = Value
   End Select
End Property

Property Get CouponAmount() As Double
   CouponAmount = p_nGCpnAmtx
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
      If Not p_bHsParent Then
         getDetail Row, Index, Value, False
      Else
         p_oBaseClas.Detail(Row, Index) = Value
      End If
   Case 2, "sdescript"
      If Not p_bHsParent Then
         'Allow an update to the description if parts id loaded is the special id and branch is LGK Dagupan
         If Not (p_oBaseClas.Detail(Row, "sPartsIDx") = "W210003452" And _
            p_oAppDrivr.ClientID = "GGC_B002") Then
            getDetail Row, Index, Value, False
         End If
      Else
         p_oBaseClas.Detail(Row, Index) = Value
      End If
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Let DisplayConfirmation(Value As Boolean)
   pbShowMsg = Value
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      Master = p_sFullName
   Case 4, "xaddressx"
      Master = p_sAddressx
   Case 8, "scardidxx"
      Master = p_sDiscCard
   Case 17, "nadvpaymx"
      Master = p_nAdvPaymx
   Case Else
      Master = p_oBaseClas.Master(Index)
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "stransnox"
   Case 1, "dtransact"
      p_oBaseClas.Master(Index) = getValidTranDate(p_oAppDrivr.SysDate, Value)
      
      RaiseEvent MasterRetrieved(1)
   Case 3, "xfullname"
      Call getCustomer(Value, False)
   Case 4, "xaddressx"
      p_sAddressx = Value
   Case 6, "ngiftcpnx"
      Dim lnNetTotal As Double
      
      p_oBaseClas.Master(Index) = Value
      
      lnNetTotal = p_oBaseClas.Master("nTranTotl") - (Value * p_nGCpnAmtx)
      p_oBaseClas.Master("nAmtPaidx") = IIf(lnNetTotal > 0, lnNetTotal, 0)
      RaiseEvent MasterRetrieved(7)
   Case 8, "sautotran"
'      If p_bNoLedger Then
'         p_oBaseClas.Master(Index) = Value
'      Else
         Select Case Value
         Case xeYes, xeNo
            If Value = xeYes Then
               If Left(p_sParent, 3) <> "cls" Then
                  If Not getAdvancePayment() Then Value = xeNo
               End If
            End If
            
            p_oBaseClas.Master(Index) = Value
            RaiseEvent MasterRetrieved(8)
            RaiseEvent MasterRetrieved(17)
         Case Else
            p_oBaseClas.Master(Index) = xeNo
         End Select
'      End If
   
   Case 10, "scardidxx"
      Call getDiscCard(Value, True)
   Case 17, "nadvpaymx"
      If p_bNoLedger Then p_nAdvPaymx = Value
   Case 18, "sadvrefer"
      If p_bNoLedger Then p_sAdvRefer = Value
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
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

Property Let Parent(ByVal Value As String)
   p_sParent = Value
   p_bHsParent = Trim(Value) <> ""
End Property

Property Let CreateLedger(Value As Boolean)
   p_bNoLedger = Not Value
End Property

Property Set AdvancePaym(Value As Recordset)
   Set p_oAdvance = Value
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   
  lsProcName = "InitTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   Set p_oBaseClas = New clsMasterDetail

   
   With p_oBaseClas
      
      
      
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "SP_SO_Master"
      .DetailTable = "SP_SO_Detail"
      .DisplayMessage = pbShowMsg
      
      .MasterQuery = "SELECT a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sSalesInv" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNULL(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName) xFullName" & _
                        ", CONCAT(IF(IFNull(b.sHouseNox, '') = '', '', CONCAT(b.sHouseNox, ' ')), b.sAddressx, ', ', c.sTownName, ', ', d.sProvName) xAddressx" & _
                        ", a.nTranTotl" & _
                        ", a.nGiftCpnx" & _
                        ", a.nAmtPaidx" & _
                        ", a.sCardIDxx" & _
                        ", a.sAutoTran" & _
                        ", a.cTranStat" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", a.sReferNox" & _
                        ", a.sClientID" & _
                        ", a.sSourceCd" & _
                        ", a.nReplAmtx" & _
                        ", a.nEntryNox" & _
                        ", a.cCardType" & _
                        ", a.sApproved"

      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN Client_Master b" & _
                              " ON a.sClientID = b.sClientID" & _
                           " JOIN TownCity c" & _
                              " ON b.sTownIDxx = c.sTownIDxx" & _
                           " JOIN Province d" & _
                              " ON c.sProvIDxx = d.sProvIDxx"

      'Include SP_Supersession in the Query
      'Kalyptus - 2012.03.17 09:27 am
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", c.nQtyOnHnd" & _
                        ", a.nQuantity" & _
                        ", a.nUnitPrce" & _
                        ", a.nDiscount" & _
                        ", a.nAddDiscx" & _
                        ", a.sTransNox" & _
                        ", a.dModified" & _
                        ", c.nQtyOnHnd + a.nQuantity xQtyOnHnd" & _
                        ", a.sPartsIDx" & _
                        ", d.cPartType" & _
                        ", b.nDelrDisc" & _
                        ", b.nSelPrice" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo" & _
                        ", a.cLaborxxx" & _
                        ", e.sLaborIDx" & _
                        ", e.sLaborNme" & _
                        ", a.nOrderQty"
      .DetailQuery = .DetailQuery & _
                        ", f.sSuperIDx" & _
                     " FROM " & .DetailTable & " a" & _
                           " LEFT JOIN Spareparts b" & _
                              " ON a.sPartsIDx = b.sPartsIDx" & _
                                 " AND a.cLaborxxx = " & xeNo & _
                           " LEFT JOIN SP_Inventory c" & _
                              " ON b.sPartsIDx = c.sPartsIDx" & _
                                 " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                           " LEFT JOIN SP_Type d" & _
                              " ON b.cPartType = d.cPartType" & _
                           " LEFT JOIN Labor e" & _
                              " ON a.sPartsIDx = e.sLaborIDx" & _
                                 " AND a.cLaborxxx = " & xeYes & _
                           " LEFT JOIN SP_Supersession f" & _
                              " ON a.sPartsIDx = f.sPartsIDx" & _
                     " ORDER BY a.nEntryNox"
      .Reference = "a.sTransNox|a.sTransNox"
      .VerifyEntryNo = True

      InitTransaction = .InitTransaction
Debug.Print .DetailQuery
      .BrowseQuery(0) = "SELECT a.sTransNox" & _
                           ", b.sCompnyNm" & _
                           ", dTransact" & _
                           ", nTranTotl" & _
                           ", sSalesInv" & _
                        " FROM " & .MasterTable & " a" & _
                              " LEFT JOIN Client_Master b" & _
                                 " ON a.sClientID = b.sClientID" & _
                        " WHERE sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                        " ORDER BY dTransact"
      .BrowseColumn(0) = "sSalesInv»sCompnyNm»dTransact»nTranTotl»sTransNox"
      .BrowseTitle(0) = "Sales Inv»Customer Name»Date»Total»Transaction No"
      .BrowseFormat(0) = "@»@»MMMM DD, YYYY»#,##0.00»@@@@-@@@@@@"

      .BrowseQuery(8) = "SELECT a.sCardIDxx" & _
                           ", a.sCardDesc" & _
                              ", b.sCompnyNm" & _
                           " FROM Discount_Card a" & _
                              ", Client_Master b" & _
                           " WHERE a.sClientID = b.sClientID" & _
                              " AND a.cRecdStat = " & xeRecStateActive
      .BrowseColumn(8) = "sCardIDxx»sCardDesc»sCompnyNm"
      .BrowseTitle(8) = "Card ID»Description»Company"
      .BrowseCriteria(8) = "a.sCardIDxx»a.sCardDesc»b.sCompnyNm"
      
      .BrowseQuery(14) = "SELECT a.sClientID" & _
                           ", a.sCompnyNm xFullName" & _
                           ", CONCAT(IF(IFNull(a.sHouseNox, '') = '', '', CONCAT(a.sHouseNox, ' ')), a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" & _
                        " FROM Client_Master a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & xeRecStateActive
      .BrowseColumn(14) = "sClientID»xFullName»xAddressx"
      .BrowseTitle(14) = "Client ID»Customer Name»Address"
      .BrowseCriteria(14) = "a.sClientID»a.sCompnyNm»CONCAT(IF(IFNull(a.sHouseNox, '') = '', '', CONCAT(a.sHouseNox, ' ')), a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode)"
      
      'Include SP_Supersession in the Query
      'Kalyptus - 2012.03.17 09:27 am
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sBarrCode" & _
                                 ", a.sDescript" & _
                                 ", b.nQtyOnHnd" & _
                                 ", a.nSelPrice" & _
                                 ", c.sModelNme" & _
                                 ", a.sPartsIDx" & _
                                 ", a.cPartType" & _
                                 ", a.nDelrDisc" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nReorderx" & _
                                 ", b.nLedgerNo" & _
                                 ", d.sSuperIDx" & _
                              " FROM Spareparts a" & _
                                    " LEFT JOIN SP_Model c" & _
                                       " ON a.sModelIDx = c.sModelIDx" & _
                                    " LEFT JOIN SP_Supersession d" & _
                                       " ON a.sPartsIDx = d.sPartsIDx" & _
                                 ", SP_Inventory b" & _
                              " WHERE a.sPartsIDx = b.sPartsIDx" & _
                                 " AND b.sBranchCd = " & strParm(p_sBranchCd)
      .BrowseDetailColumn(0) = "sBarrCode»sDescript»nQtyOnHnd»nSelPrice»sModelNme"
      .BrowseDetailTitle(0) = "Bar Code»Description»QOH»Price»Model"
      .BrowseDetailFormat(0) = "@»@»#0»#,##0.00»@"
      .BrowseDetailCriteria(0) = "a.sBarrCode»a.sDescript"
      
      p_nGCpnAmtx = p_oAppDrivr.getConfiguration(xeGiftCpn)
   End With
   
   ' initialize the client object that will handle the client info
   Set p_oClient = New clsClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      .BranchCode = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
      .ShowEntry = True
      .DisplayWarning = True
   End With
   
   p_sDiscSQL = "SELECT sCardIDxx" & _
                     ", sBrandIDx" & _
                     ", sCategrID" & _
                     ", nMinAmtxx" & _
                     ", nDiscRate" & _
                     ", nDiscAmtx" & _
                     ", nSCDiscxx" & _
                  " FROM Discount_Card_Detail" & _
                  " WHERE sDivisnID = " & strParm("MC")

   Debug.Print p_sDiscSQL
   Call initDisc
   
   Call initAdvance
   
   pbInitTran = InitTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "NewTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction
   p_oBaseClas.Branch = p_sBranchCd

   Set p_oReturn = Nothing

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "SaveTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
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
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)
   p_oBaseClas.Branch = Left(sTransNo, Len(p_oAppDrivr.BranchCode))
   Set p_oReturn = Nothing

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "DeleteTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   With p_oBaseClas
      If .Master("sReferNox") <> Empty And pbShowMsg = True Then
         MsgBox "Transaction is a created by a parent transaction!" & vbCrLf & _
                  "Delete the parent transaction for any changes!" & vbCrLf & vbCrLf & _
                  "Load other transaction then Try again!", vbInformation, "Notice"
         Exit Function
      End If

      If pbShowMsg Then
         'Enable cancel of transaction if Modification date is the same with the current date...
         'kalyptus - 2011-08-29
         If (DateDiff("h", .Master("dModified"), p_oAppDrivr.ServerDate) > 24) Then
            If Not p_oAppDrivr.UserLevel = xeEngineer Then
               MsgBox "Unable to delete " & .Master("sTransNox") & vbCrLf & _
                         "Transaction was created from the previous days...", vbCritical, "Warning"
               GoTo endProc
            End If
         Else
            If IFNull(p_oAppDrivr.getConfiguration("cRealTime", p_oAppDrivr.BranchCode), "0") <> "0" Then
               If .Master("cTranStat") <> xeStateOpen Then
                  MsgBox "Unable to delete " & .Master("sTransNox") & vbCrLf & _
                           "Only unconfirmed or unprinted transactions are allowable for this action!", vbCritical, "Warning"
                  GoTo endProc
               End If
            End If
         End If
      End If
      
      If pbInitTran Then DeleteTransaction = p_oBaseClas.DeleteTransaction
      
   End With

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lasMaster() As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String

   lsProcName = "SearchTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sSalesInv"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lasMaster = GetSplitedName(sSearch)
            lsValue = Trim(lasMaster(0)) & "»" & Trim(lasMaster(1))
            lsField = "b.sLastName" & "»" & "b.sFrstName"
            lsDescript = "b.sLastName" & "»" & "b.sFrstName"
         End If
      End If
   End If

   SearchTransaction = p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript)
   Set p_oReturn = Nothing

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function SearchMaster(ByVal Index As Variant, Optional Value As Variant = "") As Boolean
   If pbInitTran = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      SearchMaster = getCustomer(Value, True)
   End Select
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "CancelTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then
      With p_oBaseClas
         If .Master("sReferNox") <> Empty And pbShowMsg = True Then
            MsgBox "Transaction is created by a parent transaction!" & vbCrLf & _
                     "Cancel the parent transaction for any changes!" & vbCrLf & vbCrLf & _
                     "Load other transaction then Try again!", vbInformation, "Notice"
            Exit Function
         End If
            
         If pbShowMsg Then
               'Enable cancel of transaction if Modification date is the same with the current date...
               'kalyptus - 2011-08-29
            If (DateDiff("h", .Master("dModified"), p_oAppDrivr.ServerDate) > 24) Then
               If Not p_oAppDrivr.UserLevel = xeEngineer Then
                  MsgBox "Unable to cancel " & .Master("sTransNox") & vbCrLf & _
                            "Transaction was created from the other days other than today.."
                  GoTo endProc
               End If
            End If
         End If
            
         CancelTransaction = .CancelTransaction()
         
      End With
   End If

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function CloseTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "CloseTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then
      With p_oBaseClas
         If .EditMode = xeModeReady Then
            'Test first if this is not a reprint
            If .Master("cTranStat") = xeStateClosed Then
               MsgBox "Reprinting of transaction is not allowed!!!" & vbCrLf & vbCrLf & _
                        "Request not granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
            
            CloseTransaction = .CloseTransaction(.Master("sTransNox"))
         End If
      End With
   End If

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function


Function AddDetail() As Boolean
   With p_oBaseClas
      AddDetail = .AddDetail

      pnCtr = .ItemCount - 1
      .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      .Detail(pnCtr, "sBarrCode") = ""
      .Detail(pnCtr, "sDescript") = ""
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "nQtyOnHnd") = 0
      .Detail(pnCtr, "nUnitPrce") = 0#
      .Detail(pnCtr, "nDiscount") = 0#
      .Detail(pnCtr, "nAddDiscx") = 0#
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sPartsIDx") = ""
      .Detail(pnCtr, "cPartType") = 1
      .Detail(pnCtr, "nDelrDisc") = 0
      .Detail(pnCtr, "nSelPrice") = 0#
      .Detail(pnCtr, "nResvOrdr") = 0
      .Detail(pnCtr, "nBackOrdr") = 0
      .Detail(pnCtr, "nReorderx") = 0
      .Detail(pnCtr, "nLedgerNo") = 0
      .Detail(pnCtr, "cLaborxxx") = xeNo
      .Detail(pnCtr, "nOrderQty") = 0
   End With
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   DeleteDetail = p_oBaseClas.DeleteDetail(Index)
End Function

Function SearchDetail(ByVal Row As Long, _
                        ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsProcName As String

   lsProcName = "SearchDetail"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   SearchDetail = False

   If pbInitTran = False Then GoTo endProc

   Index = LCase(Index)
   Select Case Index
   Case 1, "sbarrcode"
      SearchDetail = getDetail(Row, Index, Value, True)
   Case 2, "sdescript"
      SearchDetail = getDetail(Row, Index, Value, True)
   End Select

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & Row _
                        & ", " & Index _
                        & ", " & Value & " )"
End Function

Function UpdateTransaction() As Boolean
   If pbShowMsg Then
      MsgBox "Update is no longer allowed!" & vbCrLf & _
            " If you need further info, please contact Guanzon SEG/SSG..." _
            , vbInformation, "Update Info"
      Exit Function
   End If
   
   With p_oBaseClas
      If .Master("sReferNox") <> Empty And pbShowMsg = True Then
         MsgBox "Transaction is a created by a parent transaction!" & vbCrLf & _
                  "Modify the parent transaction for any changes!" & vbCrLf & vbCrLf & _
                  "Load other transaction then Try again!", vbInformation, "Notice"
         Exit Function
      End If
      
'      'Enable cancel of transaction if Modification date is the same with the current date...
'      'kalyptus - 2011-08-29
'      If pbShowMsg Then
'         If Not (DateDiff("d", .Master("dModified"), p_oAppDrivr.ServerDate) = 0) Then
'            MsgBox "Unable to update " & .Master("sTransNox") & vbCrLf & _
'                      "Transaction was created from the other days other than today.."
'            Exit Function
'         End If
'      End If
      
      .EditMode = xeModeUpdate
   End With
   UpdateTransaction = True
   
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Function LoadByCondition(ByVal sCondition As String) As Boolean
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "LoadByCondition"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc

   If sCondition = Empty Then
      MsgBox "Invalid Condition Detected!", vbCritical, "Warning"
      GoTo endProc
   End If

   With p_oBaseClas
      lsSQL = "SELECT sTransNox" & _
               " FROM " & .MasterTable & _
               " WHERE " & sCondition
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      Debug.Print lsSQL
      If loRS.EOF = False Then
         LoadByCondition = .OpenTransaction(loRS("sTransNox"))
      End If
   End With
   
endProc:
   Set loRS = Nothing
   
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Function getDetail(ByVal lnRow As Integer, _
                        ByVal lnIndex As Integer, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsProcName As String

   lsProcName = "GetDetail"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

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

      lsDetail = p_oBaseClas.getDetail(lnRow, lsDetail)
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
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(2)
      .Detail(lnRow, "nUnitPrce") = lasDetail(3)
      
      'Part type: 0 = Replacement; 1 = Genuine
      .Detail(lnRow, "cPartType") = lasDetail(6) '9
      
'      'Set discount based on card type and part type
'      If .Master("cCardType") = "0" Then
'         If .Detail(lnRow, "cPartType") = 1 Then
'            .Detail(lnRow, "nDiscount") = p_oAppDrivr.getConfiguration("PRParts")
'         Else
'            .Detail(lnRow, "nDiscount") = p_oAppDrivr.getConfiguration("PRReplc")
'         End If
'      ElseIf .Master("cCardType") = "1" Then
'         If .Detail(lnRow, "cPartType") = 1 Then
'            .Detail(lnRow, "nDiscount") = p_oAppDrivr.getConfiguration("PLParts")
'         Else
'            .Detail(lnRow, "nDiscount") = p_oAppDrivr.getConfiguration("PLReplc")
'         End If
'      Else
'         .Detail(lnRow, "nDiscount") = 0#
'      End If
      
      .Detail(lnRow, "nDiscount") = 0#
      .Detail(lnRow, "nAddDiscx") = 0#
      .Detail(lnRow, "sPartsIDx") = lasDetail(5)
      .Detail(lnRow, "nDelrDisc") = lasDetail(7)
      .Detail(lnRow, "nResvOrdr") = lasDetail(8)
      .Detail(lnRow, "nBackOrdr") = lasDetail(9)
      .Detail(lnRow, "nReorderx") = IIf(lasDetail(10) = "", 0, lasDetail(10))
      .Detail(lnRow, "nLedgerNo") = lasDetail(11)
      .Detail(lnRow, "nSelPrice") = lasDetail(3)
      .Detail(lnRow, "cLaborxxx") = xeNo
      .Detail(lnRow, "nOrderQty") = 0
      .Detail(lnRow, "sSuperIDx") = lasDetail(12)
      
      ' XerSys - 2014-12-18
      '  Check if job order has discount card and override max discount
      If .Master("sCardIDxx") <> "" Then
         Call prcDiscount(lnRow)
      End If
   End With
   getDetail = True
   
endProc:
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)

   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nUnitPrce") = 0#
      .Detail(lnRow, "nDiscount") = 0#
      .Detail(lnRow, "nAddDiscx") = 0#
      .Detail(lnRow, "sPartsIDx") = ""
      .Detail(lnRow, "cPartType") = ""
      .Detail(lnRow, "nSelPrice") = 0#
      .Detail(lnRow, "nDelrDisc") = 0
      .Detail(lnRow, "nResvOrdr") = 0
      .Detail(lnRow, "nBackOrdr") = 0
      .Detail(lnRow, "nReorderx") = 0
      .Detail(lnRow, "nLedgerNo") = 0
      .Detail(lnRow, "cLaborxxx") = xeNo
      .Detail(lnRow, "nOrderQty") = 0
      .Detail(lnRow, "sSuperIDx") = ""
   End With
   getDetail = False
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lnIndex _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function saveSalesInvoice() As Boolean
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lnRow As Long

   lsProcName = "SaveSalesInvoice"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   saveSalesInvoice = False

   With p_oBaseClas
      If .EditMode = xeModeAddNew Or .MasterOriginal("nAmtPaidx") = 0# Then
         If Not IsNumeric(Left(Trim(.Master("sSalesInv")), 1)) Then
            saveSalesInvoice = True
            GoTo endProc
         End If
         
         lsSQL = "INSERT INTO SalesInvoice SET" & _
                     "  sTransNox = " & strParm(GetNextCode("SalesInvoice", "sTransNox", True, _
                           p_oAppDrivr.Connection, True, p_sBranchCd)) & _
                     ", sBranchCd = " & strParm(p_sBranchCd) & _
                     ", dTransact = " & dateParm(.Master("dTransact")) & _
                     ", sSalesInv = " & strParm(.Master("sSalesInv")) & _
                     ", sReferNox = " & strParm(.Master("sTransNox")) & _
                     ", sSourceCd = " & strParm(pxeSourceCd) & _
                     ", sClientID = " & strParm(.Master("sClientID")) & _
                     ", nTranTotl = " & .Master("nTranTotl") & _
                     ", nAmtPaidx = " & .Master("nAmtPaidx") & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate())
      Else
         If Not IsNumeric(Left(Trim(.Master("sSalesInv")), 1)) Then
            If Not IsNumeric(Left(Trim(.MasterOriginal("sSalesInv")), 1)) Then
               saveSalesInvoice = delSalesInvoice()
               GoTo endProc
            End If
            saveSalesInvoice = True
            GoTo endProc
         End If
         
         lsSQL = "UPDATE SalesInvoice SET" & _
                     "  dTransact = " & dateParm(.Master("dTransact")) & _
                     ", sSalesInv = " & strParm(.Master("sSalesInv")) & _
                     ", sClientID = " & strParm(.Master("sClientID")) & _
                     ", nTranTotl = " & .Master("nTranTotl") & _
                     ", nAmtPaidx = " & .Master("nAmtPaidx") & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                  " WHERE sReferNox = " & strParm(.Master("sTransNox")) & _
                     " AND sSalesInv = " & strParm(.MasterOriginal("sSalesInv"))
      End If
   
      lnRow = p_oAppDrivr.Execute(lsSQL, "SalesInvoice", p_sBranchCd)
      If lnRow = 0 Then
         MsgBox "Unable to Update Transaction Payment!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   saveSalesInvoice = True

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function getCustomer(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String

   lsProcName = "getCustomer"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   getCustomer = False
   
   With p_oBaseClas
      If lsValue <> "" Then
         If Trim(lsValue) = Trim(p_sFullName) Then GoTo endProc
         If p_oClient.SearchClient(lsValue, False) = False Then GoTo endProc
      Else
         GoTo endWithClear
      End If
      
      .Master("sClientID") = p_oClient.ClientID
      p_sFullName = p_oClient.FullName
      p_sAddressx = p_oClient.CompleteAddress
   End With
   
   getCustomer = True
   
endProc:
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   Exit Function
endWithClear:
   p_oBaseClas.Master("sClientID") = ""
   p_sFullName = ""
   p_sAddressx = ""
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getNextInvoice() As String
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lsProcName As String

   lsProcName = "getNextInvoice"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   lsSQL = "SELECT sSalesInv" & _
            " FROM SalesInvoice" & _
            " WHERE sBranchCd = " & strParm(p_sBranchCd) & _
            " ORDER BY sSalesInv * 1 DESC" & _
            " LIMIT 1"
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If loRS.EOF Then
      getNextInvoice = 1
   Else
      getNextInvoice = Format(CLng(loRS("sSalesInv")) + 1, String(Len(loRS("sSalesInv")), "0"))
   End If
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function isTotalDiscOK() As Boolean
   Dim lsProcName As String
   Dim lnItemPrice As Double, lnDiscPrice As Double
   Dim lnDiscount As Double, lnUDiscount As Double
   Dim lnADiscount As Double
   Dim lsAppvID As String, lsAppvName As String
   Dim lnAppvRights As Integer
   Dim lbGetApproval As Boolean
   Dim lnDiscPerc As Single
   Dim lnTotlPurc As Single
   Dim lsSQL As String
   
   Dim lsApprovedCD As String, lsApproveID As String, lsApproveName As String
                                                                                                          
   lsProcName = "isTotalDiscOk"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   lnAppvRights = p_oAppDrivr.UserLevel
   lsAppvID = p_oAppDrivr.UserID
   
   With p_oBaseClas
      lnUDiscount = 0
      lnADiscount = 0
      lnDiscount = 0
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sPartsIDx") = "" Then Exit For
         
         If .Detail(pnCtr, "cLaborxxx") = xeNo Then
            lnItemPrice = .Detail(pnCtr, "nSelPrice") * .Detail(pnCtr, "nQuantity")
            
'            Select Case .Detail(pnCtr, "sPartsIDx")
'            Case "W208003283", "W208004453", "W208004505", "W208003282", "W208003813"
'               If .Master("dTransact") >= CDate("2009/04/23") And _
'                     .Master("dTransact") <= CDate("2009/04/24") And _
'                     p_sBranchCd = "09" Then
'                  lnADiscount = lnADiscount + Round(lnItemPrice * 0.5, 2)
'               Else
'                  lnADiscount = lnADiscount + Round(lnItemPrice * pxeDiscount, 2)
'               End If
'            Case Else
'               lnADiscount = lnADiscount + Round(lnItemPrice * pxeDiscount, 2)
'            End Select
            
            'Replace the above with the new implementation
            'Kalyptus - 2009.06.19
'            '+++++++++++++++++++
'            lnDiscPerc = IIf(.Detail(pnCtr, "cPartType") = "1", 0.05, 0.25)
'            lnADiscount = lnADiscount + Round(lnItemPrice * lnDiscPerc, 2)
'            '+++++++++++++++++++
            
            'Kalyptus - 2010.09.22
            'Set discount based on card type and part type
            'Only parts are check in this area
            If .Master("cCardType") = "0" Then
               If .Detail(pnCtr, "cPartType") = 1 Then
                  lnDiscPerc = p_oAppDrivr.getConfiguration("PRParts") / 100
               Else
                  lnDiscPerc = p_oAppDrivr.getConfiguration("PRReplc") / 100
               End If
            ElseIf .Master("cCardType") = "1" Then
               If .Detail(pnCtr, "cPartType") = 1 Then
                  lnDiscPerc = p_oAppDrivr.getConfiguration("PLParts") / 100
               Else
                  lnDiscPerc = p_oAppDrivr.getConfiguration("PLReplc") / 100
               End If
            Else
'               'Temporarily set the check the discount accdg to the old way
'               lnDiscPerc = IIf(.Detail(pnCtr, "cPartType") = "1", 0.05, 0.25)
               ' Disable the code above and allow manager to approve discount based on
               '  spareparts department recommendation
               lnDiscPerc = .Detail(pnCtr, "nDelrDisc") / 100
            End If
            lnADiscount = lnUDiscount + Round(lnItemPrice * lnDiscPerc, 2)
            
            lnUDiscount = lnUDiscount + _
                     Round(lnItemPrice * .Detail(pnCtr, "nDelrDisc") / 100, 2)
            lnDiscPrice = .Detail(pnCtr, "nUnitPrce") * .Detail(pnCtr, "nQuantity")
            lnDiscPrice = (lnItemPrice - lnDiscPrice) + Round(lnDiscPrice * .Detail(pnCtr, "nDiscount") / 100, 2)
            lnDiscount = lnDiscount + lnDiscPrice
            
            'Compute gross sales
            lnTotlPurc = lnTotlPurc + (lnItemPrice)
         End If
      Next
      
'      'Inform the user that a purchase of
'      'less that 200 is not entitled for discount
'      If lnTotlPurc < 200 Then
'
'      End If
            
            
      '2014.05.06
      'TODO: prevents discounts from being filtered...
      If lnDiscount <= 0# Or _
         lnADiscount >= lnDiscount Then
         isTotalDiscOK = True
         GoTo endProc
      End If
      
      lbGetApproval = False
      
      If p_oAppDrivr.UserLevel <> xeEngineer Then
         lbGetApproval = True
      End If
      
'      If lbGetApproval Then
'         If GetApproval(p_oAppDrivr, lnAppvRights, lsAppvID, lsAppvName, p_oAppDrivr.MenuName) = False Then GoTo endProc
'
'         ' Remove the following code to disallow branches to give discount greater than
'         ' 5% as imposed by management
''         If lnAppvRights = xeSupervisor Or lnAppvRights = xeManager Then
''            If lnUDiscount < lnDiscount Then
''               MsgBox "Discount is not within the range of the approving officer!!!" & vbCrLf & _
''                  "Request can not be granted!!!", vbCritical, "Warning"
''               GoTo endProc
''            End If
''         ElseIf Not (lnAppvRights = xeSysOwner Or lnAppvRights = xeEngineer) Then
''            MsgBox "User is not allowed to approved discount!!!" & vbCrLf & _
''                  "Request can not be granted!!!", vbCritical, "Warning"
''            GoTo endProc
'
'         If Not (lnAppvRights = xeAudit Or lnAppvRights = xeSysOwner Or lnAppvRights = xeEngineer) Then
'            MsgBox "User is not allowed to approved discount!!!" & vbCrLf & _
'                  "Request can not be granted!!!", vbCritical, "Warning"
'            GoTo endProc
'         End If
'      End If

      'kalyptus - 2014.04.12 02:12pm
      'if discount is greater than supposed discount then asked for approval code
      If lbGetApproval Then
         If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
            MsgBox "Discount given is greater than approved discount!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
            GoTo endProc
         Else
            If isValidApproveCode( _
               p_oSCA_Const.pxePartsDiscount, _
               p_oAppDrivr.BranchCode, _
               Mid(lsApprovedCD, 4, 1), _
               .Master("dTransact"), _
               .Master("sSalesInv"), _
               lsApprovedCD) Then
               
               .Master("sApproved") = lsApproveID
               
               lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                          p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
               lsSQL = "INSERT INTO xxxSCA_Usage" & _
                      " SET sTransNox = " & strParm(lsSQL) & _
                         ", sApprCode = " & strParm(lsApprovedCD) & _
                         ", sApproved = " & strParm(lsApproveID) & _
                         ", sSystemCD = " & strParm(p_oSCA_Const.pxePartsDiscount) & _
                         ", sSourceNo = " & strParm(.Master("sTransNox")) & _
                         ", sSourceCD = " & strParm(pxeSPSales) & _
                         ", sModified = " & strParm(Encrypt(p_oAppDrivr.UserID)) & _
                         ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
               
               p_oAppDrivr.Execute lsSQL, "xxxSCA_Usage"
            
            Else
               MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                  "Verify entry then try again!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If

   End With
   
   isTotalDiscOK = True

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function delSalesInvoice() As Boolean
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lnRow As Long

   lsProcName = "delSalesInvoice"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   With p_oBaseClas
      If .Master("sSalesInv") = Empty Or _
            .Master("nAmtPaidx") = 0 Then
         delSalesInvoice = True
         GoTo endProc
      End If
      
      lsSQL = "DELETE FROM SalesInvoice " & _
               " WHERE sReferNox = " & strParm(.Master("sTransNox")) & _
                  " AND sSalesInv = " & strParm(.Master("sSalesInv"))
   
      lnRow = p_oAppDrivr.Execute(lsSQL, "SalesInvoice", p_sBranchCd)
      If lnRow = 0 Then
         MsgBox "Unable to Delete Transaction Payment!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   delSalesInvoice = True

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function saveSPTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsProcName As String
   Dim lnCtr As Integer
   
   lsProcName = "saveSPTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc
      
      pnCtr = 0
      For lnCtr = 0 To p_oBaseClas.ItemCount - 1
         'Is this an ordered item
         If p_oBaseClas.Detail(lnCtr, "nOrderQty") > 0 Then
            p_oRSIssd.Filter = "nEntryNox = " & lnCtr
            If Not p_oRSIssd.EOF Then p_oRSIssd.MoveFirst
            
            .Detail(lnCtr, "nResvOrdr") = 0
            Do Until p_oRSIssd.EOF
               .Detail(pnCtr, "sPartsIDx") = p_oRSIssd("sPartsIDx")
               .Detail(pnCtr, "nQuantity") = p_oBaseClas.Detail(lnCtr, "nQuantity")
               .Detail(pnCtr, "nResvOrdr") = p_oRSIssd("nQtyOrder")
               
               If p_oRSIssd("sPartsIDx") <> p_oRSIssd("xPartsIDx") Then
                  .Detail(pnCtr, "sReplacID") = p_oRSIssd("xPartsIDx")
               End If
                                             
               pnCtr = pnCtr + 1
               p_oRSIssd.MoveNext
            Loop
         Else
            .Detail(pnCtr, "sPartsIDx") = p_oBaseClas.Detail(lnCtr, "sPartsIDx")
            .Detail(pnCtr, "nQuantity") = p_oBaseClas.Detail(lnCtr, "nQuantity")
'            .Detail(pnCtr, "nResvOrdr") = p_oBaseClas.Detail(lnCtr, "nOrderQty")
            pnCtr = pnCtr + 1
         End If
      Next
      
      ' if spareparts is not successfully updated, cancel saving.
      If .Sales(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("dTransact"), _
            p_oBaseClas.EditMode) = False Then
         GoTo endProc
      End If
   End With
   
   saveSPTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function delSPTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsProcName As String
   Dim lnCtr As Integer
   
   lsProcName = "delSPTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc
      
      pnCtr = 0
      
      For lnCtr = 0 To .ItemCount - 1
         
         If p_oBaseClas.Detail(lnCtr, "nOrderQty") > 0 Then
            p_oRSIssd.Filter = "nEntryNox = " & lnCtr
            If Not p_oRSIssd.EOF Then p_oRSIssd.MoveFirst
            
            .Detail(lnCtr, "nResvOrdr") = 0
            Do Until p_oRSIssd.EOF
               .Detail(pnCtr, "sPartsIDx") = p_oRSIssd("sPartsIDx")
               .Detail(pnCtr, "nQuantity") = p_oRSIssd("nQtyOrder")
               If p_oRSIssd("sPartsIDx") <> p_oRSIssd("xPartsIDx") Then
                  .Detail(pnCtr, "sReplacID") = p_oRSIssd("xPartsIDx")
                  .Detail(pnCtr, "nResvOrdr") = p_oRSIssd("nQtyOrder")
               End If
                                             
               pnCtr = pnCtr + 1
               p_oRSIssd.MoveNext
            Loop
         Else
            .Detail(pnCtr, "sPartsIDx") = p_oBaseClas.Detail(lnCtr, "sPartsIDx")
            .Detail(pnCtr, "nQuantity") = p_oBaseClas.Detail(lnCtr, "nQuantity")
                     
            pnCtr = pnCtr + 1
         End If
      Next
      
      ' if a serial is not successfully updated, cancel saving.
      If .Sales(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("dTransact"), _
            xeModeDelete) = False Then
         GoTo endProc
      End If
   End With
   
   delSPTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function getOrderQty(ByVal lsPartsIDx As String, ByVal lnQuantity As Long) As Long
   Dim lsProcName As String
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lasOrder() As String
   Dim lnCtr As Integer
   Dim lnOrder As Long

   lsProcName = "getOrderQty"
   'On Error GoTo errProc

   lsSQL = ""
   p_oAdvance.MoveFirst
   Do Until p_oAdvance.EOF
      If p_oAdvance("cWithOrdr") = xeYes Then
         lsSQL = lsSQL & ", " & strParm(p_oAdvance("sTransNox"))
      End If

      p_oAdvance.MoveNext
   Loop
   
   ' this means that advance payment has no spareparts order detail
   If lsSQL = "" Then GoTo endProc

   lsSQL = "SELECT SUM(nQuantity - (nIssuedxx + nCanceled)) nOrderQty" & _
            " FROM SP_Retail_Order_Detail" & _
            " WHERE sTransNox IN (" & Mid(lsSQL, 2) & ")" & _
               " AND sPartsIDx = " & strParm(lsPartsIDx)
   Debug.Print lsSQL
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

   If Not loRS.EOF Then
      If lnQuantity > loRS("nOrderQty") Then
         getOrderQty = loRS("nOrderQty")
      Else
         getOrderQty = lnQuantity
      End If
   End If

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lsPartsIDx & " )"
End Function

Private Function saveCheck() As Boolean
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsTransNo As String
   Dim lnCtr As Integer
   
   lsProcName = "SaveCheck"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If p_oBaseClas.EditMode = xeModeAddNew Then
      If p_xChecks.CheckNox = Empty Then
         MsgBox "Invalid Check No Detected!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsSQL = "INSERT INTO Checks_Received SET" & _
                  "  sTransNox = " & _
                     strParm(GetNextCode("Checks_Received", "sTransNox", True, _
                        p_oAppDrivr.Connection, True, p_sBranchCd)) & _
                  ", sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                  ", sSourceCd = " & strParm(pxeSourceCd) & _
                  ", sCheckNox = " & strParm(p_xChecks.CheckNox) & _
                  ", sAcctNoxx = " & strParm(p_xChecks.AcctNoxx) & _
                  ", sBankIDxx = " & strParm(p_xChecks.BankIDxx) & _
                  ", dCheckDte = " & dateParm(p_xChecks.CheckDte) & _
                  ", nAmountxx = " & p_xChecks.Amountxx & _
                  ", nClearing = " & 3 & _
                  ", cDepositd = " & strParm(xeNo) & _
                  ", cTranStat = " & strParm(xeStateOpen) & _
                  ", sModified = " & strParm(Encrypt(p_oAppDrivr.UserID)) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
      
      If p_oAppDrivr.Execute(lsSQL, "Checks_Received", p_sBranchCd) <= 0 Then
         MsgBox "Unable to Update Check Info." & vbCrLf & _
               "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   Else
      If p_xChecks.Depositd Then
         MsgBox "Deposited Check is Not Allowed to be Modified!" & vbCrLf & vbCrLf & _
               "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If p_xChecks.Amountxx = 0 Then
         If p_xChecks.TransNox <> Empty Then
            lsSQL = "DELETE FROM Checks_Received" & _
                    " WHERE sTransNox = " & strParm(p_xChecks.TransNox)
             
            If p_oAppDrivr.Execute(lsSQL, "Checks_Received", p_sBranchCd) <= 0 Then
               MsgBox "Unable to Delete Check Info." & vbCrLf & _
                     "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      Else
         If p_xChecks.CheckNox = Empty Then
            MsgBox "Invalid Check No Detected!", vbCritical, "Warning"
            GoTo endProc
         End If
         
         lsSQL = "UPDATE Checks_Received SET" & _
                     "  sCheckNox = " & strParm(p_xChecks.CheckNox) & _
                     ", sAcctNoxx = " & strParm(p_xChecks.AcctNoxx) & _
                     ", sBankIDxx = " & strParm(p_xChecks.BankIDxx) & _
                     ", dCheckDte = " & dateParm(p_xChecks.CheckDte) & _
                     ", nAmountxx = " & p_xChecks.Amountxx & _
                  " WHERE sTransNox = " & p_xChecks.TransNox
                  
         If p_oAppDrivr.Execute(lsSQL, "Checks_Received", p_sBranchCd) <= 0 Then
            MsgBox "Unable to Update Check Info." & vbCrLf & _
                  "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End If
   
   saveCheck = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName
End Function

Private Function initAdvance() As Boolean
   Dim lsProcName As String

   lsProcName = "initAdvance"
   'On Error GoTo errProc

   Set p_oAdvance = New Recordset
   With p_oAdvance
      .Fields.Append "sTransNox", adVarChar, 12
      .Fields.Append "sReferNox", adVarChar, 8
      .Fields.Append "dTransact", adDate
      .Fields.Append "nTranAmtx", adCurrency, 12
      .Fields.Append "nCredtAmt", adCurrency, 12
      .Fields.Append "cTranStat", adChar, 1
      .Fields.Append "sCheckTrn", adChar, 12
      .Fields.Append "cDepositd", adChar, 1
      .Fields.Append "cChckStat", adChar, 1
      .Fields.Append "sSourceCd", adVarChar, 4
      .Fields.Append "cWithOrdr", adChar, 1
      .Open
   End With

   initAdvance = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName
End Function

Private Function getAdvancePayment() As Boolean
   Dim loAdvPaym As frmAdvancePayment
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "getAdvancePayment"
   'On Error GoTo errProc
   
   With p_oBaseClas
      If .Master("sClientID") = Empty Then GoTo endWithWarning
      
      ' XerSys 2014-04-07
      '  Include check advance payment
      ' XerSys 2014-11-07
      '  Include cash advance crediting
      lsSQL = "SELECT a.sTransNox" & _
                  ", b.sPRNoxxxx sReferNox" & _
                  ", a.dTransact" & _
                  ", b.nTranAmtx - (a.nCredtAmt + IFNull(a.nCredtChk,0)) nTranAmtx" & _
                  ", a.nCredtAmt + IFNull(a.nCredtChk,0) nCredtAmt" & _
                  ", b.cTranStat" & _
                  ", IFNULL(c.sTransNox, '') sCheckTrn" & _
                  ", IFNULL(c.cDepositd, '1') cDepositd" & _
                  ", IFNULL(c.cChckStat, '2') cChckStat" & _
                  ", " & strParm(xeSrcPReceipt) & " sSourceCd" & _
               " FROM SP_Retail_Order_Master a" & _
                  ", Provisionary_Receipt_Master b" & _
                     " LEFT JOIN Checks_Received c" & _
                        " ON b.sTransNox = c.sReferNox" & _
                           " AND c.sSourceCd = " & strParm(xeSrcPReceipt) & _
               " WHERE a.sTransNox = b.sReferNox" & _
                  " AND b.sSourceCd = " & strParm(pxeSPRetailOrder) & _
                  " AND a.sClientID = " & strParm(.Master("sClientID")) & _
                  " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                  " AND (a.nCredtAmt + IFNULL(a.nCredtChk, 0) < b.nTranAmtx)"
      lsSQL = lsSQL & _
               " UNION SELECT a.sTransNox" & _
                  ", b.sORNoxxxx sReferNox" & _
                  ", a.dTransact" & _
                  ", b.nTranAmtx - a.nCredtAmt nTranAmtx" & _
                  ", a.nCredtAmt" & _
                  ", b.cTranStat" & _
                  ", IFNULL(c.sTransNox, '') sCheckTrn" & _
                  ", IFNULL(c.cDepositd, '1') cDepositd" & _
                  ", IFNULL(c.cChckStat, '2') cChckStat" & _
                  ", " & strParm(xeSrcOReceipt) & " sSourceCd" & _
               " FROM SP_Retail_Order_Master a" & _
                  " LEFT JOIN Receipt_Master b" & _
                     " LEFT JOIN Checks_Received c" & _
                        " ON b.sTransNox = c.sReferNox" & _
                           " AND c.sSourceCd = " & strParm(xeSrcOReceipt) & _
                     " ON a.sTransNox = b.sReferNox" & _
                        " AND b.sSourceCd = " & strParm(pxeSPRetailOrder) & _
               " WHERE a.sClientID = " & strParm(.Master("sClientID")) & _
                  " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                  " AND a.nCredtAmt < b.nTranAmtx"
      lsSQL = lsSQL & _
               " UNION SELECT a.sTransNox" & _
                  ", a.sPRNoxxxx sReferNox" & _
                  ", a.dTransact" & _
                  ", a.nTranAmtx" & _
                  ", 0 nCredtAmt" & _
                  ", a.cTranStat" & _
                  ", '' sCheckTrn" & _
                  ", '1' cDepositd" & _
                  ", '2' cChckStat" & _
                  ", " & strParm(xeSrcSPCustOrder) & " sSourceCd" & _
               " FROM Provisionary_Receipt_Master a" & _
               " WHERE a.sClientID = " & strParm(.Master("sClientID")) & _
                  " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                  " AND a.cTranType = " & strParm(0) & _
                  " AND a.sReferNox = ''" & _
                  " AND a.cTranStat = " & strParm(xeStateOpen) & _
               " ORDER BY dTransact" & _
               ", sTransNox"
               Debug.Print lsSQL
   End With
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   Call initAdvance
   p_nAdvPaymx = 0#
   If loRS.EOF Then GoTo endWithWarning
   
   With p_oAdvance
      Do Until loRS.EOF
         .AddNew
         
         .Fields("sTransNox") = loRS("sTransNox")
         .Fields("sReferNox") = loRS("sReferNox")
         .Fields("dTransact") = loRS("dTransact")
         .Fields("nTranAmtx") = loRS("nTranAmtx")
         .Fields("nCredtAmt") = 0
         .Fields("cTranStat") = xeStateOpen
         .Fields("sCheckTrn") = loRS("sCheckTrn")
         .Fields("cDepositd") = loRS("cDepositd")
         .Fields("cChckStat") = loRS("cChckStat")
         .Fields("sSourceCd") = loRS("sSourceCd")
         .Fields("cWithOrdr") = IIf(loRS("sSourceCd") = pxeSPRetailOrder, xeNo, xeYes)
         
         loRS.MoveNext
      Loop
      
      If .RecordCount = 1 Then
         .Fields("cTranStat") = xeStateClosed
         p_nAdvPaymx = IIf(.Fields("nTranAmtx") > p_oBaseClas.Master("nTranTotl"), p_oBaseClas.Master("nTranTotl"), .Fields("nTranAmtx"))
         p_sAdvRefer = .Fields("sTransNox")
         
         getAdvancePayment = True
         GoTo endProc
      End If

      Set loAdvPaym = New frmAdvancePayment
      Set loAdvPaym.AppDriver = p_oAppDrivr
      Set loAdvPaym.Source = p_oAdvance

      loAdvPaym.ClientName = Master("xFullName")
      loAdvPaym.ClientAddress = Master("xAddressx")
      loAdvPaym.Show 1

      If loAdvPaym.Cancelled Then
         MsgBox "Advance Payment Confirmation was Cancelled!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      p_nAdvPaymx = 0#
      p_sAdvRefer = ""
      .MoveFirst
      Do Until .EOF
         If .Fields("cTranStat") = xeStateClosed Then
            If p_nAdvPaymx + .Fields("nTranAmtx") > Master("nTranTotl") Then
               p_nAdvPaymx = Master("nTranTotl")
            Else
               p_nAdvPaymx = p_nAdvPaymx + .Fields("nTranAmtx")
            End If
            
            p_sAdvRefer = p_sAdvRefer & "»" & p_oAdvance("sTransNox")
         End If
         .MoveNext
      Loop
      
      p_sAdvRefer = Mid(p_sAdvRefer, 2)
   End With

   getAdvancePayment = True

endProc:
   Set loRS = Nothing
   Exit Function
endWithWarning:
   MsgBox "No Record Found for this Customer!" & _
            vbCrLf & " Can Not Process Advance Payment", vbCritical, "Warning"
   GoTo endProc
errProc:
   ShowError lsProcName
End Function

Private Function updateAdvPayment() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim loSPOrder As clsRetailOrder
   Dim lsTransNox As String
   Dim lnCashAmtx As Double, lnCheckAmt As Double
   Dim lcWithOrdr As String, lcTranStat As String
   
   lsProcName = "updateAdvPayment"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   ' XerSys 2014-04-08
   '  Modify the approach of updating the affected retail order
   Set loSPOrder = New clsRetailOrder
   With loSPOrder
      Set .AppDriver = p_oAppDrivr
      If Not .InitTransaction Then GoTo endProc
   End With
   
   With p_oAdvance
      .MoveFirst
      
      lsTransNox = .Fields("sTransNox")
      lcTranStat = .Fields("cTranStat")
      lcWithOrdr = .Fields("cWithOrdr")
      lnCashAmtx = 0: lnCheckAmt = 0
      Do Until .EOF
         If lsTransNox <> .Fields("sTransNox") Then
            If lcWithOrdr = xeNo Then
               If lcTranStat = xeStateClosed Then
                  lsSQL = "UPDATE Provisionary_Receipt_Master" & _
                           " SET cTranStat = " & strParm(xeStateClosed) & _
                              ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                           " WHERE sTransNox = " & strParm(lsTransNox)
                  If p_oAppDrivr.Execute(lsSQL, "Provisionary_Receipt_Master") <= 0 Then
                     MsgBox "Unable to Update Advance Payment Info!", vbCritical, "Warning"
                     GoTo endProc
                  End If
               End If
            Else
               If lnCashAmtx + lnCheckAmt > 0 Then
                  If Not loSPOrder.CreditAdvPayment(lsTransNox, lnCashAmtx, lnCheckAmt, Master("dTransact")) Then
                     GoTo endProc
                  End If
               End If
            End If
            
            lsTransNox = .Fields("sTransNox")
            lcWithOrdr = .Fields("cWithOrdr")
            lcTranStat = .Fields("cTranStat")
            lnCashAmtx = 0: lnCheckAmt = 0
         End If
         
         If .Fields("sSourceCd") = xeSrcOReceipt Then
            lnCashAmtx = .Fields("nCredtAmt")
         Else
            lnCheckAmt = .Fields("nCredtAmt")
         End If
         
         .MoveNext
      Loop
      
      If lcWithOrdr = xeNo Then
         If lcTranStat = xeStateClosed Then
            lsSQL = "UPDATE Provisionary_Receipt_Master" & _
                     " SET cTranStat = " & strParm(xeStateClosed) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                     " WHERE sTransNox = " & strParm(lsTransNox)
            If p_oAppDrivr.Execute(lsSQL, "Provisionary_Receipt_Master") <= 0 Then
               MsgBox "Unable to Update Advance Payment Info!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      Else
         If lnCashAmtx + lnCheckAmt > 0 Then
            If Not loSPOrder.CreditAdvPayment(lsTransNox, lnCashAmtx, lnCheckAmt, Master("dTransact")) Then
               GoTo endProc
            End If
         End If
      End If
   End With
   
   ' XerSys - 2014-04-07
'   Disable the following codes
'   Set loPReceipt = New clsPRReceipt
'   With loPReceipt
'      Set .AppDriver = p_oAppDrivr
'      .Branch = p_sBranchCd
'      .Parent = pxeMODULENAME
'
'      If .InitTransaction = False Then GoTo endProc
'
'      ' only advance payment with out an order must be updated here
'      p_oAdvance.MoveFirst
'      Do Until p_oAdvance.EOF
'         If p_oAdvance("cWithOrdr") = xeNo Then
'            If .OpenTransaction(p_oAdvance("sReferNox")) = False Then GoTo endProc
'            Call .UpdateTransaction
'            .Master("cTranStat") = xeStateClosed
'            .Master("sReferNox") = p_oBaseClas.Master("sTransNox")
'            .Master("sSourceCd") = "SPSl"
'
'            If .SaveTransaction() = False Then GoTo endProc
'         End If
'         p_oAdvance.MoveNext
'      Loop
'   End With

   updateAdvPayment = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Private Function undoAdvPayment() As Boolean
   Dim lsProcName As String
'   Dim loPReceipt As clsPRReceipt
   Dim loSPOrder As clsRetailOrder
   Dim lsTransNox As String
   Dim lnCashAmtx As Double, lnCheckAmt As Double
   Dim lcWithOrdr As String
   Dim lsSQL As String

   lsProcName = "undoAdvPayment"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   ' XerSys 2014-04-08
   '  Modify the approach of resetting the affected retail order
   Set loSPOrder = New clsRetailOrder
   With loSPOrder
      Set .AppDriver = p_oAppDrivr
      If Not .InitTransaction Then GoTo endProc
   End With
   
   If p_sAdvRefer <> "" Then
      With p_oAdvance
         .MoveFirst
         lsTransNox = .Fields("sTransNox")
         lcWithOrdr = .Fields("cWithOrdr")
         lnCashAmtx = 0: lnCheckAmt = 0
         Do Until .EOF
            If lsTransNox <> .Fields("sTransNox") Then
               If lcWithOrdr = xeNo Then
                  lsSQL = "UPDATE Provisionary_Receipt_Master" & _
                           " SET cTranStat = " & strParm(xeStateOpen) & _
                              ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                           " WHERE sTransNox = " & strParm(.Fields("sTransNox"))
                  If p_oAppDrivr.Execute(lsSQL, "Provisionary_Receipt_Master") <= 0 Then
                     MsgBox "Unable to Update Advance Payment Info!", vbCritical, "Warning"
                     GoTo endProc
                  End If
               Else
                  If lnCashAmtx + lnCheckAmt > 0 Then
                     If Not loSPOrder.DebitAdvPayment(lsTransNox, lnCashAmtx, lnCheckAmt) Then
                        GoTo endProc
                     End If
                  End If
               End If
               
               lsTransNox = .Fields("sTransNox")
               lcWithOrdr = .Fields("cWithOrdr")
               lnCashAmtx = 0: lnCheckAmt = 0
            End If
            
            If .Fields("sSourceCd") = xeSrcOReceipt Then
               lnCashAmtx = .Fields("nCredtAmt")
            Else
               lnCheckAmt = .Fields("nCredtAmt")
            End If
            
            .MoveNext
         Loop
         
         If lnCashAmtx + lnCheckAmt > 0 Then
            If Not loSPOrder.DebitAdvPayment(lsTransNox, lnCashAmtx, lnCheckAmt) Then
               GoTo endProc
            End If
         End If
      End With
   End If
   
   ' XerSys 2014-04-08
   '  Disable the following codes
'   Set loPReceipt = New clsPRReceipt
'   With loPReceipt
'      Set .AppDriver = p_oAppDrivr
'      .Branch = p_sBranchCd
'      .Parent = pxeMODULENAME
'
'      If .InitTransaction = False Then GoTo endProc
'
'      ' only advance payment with out an order must be updated here
'      p_oAdvance.MoveFirst
'      Do Until p_oAdvance.EOF
'         If p_oAdvance("cWithOrdr") = xeNo Then
'            If .OpenTransaction(p_oAdvance("sReferNox")) = False Then GoTo endProc
'            Call .UpdateTransaction
'            .Master("cTranStat") = xeStateOpen
'            .Master("sReferNox") = ""
'            .Master("sSourceCd") = ""
'
'            If .SaveTransaction() = False Then GoTo endProc
'         End If
'         p_oAdvance.MoveNext
'      Loop
'   End With

   undoAdvPayment = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Private Function issueSPOrder() As Boolean
   Dim lsProcName As String
   Dim loRS As Recordset
   Dim loRSTemp As Recordset
   Dim lsSQL As String, lsCondition As String
   Dim lsPartsIDx As String
   Dim lnQuantity As Integer
   Dim lnIssueQty As Integer, lnIssueTot As Double
   Dim lnCreditAmt As Double, lnAdvPaym As Double
   Dim lnCtr As Integer, lnRow As Integer
   Dim lbCreditAll As Boolean

   lsProcName = "issueSPOrder"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   ' retrieve the existing record
   With p_oAdvance
      .MoveFirst
      lsCondition = ""
      Do Until .EOF
         If .Fields("cTranStat") = xeStateClosed Then
            If .Fields("cWithOrdr") = xeYes Then
               lsCondition = lsCondition & ", " & strParm(.Fields("sTransNox"))
            End If
         End If
         
         .MoveNext
      Loop
   End With
   ' this means that advance payment has no spareparts order detail
   If lsCondition = "" Then
      issueSPOrder = True
      GoTo endProc
   End If
   
   
   'Load the Order of customers based on selected Advance Payments of Customer
   'group ids are set to x as default so that they will be different from
   'the default group id of the detail
   lsSQL = "SELECT a.sTransNox" & _
                ", c.sPartsIDx" & _
                ", a.nQuantity - (a.nIssuedxx + a.nCanceled) nQuantity" & _
                ", a.sPartsIDx xPartsIDx" & _
                ", IFNull(b.sSuperIDx, 'x') sSuperIDx" & _
            " FROM SP_Retail_Order_Detail a" & _
                " LEFT JOIN SP_Supersession b ON a.sPartsIDx = b.sPartsIDx" & _
                " LEFT JOIN SP_Supersession c ON b.sSuperIDx = c.sSuperIDx" & _
            " WHERE a.sTransNox IN (" & Mid(lsCondition, 3) & ")" & _
               " AND a.nQuantity - (a.nIssuedxx + a.nCanceled) > 0" & _
            " ORDER BY a.sPartsIDx" & _
               ", a.sTransNox"

   Debug.Print lsSQL
   Set loRSTemp = New Recordset
   
   loRSTemp.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText

   If loRSTemp.EOF Then GoTo procAdvPaym
   
   With p_oBaseClas
      lsPartsIDx = ""
      lnQuantity = 0
      lnCtr = 0
      lnRow = 0
      lnAdvPaym = 0
      
      Call initTempIssue
            
      For lnCtr = 0 To .ItemCount - 1
         'Search from list of orders the existence of this pos sales
         'instead of using find I used filter
         'Results are stored in a temporary table to solve cases of multiple
         'usage of barcode/superseded no. for
         'kalyptus - 2012.03.17
                  
         'From this area results are stored in a temporary recordset
         '++++++++++++++++++++++++++
         loRSTemp.Filter = "sPartsIDx = " & strParm(.Detail(lnCtr, "sPartsIDx"))
         If Not loRSTemp.EOF Then loRSTemp.MoveFirst
         
         Set loRS = initTempRecd
         lnQuantity = .Detail(lnCtr, "nQuantity")
         If Not loRSTemp.EOF Then
            Do Until loRSTemp.EOF
               loRS.AddNew
               loRS("sTransNox") = loRSTemp("sTransNox")
               loRS("sPartsIDx") = loRSTemp("sPartsIDx")
               loRS("nQuantity") = loRSTemp("nQuantity")
               loRS("xPartsIDx") = loRSTemp("xPartsIDx")
               loRS("sSuperIDx") = loRSTemp("sSuperIDx")
               loRSTemp.MoveNext
            Loop
         End If
                  
         '++++++++++++++++++++++++++
         'Transfer of record is up to this area
         
         If loRS.EOF Then
            .Detail(lnCtr, "nOrderQty") = 0
         Else
            lnIssueTot = 0
            'Since we are using filter testing for equality of parts id is
            'no longer necessary
            'kalyptus - 2012.03.17 03.17 pm
            Do Until loRS.EOF
               If loRS("nQuantity") > lnQuantity Then
                  lnIssueQty = lnQuantity
                  lnIssueTot = lnIssueTot + lnIssueQty
                  lnQuantity = 0
               Else
                  lnIssueQty = loRS("nQuantity")
                  lnIssueTot = lnIssueTot + loRS("nQuantity")
                  lnQuantity = lnQuantity - lnIssueQty
               End If

               lnCreditAmt = Round(lnIssueQty * _
                              (.Detail(lnCtr, "nUnitPrce") * _
                                 (100 - .Detail(lnCtr, "nDiscount")) / 100 _
                                 - .Detail(lnCtr, "nAddDiscx")), 2)

               p_oAdvance.MoveFirst
               Call p_oAdvance.Find("sTransNox = " & strParm(loRS("sTransNox")), 0, adSearchForward, 1)
               If Not p_oAdvance.EOF Then
                  If p_oAdvance("nTranAmtx") > p_oAdvance("nCredtAmt") + lnCreditAmt Then
                     p_oAdvance("nCredtAmt") = p_oAdvance("nCredtAmt") + lnCreditAmt
                     lnAdvPaym = lnAdvPaym + lnCreditAmt
                  Else
                     lnAdvPaym = lnAdvPaym + p_oAdvance("nTranAmtx") - p_oAdvance("nCredtAmt")
                     p_oAdvance("nCredtAmt") = p_oAdvance("nTranAmtx")
                  End If
               End If

               lsSQL = "UPDATE SP_Retail_Order_Detail" & _
                        " SET nIssuedxx = nIssuedxx + " & lnIssueQty & _
                        " WHERE sTransNox = " & strParm(loRS("sTransNox")) & _
                           " AND sPartsIDx = " & strParm(loRS("xPartsIDx"))
                  
               If p_oAppDrivr.Execute(lsSQL, "SP_Retail_Order_Detail") <= 0 Then
                  MsgBox "Unable to Update Issued SP Order!", vbCritical, "Warning"
                  GoTo endProc
               End If
               
               'Save the issued parts
               p_oRSIssd.AddNew
               p_oRSIssd("nEntryNox") = lnCtr
               p_oRSIssd("sPartsIDx") = loRS("sPartsIDx")
               p_oRSIssd("nQtyOrder") = lnIssueQty
               p_oRSIssd("xPartsIDx") = loRS("xPartsIDx")
               
               If lnQuantity = 0 Then Exit Do
               loRS.MoveNext
            Loop
            
            .Detail(lnCtr, "nOrderQty") = lnIssueTot
         End If
      Next
   End With

procAdvPaym:
   ' XerSys 2014-04-08
   '  Disable the following code and move the updating of advance payment
   '     to updateAdvPaym procedure
   ' finally update the issued amount of the ordered parts
   With p_oAdvance
      ' »»» XerSys - 01-18-2012
      '  Assume that the amount paid less the tran total will be the
      '     amount to be credited on the advance payment trans
      lnIssueTot = p_oBaseClas.Master("nTranTotl") - p_oBaseClas.Master("nAmtPaidx")
      .MoveFirst

      If lnAdvPaym < p_nAdvPaymx Then
         If MsgBox("Advance payment amounting to Php" & Trim(Format(p_nAdvPaymx, "#,##0.00")) & " exceeds " & vbCrLf & _
                   "   Php" & Trim(Format(lnAdvPaym, "#,##0.00")) & ", the amount issued spare parts!" & vbCrLf & vbCrLf & _
            "Credit whole Amount of Advance Payment?", vbCritical + vbYesNo, "Warning") = vbYes Then
            lbCreditAll = True
         Else
            lbCreditAll = False
         End If
      End If

      p_nAdvPaymx = 0
      Do Until .EOF
         lnCreditAmt = 0
         If .Fields("cTranStat") = xeStateClosed Then
            If .Fields("cWithOrdr") = xeYes Then
               If .Fields("nTranAmtx") > .Fields("nCredtAmt") Then
                  If lbCreditAll Then
                     lnCreditAmt = .Fields("nTranAmtx")
                  Else
                     lnCreditAmt = .Fields("nCredtAmt")
                  End If
               Else
                  lnCreditAmt = .Fields("nCredtAmt")
               End If
   
               ' »»» XerSys 01-18-2012
               '  Always credit what is only used
               If lnIssueTot > lnCreditAmt Then
                  lnIssueTot = lnIssueTot - lnCreditAmt
               Else
                  lnCreditAmt = lnIssueTot
                  lnIssueTot = 0
               End If
   
'               lsSQL = "UPDATE SP_Retail_Order_Master" & _
'                        " SET nCredtAmt = nCredtAmt + " & lnCreditAmt & _
'                           ", cTranStat = " & strParm(xeStatePosted) & _
'                        " WHERE sTransNox = " & strParm(.Fields("sReferNox"))
'
'               If p_oAppDrivr.Execute(lsSQL, "SP_Retail_Order_Master") <= 0 Then
'                  MsgBox "Unable to Update Reservation Info!", vbCritical, "Warning"
'                  GoTo endProc
'               End If
'
               .Fields("nCredtAmt") = lnCreditAmt
            Else
               ' »»» XerSys 01-18-2012
               '  I can't do anything about this now, so leave it as it is.
               lnCreditAmt = .Fields("nTranAmtx")
               lnIssueTot = lnIssueTot - lnCreditAmt
               If lnIssueTot < 0 Then lnIssueTot = 0
            End If
         End If
         p_nAdvPaymx = p_nAdvPaymx + lnCreditAmt

         .MoveNext
      Loop
   End With

   issueSPOrder = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Private Function unIssueSPOrder() As Boolean
   Dim lsProcName As String
   Dim loRS As Recordset
   Dim loRSTemp As Recordset
   Dim lsSQL As String, lsCondition As String
   Dim lsPartsIDx As String
   Dim lnQuantity As Integer
   Dim lnIssueTot As Integer
   Dim lnCreditAmt As Double
   Dim lnCtr As Integer

   lsProcName = "unIssueSPOrder"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   ' retrieve the existing record
   With p_oAdvance
      .MoveFirst
      If .Fields("sTransNox") <> "" Then
         lsCondition = ""
         Do Until .EOF
            If .Fields("cWithOrdr") = xeYes Then
               lsCondition = lsCondition & ", " & strParm(.Fields("sTransNox"))

               ' XerSys 2014-04-08
               '  Disable the following codes
'               lsSQL = "UPDATE SP_Retail_Order_Master" & _
'                        " SET nCredtAmt = nCredtAmt - " & .Fields("nCredtAmt") & _
'                           ", cTranStat = " & IIf(.Fields("nTotCredt") = .Fields("nCredtAmt"), strParm(xeStateOpen), strParm(xeStatePosted)) & _
'                        " WHERE sTransNox = " & strParm(.Fields("sTransNox"))
'
'               If p_oAppDrivr.Execute(lsSQL, "SP_Retail_Order_Master") < 0 Then
'                  MsgBox "Unable to Update Reservation Info!", vbCritical, "Warning"
'                  GoTo endProc
'               End If
            End If

            .Fields("nCredtAmt") = 0
            .MoveNext
         Loop

         If lsCondition <> "" Then
            lsCondition = "sTransNox IN (" & Mid(lsCondition, 3) & ")"
         End If
      End If
   End With

   ' this means that advance payment has no spareparts order detail
   If lsCondition = "" Then
      unIssueSPOrder = True
      GoTo endProc
   End If
   
   'join the table to a group
   'see remarks from issueSPOrder for more info...
   'kalyptus - 2012.03.17
   lsSQL = "SELECT a.sTransNox" & _
                ", c.sPartsIDx" & _
                ", a.nIssuedxx" & _
                ", a.sPartsIDx xPartsIDx" & _
                ", IFNull(b.sSuperIDx, 'x') sSuperIDx" & _
            " FROM SP_Retail_Order_Detail a" & _
                " LEFT JOIN SP_Supersession b ON a.sPartsIDx = b.sPartsIDx" & _
                " LEFT JOIN SP_Supersession c ON b.sSuperIDx = c.sSuperIDx" & _
            " WHERE " & lsCondition & _
               " AND a.nIssuedxx > 0" & _
            " ORDER BY a.sPartsIDx" & _
               ", a.sTransNox DESC"

   Set loRSTemp = New Recordset
   loRSTemp.Open AddCondition(lsSQL, lsCondition), _
         p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText

   ' XerSys - 2013-10-16
   '  Advance payment are sometimes used to other parts
   '     which are not included in the original order
   If loRSTemp.EOF Then
      unIssueSPOrder = True
      GoTo endProc
'      MsgBox "Unable to Retrieve Issued Parts for Deletion!", vbCritical, "Warning"
'      GoTo endProc
   End If
   
   With p_oBaseClas
      lsPartsIDx = ""
      lnIssueTot = 0
      
      Call initTempIssue
      
      For lnCtr = 0 To .ItemCount - 1
         'Process only sales that have an SP Order
         If .Detail(lnCtr, "nOrderQty") > 0 Then
            loRSTemp.Filter = "sPartsIDx = " & strParm(.Detail(lnCtr, "sPartsIDx"))
            If Not loRSTemp.EOF Then loRSTemp.MoveFirst
            
            Set loRS = initTempRecd
            
            lnIssueTot = .Detail(lnCtr, "nOrderQty")
            
            If Not loRSTemp.EOF Then
               Do Until loRSTemp.EOF
                  loRS.AddNew
                  loRS("sTransNox") = loRSTemp("sTransNox")
                  loRS("sPartsIDx") = loRSTemp("sPartsIDx")
                  loRS("nQuantity") = loRSTemp("nIssuedxx")
                  loRS("xPartsIDx") = loRSTemp("xPartsIDx")
                  loRS("sSuperIDx") = loRSTemp("sSuperIDx")
                  loRSTemp.MoveNext
               Loop
            End If
                        
            Do Until loRS.EOF
               If loRS("nQuantity") > lnIssueTot Then
                  lnQuantity = lnIssueTot
                  lnIssueTot = 0
               Else
                  lnQuantity = loRS("nQuantity")
                  lnIssueTot = lnIssueTot - lnQuantity
               End If
               
               lsSQL = "UPDATE SP_Retail_Order_Detail" & _
                        " SET nIssuedxx = nIssuedxx - " & lnQuantity & _
                        " WHERE sTransNox = " & strParm(loRS("sTransNox")) & _
                           " AND sPartsIDx = " & strParm(loRS("xPartsIDx"))

               If p_oAppDrivr.Execute(lsSQL, "SP_Retail_Order_Detail") <= 0 Then
                  MsgBox "Unable to Update Issued SP Order!", vbCritical, "Warning"
                  GoTo endProc
               End If
               
               'Save the issued parts
               p_oRSIssd.AddNew
               p_oRSIssd("nEntryNox") = lnCtr
               p_oRSIssd("sPartsIDx") = loRS("sPartsIDx")
               p_oRSIssd("nQtyOrder") = lnQuantity
               p_oRSIssd("xPartsIDx") = loRS("xPartsIDx")

               If lnIssueTot = 0 Then Exit Do
               
               loRS.MoveNext
            Loop
         End If
      Next
   End With

   unIssueSPOrder = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Private Sub Class_Initialize()
   pbShowMsg = True
   p_bHsParent = False
   p_nGCpnAmtx = 0
   p_nAdvPaymx = 0
End Sub

Private Sub Class_Terminate()
   Set p_oBaseClas = Nothing
   Set p_oClient = Nothing
End Sub

Private Sub p_oBaseClas_CancelOthers(Cancel As Boolean)
   Call undoAdvPayment
   RaiseEvent CancelOthers(Cancel)
End Sub

Private Sub p_oBaseClas_DeleteOthers(Cancel As Boolean)
   Call undoAdvPayment
   RaiseEvent DeleteOthers(Cancel)
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_InitMaster"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = Format(p_oAppDrivr.ServerDate, "MM/DD/YYYY")
      .Master("sSalesInv") = getNextInvoice
      .Master("nTranTotl") = 0#
      .Master("nGiftCpnx") = 0
      .Master("nAmtPaidx") = 0#
      .Master("cTranStat") = xeStateOpen
      .Master("sAutoTran") = xeNo
      .Master("sReferNox") = ""
      .Master("sSourceCd") = ""
      .Master("nReplAmtx") = 0
      .Master("sClientID") = ""
      .Master("cCardType") = 2

      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "nQuantity") = 0
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nUnitPrce") = 0#
      .Detail(0, "nDiscount") = 0#
      .Detail(0, "nAddDiscx") = 0#
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sPartsIDx") = ""
      .Detail(0, "cPartType") = 1
      .Detail(0, "nSelPrice") = 0#
      .Detail(0, "nDelrDisc") = 1
      .Detail(0, "nResvOrdr") = 1
      .Detail(0, "nBackOrdr") = 1
      .Detail(0, "nReorderx") = 1
      .Detail(0, "nLedgerNo") = 1
      .Detail(0, "nOrderQty") = 0
      .Detail(0, "sSuperIDx") = 0
   End With
   
   p_xChecks.TransNox = ""
   p_xChecks.AcctNoxx = ""
   p_xChecks.BankIDxx = ""
   p_xChecks.CheckNox = ""
   p_xChecks.CheckDte = Date
   p_xChecks.Amountxx = 0
   
   p_sAddressx = ""
   p_sFullName = ""
   p_nAdvPaymx = 0#
      
   Call initDisc
   
   Call initAdvance
   
   p_oClient.InitClient

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "p_oBaseClas_LoadOthers()"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   With p_oBaseClas
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "cLaborxxx") = 1 Then
            .Detail(pnCtr, "sBarrCode") = .Detail(pnCtr, "sLaborIDx")
            .Detail(pnCtr, "sDescript") = .Detail(pnCtr, "sLaborNme")
            .Detail(pnCtr, "nQtyOnHnd") = 0
            .Detail(pnCtr, "cPartType") = 1
            .Detail(pnCtr, "nSelPrice") = .Detail(pnCtr, "nUnitPrce")
            .Detail(pnCtr, "nDelrDisc") = .Detail(pnCtr, "nDiscount")
            .Detail(pnCtr, "nResvOrdr") = 0
            .Detail(pnCtr, "nBackOrdr") = 0
            .Detail(pnCtr, "nReorderx") = 0
            .Detail(pnCtr, "nLedgerNo") = 0
'         Else
'            .Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "xQtyOnHnd")
         End If
      Next
   
      p_sAddressx = IIf(IsNull(.Master("xAddressx")), "", .Master("xAddressx"))
      p_sFullName = IIf(IsNull(.Master("xFullName")), "", .Master("xFullName"))
      
      lsSQL = "SELECT" & _
                  "  sTransNox" & _
                  ", sCheckNox" & _
                  ", sAcctNoxx" & _
                  ", sBankIDxx" & _
                  ", dCheckDte" & _
                  ", nAmountxx" & _
                  ", cDepositd" & _
               " FROM Checks_Received" & _
               " WHERE sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                  " AND sSourceCd = " & strParm(pxeSourceCd)
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
      With loRS
         If loRS.EOF Then
            p_xChecks.TransNox = ""
            p_xChecks.AcctNoxx = ""
            p_xChecks.BankIDxx = ""
            p_xChecks.CheckNox = ""
            p_xChecks.CheckDte = Date
            p_xChecks.Amountxx = 0
         Else
            p_xChecks.TransNox = .Fields("sTransNox")
            p_xChecks.CheckNox = .Fields("sCheckNox")
            p_xChecks.AcctNoxx = .Fields("sAcctNoxx")
            p_xChecks.BankIDxx = .Fields("sBankIDxx")
            p_xChecks.CheckDte = .Fields("dCheckDte")
            p_xChecks.Amountxx = .Fields("nAmountxx")
            p_xChecks.Depositd = .Fields("cDepositd")
         End If
      End With
  
      ' check if advance payment exist
      If .Master("sAutoTran") = CStr(xeYes) Then
         Call initAdvance
         lsSQL = "SELECT b.sTransNox" & _
                     ", c.sORNoxxxx sReferNox" & _
                     ", c.dTransact" & _
                     ", a.nCredtAmt nTranAmtx" & _
                     ", a.nCredtAmt" & _
                     ", IFNULL(d.sTransNox, '') sCheckTrn" & _
                     ", IFNULL(d.cDepositd, '1') cDepositd" & _
                     ", IFNULL(d.cChckStat, '2') cChckStat" & _
                     ", a.sSourceCd" & _
                     ", a.cWithOrdr" & _
                  " FROM SP_SO_Reservation a" & _
                     ", SP_Retail_Order_Master b" & _
                     ", Receipt_Master c" & _
                        " LEFT JOIN Checks_Received d" & _
                           " ON c.sTransNox = d.sReferNox" & _
                              " AND d.sSourceCd = " & strParm(xeSrcOReceipt) & _
                  " WHERE a.sTransNox = " & strParm(.Master("sTransNox")) & _
                     " AND a.cWithOrdr = " & strParm(xeYes) & _
                     " AND a.sReferNox = b.sTransNox" & _
                     " AND a.sSourceCd = " & strParm(xeSrcOReceipt) & _
                     " AND b.sTransNox = c.sReferNox" & _
                     " AND c.sSourceCd = " & strParm(xeSrcSPCustOrder)
         lsSQL = lsSQL & _
                  " UNION SELECT b.sTransNox" & _
                     ", c.sPRNoxxxx sReferNox" & _
                     ", c.dTransact" & _
                     ", a.nCredtAmt nTranAmtx" & _
                     ", a.nCredtAmt" & _
                     ", IFNULL(d.sTransNox, '') sCheckTrn" & _
                     ", IFNULL(d.cDepositd, '1') cDepositd" & _
                     ", IFNULL(d.cChckStat, '2') cChckStat" & _
                     ", a.sSourceCd" & _
                     ", a.cWithOrdr" & _
                  " FROM SP_SO_Reservation a" & _
                     ", SP_Retail_Order_Master b" & _
                     ", Provisionary_Receipt_Master c" & _
                        " LEFT JOIN Checks_Received d" & _
                           " ON c.sTransNox = d.sReferNox" & _
                              " AND d.sSourceCd = " & strParm(xeSrcPReceipt) & _
                  " WHERE a.sTransNox = " & strParm(.Master("sTransNox")) & _
                     " AND a.cWithOrdr = " & strParm(xeYes) & _
                     " AND a.sReferNox = b.sTransNox" & _
                     " AND a.sSourceCd = " & strParm(xeSrcPReceipt) & _
                     " AND b.sTransNox = c.sReferNox" & _
                     " AND c.sSourceCd = " & strParm(xeSrcSPCustOrder)
         lsSQL = lsSQL & _
                  " UNION SELECT c.sTransNox" & _
                     ", c.sPRNoxxxx sReferNox" & _
                     ", c.dTransact" & _
                     ", a.nCredtAmt nTranAmtx" & _
                     ", a.nCredtAmt" & _
                     ", '' sCheckTrn" & _
                     ", '1' cDepositd" & _
                     ", '2' cChckStat" & _
                     ", a.sSourceCd" & _
                     ", a.cWithOrdr" & _
                  " FROM SP_SO_Reservation a" & _
                     ", Provisionary_Receipt_Master c" & _
                  " WHERE a.sTransNox = " & strParm(.Master("sTransNox")) & _
                     " AND a.cWithOrdr = " & strParm(xeNo) & _
                     " AND a.sReferNox = c.sTransNox" & _
                     " AND c.cTranType = " & strParm(0) & _
                     " AND c.sReferNox = ''" & _
                     " AND c.cTranStat = " & strParm(xeStateClosed)

'         lsSQL = "SELECT a.sReferNox" & _
'                     ", IF(a.cWithOrdr = " & strParm(xeNo) & ", b.nTranAmtx, c.nTranTotl) nTranAmtx" & _
'                     ", IF(a.cWithOrdr = " & strParm(xeNo) & ", a.nCredtAmt, c.nCredtAmt) nTotCredt" & _
'                     ", a.nCredtAmt" & _
'                     ", a.cWithOrdr" & _
'                  " FROM SP_SO_Reservation a" & _
'                     " LEFT JOIN Provisionary_Receipt_Master b" & _
'                        " ON a.sReferNox = b.sTransNox" & _
'                           " AND b.sSourceCd = " & strParm(pxeSPSales) & _
'                           " AND a.cWithOrdr = " & strParm(xeNo) & _
'                     " LEFT JOIN SP_Retail_Order_Master c" & _
'                        " ON a.sReferNox = c.sTransNox" & _
'                           " AND a.cWithOrdr = " & strParm(xeYes) & _
'                  " WHERE a.sTransNox = " & strParm(.Master("sTransNox"))
'
'         ' this will be for the old implementation
'         lsSQL = lsSQL & _
'                  " UNION SELECT sTransNox sReferNox" & _
'                     ", nTranAmtx" & _
'                     ", nTranAmtx nTotCredt" & _
'                     ", nTranAmtx nCredtAmt" & _
'                     ", 0 cWithOrdr" & _
'                  " FROM Provisionary_Receipt_Master" & _
'                  " WHERE sReferNox = " & strParm(.Master("sTransNox")) & _
'                     " AND sSourceCd = " & strParm(pxeSPSales) & _
'                  " ORDER BY sReferNox"

         Set loRS = New Recordset
         loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
         Debug.Print lsSQL, loRS.EOF
            
         p_nAdvPaymx = 0
         p_sAdvRefer = ""
         With p_oAdvance
            Do Until loRS.EOF
               .AddNew
               .Fields("sTransNox") = loRS("sTransNox")
               .Fields("sReferNox") = loRS("sReferNox")
               .Fields("dTransact") = loRS("dTransact")
               .Fields("nTranAmtx") = loRS("nTranAmtx")
               .Fields("nCredtAmt") = loRS("nCredtAmt")
               .Fields("sCheckTrn") = loRS("sCheckTrn")
               .Fields("sSourceCd") = loRS("sSourceCd")
               .Fields("cWithOrdr") = loRS("cWithOrdr")
               p_nAdvPaymx = p_nAdvPaymx + loRS("nCredtAmt")
               
               p_sAdvRefer = p_sAdvRefer & ", " & .Fields("sTransNox")
               loRS.MoveNext
            Loop
            p_sAdvRefer = Mid(p_sAdvRefer, 2)
         End With
         
         If .Master("nTranTotl") < (.Master("nAmtPaidx") + p_nAdvPaymx) Then
            p_nAdvPaymx = .Master("nTranTotl") - .Master("nAmtPaidx")
         End If
      End If
   End With
   
   p_oClient.InitClient

endProc:
   Set loRS = Nothing
   
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   Dim lsProcName As String
   Dim lbCancel As Boolean

   lsProcName = "p_oBaseClass_SaveOthers"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
      
   'Possible solution is this event
   RaiseEvent SaveOthers(Cancel)
      
   If Cancel Then GoTo endProc
   
   Cancel = True
   
   With p_oBaseClas
      If p_bNoLedger = False Then
         If saveSPTransaction() = False Then GoTo endProc
      End If
      
'      If .Master("nAmtPaidx") > 0 Then
'         Cancel = Not saveSalesInvoice()
'         If Cancel Then GoTo endProc
'      End If
      
      If Not saveSalesInvoice() Then GoTo endProc
      
      If p_sAdvRefer <> "" Then
         If Not updateAdvPayment() Then GoTo endProc
      End If
      
      If p_xChecks.Amountxx > 0 Then
         If saveCheck = False Then GoTo endProc
      End If
   End With
   
   If TypeName(p_oReturn) <> "Nothing" Then
      Cancel = Not p_oReturn.SaveTransaction
   Else
      Cancel = False
   End If
   
endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillCancel(Cancel As Boolean)
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_WillCancel"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Call cancelSPSOReturn
   
   Cancel = True
   
   If p_sAdvRefer <> "" Then
      If Not undoAdvPayment() Then GoTo endProc
   
      If Not unIssueSPOrder() Then GoTo endProc
   End If
   
   If Not p_bNoLedger Then
      If Not delSPTransaction() Then GoTo endProc
   End If
   
   ' delete salesinvoice
   If p_oBaseClas.Master("nAmtPaidx") > 0 Then
      If Not delSalesInvoice() Then GoTo endProc
   End If
   
   Cancel = False

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillDelete(Cancel As Boolean)
   Dim lsProcName As String
   Dim lbCancel As Boolean
   Dim lsSQL As String
   
   lsProcName = "p_oBaseClas_WillDelete"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   Cancel = True
   
   Call delSPSOReturn
   
   If p_sAdvRefer <> "" Then
      If Not undoAdvPayment() Then GoTo endProc

      If Not unIssueSPOrder() Then GoTo endProc
      
      ' finally remove the reservation info in the sales table
      lsSQL = "DELETE FROM SP_SO_Reservation" & _
               " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
      
      If p_oAppDrivr.Execute(lsSQL, "SP_PO_Reservation") <= 0 Then
         MsgBox "Unable to Remove Reservation History!", vbCritical, "Warning!"
         GoTo endProc
      End If
   End If
   
   If Not p_bNoLedger Then
      If Not delSPTransaction() Then GoTo endProc
   End If
   
   ' delete salesinvoice
   If p_oBaseClas.Master("nAmtPaidx") > 0 Then
      If Not delSalesInvoice() Then GoTo endProc
   End If
   
   Cancel = False

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lbCancel As Boolean
   Dim lbWithAdvance As Boolean

   lsProcName = "p_oBaseClass_WillSave"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Cancel = True
   
   pbModified = True
   With p_oBaseClas
      If .isMasterModify = False And .isDetailModify = False Then
         pbModified = False
         Cancel = False
         GoTo endProc
      End If

      RaiseEvent WillSave(lbCancel)
      If lbCancel Then GoTo endProc
      
      If isTotalDiscOK = False Then GoTo endProc
      
      If p_sAdvRefer <> "" Then
         ' XerSys 2014-05-30
         '  Disable the following codes coz it will not process transaction
         '     with multiple advance payment and the latest transactions was used
'         p_oAdvance.MoveFirst
'         lbWithAdvance = p_oAdvance("sReferNox") <> ""
'         If lbWithAdvance Then
            If Not issueSPOrder() Then GoTo endProc
'         End If
      End If
      
      If .Master("nTranTotl") > _
            (.Master("nAmtPaidx") _
               + p_nGCpnAmtx * .Master("nGiftCpnx") _
               + .Master("nReplAmtx") _
               + p_nAdvPaymx) Then
               
         If .Master("nAmtPaidx") >= 0 Then
            MsgBox "Invalid Amount Paid Detected!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
      
      If pbShowMsg Then
         p_oClient.SPClient = 1
         If p_oClient.SaveClient = False Then GoTo endProc
         If Trim(p_oClient.ClientID) <> "" Then .Master("sClientID") = p_oClient.ClientID
      End If
      
      'MS-2007.12.18
      If .EditMode = xeModeAddNew Then
         .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
      End If
      
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sPartsIDx") = "" Then Exit For
         
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1

'         If lbWithAdvance Then
'            .Detail(pnCtr, "nOrderQty") = getOrderQty(.Detail(pnCtr, "sPartsIDx"), .Detail(pnCtr, "nQuantity"))
'         Else
'            .Detail(pnCtr, "nOrderQty") = 0
'         End If
      Next
   End With
   
   
   If p_sAdvRefer <> "" Then
      With p_oAdvance
         .MoveFirst
         pnCtr = 1
         Do Until .EOF
            If .Fields("cTranStat") = xeStateClosed Then
               lsSQL = "INSERT INTO SP_SO_Reservation" & _
                        " SET sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                           ", sReferNox = " & strParm(.Fields("sTransNox")) & _
                           ", sSourceCd = " & strParm(.Fields("sSourceCd")) & _
                           ", nCredtAmt = " & .Fields("nCredtAmt") & _
                           ", cWithOrdr = " & .Fields("cWithOrdr") & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate())
               
               If p_oAppDrivr.Execute(lsSQL, "SP_PO_Reservation") <= 0 Then
                  MsgBox "Unable to Create Reservation History!", vbCritical, "Warning!"
                  GoTo endProc
               End If
            End If
            
            .MoveNext
         Loop
      End With
   End If
   Cancel = False
   
endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Sub

' --------------------------
' SPSalesReturn implementation starts here!
' --------------------------
Sub getSPSOReturn()
   If TypeName(p_oReturn) = "Nothing" Then
      Set p_oReturn = New clsSPSalesReturn
      With p_oReturn
         Set .AppDriver = p_oAppDrivr
         .TransNo = p_oBaseClas.Master("sTransNox")
         .InitTransaction
         .ClientID = p_oBaseClas.Master("sClientID")
         .ClientNm = p_sFullName
         .Address = p_sAddressx
         .TransactDate = p_oBaseClas.Master("dTransact")
               
         .OpenTransaction
      End With
   Else
      With p_oReturn
         .ClientID = p_oBaseClas.Master("sClientID")
         .ClientNm = p_sFullName
         .Address = p_sAddressx
         .TransactDate = p_oBaseClas.Master("dTransact")
      End With
   End If
   
   p_oReturn.ShowReturn

   p_oBaseClas.Master("nReplAmtx") = p_oReturn.Total
   RaiseEvent MasterRetrieved(15)
End Sub

Private Sub delSPSOReturn()
   If TypeName(p_oReturn) = "Nothing" Then
      Set p_oReturn = New clsSPSalesReturn
      With p_oReturn
         Set .AppDriver = p_oAppDrivr
         .TransNo = p_oBaseClas.Master("sTransNox")
         .InitTransaction
         .ClientID = p_oBaseClas.Master("sClientID")
         .ClientNm = p_oBaseClas.Master("xFullName")
         .Address = p_oBaseClas.Master("xAddressx")
         .TransactDate = p_oBaseClas.Master("dTransact")
               
         .OpenTransaction
      End With
   Else
      With p_oReturn
         .ClientID = p_oBaseClas.Master("sClientID")
         .ClientNm = p_oBaseClas.Master("xFullName")
         .Address = p_oBaseClas.Master("xAddressx")
         .TransactDate = p_oBaseClas.Master("dTransact")
      End With
   End If
   
   p_oReturn.DeleteTransaction
End Sub

Private Sub cancelSPSOReturn()
   If TypeName(p_oReturn) = "Nothing" Then
      Set p_oReturn = New clsSPSalesReturn
      With p_oReturn
         Set .AppDriver = p_oAppDrivr
         .TransNo = p_oBaseClas.Master("sTransNox")
         .InitTransaction
         .ClientID = p_oBaseClas.Master("sClientID")
         .ClientNm = p_oBaseClas.Master("xFullName")
         .Address = p_oBaseClas.Master("xAddressx")
         .TransactDate = p_oBaseClas.Master("dTransact")
               
         .OpenTransaction
      End With
   Else
      With p_oReturn
         .ClientID = p_oBaseClas.Master("sClientID")
         .ClientNm = p_oBaseClas.Master("xClientNm")
         .Address = p_oBaseClas.Master("xAddressx")
         .TransactDate = p_oBaseClas.Master("dTransact")
      End With
   End If
   
   p_oReturn.CancelTransaction
End Sub


Private Function initTempRecd() As Recordset
   Dim loRS As Recordset
   Set loRS = New ADODB.Recordset
   
   With loRS
      .Fields.Append "sTransNox", adVarChar, 12
      .Fields.Append "sPartsIDx", adVarChar, 12
      .Fields.Append "nQuantity", adInteger
      .Fields.Append "xPartsIDx", adVarChar, 12
      .Fields.Append "sSuperIDx", adVarChar, 12
      .Open
   End With
   
   Set initTempRecd = loRS
   
End Function

Private Sub initTempIssue()
   Set p_oRSIssd = New ADODB.Recordset
   
   With p_oRSIssd
      .Fields.Append "nEntryNox", adInteger
      .Fields.Append "sPartsIDx", adVarChar, 12
      .Fields.Append "nQtyOrder", adInteger
      .Fields.Append "xPartsIDx", adVarChar, 12
      .Open
   End With
End Sub

'kalyptus - 08/23/2012 03:02 pm
'Used in resetting the status of provisionary receipt
Private Sub resetAdvance()
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String

   lsProcName = "resetAdvance"
   'On Error GoTo errProc
   
   With p_oBaseClas
      If .Master("sClientID") = Empty Then GoTo endWithWarning
      
      lsSQL = "SELECT" & _
                  "  sTransNox" & _
                  ", sPRNoxxxx" & _
                  ", dTransact" & _
                  ", nTranAmtx" & _
                  ", cTranStat" & _
               " FROM Provisionary_Receipt_Master" & _
               " WHERE sClientID = " & strParm(.Master("sClientID")) & _
                  " AND sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                  " AND cTranType = " & strParm(0) & _
                  " AND (cTranStat = " & strParm(xeStateClosed) & _
                   " AND sReferNox = " & strParm(.Master("sTransNox")) & _
                   " AND sSourceCd = " & strParm("SPSl") & ")" & _
               " ORDER BY sPRNoxxxx"
      
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set loRS.ActiveConnection = Nothing
      
      Do Until loRS.EOF
         lsSQL = "UPDATE Provisionary_Receipt_Master" & _
                " SET cTranStat = " & strParm(xeStateOpen) & _
                   ", sReferNox = ''" & _
                   ", sSourceCD = ''" & _
                " WHERE sTransNox = " & strParm(loRS("sTransNox"))
         p_oAppDrivr.Execute lsSQL, "Provisionary_Receipt_Master"
         
         loRS.MoveNext
      Loop
   End With
   
endProc:
   Set loRS = Nothing
   Exit Sub
endWithWarning:
   MsgBox "No Record Found for this Customer!" & _
            vbCrLf & " Can Not Process Advance Payment", vbCritical, "Warning"
   GoTo endProc
errProc:
   ShowError lsProcName
End Sub

' XerSys 2014-12-18
'  This will be in-charge in setting default discount
Private Function prcDiscount(ByVal lnRow As Long) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lbDiscount As Boolean
   
   lsProcName = "prcDiscount"
   'On Error GoTo errProc
   
   With p_oBaseClas
      lsSQL = "SELECT sPartsIDx" & _
                  ", IFNULL(sBrandIDx, '') sBrandIDx" & _
                  ", IFNULL(sCategIDx, '') sCategIDx" & _
               " FROM Spareparts" & _
               " WHERE sPartsIDx = " & strParm(.Detail(lnRow, "sPartsIDx"))
               
      Debug.Print lsSQL
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If loRS.EOF Then GoTo endProc
      
      ' priority discount will be the category
      If loRS("sCategIDx") <> "" Then
         p_oDiscount.MoveFirst
         Do Until p_oDiscount.EOF
            If loRS("sCategIDx") = p_oDiscount("sCategrID") Then
               ' XerSys - 2015-06-08
               '  Check if their is a minimum unit price
               lbDiscount = False
               If p_oDiscount("nMinAmtxx") <> 0 Then
                  If p_oDiscount("nMinAmtxx") <= .Detail(lnRow, "nUnitPrce") Then
                     lbDiscount = True
                  End If
               Else
                  lbDiscount = True
               End If
            
               If lbDiscount Then
                  If p_oDiscount("nDiscRate") > 0 Then
                     .Detail(lnRow, "nDiscount") = p_oDiscount("nDiscRate")
                     If .Detail(lnRow, "nSMaxDisc") < p_oDiscount("nDiscRate") Then
                        .Detail(lnRow, "nSMaxDisc") = p_oDiscount("nDiscRate")
                     End If
                     
                     If .Detail(lnRow, "nMMaxDisc") < p_oDiscount("nDiscRate") Then
                        .Detail(lnRow, "nMMaxDisc") = p_oDiscount("nDiscRate")
                     End If
                     Exit Do
                  End If
                  
                  If p_oDiscount("nDiscAmtx") > 0 Then
                     .Detail(lnRow, "nAddDiscx") = p_oDiscount("nDiscAmtx")
                     Exit Do
                  End If
               End If
            End If
            
            p_oDiscount.MoveNext
         Loop
      ElseIf loRS("sBrandIDx") <> "" Then
         p_oDiscount.MoveFirst
         Do Until p_oDiscount.EOF
            ' check for per brand discount
            If loRS("sBrandIDx") = p_oDiscount("sBrandIDx") Then
               ' XerSys - 2015-06-08
               '  Check if their is a minimum unit price
               lbDiscount = False
               If p_oDiscount("nMinAmtxx") <> 0 Then
                  If p_oDiscount("nMinAmtxx") <= .Detail(lnRow, "nUnitPrce") Then
                     lbDiscount = True
                  End If
               Else
                  lbDiscount = True
               End If
            
               If lbDiscount Then
                  If p_oDiscount("nDiscRate") > 0 Then
                     .Detail(lnRow, "nDiscount") = p_oDiscount("nDiscRate")
                     If .Detail(lnRow, "nSMaxDisc") < p_oDiscount("nDiscRate") Then
                        .Detail(lnRow, "nSMaxDisc") = p_oDiscount("nDiscRate")
                     End If
                     
                     If .Detail(lnRow, "nMMaxDisc") < p_oDiscount("nDiscRate") Then
                        .Detail(lnRow, "nMMaxDisc") = p_oDiscount("nDiscRate")
                     End If
                     Exit Do
                  End If
                  
                  If p_oDiscount("nDiscAmtx") > 0 Then
                     .Detail(lnRow, "nAddDiscx") = p_oDiscount("nDiscAmtx")
                     Exit Do
                  End If
               End If
            End If
            p_oDiscount.MoveNext
         Loop
      End If
   End With
   prcDiscount = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lnRow & " )"
End Function

Private Function getDiscCard(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsOldProc As String
   Dim lsMaster As String
   Dim lasMaster() As String

   lsOldProc = "getDiscCard"
   'On Error GoTo errProc
   getDiscCard = False

   With p_oBaseClas
      If lsValue <> "" Then
         If Trim(lsValue) = Trim(.Master("sCardIDxx")) Then GoTo endProc
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
      
      If lbSearch Then
         lsMaster = "a.sCardDesc LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsMaster = "a.sCardDesc = " & strParm(Trim(lsValue))
      End If
         
      lsMaster = .getMaster(8, lsMaster)
      If lsMaster = "" Then
         If lbSearch Then
            GoTo endProc
         End If
         GoTo endWithClear
      End If

      lasMaster = Split(lsMaster, "»")
      p_sDiscCard = lasMaster(1)
      .Master("sCardIDxx") = lasMaster(0)
      
      ' Retrieve discounts
      Call getDiscounts(lasMaster(0))
      
      ' check if available parts exists
      If .Detail(0, "sPartsIDx") <> "" Then
         For pnCtr = 0 To .ItemCount - 1
           Call prcDiscount(pnCtr)
         Next
      End If
   End With

   getDiscCard = True

endProc:
   RaiseEvent MasterRetrieved(8)
   Exit Function
endWithClear:
   p_oBaseClas.Master("sCardIDxx") = ""
   p_sDiscCard = ""
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function initDisc() As Boolean
   Dim lsProcName As String
   
   lsProcName = "initDisc"
   'On Error GoTo errProc
   
   Set p_oDiscount = New Recordset
   With p_oDiscount
      .Open AddCondition(p_sDiscSQL, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      .ActiveConnection = Nothing
      
      .AddNew
      .Fields("sCardIDxx") = ""
      .Fields("sBrandIDx") = ""
      .Fields("sCategrID") = ""
      .Fields("nMinAmtxx") = 0#
      .Fields("nDiscRate") = 0#
      .Fields("nDiscAmtx") = 0#
      .Fields("nSCDiscxx") = 0#
   End With
   
   p_sDiscCard = ""
   initDisc = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function getDiscounts(ByVal lsCardID) As Boolean
   Dim lsProcName As String
   
   lsProcName = "getDiscounts"
   'On Error GoTo errProc
   
   Set p_oDiscount = New Recordset
   p_oDiscount.Open AddCondition(p_sDiscSQL, "sCardIDxx = " & strParm(lsCardID)), _
         p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   Debug.Print p_sDiscSQL
   If p_oDiscount.EOF Then Call initDisc
   Set p_oDiscount.ActiveConnection = Nothing
   
   getDiscounts = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lsCardID & " )"
End Function
