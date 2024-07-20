VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Rex S. Adversalo
' XerSys Computing
' Canaoalan, Binmaley, Pangasinan
'
' Copyright 2003-2005 and beyond
' All Rights Reserved
'
'     Spareparts Branch Delivery Object
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
'  XerSys [ 01/29/2005 05:28 pm ]
'     Start creating this object.
'  XerSys [ 11/03/2005 04:49 pm ]
'     Add the issuance of ordered parts...
'  Kalyptus [ 12/11/2006 10:08 am]
'     Implements the BugTracker
'  Kalyptus [ 04/02/2007 10:25 am]
'     Updated the p_oBase_Clas to conform it with the clsMasterDetail specification
'  XerSys [ 08/11/2007 03:14 pm ]
'     Update this object to adopt the new implementation of SP Inventory Transaction
'  XerSys [ 02/16/2007 08:47 am ]
'     Verify Order Issuance policy
'  XerSys [ 05/26/2008 05:33 pm ]
'     Add update of backorder if issuing branch lessen requested parts
'  XerSys [ 07/08/2011 09:19 am ]
'     Add display message confirmation and reference field in master
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME = "clsSPTransfer"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oDetail As Recordset

Private p_sBranchCd As String
Private p_sAddressx As String
Private p_nTranStat As Integer
Private p_sParent As String

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pbShowMsg As Boolean
Private pnCtr As Integer

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(Value As String)
   p_sBranchCd = Value
End Property

Property Let DisplayConfirmation(Value As Boolean)
   pbShowMsg = Value
End Property

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   'On error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   Detail = p_oBaseClas.Detail(Row, Index)
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   'On error Resume Next
   
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
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xaddressx"
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
      p_oBaseClas.Master(Index) = Value
   Case 2, "sdestinat"
      getMaster Value, False
   Case 3, "xaddressx"
      p_sAddressx = Value
   Case 12, 13, "srefernox", "ssourcecd"
      If Left(p_sParent, 3) = "cls" Then
         p_oBaseClas.Master(Index) = Value
      End If
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oBaseClas.ItemCount
End Property

Property Let Parent(ByVal Value As String)
   p_sParent = Value
End Property

Property Get MasFldSize(ByVal Index As Integer) As Variant
   'On error Resume Next

   If pbInitTran = False Then Exit Property
   
   MasFldSize = p_oBaseClas.MasFldSize(Index)
End Property

Property Let TransStatus(ByVal Value As Integer)
   p_nTranStat = Value
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsCondition As String
   
   lsProcName = "InitTransaction"
   'On Error Goto errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail
   
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "SP_Transfer_Master"
      .DetailTable = "SP_Transfer_Detail"
      .DisplayMessage = pbShowMsg
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDestinat" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                        ", a.sRequestx" & _
                        ", a.sRemarksx" & _
                        ", a.sReceived" & _
                        ", a.dReceived" & _
                        ", a.sApproved" & _
                        ", a.nGrossAmt" & _
                        ", a.nDiscount" & _
                        ", a.nTranTotl" & _
                        ", a.sReferNox" & _
                        ", a.sSourceCd" & _
                        ", a.cTranStat" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") xSourcexx" & _
                        ", a.nEntryNox"
      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                        ", Branch b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE a.sDestinat = b.sBranchCd" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx"
                     
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", d.sTypeCode" & _
                        ", e.sModelNme" & _
                        ", c.nQtyOnHnd" & _
                        ", a.nQuantity" & _
                        ", a.nUnitPrce" & _
                        ", a.dModified" & _
                        ", a.sTransNox" & _
                        ", a.sPartsIDx" & _
                        ", c.nQtyOnHnd  + a.nQuantity xQtyOnHnd" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo"
      .DetailQuery = .DetailQuery & _
                     " FROM " & .DetailTable & " a" & _
                        ", Spareparts b" & _
                              " LEFT JOIN SP_Model e" & _
                                 " ON b.sModelIDx = e.sModelIDx" & _
                        ", SP_Inventory c" & _
                        ", SP_Type d" & _
                     " WHERE a.sPartsIDx = b.sPartsIDx" & _
                        " AND b.sPartsIDx = c.sPartsIDx" & _
                        " AND b.cPartType = d.cPartType" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                     " ORDER BY a.nEntryNox"
   
      .Reference = "a.sTransNox|a.sTransNox"
      .VerifyEntryNo = True
      
      pbInitTran = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sBranchNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Branch b" & _
                        " WHERE a.sDestinat = b.sBranchCd" & _
                           " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                        " ORDER BY b.sBranchNm" & _
                           ", a.dTransact"
      
      If p_nTranStat <> xeStateUnknown Then
         If p_nTranStat > xeStateUnknown Then
            lsCondition = "("
            For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
               lsCondition = lsCondition & " cTranStat = " & _
                                 strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
            Next
            lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
         Else
            lsCondition = "cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      .BrowseQuery(0) = AddCondition(.BrowseQuery(0), lsCondition)
      
      .BrowseColumn(0) = "sTransNox»sBranchNm»dTransact"
      .BrowseCriteria(0) = "a.sTransNox»b.sBranchNm»a.dTransact"
      .BrowseTitle(0) = "Transaction No»Branch»Date"
      .BrowseFormat(0) = "@@-@@@@@@@@»@»MMMM DD, YYYY"
      
      ' Query statement for the branch destination field
      .BrowseQuery(2) = "SELECT" & _
                           "  a.sBranchCd" & _
                           ", a.sBranchNm" & _
                           ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) AS xAddressx" & _
                        " FROM Branch a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & xeRecStateActive & _
                           " AND a.sBranchCd <> " & strParm(p_sBranchCd)
      .BrowseColumn(2) = "sBranchCd»sBranchNm»xAddressx"
      .BrowseTitle(2) = "Code»Branch Name»Address"
      
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sBarrCode" & _
                                 ", a.sDescript" & _
                                 ", c.sTypeCode" & _
                                 ", d.sModelNme" & _
                                 ", b.nQtyOnHnd" & _
                                 ", a.nSelPrice" & _
                                 ", a.sPartsIDx" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nReorderx" & _
                                 ", b.nLedgerNo" & _
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
   End With
   
   InitTransaction = pbInitTran
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "NewTransaction"
   'On Error Goto errProc

   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "SaveTransaction"
   'On Error Goto errProc
   
   If pbInitTran Then SaveTransaction = p_oBaseClas.SaveTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "OpenTransaction"
   'On Error Goto errProc
   
   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "DeleteTransaction"
   'On Error Goto errProc
   
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
   'On Error Goto errProc
   SearchTransaction = False
   
   If pbInitTran = False Then GoTo endProc
   
   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sTransNox"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "b.sBranchNm"
            lsDescript = "b.sBranchNm"
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

Function SearchAcceptance(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lsRecord As String
   Dim lsValue As String, lsField As String
   Dim lsDescript As String, lsCondition As String
   Dim lsMaster As String, lsDetail As String
   Dim lsQuery As String
   
   lsProcName = "SearchAcceptance"
   'On Error Goto errProc
   
   If pbInitTran = False Then GoTo endProc
   
   With p_oBaseClas
      lsMaster = .MasterQuery
      lsDetail = .DetailQuery
      lsQuery = .BrowseQuery(0)

      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDestinat" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                        ", a.sRequestx" & _
                        ", a.sRemarksx" & _
                        ", a.sReceived" & _
                        ", a.dReceived" & _
                        ", a.sApproved" & _
                        ", a.nGrossAmt" & _
                        ", a.nDiscount" & _
                        ", a.nTranTotl" & _
                        ", a.cTranStat" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", b.sBranchNm xSourcexx" & _
                        ", a.nEntryNox" & _
                     " FROM " & .MasterTable & " a" & _
                        ", Branch b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = b.sBranchCd" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx"
      
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", d.sTypeCode" & _
                        ", e.sModelNme" & _
                        ", c.nQtyOnHnd" & _
                        ", a.nQuantity" & _
                        ", a.nUnitPrce" & _
                        ", a.dModified" & _
                        ", a.sTransNox" & _
                        ", a.sPartsIDx" & _
                        ", c.nQtyOnHnd  + a.nQuantity xQtyOnHnd" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo"
      .DetailQuery = .DetailQuery & _
                     " FROM " & .DetailTable & " a" & _
                        ", Spareparts b" & _
                              " LEFT JOIN SP_Model e" & _
                                 " ON b.sModelIDx = e.sModelIDx" & _
                        ", SP_Inventory c" & _
                        ", SP_Type d" & _
                     " WHERE a.sPartsIDx = b.sPartsIDx" & _
                        " AND b.sPartsIDx = c.sPartsIDx" & _
                        " AND b.cPartType = d.cPartType" & _
                        " AND c.sBranchCd = LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ")" & _
                     " ORDER BY a.nEntryNox"

      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sBranchNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Branch b" & _
                        " WHERE LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = b.sBranchCd" & _
                           " AND a.sDestinat = " & strParm(p_sBranchCd) & _
                        " ORDER BY b.sBranchNm" & _
                           ", a.dTransact"
      
      If p_nTranStat <> xeStateUnknown Then
         If p_nTranStat > xeStateUnknown Then
            lsCondition = "("
            For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
               lsCondition = lsCondition & " cTranStat = " & _
                                 strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
            Next
            lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
         Else
            lsCondition = "cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      .BrowseQuery(0) = AddCondition(.BrowseQuery(0), lsCondition)
   End With
   
   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sTransNox"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "b.sBranchNm"
            lsDescript = "b.sBranchNm"
         End If
      End If
   End If
   
   SearchAcceptance = p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript)
   
endProc:
   With p_oBaseClas
      .MasterQuery = lsMaster
      .DetailQuery = lsDetail
      .BrowseQuery(0) = lsQuery
   End With

   Exit Function
errProc:
    ShowError lsProcName & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function PostTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "PostTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   
   If pbInitTran Then PostTransaction = p_oBaseClas.PostTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String)
   Dim lsProcName As String
   
   lsProcName = "CloseTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   
   If pbInitTran Then CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "CancelTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   
   If pbInitTran Then CancelTransaction = p_oBaseClas.CancelTransaction()

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
      .Detail(pnCtr, "sTypeCode") = ""
      .Detail(pnCtr, "sModelNme") = ""
      .Detail(pnCtr, "nQtyOnHnd") = 0
      .Detail(pnCtr, "nUnitPrce") = 0#
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sPartsIDx") = ""
      .Detail(pnCtr, "nResvOrdr") = 0
      .Detail(pnCtr, "nBackOrdr") = 0
      .Detail(pnCtr, "nResvOrdr") = 0
      .Detail(pnCtr, "nBackOrdr") = 0
      .Detail(pnCtr, "nReorderx") = 0
      .Detail(pnCtr, "nLedgerNo") = 0
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
   'On Error Goto errProc
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
   ShowError lsProcName & "( " & " )"
End Function

Function SearchMaster(ByVal Index As Variant, Optional Value As Variant = "") As Boolean
   If pbInitTran = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "sdestinat"
      getMaster Value, True
   Case Else
      
   End Select
   SearchMaster = True
End Function

Function UpdateTransaction() As Boolean
   With p_oBaseClas
      If .Master("sReferNox") <> Empty And pbShowMsg Then
         MsgBox "Transaction is a created by a parent transaction!" & vbCrLf & _
                  "Modify the parent transaction for any changes!" & vbCrLf & vbCrLf & _
                  "Load other transaction then Try again!", vbInformation, "Notice"
         Exit Function
      End If
      
      .EditMode = xeModeUpdate
   End With
   UpdateTransaction = True
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Function AcceptDelivery(dReceived As Date) As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnRow As Long
   
   lsProcName = "AcceptDelivery"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   AcceptDelivery = False
   
   If Not pbInitTran Then GoTo endProc
   
   With p_oBaseClas
      If .Master("sTransNox") = "" Then GoTo endProc
      
      'Test validity of the acceptance date
      '++++++++++++++++++++++
      'Validity #1: Date Transferred < Date Received
      If .Master("dTransact") > dReceived Then
         If p_oAppDrivr.UserLevel = xeEngineer Then
            If MsgBox("[Date Received] is less than the [Date Transferred]." & vbCrLf & _
                      "Do you want to continue?", vbOKCancel, "Warning") = vbCancel Then GoTo endProc
         Else
            If MsgBox("[Date Received] is less than the [Date Transferred]." & vbCrLf & _
                      "Please inform SEG/SSG if this is correct!", vbOKCancel, "Warning") = vbCancel Then GoTo endProc
         End If
      End If
      
      'Validity #2: Date Received is more than 15 days after Date Transferred
      If DateDiff("d", .Master("dTransact"), dReceived) > 15 Then
         If p_oAppDrivr.UserLevel = xeEngineer Then
            If MsgBox("[Date Received] is more than 15 days after [Date Transferred]." & vbCrLf & _
                      "Do you want to continue?", vbOKCancel, "Warning") = vbCancel Then GoTo endProc
         Else
            If MsgBox("[Date Received] is more than 15 days after [Date Transferred]." & vbCrLf & _
                      "Please inform SEG/SSG if this is correct!", vbOKCancel, "Warning") = vbCancel Then GoTo endProc
         End If
      End If
      '++++++++++++++++++++++
     
      lsSQL = "SELECT" & _
                  "  a.sPartsIDx" & _
                  ", a.nQuantity" & _
                  ", b.nQtyOnHnd" & _
                  ", b.nResvOrdr" & _
                  ", b.nBackOrdr" & _
                  ", b.nReorderx" & _
                  ", b.nLedgerNo" & _
               " FROM " & .DetailTable & " a" & _
                     " LEFT JOIN SP_Inventory b" & _
                        " ON a.sPartsIDx = b.sPartsIDx" & _
                           " AND b.sBranchCd = " & strParm(.Master("sDestinat")) & _
               " WHERE a.sTransNox = " & strParm(.Master("sTransNox")) & _
               " ORDER BY a.nEntryNox"
      
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If loRS.RecordCount <> .ItemCount Then
         MsgBox "Data Discrepancy Detected!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      Set loSPTrans = New clsSPInventoryTrans
      With loSPTrans
         Set .AppDriver = p_oAppDrivr
         .Branch = p_oBaseClas.Master("sDestinat")
         If .InitTransaction() = False Then GoTo endProc
      End With
      
      p_oAppDrivr.BeginTrans
      For pnCtr = 0 To .ItemCount - 1
         loSPTrans.Detail(pnCtr, "sPartsIDx") = loRS("sPartsIDx")
         loSPTrans.Detail(pnCtr, "nQuantity") = loRS("nQuantity")
         If IsNull(loRS("nQtyOnHnd")) Then
            loSPTrans.Detail(pnCtr, "cNewParts") = xeYes
            loSPTrans.Detail(pnCtr, "nQtyOnHnd") = 0
            loSPTrans.Detail(pnCtr, "nResvOrdr") = 0
            loSPTrans.Detail(pnCtr, "nBackOrdr") = 0
         Else
            loSPTrans.Detail(pnCtr, "nQtyOnHnd") = loRS("nQtyOnHnd")
            loSPTrans.Detail(pnCtr, "nResvOrdr") = loRS("nResvOrdr")
            loSPTrans.Detail(pnCtr, "nBackOrdr") = loRS("nBackOrdr")
            loSPTrans.Detail(pnCtr, "nLedgerNo") = loRS("nLedgerNo")
         End If
         
         loRS.MoveNext
      Next
      
      ' if a spareparts is not successfully updated, cancel saving.
      If loSPTrans.AcceptDelivery(.Master("sTransNox"), _
            dReceived, _
            xeModeAddNew) = False Then
         GoTo endProcWithRoll
      End If
      
      lsSQL = "UPDATE " & .MasterTable & " SET" & _
                  "  sReceived = " & strParm(p_oAppDrivr.UserID) & _
                  ", dReceived = " & dateParm(dReceived) & _
                  ", cTranStat = " & strParm(xeStatePosted) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sTransNox = " & strParm(.Master("sTransNox"))
               
      lnRow = p_oAppDrivr.Execute(lsSQL, _
                  .MasterTable, _
                  p_sBranchCd, _
                  Left(.Master("sTransNox"), Len(p_sBranchCd)))
      
      If lnRow = 0 Then
         MsgBox "Unable to Update Delivery Transaction!!!", vbCritical, "Warning"
         GoTo endProcWithRoll
      End If
      
      p_oAppDrivr.CommitTrans
      .Master("cTranStat") = xeStatePosted
   End With
   
   AcceptDelivery = True
   
endProc:
   Set loSPTrans = Nothing

   Exit Function
endProcWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
   p_oAppDrivr.RollbackTrans
   ShowError lsProcName & "( " & dReceived & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Sub getMaster(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "GetMaster"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   
   With p_oBaseClas
      If lsValue <> "" Then
         If lsValue = .Master(2) Then GoTo endProc
         
         If lbSearch Then
            lsMaster = "a.sBranchNm LIKE " & strParm(Trim(lsValue) & "%")
         Else
            lsMaster = "a.sBranchNm = " & strParm(Trim(lsValue))
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
   
      lsMaster = .getMaster(2, lsMaster)
      If lsMaster = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")
      
      .Master(2) = lasMaster(1)
      p_sAddressx = lasMaster(2)
   End With

endProc:
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master(2) = ""
      p_sAddressx = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
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
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   
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
      .Detail(lnRow, "nQtyOnHnd") = CLng(lasDetail(4))
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "nUnitPrce") = CDbl(lasDetail(5))
      .Detail(lnRow, "sPartsIDx") = lasDetail(6)
      .Detail(lnRow, "nResvOrdr") = lasDetail(7)
      .Detail(lnRow, "nBackOrdr") = lasDetail(8)
      .Detail(lnRow, "nReorderx") = lasDetail(9)
      .Detail(lnRow, "nLedgerNo") = lasDetail(10)
   End With

endProc:
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)
   RaiseEvent DetailRetrieved(7)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "sTypeCode") = ""
      .Detail(lnRow, "sModelNme") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nUnitPrce") = 0#
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "sPartsIDx") = ""
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

Private Function saveIssuedOrder() As Boolean
   Dim lsProcName As String
   
   lsProcName = "SaveIssuedOrder"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   
   With p_oDetail
      .MoveFirst
      Do While .EOF = False
         If .Fields("nNewQtyxx") > .Fields("nOldQtyxx") Then
            If addIssuedOrder(.Fields("sPartsIDx"), _
                  .Fields("nNewQtyxx") - .Fields("nOldQtyxx")) = False Then GoTo endProc
         ElseIf .Fields("nNewQtyxx") < .Fields("nOldQtyxx") Then
            If delIssuedOrder(.Fields("sPartsIDx"), _
                  .Fields("nOldQtyxx") - .Fields("nNewQtyxx")) = False Then GoTo endProc
         End If
         
         .MoveNext
      Loop
   End With
   
   saveIssuedOrder = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function addIssuedOrder(lsPartsIDx As String, _
      lnQuantity As Integer) As Boolean
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnIssuedxx As Long
   
   lsProcName = "addIssuedOrder"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   
   lsSQL = "SELECT" & _
               "  a.sTransNox" & _
               ", a.nEntryNox" & _
               ", a.sPartsIDx" & _
               ", a.nApproved - a.nIssuedxx xOrderQty" & _
            " FROM SP_Branch_Order_Detail a" & _
               ", SP_Branch_Order_Master b" & _
            " WHERE a.sTransNox = b.sTransNox" & _
               " AND a.cCanceled = " & strParm(xeNo) & _
               " AND b.cTranStat = " & strParm(xeStatePosted) & _
               " AND a.nApproved > a.nIssuedxx" & _
               " AND b.sTransNox LIKE " & strParm(p_oBaseClas.Master("sDestinat") & "%") & _
               " AND a.sPartsIDx = " & strParm(lsPartsIDx) & _
            " ORDER BY a.sTransNox"
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If loRS.EOF Then
      addIssuedOrder = True
      GoTo endProc
   End If
   
   Do
      lnIssuedxx = loRS("xOrderQty")
      
      If lnQuantity < lnIssuedxx Then lnIssuedxx = lnQuantity
      
      lnQuantity = lnQuantity - lnIssuedxx
      lsSQL = "UPDATE SP_Branch_Order_Detail SET" & _
                  "  nIssuedxx = nIssuedxx + " & lnIssuedxx & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
               " WHERE sTransNox = " & strParm(loRS("sTransNox")) & _
                  " AND sPartsIDx = " & strParm(loRS("sPartsIDx")) & _
                  " AND nEntryNox = " & loRS("nEntryNox")
                  
      If p_oAppDrivr.Execute(lsSQL, "SP_Branch_Order_Detail", _
                  p_sBranchCd, p_oBaseClas.Master("sDestinat")) <= 0 Then
         MsgBox lsSQL & vbCrLf & "Unable to Update Spareparts Pending Branch Order Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If lnQuantity = 0 Then Exit Do
      loRS.MoveNext
   Loop Until loRS.EOF

   addIssuedOrder = True
   
endProc:
   Set loRS = Nothing

   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function delIssuedOrder(lsPartsIDx As String, _
      lnQuantity As Integer) As Boolean
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnIssuedxx As Long
   
   lsProcName = "delIssuedOrder"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   delIssuedOrder = False
   
   With p_oBaseClas
      lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", a.nEntryNox" & _
                  ", a.sPartsIDx" & _
                  ", a.nIssuedxx" & _
               " FROM SP_Branch_Order_Detail a" & _
                  ", SP_Branch_Order_Master b" & _
               " WHERE a.sTransNox = b.sTransNox" & _
                  " AND a.cCanceled = " & strParm(xeNo) & _
                  " AND b.cTranStat = " & strParm(xeStatePosted) & _
                  " AND LEFT(b.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = " & strParm(.MasterOriginal("sDestinat")) & _
                  " AND a.sPartsIDx = " & strParm(lsPartsIDx) & _
                  " AND a.nIssuedxx > 0" & _
               " ORDER BY a.sTransNox DESC"
      
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If loRS.EOF Then
         delIssuedOrder = True
         GoTo endProc
      End If
      
      Do
         lnIssuedxx = loRS("nIssuedxx")
         
         If lnQuantity < lnIssuedxx Then lnIssuedxx = lnQuantity
         
         lnQuantity = lnQuantity - lnIssuedxx
         lsSQL = "UPDATE SP_Branch_Order_Detail SET" & _
                     " nIssuedxx = nIssuedxx - " & lnIssuedxx & _
                  " WHERE sTransNox = " & strParm(loRS("sTransNox")) & _
                     " AND sPartsIDx = " & strParm(loRS("sPartsIDx")) & _
                     " AND nEntryNox = " & loRS("nEntryNox")
                     
         If p_oAppDrivr.Execute(lsSQL, "SP_Branch_Order_Detail", _
                     p_sBranchCd, .Master("sDestinat")) <= 0 Then
            MsgBox "Unable to Update Spareparts Pending Branch Order Info!!!", vbCritical, "Warning"
            GoTo endProc
         End If
         
         If lnQuantity = 0 Then Exit Do
         loRS.MoveNext
      Loop Until loRS.EOF
  End With
   
   delIssuedOrder = True
   
endProc:
   Set loRS = Nothing

   Exit Function
errProc:
    ShowError lsProcName & "( " & lnQuantity _
                        & ", " & lsPartsIDx & " )"
End Function

Private Function delSPTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsProcName As String
   
   lsProcName = "delSPTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   
   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc
      
      For pnCtr = 0 To .ItemCount - 1
         .Detail(pnCtr, "sPartsIDx") = p_oBaseClas.Detail(pnCtr, "sPartsIDx")
         .Detail(pnCtr, "nQtyOnHnd") = p_oBaseClas.Detail(pnCtr, "nQtyOnHnd")
         .Detail(pnCtr, "nResvOrdr") = p_oBaseClas.Detail(pnCtr, "nResvOrdr")
         .Detail(pnCtr, "nBackOrdr") = p_oBaseClas.Detail(pnCtr, "nBackOrdr")
         .Detail(pnCtr, "nLedgerNo") = p_oBaseClas.Detail(pnCtr, "nLedgerNo")
         
      Next
      
      Call clearTempRec
      
      ' remove branch order issuance
      If saveIssuedOrder() = False Then GoTo endProc
            
      ' if a serial is not successfully updated, cancel saving.
      If .Delivery(p_oBaseClas.Master("sTransNox"), _
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

Private Sub createTempTable()
   Dim lsProcName As String
   
   lsProcName = "createTempTable"
   'On Error Goto errProc
   
   Set p_oDetail = New Recordset
   With p_oDetail
      .Fields.Append "sPartsIDx", adVarChar, 25
      .Fields.Append "nOldQtyxx", adInteger, 4
      .Fields.Append "nNewQtyxx", adInteger, 4
      .Open
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub addTempRec(lsPartsIDx As String, _
      lnOldQtyxx As Integer, _
      lnNewQtyxx As Integer)
   Dim lsProcName As String
   
   lsProcName = "addTempRec"
   'On Error Goto errProc

   With p_oDetail
      If .RecordCount = 0 Then
         .AddNew
         .Fields("sPartsIDx") = lsPartsIDx
         .Fields("nOldQtyxx") = 0
         .Fields("nNewQtyxx") = 0
      Else
         .MoveFirst
         .Find "sPartsIDx = " & strParm(lsPartsIDx), 0, adSearchForward
         
         If .EOF Then
            .AddNew
            .Fields("sPartsIDx") = lsPartsIDx
            .Fields("nOldQtyxx") = 0
            .Fields("nNewQtyxx") = 0
         End If
      End If
      .Fields("nOldQtyxx") = .Fields("nOldQtyxx") + lnOldQtyxx
      .Fields("nNewQtyxx") = .Fields("nNewQtyxx") + lnNewQtyxx
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub clearTempRec()
   Dim lsProcName As String
   
   lsProcName = "clearTempRec"
   'On Error Goto errProc

   With p_oDetail
      If Not .EOF Then .MoveFirst
      Do While .EOF = False
         .Fields("nNewQtyxx") = 0
         
         .MoveNext
      Loop
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub Class_Initialize()
   p_nTranStat = xeStateUnknown
   pbShowMsg = True
End Sub

Private Sub Class_Terminate()
   Set p_oBaseClas = Nothing
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_InitMaster"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
      
   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sDestinat") = ""
      .Master("sRequestx") = ""
      .Master("sRemarksx") = ""
      .Master("sReceived") = ""
      .Master("dReceived") = p_oAppDrivr.ServerDate
      .Master("sApproved") = ""
      .Master("nGrossAmt") = 0
      .Master("nDiscount") = 0
      .Master("nTranTotl") = 0
      .Master("sReferNox") = ""
      .Master("sSourceCd") = ""
      .Master("cTranStat") = xeStateOpen
      .Master("nEntryNox") = 1

      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "sTypeCode") = ""
      .Detail(0, "sModelNme") = ""
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nUnitPrce") = 0#
      .Detail(0, "nQuantity") = 0
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sPartsIDx") = ""
      .Detail(0, "nResvOrdr") = 0
      .Detail(0, "nBackOrdr") = 0
      .Detail(0, "nLedgerNo") = 0
   End With

   Call createTempTable
   
endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   Call createTempTable
   
   With p_oBaseClas
      p_sAddressx = .Master("xAddressx")
      For pnCtr = 0 To .ItemCount - 1
         .Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "xQtyOnHnd")
         
         Call addTempRec(.Detail(pnCtr, "sPartsIDx"), .Detail(pnCtr, "nQuantity"), 0)
      Next
   End With
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = Not pbModified
End Sub

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "p_oBaseClas_SaveOthers"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   Cancel = True

   With p_oBaseClas
      Set loSPTrans = New clsSPInventoryTrans
      With loSPTrans
         .Branch = p_sBranchCd
         Set .AppDriver = p_oAppDrivr
         If .InitTransaction() = False Then GoTo endProc
      End With
      
      Call clearTempRec
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sPartsIDx") = "" Then Exit For
         
         loSPTrans.Detail(pnCtr, "sPartsIDx") = .Detail(pnCtr, "sPartsIDx")
         loSPTrans.Detail(pnCtr, "nQuantity") = .Detail(pnCtr, "nQuantity")
         loSPTrans.Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
         loSPTrans.Detail(pnCtr, "nResvOrdr") = .Detail(pnCtr, "nResvOrdr")
         loSPTrans.Detail(pnCtr, "nBackOrdr") = .Detail(pnCtr, "nBackOrdr")
         loSPTrans.Detail(pnCtr, "nReorderx") = .Detail(pnCtr, "nReorderx")
'         loSPTrans.Detail(pnCtr, "nLedgerNo") = .Detail(pnCtr, "nLedgerNo")
         loSPTrans.Detail(pnCtr, "nLedgerNo") = IIf(.Detail(pnCtr, "nLedgerNo") = "", 0, .Detail(pnCtr, "nLedgerNo"))
            
         Call addTempRec(.Detail(pnCtr, "sPartsIDx"), 0, .Detail(pnCtr, "nQuantity"))
      Next
      
      ' update branch order info
      If saveIssuedOrder() = False Then GoTo endProc
      
      ' if a spareparts is not successfully updated, cancel saving.
      If loSPTrans.Delivery(.Master("sTransNox"), _
            .Master("dTransact"), _
            .EditMode) = False Then
         GoTo endProc
      End If
      
      ' check if reorder level of some parts has reached
      If loSPTrans.ReOrderParts Then
         MsgBox "Some Parts Included in the Transaction has Reached it Reorder Level Limit!!!" & _
                  vbCrLf & vbCrLf & "Please Inform the Spareparts Manager on this Alert!!!", _
                  vbInformation, "Notice"
      End If
      
      ' after saving

   End With
   
   Cancel = False

endProc:
   Set loSPTrans = Nothing
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillCancel(Cancel As Boolean)
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_WillCancel"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   
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
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   
   Cancel = Not delSPTransaction

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_WillSave"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error Goto errProc
   Cancel = True

   pbModified = True
   With p_oBaseClas
      If .isMasterModify = False And .isDetailModify = False Then
         pbModified = False
         Cancel = False
         GoTo endProc
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
      Next
   End With
   
   Cancel = False

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub
