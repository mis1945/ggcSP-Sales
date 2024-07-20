VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls2HSPTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     2H Job Order Class Object
'
' Copyright 2013 and beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  XerSys [ 11/26/2013 05:23 pm ]
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxe2HBranch As String = "M0W1"
Private Const pxeMODULENAME As String = "cls2HSPTransfer"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1

Private p_sBranchCd As String
Private p_sAddressx As String
Private p_nTranStat As Integer
Private p_sSystemCd As String

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

Property Let SystemCd(Value As String)
   p_sSystemCd = Value
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
   Case 2, "sbarrcode"
      getDetail Row, 2, Value, False
   Case 3, "sdescript"
      getDetail Row, 3, Value, False
   Case 6, 7, 8, "nquantity", "nissueqty", "cusedpart"
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
   Case 2, "sdestinat"
      Call getMaster(Value, False)
   Case 4, "srefernox"
      Call getJobOrder(Value, False)
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
End Property

Property Let TransStatus(ByVal Value As Integer)
   p_nTranStat = Value
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property

   ItemCount = p_oBaseClas.ItemCount
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String, lsCondition As String

   lsProcName = "InitTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode

   Set p_oBaseClas = New clsMasterDetail

   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "SP_2H_Transfer_Master"
      .DetailTable = "SP_2H_Transfer_Detail"
      .DisplayMessage = pbShowMsg

      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDestinat" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                        ", a.sReferNox" & _
                        ", a.nEntryNox" & _
                        ", a.cTranStat" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", (SELECT sBranchNm FROM Branch WHERE sBranchCd = LEFT(a.sTransNox, 4)) xSourcexx" & _
                        ", a.sRemarksx" & _
                     " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN 2H_JobOrder_Master e" & _
                              " ON a.sReferNox = e.sTransNox" & _
                              " AND LEFT(a.sReferNox, " & Len(p_oAppDrivr.BranchCode) & ") = " & strParm("M0W1") & _
                        ", Branch b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE a.sDestinat = b.sBranchCd" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx"

      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", d.sBarrCode sOrigBrCd" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", g.sModelNme" & _
                        ", f.nQtyOnHnd" & _
                        ", c.nQuantity" & _
                        ", a.nQuantity nIssueQty" & _
                        ", a.cUsedPart" & _
                        ", d.sDescript sOrigDesc" & _
                        ", a.dModified" & _
                        ", a.sTransNox" & _
                        ", a.sPartsIDx" & _
                        ", a.sOrigPart"
                        
      .DetailQuery = .DetailQuery & _
                        ", h.nQtyOnHnd nUsedQOHx" & _
                        ", b.nSelPrice" & _
                        ", f.nLedgerNo" & _
                        ", f.nResvOrdr" & _
                        ", f.nBackOrdr"
                        
'      .DetailQuery = .DetailQuery & _
'                        " FROM (" & _
'                               " SELECT xx.sTransNox," & _
'                                 " xx.nEntryNox," & _
'                                 " xx.nQuantity," & _
'                                 " xx.cUsedPart," & _
'                                 " xx.dModified," & _
'                                 " xx.sPartsIDx," & _
'                                 " xx.sOrigPart," & _
'                                 " yy.sReferNox" & _
'                               " FROM SP_2H_Transfer_Detail xx" & _
'                                  ",   SP_2H_Transfer_Master yy" & _
'                               " WHERE xx.sTransNox = yy.sTransNox ) a" & _
'                              " LEFT JOIN 2H_JobOrder_Detail c" & _
'                                " ON a.sReferNox = c.sTransNox" & _
'                                  " AND a.sPartsIDx = c.sPartsIDx" & _
'                              " LEFT JOIN Spareparts d" & _
'                                " ON a.sOrigPart = d.sPartsIDx" & _
'                              " LEFT JOIN SP_Inventory e" & _
'                                " ON d.sPartsIDx = e.sPartsIDx" & _
'                                  " AND e.sBranchCd = " & strParm(p_sBranchCd)

      .DetailQuery = .DetailQuery & _
                        " FROM SP_2H_Transfer_Master i" & _
                              " LEFT JOIN SP_2H_Transfer_Detail a" & _
                                 " ON a.sTransNox = i.sTransNox" & _
                              " LEFT JOIN 2H_JobOrder_Detail c" & _
                                 " ON c.sTransNox = i.sReferNox" & _
                                  " AND a.sPartsIDx = c.sPartsIDx" & _
                              " LEFT JOIN Spareparts d" & _
                                 " ON a.sOrigPart = d.sPartsIDx" & _
                              " LEFT JOIN SP_Inventory e" & _
                                 " ON d.sPartsIDx = e.sPartsIDx" & _
                                    " AND e.sBranchCd = " & strParm(p_sBranchCd)

      .DetailQuery = .DetailQuery & _
                                 ", Spareparts b" & _
                             " LEFT JOIN SP_Inventory f" & _
                               " ON b.sPartsIDx = f.sPartsIDx" & _
                                 " AND f.sBranchCd = " & strParm(p_sBranchCd) & _
                             " LEFT JOIN SP_Model g" & _
                               " ON b.sModelIDx = g.sModelIDx" & _
                             " LEFT JOIN SP_Salvage h" & _
                               " ON b.sPartsIDx = h.sPartsIDx" & _
                                 " AND h.sBranchCd = " & strParm(p_sBranchCd) & _
                           " WHERE a.sPartsIDx = b.sPartsIDx" & _
                           " ORDER BY a.nEntryNox"
                        
'      .DetailQuery = .DetailQuery & _
'                     " FROM " & .DetailTable & " a" & _
'                           " LEFT JOIN Spareparts d" & _
'                              " ON a.sOrigPart = d.sPartsIDx" & _
'                           " LEFT JOIN SP_Inventory e" & _
'                              " ON d.sPartsIDx = e.sPartsIDx" & _
'                                 " AND e.sBranchCd = " & strParm(p_sBranchCd) & _
'                        ", Spareparts b" & _
'                           " LEFT JOIN SP_Inventory f" & _
'                              " ON b.sPartsIDx = f.sPartsIDx" & _
'                                 " AND f.sBranchCd = " & strParm(p_sBranchCd) & _
'                           " LEFT JOIN SP_Model g" & _
'                              " ON b.sModelIDx = g.sModelIDx" & _
'                           " LEFT JOIN SP_Salvage h" & _
'                              " ON b.sPartsIDx = h.sPartsIDx" & _
'                                 " AND h.sBranchCd = " & strParm(p_sBranchCd) & _
'                        ", " & .MasterTable & " i" & _
'                           " LEFT JOIN 2H_JobOrder_Detail c" & _
'                              " ON i.sReferNox = c.sTransNox" & _
'                     " WHERE a.sPartsIDx = b.sPartsIDx" & _
'                        " AND a.sTransNox = i.sTransNox" & _
'                     " ORDER BY a.nEntryNox"

      .Reference = "a.sTransNox|a.sTransNox"

      pbInitTran = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sReferNox" & _
                           ", b.sBranchNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                              " LEFT JOIN 2H_JobOrder_Master c" & _
                                 " ON a.sReferNox = c.sTransNox" & _
                           ", Branch b" & _
                        " WHERE a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                           " AND a.sTransNox LIKE CONCAT(b.sBranchCd, '%')" & _
                        " ORDER BY b.sBranchNm" & _
                           ", a.dTransact"

      If p_nTranStat <> xeStateUnknown Then
         If p_nTranStat > xeStateUnknown Then
            lsCondition = "("
            For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
               lsCondition = lsCondition & " a.cTranStat = " & _
                           strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
            Next
            lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
         Else
            lsCondition = "a.cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      .BrowseQuery(0) = AddCondition(.BrowseQuery(0), lsCondition)
      Debug.Print .BrowseQuery(0)
      .BrowseColumn(0) = "sReferNox»sBranchNm»sTransNox»dTransact"
      .BrowseTitle(0) = "Reference No»Branch»Transaction No»Date"
      .BrowseFormat(0) = "@»@»@@-@@@@@@@@»MMMM DD, YYYY"

      ' Query statement for the branch destination field
      .BrowseQuery(2) = "SELECT a.sBranchCd" & _
                           ", a.sBranchNm" & _
                           ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" & _
                        " FROM Branch a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & xeRecStateActive & _
                           " AND a.sBranchCd <> " & strParm(p_sBranchCd) & _
                        " ORDER BY a.sBranchNm"
      .BrowseColumn(2) = "sBranchCd»sBranchNm»xAddressx"
      .BrowseTitle(2) = "Code»Branch Name»Address"

      ' Query statement for the branch destination field
      .BrowseQuery(3) = "SELECT DISTINCT a.sTransNox" & _
                           ", d.sEngineNo" & _
                           ", a.dTransact" & _
                           ", c.sBranchNm" & _
                        " FROM 2H_JobOrder_Master a" & _
                              " LEFT JOIN MC_Serial d" & _
                                 " ON a.sSerialID = d.sSerialID" & _
                           ", 2H_JobOrder_Detail b" & _
                           ", Branch c" & _
                        " WHERE a.sTransNox = b.sTransNox" & _
                           " AND a.sTransNox LIKE CONCAT(c.sBranchCd, '%')" & _
                           " AND a.cTranStat NOT IN (" & _
                              strParm(xeStateCancelled) & ", " & strParm(xeStatePosted) & ", " & strParm(xeStateUnknown) & " )" & _
                           " AND b.nQuantity > ( b.nIssueNew + b.nIssueUse )"
      .BrowseColumn(3) = "sTransNox»sEngineNo»dTransact»sBranchNm"
      .BrowseTitle(3) = "JO Number»Engine No»Date»Branch Name"
      .BrowseCriteria(3) = "a.sTransNox»d.sEngineNo»a.dTransact»c.sBranchNm"

      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sBarrCode" & _
                                 ", a.sDescript" & _
                                 ", d.sModelNme" & _
                                 ", b.nQtyOnHnd" & _
                                 ", c.nQtyOnHnd nUsedQOHx" & _
                                 ", a.sPartsIDx" & _
                                 ", a.nSelPrice" & _
                              " FROM Spareparts a" & _
                                    " LEFT JOIN SP_Salvage c" & _
                                       " ON a.sPartsIDx = c.sPartsIDx" & _
                                          " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                                    " LEFT JOIN SP_Model d" & _
                                       " ON a.sModelIDx = d.sModelIDx" & _
                                 ", SP_Inventory b" & _
                              " WHERE a.sPartsIDx = b.sPartsIDx" & _
                                 " AND b.sBranchCd = " & strParm(p_sBranchCd)
      .BrowseDetailColumn(0) = "sBarrCode»sDescript»sModelNme»nQtyOnHnd»nSelPrice"
      .BrowseDetailTitle(0) = "Bar Code»Description»Model»QOH»»Price"
      .BrowseDetailFormat(0) = "@»@»@»#0»#,##0.00"
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
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction

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

   If Not pbInitTran Then GoTo endProc
   With p_oBaseClas
      OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)
   End With

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
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

Function CloseTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "CloseTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then CloseTransaction = p_oBaseClas.CloseTransaction(p_oBaseClas.Master("sTransNox"))
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "CancelTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then CancelTransaction = p_oBaseClas.CancelTransaction()

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "DeleteTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

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
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sReferNox"
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

Function SearchByReference(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lsRecord As String
   Dim lsValue As String, lsField As String
   Dim lsDescript As String, lsCondition As String
   Dim lsQuery As String

   lsProcName = "SearchByReference"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc

   With p_oBaseClas
      lsQuery = .BrowseQuery(0)
      
      .BrowseQuery(0) = "SELECT DISTINCT a.sTransNox" & _
                           ", c.sBranchNm" & _
                           ", a.dTransact" & _
                           ", d.sEngineNo" & _
                        " FROM " & .MasterTable & " a" & _
                              " LEFT JOIN 2H_JobOrder_Master b" & _
                                 " ON a.sReferNox = b.sTransNox" & _
                              " LEFT JOIN MC_Serial d" & _
                                 " ON b.sSerialID = d.sSerialID" & _
                           ", Branch c" & _
                        " WHERE a.sDestinat = c.sBranchCd" & _
                           " AND a.sDestinat = " & strParm(p_oAppDrivr.BranchCode) & _
                        " ORDER BY c.sBranchNm" & _
                           ", a.dTransact"

'      .BrowseQuery(0) = "SELECT DISTINCT a.sTransNox" & _
'                           ", c.sBranchNm" & _
'                           ", a.dTransact" & _
'                           ", d.sEngineNo" & _
'                        " FROM " & .MasterTable & " a" & _
'                           ", 2H_JobOrder_Master b" & _
'                              " LEFT JOIN MC_Serial d" & _
'                                 " ON b.sSerialID = d.sSerialID" & _
'                           ", Branch c" & _
'                        " WHERE a.sReferNox = b.sTransNox" & _
'                           " AND a.sReferNox LIKE CONCAT(c.sBranchCd, '%')" & _
'                        " ORDER BY c.sBranchNm" & _
'                           ", a.dTransact"

      If p_nTranStat <> xeStateUnknown Then
         If p_nTranStat > xeStateUnknown Then
            lsCondition = "("
            For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
               lsCondition = lsCondition & " a.cTranStat = " & _
                           strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
            Next
            lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
         Else
            lsCondition = "a.cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      
      .BrowseQuery(0) = AddCondition(.BrowseQuery(0), lsCondition)
      Debug.Print .BrowseQuery(0)
      .BrowseColumn(0) = "sTransNox»sEngineNo»sBranchNm»dTransact"
      .BrowseTitle(0) = "Transsaction No»Engine No»Branch»Date"
      .BrowseFormat(0) = "@@-@@@@@@@@»@»@»MMMM DD, YYYY"
   End With

   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "d.sEngineNo"
      lsDescript = "d.sEngineNo"
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "c.sBranchNm"
            lsDescript = "c.sBranchNm"
         End If
      End If
   End If

   SearchByReference = p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript)

endProc:
   p_oBaseClas.BrowseQuery(0) = lsQuery

   Exit Function
errProc:
    ShowError lsProcName & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function PostTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String

   lsProcName = "PostTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then PostTransaction = p_oBaseClas.PostTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function AddDetail() As Boolean
   With p_oBaseClas
      AddDetail = .AddDetail

      pnCtr = .ItemCount - 1
      .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      .Detail(pnCtr, "sOrigBrCd") = ""
      .Detail(pnCtr, "sBarrCode") = ""
      .Detail(pnCtr, "sDescript") = ""
      .Detail(pnCtr, "nQtyOnHnd") = 0
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "nIssueQty") = 0
      .Detail(pnCtr, "cUsedPart") = xeNo
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sPartsIDx") = ""
      .Detail(pnCtr, "sOrigPart") = ""
      .Detail(pnCtr, "nUsedQOHx") = 0
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
   Case 2, "sbarrcode"
      getDetail Row, 2, Value, True
   Case 3, "sdescript"
      getDetail Row, 3, Value, True
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

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "sdestinat"
      Call getMaster(Value, True)
   Case 4, "srefernox"
      Call getJobOrder(Value, True)
   End Select
   SearchMaster = True
End Function

Function UpdateTransaction() As Boolean
   With p_oBaseClas
      .EditMode = xeModeUpdate
   End With
   UpdateTransaction = True
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Function AcceptDelivery(dReceived As Date) As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim loSPSalvg As clsSPSlvgInvTrans
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String

   Dim lnNew As Integer
   Dim lnUse As Integer

   lsProcName = "AcceptDelivery"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   AcceptDelivery = False

   If Not pbInitTran Then GoTo endProc

   With p_oBaseClas
      If .Master("sTransNox") = "" Then GoTo endProc

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

      lsSQL = "SELECT" & _
                  "  a.sPartsIDx" & _
                  ", a.sOrigPart" & _
                  ", a.nQuantity" & _
                  ", a.cUsedPart" & _
                  ", b.nQtyOnHnd" & _
               " FROM " & .DetailTable & " a" & _
                     " LEFT JOIN SP_Inventory b" & _
                        " ON a.sPartsIDx = b.sPartsIDx" & _
                           " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
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

      Set loSPSalvg = New clsSPSlvgInvTrans
      With loSPSalvg
         Set .AppDriver = p_oAppDrivr
         .Branch = p_oBaseClas.Master("sDestinat")
         If .InitTransaction() = False Then GoTo endProc
      End With
      p_oAppDrivr.BeginTrans

      Do Until loRS.EOF()
         If loRS("cUsedPart") = xeYes Then
            loSPSalvg.Detail(lnUse, "sPartsIDx") = loRS("sPartsIDx")
            loSPSalvg.Detail(lnUse, "nQuantity") = loRS("nQuantity")
            lnUse = lnUse + 1

            lsSQL = ", nIssueUse = " & loRS("nQuantity")
         Else
            loSPTrans.Detail(lnNew, "sPartsIDx") = loRS("sPartsIDx")
            loSPTrans.Detail(lnNew, "nQuantity") = loRS("nQuantity")
            lnNew = lnNew + 1

            lsSQL = ", nIssueNew = " & loRS("nQuantity")
         End If

         If loRS("sOrigPart") <> "" Then
            lsSQL = lsSQL & ", sIssParts = " & strParm(loRS("sPartsIDx"))
         End If
         
         If .Master("sReferNox") <> "" Then
            lsSQL = "UPDATE 2H_JobOrder_Detail" & _
                     " SET " & Trim(Mid(lsSQL, 2)) & _
                     " WHERE sTransNox = " & strParm(.Master("sReferNox")) & _
                        " AND sPartsIDx = " & strParm(IIf(loRS("sOrigPart") = "", loRS("sPartsIDx"), loRS("sOrigPart")))
   
            If p_oAppDrivr.Execute(lsSQL, "2H_JobOrder_Detail", p_sBranchCd, _
                        Left(.Master("sTransNox"), 4)) <= 0 Then
               MsgBox "Unable to Update Delivery Transaction!!!", vbCritical, "Warning"
               GoTo endProcWithRoll
            End If
         End If
         loRS.MoveNext
      Loop

      ' now check if it still needs to be save
      If lnNew > 0 Then
         If loSPTrans.XHAcceptDelivery(.Master("sTransNox"), _
               dReceived, _
               xeModeAddNew) = False Then
            GoTo endProcWithRoll
         End If
      End If

      If lnUse > 0 Then
         If loSPSalvg.XHAcceptDelivery(.Master("sTransNox"), _
               dReceived, _
               xeModeAddNew) = False Then
            GoTo endProcWithRoll
         End If
      End If

      lsSQL = "UPDATE " & .MasterTable & _
               " SET sReceived = " & strParm(p_oAppDrivr.UserID) & _
                  ", dReceived = " & dateParm(dReceived) & _
                  ", cTranStat = " & strParm(xeStatePosted) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sTransNox = " & strParm(.Master("sTransNox"))

      If p_oAppDrivr.Execute(lsSQL, .MasterTable, p_sBranchCd, _
                  Left(.Master("sTransNox"), 4)) <= 0 Then
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
   'On Error GoTo errProc

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
   'On Error GoTo errProc

   With p_oBaseClas
      If lsValue <> "" Then
         If lnIndex = 2 Then     ' Bar Code
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

      If lasDetail(0) = .Detail(lnRow, "sBarrCode") Then GoTo endProc

      ' move original barcode to OrigPart
      If .Detail(lnRow, "sOrigPart") = "" Then
         .Detail(lnRow, "sOrigPart") = .Detail(lnRow, "sPartsIDx")
         .Detail(lnRow, "sOrigBrCd") = .Detail(lnRow, "sBarrCode")
         .Detail(lnRow, "sOrigDesc") = .Detail(lnRow, "sDescript")
      End If

      .Detail(lnRow, "sBarrCode") = lasDetail(0)
      .Detail(lnRow, "sDescript") = lasDetail(1)
      .Detail(lnRow, "sModelNme") = lasDetail(2)
      .Detail(lnRow, "nQtyOnHnd") = IIf(lasDetail(3) = "", 0, CLng(lasDetail(3)))
      .Detail(lnRow, "nUsedQOHx") = IIf(lasDetail(4) = "", 0, CLng(IIf(lasDetail(4) = "", 0, lasDetail(4))))
      .Detail(lnRow, "nSelPrice") = CDbl(lasDetail(6))
      .Detail(lnRow, "sPartsIDx") = lasDetail(5)
   End With

endProc:
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)

   Exit Sub
endWithClear:
   With p_oBaseClas
      If .Detail(lnRow, "sOrigPart") = "" Then
         .Detail(lnRow, "sBarrCode") = ""
         .Detail(lnRow, "sDescript") = ""
         .Detail(lnRow, "sModelNme") = ""
      Else
         .Detail(lnRow, "sPartsIDx") = .Detail(lnRow, "sOrigPart")
         .Detail(lnRow, "sBarrCode") = .Detail(lnRow, "sOrigBrCd")
         .Detail(lnRow, "sDescript") = .Detail(lnRow, "sOrigDesc")
         .Detail(lnRow, "sOrigPart") = ""
         .Detail(lnRow, "sOrigBrCd") = ""
         .Detail(lnRow, "sOrigDesc") = ""
      End If

      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "sModelNme") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nUsedQOHx") = 0
      .Detail(lnRow, "nUnitPrce") = 0#
      .Detail(lnRow, "sPartsIDx") = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lnIndex _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"

End Sub

Private Function getJobOrder(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lnCtr As Integer
   Dim lbAdd As Boolean

   lsProcName = "loadRecord"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   With p_oBaseClas
      If lsValue <> "" Then
         If lsValue = .Master("sReferNox") Then GoTo endProc

         ' Check if record was previously loaded
         If Trim(.Master("sReferNox")) <> "" Then
            If Trim(.Master("sReferNox")) <> lsValue Then
               If MsgBox("Loading different transaction will override changes made." & vbCrLf & _
                        "Continue Loading Transaction Anyway!", vbQuestion + vbYesNo) <> vbYes Then GoTo endProc
            Else
               GoTo endProc
            End If
         End If

         If lbSearch Then
            lsMaster = "d.sEngineNo LIKE " & strParm("%" & Trim(lsValue))
         Else
            lsMaster = "d.sEngineNo = " & strParm(Trim(lsValue))
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If

      lsMaster = .getMaster(3, lsMaster)
      If lsMaster = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")

      .Master(4) = Trim(lasMaster(0))

      Call loadRecord
   End With

   getJobOrder = True

endProc:
   RaiseEvent MasterRetrieved(4)

   Exit Function
endWithClear:
   p_oBaseClas.Master(4) = ""
   GoTo endProc
errProc:
   ShowError lsProcName & " ( " & " ) "
End Function

Private Function loadRecord() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lnCtr As Integer
   Dim lbAdd As Boolean

   lsProcName = "loadRecord"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   With p_oBaseClas
      lsSQL = "SELECT a.sPartsIDx" & _
                  ", b.sBarrCode" & _
                  ", b.sDescript" & _
                  ", e.sModelNme" & _
                  ", a.nQuantity - ( a.nIssueNew + a.nIssueUse ) nQuantity" & _
                  ", c.nQtyOnHnd" & _
                  ", d.nQtyOnHnd nUsedQOHx" & _
                  ", c.nResvOrdr" & _
                  ", c.nBackOrdr" & _
                  ", c.nLedgerNo" & _
               " FROM 2H_JobOrder_Detail a" & _
                  ", Spareparts b" & _
                     " LEFT JOIN SP_Inventory c" & _
                        " ON b.sPartsIDx = c.sPartsIDx" & _
                           " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                     " LEFT JOIN SP_Salvage d" & _
                        " ON b.sPartsIDx = d.sPartsIDx" & _
                           " AND d.sBranchCd = " & strParm(p_sBranchCd) & _
                     " LEFT JOIN SP_Model e" & _
                        " ON b.sModelIDx = e.sModelIDx" & _
               " WHERE a.sPartsIDx = b.sPartsIDx" & _
                  " AND a.sTransNox = " & strParm(.Master("sReferNox")) & _
                  " AND a.nQuantity > ( a.nIssueNew + a.nIssueUse )"
      Set loRS = New Recordset
      Debug.Print lsSQL
      loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText

      If loRS.EOF Then GoTo endProc

      lnCtr = 0
      Do Until loRS.EOF
         lbAdd = .Detail(lnCtr, "sPartsIDx") <> ""
         .Detail(lnCtr, "sPartsIDx") = loRS("sPartsIDx")
         .Detail(lnCtr, "sBarrCode") = loRS("sBarrCode")
         .Detail(lnCtr, "sDescript") = loRS("sDescript")
         .Detail(lnCtr, "sModelNme") = loRS("sModelNme")
         .Detail(lnCtr, "nQuantity") = loRS("nQuantity")
         .Detail(lnCtr, "nQtyOnHnd") = loRS("nQtyOnHnd")
         .Detail(lnCtr, "nUsedQOHx") = loRS("nUsedQOHx")
         .Detail(lnCtr, "nResvOrdr") = loRS("nResvOrdr")
         .Detail(lnCtr, "nBackOrdr") = loRS("nBackOrdr")
         .Detail(lnCtr, "nLedgerNo") = loRS("nLedgerNo")
         If lbAdd Then
            Call AddDetail
         Else
            .Detail(lnCtr, "sOrigBrCd") = ""
            .Detail(lnCtr, "nIssueQty") = 0
            .Detail(lnCtr, "cUsedPart") = xeNo
         End If

         lnCtr = lnCtr + 1
         loRS.MoveNext
      Loop

      Do Until lnCtr > .ItemCount
         If .Detail(lnCtr, "sPartsIDx") = "" Then Exit Do

         Call DeleteDetail(lnCtr)
         lnCtr = lnCtr + 1
      Loop
      loadRecord = True
   End With
endProc:
   Exit Function
errProc:
   ShowError lsProcName & " ( " & " ) "
End Function

Private Function saveSPTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim loSPSalvg As clsSPSlvgInvTrans
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lnQty As Integer

   lsProcName = "saveSPTransaction"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If Not pbInitTran Then GoTo endProc
   If p_oBaseClas.ItemCount = 1 And IFNull(p_oBaseClas.Detail(0, "sPartsIDx")) = "" Then
      saveSPTransaction = True
      GoTo endProc
   End If

   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc
        
      lnCtr = 0
      For pnCtr = 0 To p_oBaseClas.ItemCount - 1
         If p_oBaseClas.Detail(pnCtr, "sPartsIDx") = "" Then Exit For
         
         .Detail(pnCtr, "sPartsIDx") = p_oBaseClas.Detail(pnCtr, "sPartsIDx")
         .Detail(pnCtr, "sReplacID") = Trim(IFNull(p_oBaseClas.Detail(pnCtr, "sOrigPart"), ""))
         .Detail(pnCtr, "nQuantity") = p_oBaseClas.Detail(pnCtr, "nQuantity")
         lnCtr = lnCtr + 1
      Next

      ' check if record exist for new items
      If lnCtr > 0 Then
         If Not .XHDelivery(p_oBaseClas.Master("sTransNox"), p_oBaseClas.Master("dTransact"), xeModeAddNew) Then
            GoTo endProc
         End If
      End If
   End With

   Set loSPSalvg = New clsSPSlvgInvTrans
   With loSPSalvg
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc

      lnCtr = 0
      For pnCtr = 0 To p_oBaseClas.ItemCount - 1
         If p_oBaseClas.Detail(pnCtr, "cUsedPart") = xeYes Then
            .Detail(lnCtr, "sPartsIDx") = p_oBaseClas.Detail(pnCtr, "sPartsIDx")
            .Detail(lnCtr, "nQuantity") = p_oBaseClas.Detail(pnCtr, "nIssueQty")
            lnCtr = lnCtr + 1
         End If
      Next

      ' check if record exist for new items
      If lnCtr > 0 Then
         If Not .XHDelivery(p_oBaseClas.Master("sTransNox"), p_oBaseClas.Master("dTransact"), xeModeAddNew) Then
            GoTo endProc
         End If
      End If
   End With

   saveSPTransaction = True

endProc:
   Set loSPTrans = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function delSPTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim loSPSalvg As clsSPSlvgInvTrans
   Dim lsProcName As String
   Dim lnCtr As Integer

   lsProcName = "delSPTransaction"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc


   If Not pbInitTran Then GoTo endProc
   If p_oBaseClas.ItemCount = 1 And IFNull(p_oBaseClas.Detail(0, "sPartsIDx")) = "" Then
      delSPTransaction = True
      GoTo endProc
   End If

   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      If .InitTransaction() = False Then GoTo endProc

      For pnCtr = 0 To p_oBaseClas.ItemCount - 1
         If p_oBaseClas.Detail(pnCtr, "cUsedPart") = xeNo Then
            If p_oBaseClas.Detail(pnCtr, "sIssParts") <> "" Then
               .Detail(lnCtr, "sPartsIDx") = p_oBaseClas.Detail(pnCtr, "sPartsIDx")
               .Detail(lnCtr, "nQuantity") = p_oBaseClas.Detail(pnCtr, "nIssueNew")
               lnCtr = lnCtr + 1
            End If
         End If
      Next

      ' check if record exist for new items
      If lnCtr > 0 Then
         If Not .Delivery(p_oBaseClas.Master("sTransNox"), p_oBaseClas.Master("dTransact"), xeModeDelete) Then
            GoTo endProc
         End If
      End If
   End With

   Set loSPSalvg = New clsSPSlvgInvTrans
   With loSPSalvg
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc

      lnCtr = 0
      For pnCtr = 0 To p_oBaseClas.ItemCount - 1
         If p_oBaseClas.Detail(pnCtr, "cUsedPart") = xeNo Then
            .Detail(lnCtr, "sPartsIDx") = p_oBaseClas.Detail(pnCtr, "sPartsIDx")
            .Detail(lnCtr, "nQuantity") = p_oBaseClas.Detail(pnCtr, "nIssueUse")
            lnCtr = lnCtr + 1
         End If
      Next

      ' check if record exist for new items
      If lnCtr > 0 Then
         If Not .Delivery(p_oBaseClas.Master("sTransNox"), p_oBaseClas.Master("dTransact"), xeModeDelete) Then
            GoTo endProc
         End If
      End If
   End With

   delSPTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

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
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sDestinat") = pxe2HBranch
      .Master("sReferNox") = ""
      .Master("nEntryNox") = 0
      .Master("cTranStat") = xeStateOpen

      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sOrigBrCd") = ""
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "sModelNme") = ""
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nQuantity") = 0
      .Detail(0, "nIssueQty") = 0
      .Detail(0, "cUsedPart") = xeNo
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sPartsIDx") = ""
      .Detail(0, "sOrigPart") = ""
   End With

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      p_sAddressx = .Master("xAddressx")

'      For lnCtr = 0 To .ItemCount - 1
'         If .Detail(lnCtr, "sIssuedxx") <> "" Then
'            .Detail("sDescript") = .Detail(lnCtr, "sIssuedxx")
'         End If
'      Next
   End With
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_SaveOthers"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   Cancel = Not saveSPTransaction

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
   
   Cancel = Not delSPTransaction

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillDelete(Cancel As Boolean)
   Dim loSPTrans As clsSPInventoryTrans
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_WillDelete"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = delSPTransaction

endProc:
   Set loSPTrans = Nothing
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsProcName As String
   Dim lsAppvID As String, lsAppvName As String, lsAppvCD As String
   Dim lnAppvRights As Integer
   Dim lbGetApproval As Boolean

   lsProcName = "p_oBaseClas_WillSave"
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

      .Master("nEntryNox") = .ItemCount
      
      lnAppvRights = p_oAppDrivr.UserLevel
      lsAppvID = p_oAppDrivr.UserID
      lbGetApproval = False
      
      If Trim(.Master("sReferNox")) = "" Then
         If p_oAppDrivr.UserLevel <> xeEngineer Then
            lbGetApproval = True
         End If
      End If
      
      If lbGetApproval Then
         If GetApproval(p_oAppDrivr, lnAppvRights, lsAppvID, lsAppvName, p_oAppDrivr.MenuName) = False Then GoTo endProc
         
         If lnAppvRights <> xeEngineer Then
            MsgBox "Approving Officer is not allowed to direct transfer!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End With

   Cancel = False
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub

