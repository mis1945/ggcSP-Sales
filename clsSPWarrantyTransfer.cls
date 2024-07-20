VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPWarrantyTransfer"
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
'     Warranty Spareparts Transfer Object
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
'  XerSys [ 01/31/2005 01:31 pm ]
'     Start creating this object.
'  Kalyptus [ 12/11/2006 10:08 am]
'     Implements the BugTracker
'  Kalyptus [ 04/02/2007 10:27 am]
'     Updated the p_oBase_Clas to conform it with the clsMasterDetail specification
'  XerSys [ 08/14/2007 05:12 pm ]
'     Update this object to adopt the new implementation of SP Inventory Transaction
'  XerSys [ 03/05/2014 10:35 am ]
'     Remove connection of spareparts to inventory of the receiving branch
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME = "clsSPWarrantyTransfer"
Private Const pxeWHouseID As String = "W2"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1

Private p_sBranchCd As String
Private p_sAddressx As String
Private p_nTranStat As Integer

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pbClosed As Boolean
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

Property Let TransStatus(ByVal Value As Integer)
   p_nTranStat = Value
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
   Case 1, 2, "sbarrcode", "sdescript"
      getDetail Row, Index, Value, False
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
   Case 2, "sdestinat"
      getMaster Value, False
   Case 3, "xaddressx"
      p_sAddressx = Value
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oBaseClas.ItemCount
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsCondition As String
   
   lsProcName = "InitTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail
   
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "SP_Warranty_Transfer_Master"
      .DetailTable = "SP_Warranty_Transfer_Detail"

      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDestinat" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                        ", a.sRemarksx" & _
                        ", a.cReceived" & _
                        ", a.dReceived" & _
                        ", a.sApproved" & _
                        ", a.cTranStat" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") xSourcexx" & _
                        ", a.nEntryNox" & _
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
                        ", d.sModelNme" & _
                        ", c.nQtyOnHnd" & _
                        ", a.nQuantity" & _
                        ", a.sWCRNoxxx" & _
                        ", IFNULL(c.nQtyOnHnd, 0) + a.nQuantity xQtyOnHnd" & _
                        ", a.dModified" & _
                        ", a.sTransNox" & _
                        ", a.sPartsIDx" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo" & _
                     " FROM " & .DetailTable & " a" & _
                           " LEFT JOIN SP_Inventory c" & _
                              " ON a.sPartsIDx = c.sPartsIDx" & _
                                 " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                        ", Spareparts b" & _
                              " LEFT JOIN SP_Model d" & _
                                 " ON b.sModelIDx = d.sModelIDx" & _
                     " WHERE a.sPartsIDx = b.sPartsIDx" & _
                     " ORDER BY a.nEntryNox"
   
      .Reference = "sTransNox|a.sTransNox"
      
      pbInitTran = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sBranchNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Branch b" & _
                        " WHERE a.sDestinat = b.sBranchCd" & _
                           " AND LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = " & strParm(p_sBranchCd) & _
                        " ORDER BY b.sBranchNm" & _
                           ", a.dTransact"
      
      .BrowseQuery(0) = AddCondition(.BrowseQuery(0), lsCondition)
                           
      .BrowseColumn(0) = "sTransNox»sBranchNm»dTransact"
      .BrowseCriteria(0) = "a.sTransNox»b.sBranchNm»a.dTransact"
      .BrowseTitle(0) = "Transaction No»Branch»Date"
      .BrowseFormat(0) = "@@@@-@@@@@@»@»MMMM DD, YYYY"
      
      ' Query statement for the branch destination field
      .BrowseQuery(2) = "SELECT" & _
                           "  a.sBranchCd" & _
                           ", a.sBranchNm" & _
                           ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" & _
                        " FROM Branch a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
                           " AND a.sBranchCd <> " & strParm(p_sBranchCd)
      .BrowseColumn(2) = "sBranchCd»sBranchNm»xAddressx"
      .BrowseTitle(2) = "Code»Branch Name»Address"
      
      .BrowseQuery(11) = "SELECT" & _
                           "  a.sBranchCd" & _
                           ", a.sBranchNm" & _
                           ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" & _
                        " FROM Branch a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & strParm(xeRecStateActive)
      .BrowseColumn(11) = "sBranchCd»sBranchNm»xAddressx"
      .BrowseTitle(11) = "Code»Branch Name»Address"
      
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sBarrCode" & _
                                 ", a.sDescript" & _
                                 ", e.sModelNme" & _
                                 ", b.nQtyOnHnd" & _
                                 ", a.sPartsIDx" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nReorderx" & _
                                 ", b.nLedgerNo" & _
                              " FROM Spareparts a" & _
                                    " LEFT JOIN SP_Model e" & _
                                       " ON a.sModelIDx = e.sModelIDx" & _
                                 ", SP_Inventory b" & _
                              " WHERE a.sPartsIDx = b.sPartsIDx" & _
                                 " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                                 " AND b.cRecdStat = " & strParm(xeRecStateActive)
      .BrowseDetailColumn(0) = "sBarrCode»sDescript»sModelNme»nQtyOnHnd"
      .BrowseDetailTitle(0) = "Bar Code»Description»Model»QtyOnHand"
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
   Debug.Print pxeMODULENAME & "." & lsProcName
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
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   p_oBaseClas.Master("nEntryNox") = p_oBaseClas.ItemCount
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
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
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
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String
   Dim lsMaster As String
   Dim lsQuery As String
   Dim lsCondition As String
   
   lsProcName = "SearchAcceptance"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   SearchAcceptance = False
   
   If pbInitTran = False Then GoTo endProc
   
   With p_oBaseClas
      lsMaster = .MasterQuery
      lsQuery = .BrowseQuery(0)
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDestinat" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                        ", a.sRemarksx" & _
                        ", a.cReceived" & _
                        ", a.dReceived" & _
                        ", a.sApproved" & _
                        ", a.cTranStat" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") xSourcexx" & _
                        ", a.nEntryNox" & _
                     " FROM " & .MasterTable & " a" & _
                        ", Branch b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE a.sTransNox LIKE CONCAT(b.sBranchCd, '%')" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx"
                        
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sBranchNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Branch b" & _
                        " WHERE a.sTransNox LIKE CONCAT(b.sBranchCd, '%')" & _
                           " AND a.sDestinat = " & strParm(p_sBranchCd) & _
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
      .BrowseQuery(0) = lsQuery
   End With

   Exit Function
errProc:
    ShowError lsProcName & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function SearchApprove(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String
   Dim lsMaster As String
   Dim lsQuery As String
   Dim lsCondition As String
   
   lsProcName = "SearchApprove"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   SearchApprove = False
   
   If pbInitTran = False Then GoTo endProc
   
   With p_oBaseClas
      lsMaster = .MasterQuery
      lsQuery = .BrowseQuery(0)
'
'      .MasterQuery = "SELECT" & _
'                        "  a.sTransNox" & _
'                        ", a.dTransact" & _
'                        ", a.sDestinat" & _
'                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
'                        ", a.sRemarksx" & _
'                        ", a.cReceived" & _
'                        ", a.dReceived" & _
'                        ", a.sApproved" & _
'                        ", a.cTranStat" & _
'                        ", a.sModified" & _
'                        ", a.dModified" & _
'                     " FROM " & .MasterTable & " a" & _
'                        ", Branch b" & _
'                        ", TownCity c" & _
'                        ", Province d" & _
'                     " WHERE a.sTransNox LIKE CONCAT(b.sBranchCd, '%')" & _
'                        " AND b.sTownIDxx = c.sTownIDxx" & _
'                        " AND c.sProvIDxx = d.sProvIDxx"
'
'      .BrowseQuery(0) = "SELECT" & _
'                           "  a.sTransNox" & _
'                           ", b.sBranchNm" & _
'                           ", a.dTransact" & _
'                        " FROM " & .MasterTable & " a" & _
'                           ", Branch b" & _
'                        " WHERE a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
'                           " AND a.sDestinat LIKE CONCAT(b.sBranchCd, '%')" & _
'                           " AND a.sApproved = ''" & _
'                        " ORDER BY b.sBranchNm" & _
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
   
   SearchApprove = p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript)
   
endProc:
   With p_oBaseClas
      .MasterQuery = lsMaster
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
   'On Error GoTo errProc
   
   If pbInitTran Then PostTransaction = p_oBaseClas.PostTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function ClosedTransaction(ByVal sTransNo As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   
   lsOldProc = "ClosedTransaction"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If Not pbInitTran Then GoTo endProc
   
   pbClosed = False
   p_oAppDrivr.BeginTrans
   With p_oBaseClas
      If .OpenTransaction(sTransNo) = False Then
         GoTo endProcWithRoll
      End If
      
      'Verify if Qty to transfer is okey
      For lnCtr = 0 To .ItemCount - 1
         If .Detail(lnCtr, "sPartsIDx") = "" Then Exit For
         If .Detail(lnCtr, "nQuantity") > .Detail(lnCtr, "nQtyOnHnd") Then
            MsgBox "Trans QTY is higher than QTY-ON-Hand.", vbCritical, "Warning"
            GoTo endProcWithRoll
         ElseIf .Detail(lnCtr, "nQuantity") < 1 Then
            MsgBox "Quantity Transferred is less than 1!", vbCritical, "Warning"
            GoTo endProcWithRoll
         End If
      Next
      
      'Save SP Transaction/Previously from ApproveTransaction
      If saveSPTransaction() = False Then
         GoTo endProcWithRoll
      End If
      
      .Master("cTranStat") = xeStateClosed
         
      'sApproved: Previously from ApproveTransaction
      lsSQL = "UPDATE " & .MasterTable & " SET" & _
                  "  sApproved = " & strParm(Decrypt(p_oAppDrivr.UserID)) & _
                  ", cTranStat = " & .Master("cTranStat") & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
               " WHERE sTransNox = " & strParm(.Master("sTransNox"))

      If p_oAppDrivr.Execute(lsSQL, .MasterTable, p_oAppDrivr.BranchCode) = 0 Then
         MsgBox "Unable to Close Transaction!!!", vbCritical, "Warning"
         GoTo endProcWithRoll
      End If

      p_oAppDrivr.CommitTrans
   End With
   pbClosed = True
   
   ClosedTransaction = True
   
endProc:
   Exit Function
endProcWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
   p_oAppDrivr.RollbackTrans
   ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function ApprovedTransaction(ByVal sTransNo As String) As Boolean
'   Dim lsOldProc As String
'   Dim lsSQL As String
'
'   lsOldProc = "ApprovedTransaction"
'   Debug.Print pxeMODULENAME & "." & lsOldProc
'   'On Error GoTo errProc
'
'   If Not pbInitTran Then GoTo endProc
'
'   p_oAppDrivr.BeginTrans
'   With p_oBaseClas
'      If .OpenTransaction(sTransNo) = False Then
'         GoTo endProcWithRoll
'      End If
'
'      If saveSPTransaction() = False Then
'         GoTo endProcWithRoll
'      End If
'
'      lsSQL = "UPDATE " & .MasterTable & " SET" & _
'                  "  sApproved = " & strParm(Decrypt(p_oAppDrivr.UserID)) & _
'                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
'               " WHERE sTransNox = " & strParm(.Master("sTransNox"))
'
'      If p_oAppDrivr.Execute(lsSQL, .MasterTable, p_sBranchCd) = 0 Then
'         MsgBox "Unable to Approved Transaction!!!", vbCritical, "Warning"
'         GoTo endProcWithRoll
'      End If
'
'      p_oAppDrivr.CommitTrans
'   End With
'
'   ApprovedTransaction = True
'
'endProc:
'   Exit Function
'endProcWithRoll:
'   p_oAppDrivr.RollbackTrans
'   GoTo endProc
'errProc:
'   p_oAppDrivr.RollbackTrans
'   ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Private Function saveSPTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsSQL As String
   Dim lsProcName As String

   lsProcName = "saveSPTransaction"
   'On Error GoTo errProc
   saveSPTransaction = False

   With p_oBaseClas
      Set loSPTrans = New clsSPInventoryTrans
      With loSPTrans
         .Branch = p_sBranchCd
         Set .AppDriver = p_oAppDrivr
         If .InitTransaction() = False Then GoTo endProc
      End With
      
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sPartsIDx") = "" Then Exit For
         
         loSPTrans.Detail(pnCtr, "sPartsIDx") = .Detail(pnCtr, "sPartsIDx")
         loSPTrans.Detail(pnCtr, "nQuantity") = .Detail(pnCtr, "nQuantity")
         loSPTrans.Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
         loSPTrans.Detail(pnCtr, "nResvOrdr") = .Detail(pnCtr, "nResvOrdr")
         loSPTrans.Detail(pnCtr, "nBackOrdr") = .Detail(pnCtr, "nBackOrdr")
         loSPTrans.Detail(pnCtr, "nLedgerNo") = .Detail(pnCtr, "nLedgerNo")
      Next
      
      ' if a spareparts is not successfully updated, cancel saving.
      If loSPTrans.WarrantyTransfer(.Master("sTransNox"), _
            .Master("dTransact"), _
            IIf(.EditMode = xeModeReady, xeModeAddNew, .EditMode)) = False Then
         GoTo endProc
      End If
      
      ' check if reorder level of some parts has reached
      If loSPTrans.ReOrderParts Then
         MsgBox "Some Parts Included in the Transaction has Reached it Reorder Level Limit!!!" & _
                  vbCrLf & vbCrLf & "Please Inform the Spareparts Manager on this Alert!!!", _
                  vbInformation, "Notice"
      End If
   End With
   
   saveSPTransaction = True

endProc:
   Set loSPTrans = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & "" & " )"
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

Function AddDetail() As Boolean
   With p_oBaseClas
      AddDetail = .AddDetail
   
      pnCtr = .ItemCount - 1
      .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      .Detail(pnCtr, "sBarrCode") = ""
      .Detail(pnCtr, "sDescript") = ""
      .Detail(pnCtr, "sModelNme") = ""
      .Detail(pnCtr, "nQtyOnHnd") = 0
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "sWCRNoxxx") = ""
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
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
   'On Error GoTo errProc
   SearchDetail = False

   If pbInitTran = False Then GoTo endProc

   Index = LCase(Index)
   Select Case Index
   Case 1, "sbarrcode"
      getDetail Row, Index, Value, True
   Case 2, "sdescript"
      getDetail Row, Index, Value, True
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
      getMaster Value, True
   End Select
   SearchMaster = True
End Function

Function UpdateTransaction() As Boolean
   With p_oBaseClas
      If .Master("cReceived") = xeYes Then Exit Function
         
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
   'On Error GoTo errProc
   
   If Not pbInitTran Then GoTo endProc
   
   With p_oBaseClas
      If .Master("sTransNox") = "" Then GoTo endProc
      
      'Test validity of the acceptance date
      '++++++++++++++++++++++
      'Validity #1: Date Transferred < Date Received
      If .Master("dTransact") < dReceived Then
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
               " WHERE a.sTransNox = " & strParm(.Master("sTransNox"))
      
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If loRS.RecordCount <> .ItemCount Then
         MsgBox "Data Discrepancy Detected!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      Set loSPTrans = New clsSPInventoryTrans
      With loSPTrans
         .Branch = p_sBranchCd
         Set .AppDriver = p_oAppDrivr
         If .InitTransaction() = False Then GoTo endProc
      End With
      
      p_oAppDrivr.BeginTrans
      For pnCtr = 0 To loRS.RecordCount - 1
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
      If loSPTrans.AcceptWarrantyTransfer(.Master("sTransNox"), _
            dReceived, _
            xeModeAddNew) = False Then
         GoTo endProcWithRoll
      End If
      
      lsSQL = "UPDATE " & .MasterTable & " SET" & _
                  "  cReceived = " & strParm(xeYes) & _
                  ", dReceived = " & dateParm(dReceived) & _
                  ", cDeliverx = " & strParm(xeStateClosed) & _
                  ", cTranStat = " & strParm(xeStatePosted) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
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

   lsProcName = "getMaster"
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
      
      RaiseEvent MasterRetrieved(2)
      RaiseEvent MasterRetrieved(3)
   End With

endProc:

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master(2) = ""
      p_sAddressx = ""
   End With
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)
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
         If pbClosed Then
            If .Detail(lnRow, "sBarrCode") <> Empty Then GoTo endProc
         End If
      
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
      .Detail(lnRow, "sModelNme") = lasDetail(2)
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(3)
      .Detail(lnRow, "sPartsIDx") = lasDetail(4)
      .Detail(lnRow, "nResvOrdr") = lasDetail(5)
      .Detail(lnRow, "nBackOrdr") = lasDetail(6)
      .Detail(lnRow, "nReorderx") = lasDetail(7)
      .Detail(lnRow, "nLedgerNo") = lasDetail(8)
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "sWCRNoxxx") = ""
   End With
   
endProc:
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "sModelNme") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "sWCRNoxxx") = ""
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

Private Function delSPTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsProcName As String
   
   lsProcName = "delSPTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
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
      
      ' if a serial is not successfully updated, cancel saving.
      If .WarrantyTransfer(p_oBaseClas.Master("sTransNox"), _
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
      .Master("sDestinat") = ""
      .Master("sRemarksx") = ""
      .Master("cReceived") = ""
      .Master("sApproved") = ""
      .Master("dReceived") = "01/01/1900"
      .Master("cTranStat") = xeStateOpen

      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "sModelNme") = ""
      .Detail(0, "nQuantity") = 0
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "sWCRNoxxx") = ""
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sPartsIDx") = ""
      .Detail(0, "nResvOrdr") = 0
      .Detail(0, "nBackOrdr") = 0
      .Detail(0, "nReorderx") = 0
      .Detail(0, "nLedgerNo") = 0
      
      p_sAddressx = Empty
   End With

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      p_sAddressx = .Master("xAddressx")
      For pnCtr = 0 To .ItemCount - 1
         .Detail(pnCtr, "nQtyOnHnd") = IIf(.Master("sApproved") = "", .Detail(pnCtr, "nQtyOnHnd"), .Detail(pnCtr, "xQtyOnHnd"))
      Next
      
      pbClosed = .Master("cTranStat") = xeStateClosed
   End With
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = Not pbModified
End Sub

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   Dim lsOldProc As String
   Dim lbCancel As Boolean

   lsOldProc = "p_oBaseClas_Others"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   Cancel = True
   
   With p_oBaseClas
      If .Master("cTranStat") = xeStateClosed Or .Master("cTranStat") = xeStatePosted Then
         If .Master("sApproved") <> "" Then
            If saveSPTransaction() = False Then GoTo endProc
         End If
      End If
   End With
   
   Cancel = False

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillCancel(Cancel As Boolean)
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_WillCancel"
   'On Error GoTo errProc
   
   Cancel = Not delSPTransaction()

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
   
   Cancel = Not delSPTransaction()

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

'This function has the same logic with the clsJobOrder so be sure
'to update this function here or of the job order
Function RequestItem() As Boolean
   Dim lsOldProc As String
   Dim loSPOrder As clsBranchOrder
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lnWCCtr As Integer
   Dim lnBOCtr As Integer
   Dim lbNew As Boolean
      
   lsOldProc = "RequestItem()"
   'Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
      
   'Search for the existence of a branch order for this record
   lsSQL = "SELECT sTransNox, cTranStat" & _
          " FROM SP_Branch_Order_Master" & _
          " WHERE sRequestx = " & strParm(p_oBaseClas.Master("sTransNox") & "-WC") & _
          " ORDER BY sTransNox DESC"
   Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   Debug.Print lsSQL
   
   If Not loRS.EOF Then
      If MsgBox("You have already created a request order! Continue anyway?", vbYesNo + vbInformation, "Request Confirmation") = vbNo Then
         GoTo endProc
      End If
   End If
   
   Set loSPOrder = New clsBranchOrder
   
   With loSPOrder
      'Initialize what we need to initialize here
      Set .AppDriver = p_oAppDrivr
      If .InitTransaction() = False Then GoTo endProc
      .DisplayMessage = False
      
      If loRS.EOF Then
         'Automatically create a new transaction if no record was found
         If .NewTransaction() = False Then GoTo endProc
         lbNew = True
      Else
         'Ask Nildz first if she has to print the Branch Order/If so then the following logic is okay
         'If we already have an open request
         If loRS("cTranStat") = "0" Then
            If .OpenTransaction(loRS("sTransNox")) = False Then GoTo endProc
            If .UpdateTransaction = False Then GoTo endProc
         Else
            'Create a new request if previous request are printed/closed or confirmed/posted or cancelled
            If .NewTransaction() = False Then GoTo endProc
            lbNew = True
         End If
      End If
         
      'Fill up the master info if this is a new record
      If lbNew Then
         .Master("dTransact") = p_oAppDrivr.ServerDate
         .Master("sDestinat") = pxeWHouseID
         .Master("sRequestx") = p_oBaseClas.Master("sTransNox") & "-WC"
      End If
                           
      'Transfer all items that needs to be requested
      For lnBOCtr = 0 To p_oBaseClas.ItemCount - 1
         'Be sure that we have a detail to fill
         If lnWCCtr + 1 > .ItemCount Then
            .AddDetail
         End If
         
         'If qty on hand > quantity to be used and not a labor then add this item as to be requested
         If p_oBaseClas.Detail(lnBOCtr, "nQuantity") > p_oBaseClas.Detail(lnBOCtr, "nQtyOnHnd") And p_oBaseClas.Detail(pnCtr, "cLaborxxx") = 0 Then
            Call .SearchDetail(lnWCCtr, 1, p_oBaseClas.Detail(lnBOCtr, "sBarrCode"))
            .Detail(lnWCCtr, "nQuantity") = p_oBaseClas.Detail(lnBOCtr, "nQuantity") - p_oBaseClas.Detail(lnBOCtr, "nQtyOnHnd")
            lnWCCtr = lnWCCtr + 1
         End If
      Next
      
      'Remove items that are not a part of the items to be requested
      Do Until lnWCCtr = .ItemCount
         .DeleteDetail lnWCCtr
      Loop
      
      'Save the transaction
      If lnWCCtr > 0 Then
         RequestItem = .SaveTransaction
         If RequestItem Then
            If lbNew Then
               MsgBox "Creation of request for items successful"
            Else
               MsgBox "Update of request for items successful"
            End If
         End If
      Else
                  
         RequestItem = False
      End If
   End With

endProc:
   Exit Function
errProc:
   ShowError lsOldProc
   p_oAppDrivr.RollbackTrans
End Function


