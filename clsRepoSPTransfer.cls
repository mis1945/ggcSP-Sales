VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRepoSPTransfer"
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
'  Kalyptus [ 10/18/2008 03:58 pm ]
'     Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME = "clsRepoSPTransfer"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oDetail As Recordset

Private p_sBranchCd As String
Private p_sAddressx As String
Private p_nTranStat As Integer

Private pbInitTran As Boolean
Private pbModified As Boolean
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

Property Let TransStatus(ByVal Value As Integer)
   p_nTranStat = Value
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsCondition As String
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail
   
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "RepoSP_Transfer_Master"
      .DetailTable = "RepoSP_Transfer_Detail"
      
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
                        ", d.sTypeCode" & _
                        ", e.sModelNme" & _
                        ", c.nQtyOnHnd" & _
                        ", a.nQuantity" & _
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
                              " LEFT JOIN SP_Inventory c" & _
                                 " ON b.sPartsIDx = c.sPartsIDx" & _
                              " LEFT JOIN SP_Type d" & _
                                 " ON b.cPartType = d.cPartType" & _
                     " WHERE a.sPartsIDx = b.sPartsIDx" & _
                     " ORDER BY a.nEntryNox"
                     
      If p_oAppDrivr.BranchCode = "W2" Then
         .DetailQuery = AddCondition(.DetailQuery, "c.sBranchCd = " & strParm(p_sBranchCd))
      End If
      
      Debug.Print .DetailQuery
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
                                 ", a.sPartsIDx" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nReorderx" & _
                                 ", b.nLedgerNo" & _
                              " FROM Spareparts a" & _
                                    " LEFT JOIN SP_Model d" & _
                                       " ON a.sModelIDx = d.sModelIDx" & _
                                    " LEFT JOIN SP_Inventory b" & _
                                       " ON a.sPartsIDx = b.sPartsIDx" & _
                                 ", SP_Type c" & _
                              " WHERE a.cPartType = c.cPartType" & _
                                " AND b.sBranchCd = " & strParm(p_sBranchCd)
      .BrowseDetailColumn(0) = "sBarrCode»sDescript»sTypeCode»sModelNme»nQtyOnHnd"
      .BrowseDetailTitle(0) = "Bar Code»Description»Type»Model»QtyOnHand"
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

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
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
   'On Error GoTo errProc
   
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
                              " LEFT JOIN SP_Inventory c" & _
                                 " ON b.sPartsIDx = c.sPartsIDx" & _
                        ", SP_Type d" & _
                     " WHERE a.sPartsIDx = b.sPartsIDx" & _
                        " AND b.cPartType = d.cPartType" & _
                     " ORDER BY a.nEntryNox"

      If p_oAppDrivr.BranchCode = "W2" Then
         .DetailQuery = AddCondition(.DetailQuery, "c.sBranchCd = LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ")")
      End If
      
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
   'On Error GoTo errProc
   
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
   'On Error GoTo errProc
   
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
      .Detail(pnCtr, "sTypeCode") = ""
      .Detail(pnCtr, "sModelNme") = ""
      .Detail(pnCtr, "nQtyOnHnd") = 0
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
   'On Error GoTo errProc
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
      If .Master("sReceived") <> Empty Then Exit Function
         
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
   AcceptDelivery = False
   
   If Not pbInitTran Then GoTo endProc
   
   If p_oAppDrivr.BranchCode = "W2" Then
      With p_oBaseClas
         If .Master("sTransNox") = "" Then GoTo endProc
         
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
                     Left(.Master("sTransNox"), Len(p_oAppDrivr.BranchCode)))
         
         If lnRow = 0 Then
            MsgBox "Unable to Update Delivery Transaction!!!", vbCritical, "Warning"
            GoTo endProcWithRoll
         End If
         
         p_oAppDrivr.CommitTrans
         .Master("cTranStat") = xeStatePosted
      End With
   End If
   
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
      .Detail(lnRow, "sPartsIDx") = lasDetail(5)
      .Detail(lnRow, "nResvOrdr") = lasDetail(6)
      .Detail(lnRow, "nBackOrdr") = lasDetail(7)
      .Detail(lnRow, "nReorderx") = lasDetail(8)
      .Detail(lnRow, "nLedgerNo") = lasDetail(9)
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
      
      Call clearTempRec
      
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
   'On Error GoTo errProc
   
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
   'On Error GoTo errProc

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
   'On Error GoTo errProc

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
   'On Error GoTo errProc
      
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
      .Master("cTranStat") = xeStateOpen
      .Master("nEntryNox") = 1

      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "sTypeCode") = ""
      .Detail(0, "sModelNme") = ""
      .Detail(0, "nQtyOnHnd") = 0
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
'         .Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "xQtyOnHnd")
         
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
   'On Error GoTo errProc
   Cancel = True

   
   If p_oAppDrivr.BranchCode = "W2" Then
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
            'loSPTrans.Detail(pnCtr, "nReorderx") = .Detail(pnCtr, "nReorderx")
            loSPTrans.Detail(pnCtr, "nLedgerNo") = IIf(.Detail(pnCtr, "nLedgerNo") = "", 0, .Detail(pnCtr, "nLedgerNo"))
               
            Call addTempRec(.Detail(pnCtr, "sPartsIDx"), 0, .Detail(pnCtr, "nQuantity"))
         Next
         
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
      End With
   End If
   
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
   Debug.Print pxeMODULENAME & "." & lsProcName
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




