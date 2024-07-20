VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPAutoTransfer"
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
'     Spareparts Automatic Stock Transfer Object
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
'  XerSys [ 07/01/2005 05:11 pm ]
'     Start creating this object.
'  Kalyptus [ 12/11/2006 10:08 am]
'     Implements the BugTracker
'  Kalyptus [ 04/02/2007 10:23 am]
'     Updated the p_oBase_Clas to conform it with the clsMasterDetail specification
'  XerSys [ 08/14/2007 05:33 pm ]
'     Update this object to adopt the new implementation of SP Inventory Transaction
'  XerSys [ 09/24/2008 01:53 pm ]
'     Posted Autotransfer will not appear to be for posting even if new parts were added,
'        so adjust the select criteria to include those transaction.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeWHouseID As String = "W2"
Private Const pxeMODULENAME As String = "clsSPAutoTransfer"

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
   Case 1, "sbarrcode"
      getDetail Row, 1, Value, False
   Case 2, "sdescript"
      getDetail Row, 2, Value, False
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
   Case 3, "xaddressx"
      p_sAddressx = Value
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
      .MasterTable = "SP_AutoTransfer_Master"
      .DetailTable = "SP_AutoTransfer_Detail"
      .DisplayMessage = pbShowMsg
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDestinat" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                        ", a.sReferNox" & _
                        ", a.sSourceCd" & _
                        ", a.nEntryNox" & _
                        ", a.cTranStat" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", e.sEstimate" & _
                        ", LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") xSourcexx" & _
                     " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN JobOrder_Master e" & _
                              " ON a.sTransNox = e.sAutoTran" & _
                        ", Branch b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE a.sDestinat = b.sBranchCd" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx" & _
                        " AND LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = " & strParm(p_sBranchCd)
                     
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", d.sTypeCode" & _
                        ", e.sModelNme" & _
                        ", a.nQuantity" & _
                        ", a.nIssuedxx" & _
                        ", a.dModified" & _
                        ", a.sTransNox" & _
                        ", a.sPartsIDx" & _
                        ", c.nQtyOnHnd" & _
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
            
      pbInitTran = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sReferNox" & _
                           ", b.sBranchNm" & _
                           ", a.dTransact" & _
                           ", c.sEstimate" & _
                        " FROM " & .MasterTable & " a" & _
                              " LEFT JOIN JobOrder_Master c" & _
                                 " ON a.sTransNox = c.sAutoTran" & _
                           ", Branch b" & _
                        " WHERE LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & " ) = b.sBranchCd" & _
                           " AND LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = " & strParm(p_sBranchCd) & _
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
            ' assume that this will be used for posting quantity issue
            .DetailQuery = AddCondition(.DetailQuery, "a.nQuantity > a.nIssuedxx")
         Else
            lsCondition = "a.cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      .BrowseQuery(0) = AddCondition(.BrowseQuery(0), lsCondition)
      
      .BrowseColumn(0) = "sReferNox»sBranchNm»sTransNox»dTransact"
      .BrowseTitle(0) = "Reference No»Branch»Transaction No»Date"
      .BrowseFormat(0) = "@»@»@@-@@@@@@@@»MMMM DD, YYYY"
      
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
                           " AND a.cRecdStat = " & xeRecStateActive & _
                           " AND a.sBranchCd <> " & strParm(p_sBranchCd) & _
                        " ORDER BY a.sBranchNm"
      .BrowseColumn(2) = "sBranchCd»sBranchNm»xAddressx"
      .BrowseTitle(2) = "Code»Branch Name»Address"
      
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sBarrCode" & _
                                 ", a.sDescript" & _
                                 ", c.sTypeCode" & _
                                 ", d.sModelNme" & _
                                 ", e.sSizeName" & _
                                 ", f.sFormName" & _
                                 ", b.nQtyOnHnd" & _
                                 ", a.sPartsIDx" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nReorderx" & _
                                 ", b.nLedgerNo" & _
                              " FROM Spareparts a" & _
                                    " LEFT JOIN SP_Model d" & _
                                       " ON a.sModelIDx = d.sModelIDx" & _
                                    " LEFT JOIN Size e" & _
                                       " ON a.sSizeIDxx = e.sSizeIDxx" & _
                                    " LEFT JOIN Form f" & _
                                       " ON a.sFormIDxx = f.sFormIDxx" & _
                                 ", SP_Inventory b" & _
                                 ", SP_Type c" & _
                              " WHERE a.sPartsIDx = b.sPartsIDx" & _
                                 " AND a.cPartType = c.cPartType" & _
                                 " AND b.sBranchCd = " & strParm(p_sBranchCd)
      .BrowseDetailColumn(0) = "sBarrCode»sDescript»sTypeCode»sModelNme»sSizeName»sFormName»nQtyOnHnd»nSelPrice"
      .BrowseDetailTitle(0) = "Bar Code»Description»Type»Model»Size»Form»QtyOnHand»Price"
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

Function SearchByEstimate(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lsRecord As String
   Dim lsValue As String, lsField As String
   Dim lsDescript As String, lsCondition As String
   Dim lsQuery As String
   
   lsProcName = "SearchByEstimate"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   With p_oBaseClas
      lsQuery = .BrowseQuery(0)
      
      .BrowseQuery(0) = "SELECT DISTINCT" & _
                           "  a.sTransNox" & _
                           ", c.sEstimate" & _
                           ", b.sBranchNm" & _
                           ", a.dTransact" & _
                           ", c.sEstimate" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Branch b" & _
                           ", JobOrder_Master c" & _
                           ", " & .DetailTable & " d" & _
                        " WHERE LEFT(a.sReferNox, " & Len(p_oAppDrivr.BranchCode) & ") = b.sBranchCd" & _
                           " AND LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = " & strParm(p_sBranchCd) & _
                           " AND a.sTransNox = c.sAutoTran" & _
                           " AND a.sTransNox = d.sTransNox" & _
                           " AND d.nQuantity > d.nIssuedxx" & _
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
      .BrowseColumn(0) = "sEstimate»sBranchNm»sTransNox»dTransact"
      .BrowseTitle(0) = "Estimate No»Branch»Transaction No»Date"
      .BrowseFormat(0) = "@»@»@@-@@@@@@@@»MMMM DD, YYYY"
   End With
   
   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "c.sEstimate"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "b.sBranchNm"
            lsDescript = "b.sBranchNm"
         End If
      End If
   End If
   
   Debug.Print p_oBaseClas.BrowseQuery(0)
   SearchByEstimate = p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript)
   
endProc:
   p_oBaseClas.BrowseQuery(0) = lsQuery
   
   Exit Function
errProc:
    ShowError lsProcName & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function PostTransaction(ByVal sTransNo As String) As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "PostTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   PostTransaction = False
   
   If Not pbInitTran Then GoTo endProc
   
   With p_oBaseClas
      If sTransNo = "" Then GoTo endProc
      
      If sTransNo <> .Master("sTransNox") Then
         If .OpenTransaction(sTransNo) = False Then GoTo endProc
      End If
   
      Set loSPTrans = New clsSPInventoryTrans
      With loSPTrans
         Set .AppDriver = p_oAppDrivr
         .Branch = p_sBranchCd
         If .InitTransaction() = False Then GoTo endProc
      End With
      
      p_oAppDrivr.BeginTrans
      For pnCtr = 0 To .ItemCount - 1
         loSPTrans.Detail(pnCtr, "sPartsIDx") = .Detail(pnCtr, "sPartsIDx")
         loSPTrans.Detail(pnCtr, "nQuantity") = .Detail(pnCtr, "nQuantity")
         loSPTrans.Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
         loSPTrans.Detail(pnCtr, "nResvOrdr") = .Detail(pnCtr, "nResvOrdr")
         loSPTrans.Detail(pnCtr, "nBackOrdr") = .Detail(pnCtr, "nBackOrdr")
         loSPTrans.Detail(pnCtr, "nReorderx") = .Detail(pnCtr, "nReorderx")
         loSPTrans.Detail(pnCtr, "nLedgerNo") = .Detail(pnCtr, "nLedgerNo")
         
         lsSQL = "UPDATE " & .DetailTable & " SET" & _
                     "  nIssuedxx = " & .Detail(pnCtr, "nQuantity") & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                  " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                     " AND nEntryNox = " & .Detail(pnCtr, "nEntryNox")
         If p_oAppDrivr.Execute(lsSQL, .DetailTable, p_sBranchCd) = 0 Then
            MsgBox "Unable to Update Issued Quantity!", vbCritical, "Warning"
            p_oAppDrivr.RollbackTrans
            GoTo endProc
         End If
      Next
      
      If loSPTrans.AutoDelivery(.Master("sTransNox"), _
                        .Master("dTransact"), _
                        xeModeAddNew) = False Then
         p_oAppDrivr.RollbackTrans
         GoTo endProc
      End If
                             
      ' check if reorder level of some parts has reached
      If loSPTrans.ReOrderParts Then
         MsgBox "Some Parts Included in the Transaction has Reached it Reorder Level Limit!!!" & _
                  vbCrLf & vbCrLf & "Please Inform the Spareparts Manager on this Alert!!!", _
                  vbInformation, "Notice"
      End If
      
      lsSQL = "UPDATE " & .MasterTable & " SET" & _
                  "  dTransact = " & dateParm(.Master("dTransact")) & _
                  ", cTranStat = " & xeStatePosted & _
                  ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
               " WHERE sTransNox = " & strParm(sTransNo)
               
      If p_oAppDrivr.Execute(lsSQL, .MasterTable, p_sBranchCd) = 0 Then
         MsgBox "Unable to Post Transaction!", vbCritical, "Warning"
         p_oAppDrivr.RollbackTrans
         GoTo endProc
      End If
   End With
   
   p_oAppDrivr.CommitTrans
   
   PostTransaction = True
   
endProc:
   Set loSPTrans = Nothing
   Exit Function
errProc:
   p_oAppDrivr.RollbackTrans
    ShowError lsProcName & "( " & sTransNo & " )"
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
   Dim lasSelect() As String
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
               'she 2021-08-13
               'replace "-" to "" for barcode reader use
'               lsDetail = "a.sBarrCode LIKE " & strParm(Trim(lsValue) & "%")
               lsDetail = "(a.sBarrCode LIKE " & strParm(Trim(lsValue) + "%") & _
                           " OR REPLACE(a.sBarrcode, '-', '') LIKE " & strParm(Trim(lsValue) + "%") & ")"
            Else
'               lsDetail = "a.sBarrcode = " & strParm(Trim(lsValue))
               lsDetail = "(a.sBarrCode = " & strParm(Trim(lsValue)) & _
                           " OR REPLACE(a.sBarrcode, '-', '') = " & strParm(Trim(lsValue)) & ")"
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
      .Detail(lnRow, "nQtyOnHnd") = CLng(lasDetail(6))
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "sPartsIDx") = lasDetail(7)
      .Detail(lnRow, "nResvOrdr") = lasDetail(8)
      .Detail(lnRow, "nBackOrdr") = lasDetail(9)
      .Detail(lnRow, "nReorderx") = lasDetail(10)
      .Detail(lnRow, "nLedgerNo") = lasDetail(11)
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
      .Master("sReferNox") = ""
      .Master("sSourceCd") = ""
      .Master("nEntryNox") = 0
      .Master("cTranStat") = xeStateOpen

      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "sTypeCode") = ""
      .Detail(0, "sModelNme") = ""
      .Detail(0, "nQuantity") = 0
      .Detail(0, "nIssuedxx") = 0
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sPartsIDx") = ""
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nResvOrdr") = 0
      .Detail(0, "nBackOrdr") = 0
      .Detail(0, "nReorderx") = 0
      .Detail(0, "nLedgerNo") = 0
   End With

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      p_sAddressx = .Master("xAddressx")
   End With
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   Dim loSPTrans As clsSPInventoryTrans
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_SaveOthers"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = True

   With p_oBaseClas
      Set loSPTrans = New clsSPInventoryTrans
      With loSPTrans
         .Branch = p_oBaseClas.Master("sDestinat")
         Set .AppDriver = p_oAppDrivr
         If .InitTransaction() = False Then GoTo endProc
      End With
      
      Set loRS = New Recordset
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sPartsIDx") = "" Then Exit For
'
'         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
'         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
         
         lsSQL = "SELECT" & _
                     "  sPartsIDx" & _
                     ", nQtyOnHnd" & _
                     ", nResvOrdr" & _
                     ", nBackOrdr" & _
                     ", nReorderx" & _
                     ", nLedgerNo" & _
                  " FROM SP_Inventory" & _
                  " WHERE sPartsIDx = " & strParm(.Detail(pnCtr, "sPartsIDx")) & _
                     " AND sBranchCd = " & strParm(.Master("sDestinat"))
         
         loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
         
         loSPTrans.Detail(pnCtr, "sPartsIDx") = .Detail(pnCtr, "sPartsIDx")
         loSPTrans.Detail(pnCtr, "nQuantity") = .Detail(pnCtr, "nQuantity")
   
         If loRS.EOF Then
            loSPTrans.Detail(pnCtr, "cNewParts") = xeYes
            loSPTrans.Detail(pnCtr, "nQtyOnHnd") = 0
            loSPTrans.Detail(pnCtr, "nResvOrdr") = 0
            loSPTrans.Detail(pnCtr, "nBackOrdr") = 0
         Else
            loSPTrans.Detail(pnCtr, "nQtyOnHnd") = loRS("nQtyOnHnd")
            loSPTrans.Detail(pnCtr, "nResvOrdr") = loRS("nResvOrdr")
            loSPTrans.Detail(pnCtr, "nBackOrdr") = loRS("nBackOrdr")
'            loSPTrans.Detail(pnCtr, "nReorderx") = loRS("nReorderx")
            loSPTrans.Detail(pnCtr, "nLedgerNo") = loRS("nLedgerNo")
         End If
         loRS.Close
      Next
      
      ' if a spareparts is not successfully updated, cancel saving.
      If loSPTrans.AcceptAutoDelivery(.Master("sTransNox"), _
            .Master("dTransact"), _
            .EditMode) = False Then
         GoTo endProc
      End If
   End With
   
   Cancel = False

endProc:
   Set loSPTrans = Nothing
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
   Cancel = True

   With p_oBaseClas
      Set loSPTrans = New clsSPInventoryTrans
      With loSPTrans
         .Branch = p_oBaseClas.Master("sDestinat")
         Set .AppDriver = p_oAppDrivr
         If .InitTransaction() = False Then GoTo endProc
      End With
      
      Set loRS = New Recordset
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sPartsIDx") = "" Then Exit For
         
         lsSQL = "SELECT" & _
                     "  sPartsIDx" & _
                     ", nQtyOnHnd" & _
                     ", nResvOrdr" & _
                     ", nBackOrdr" & _
                     ", nReorderx" & _
                     ", nLedgerNo" & _
                  " FROM SP_Inventory" & _
                  " WHERE sPartsIDx = " & strParm(.Detail(pnCtr, "sPartsIDx")) & _
                     " AND sBranchCd = " & strParm(.Master("sDestinat"))
         
         loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
         
         loSPTrans.Detail(pnCtr, "sPartsIDx") = .Detail(pnCtr, "sPartsIDx")
         loSPTrans.Detail(pnCtr, "nQuantity") = .Detail(pnCtr, "nQuantity")
   
         If loRS.EOF = False Then
            loSPTrans.Detail(pnCtr, "nQtyOnHnd") = loRS("nQtyOnHnd")
            loSPTrans.Detail(pnCtr, "nResvOrdr") = loRS("nResvOrdr")
            loSPTrans.Detail(pnCtr, "nBackOrdr") = loRS("nBackOrdr")
            loSPTrans.Detail(pnCtr, "nReorderx") = loRS("nReorderx")
            loSPTrans.Detail(pnCtr, "nLedgerNo") = loRS("nLedgerNo")
         End If
         loRS.Close
      Next
      
      ' if a spareparts is not successfully updated, cancel saving.
      If loSPTrans.AcceptAutoDelivery(.Master("sTransNox"), _
            .Master("dTransact"), _
            xeModeDelete) = False Then
         GoTo endProc
      End If
   End With
   
   Cancel = False

endProc:
   Set loSPTrans = Nothing
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
