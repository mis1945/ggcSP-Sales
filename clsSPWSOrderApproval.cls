VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPWSOrderApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Spareparts Stock Request Approval object
'
' Copyright 2010 and Beyond
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
'  XerSys [ 09/19/2011 10:36 pm ]
'     Start creating this object.
'  XerSys [ 06/04/2014 09:54 am ]
'     Add area as filter
'  XerSys [ 10/03/2014 09:57 am ]
'     Update reserve order once request is approved as per request by Jojo Jose
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsSPReqApproval"
Private Const pxeBranchCodeLen As Integer = 4
Private Const pxeAllowed As Integer = 236

Private p_oAppDrivr As clsAppDriver
Private p_oSPOrder As clsSPWSOrder
Private p_oDetail As Recordset

Private p_sSQLDetail As String

Private p_bDispMsg As Boolean
Private p_sParent As String
Private p_sBranchCd As String

Private pbInitTran As Boolean
Private pbRecExist As Boolean

Private pbModified As Boolean
Private pnCtr As Integer
Private pbShwMsg As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   'on error Resume Next

   If Not pbInitTran Then
      Call showNotInitMsg
      Exit Property
   End If
   
   If Row > p_oDetail.RecordCount - 1 Then
'      MsgBox "Invalid Row Value Detected!!!", vbCritical, "Warning"
      Exit Property
   End If
   p_oDetail.Move Row, adBookmarkFirst
   
   Row = p_oDetail("nEntryNox") - 1
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, 1, 2, "nentrynox", "sbarrcode", "sdescript"
      Detail = p_oSPOrder.Detail(Row, Index)
   Case 7, "ncancelld"
      Detail = p_oSPOrder.Detail(Row, Index)
   Case 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, _
         "nquantity", "nrecorder", "nissuedxx", _
         "norderqty", "nqtyonhnd", "nresvordr", _
         "nbackordr", "nfloatqty", "navemonsl", _
         "cclassify"
      If IsNumeric(Index) Then Index = Index + 2
      
      Detail = p_oSPOrder.Detail(Row, Index)
   Case 14, "xresvordr"
      Detail = p_oDetail("nResvOrdr")
   Case 15, "xbackordr"
      Detail = p_oDetail("nBackOrdr")
   Case 16, "xfloatqty"
      Detail = p_oDetail("nFloatQty")
   Case 17, "xqtyonhnd"
      Detail = p_oDetail("nQtyOnHnd")
   Case 18, "smodelnme"
      Detail = p_oDetail("sModelNme")
   Case 19, "spartsidx"
      Detail = p_oSPOrder.Detail(Row, "sPartsIDx")
   Case 20
      Detail = p_oDetail("sRemarksx")
   Case Else
'      MsgBox "Invalid property index detected!", vbCritical, "Warning"
      Exit Property
   End Select
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   'on error Resume Next

   If pbInitTran = False Then
      Call showNotInitMsg
      Exit Property
   End If

   If Row > p_oDetail.RecordCount - 1 Then
      MsgBox "Invalid Row Value Detected!!!", vbCritical, "Warning"
      Exit Property
   End If
         
   p_oDetail.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, 1, 2, 3, 4, "nentrynox", "sbarrcode", "sdescript", "nquantity", "nrecorder"
   Case 7, "ncancelld"
      If Not IsNumeric(Value) Then Exit Property
      
      p_oSPOrder.Detail(Row, "nCancelld") = Value
   Case 5, 6, 8, 9, 10, 11, "nissuedxx", "norderqty", "nqtyonhnd", "nresvordr", "nbackordr", "nfloatqty"
   Case 12, 13, "navemonsl", "cclassify"
   Case 14, 15, 16, 17, 18, "xresvordr", "xbackordr", "xfloatqty", "xqtyonhnd", "smodelnme"
   Case 19, "spartsidx"
   Case Else
      MsgBox "Invalid property index detected!", vbCritical, "Warning"
      Exit Property
   End Select
End Property

Property Get Master(Index As Variant) As Variant
   If Not pbInitTran Then
      Call showNotInitMsg
      Exit Property
   End If
   
   Master = p_oSPOrder.Master(Index)
End Property

Property Let Master(Index As Variant, Value As Variant)
End Property

Property Get ItemCount() As Long
   If Not pbInitTran Then
      Call showNotInitMsg
      Exit Property
   End If

   ItemCount = p_oSPOrder.ItemCount
End Property

Property Let Parent(ByVal Value As String)
   p_sParent = Value
End Property

Function InitTransaction(Optional oAppDriver As Variant) As Boolean
   Dim lsProcName As String
   Dim lsCondition As String
   
   lsProcName = "InitTransaction"
   On Error GoTo errProc

   If pbInitTran = False Then
      If IsMissing(oAppDriver) Then
         If isAppDriverOK(p_oAppDrivr) = False Then
            Debug.Print "Uninitialized application driver!"
            GoTo endProc
         End If
      Else
         If TypeName(oAppDriver) <> "clsAppDriver" Then
            Debug.Print "Invalid application driver detected!"
            GoTo endProc
         End If
         Set p_oAppDrivr = oAppDriver
      End If
      If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
      pbInitTran = True
   End If
   
   p_sSQLDetail = "SELECT a.nEntryNox" & _
                     ", a.sPartsIDx" & _
                     ", IFNULL(b.nResvOrdr, 0) nResvOrdr" & _
                     ", IFNULL(b.nBackOrdr, 0) nBackOrdr" & _
                     ", IFNULL(b.nFloatQty, 0) nFloatQty" & _
                     ", IFNULL(b.nQtyOnHnd, 0) nQtyOnHnd" & _
                     ", d.sModelNme" & _
                     ", c.nSelPrice" & _
                     ", e.sRemarksx" & _
                  " FROM SP_WS_Order_Detail a" & _
                        " LEFT JOIN SP_Inventory b" & _
                           " ON a.sPartsIDx = b.sPartsIDx" & _
                              " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                     ", Spareparts c" & _
                        " LEFT JOIN SP_Model d" & _
                           " ON c.sModelIDx = d.sModelIDx" & _
                     ", SP_WS_Order_Master e" & _
                  " WHERE a.sPartsIDx = c.sPartsIDx" & _
                     " AND a.sTransNox = e.sTransNox" & _
                     " AND a.nQuantity - a.nCancelld > a.nIssuedxx" & _
                  " ORDER BY nEntryNox"
            
   Set p_oDetail = New Recordset
   Debug.Print p_sSQLDetail
   p_oDetail.Open AddCondition(p_sSQLDetail, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Debug.Print "Record Count:", p_oDetail.RecordCount
      
   Set p_oSPOrder = New clsSPWSOrder
   With p_oSPOrder
      Set .AppDriver = p_oAppDrivr
      .DisplayMessage = False
      .Parent = pxeMODULENAME
      .Distribute = True
      
      If Not .InitTransaction Then GoTo endProc
   End With
   
   InitTransaction = True
   pbInitTran = InitTransaction

endProc:
   Exit Function
errProc:
  ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   Dim lnCtr As Integer
   Dim lbUpdate As Boolean
   Dim lbApproval As Boolean
   Dim lsSupplier As String
   Dim loSPTrans As clsSPInventoryTrans

   lsProcName = "SaveTransaction"
   On Error GoTo errProc

   If Not pbInitTran Then
      Call showNotInitMsg
      GoTo endProc
   End If
   
   If Not pbRecExist Then
      Call showNoRecMsg
      GoTo endProc
   End If
   
   With p_oSPOrder
      ' check now if their is an item that has greater order
      '  than the recommended
      For lnCtr = 0 To .ItemCount - 1
         If Not lbApproval Then
            If .Detail(lnCtr, "nQuantity") > .Detail(lnCtr, "nRecOrder") Then
               If .Detail(lnCtr, "nQuantity") - .Detail(lnCtr, "nCancelld") > .Detail(lnCtr, "nRecOrder") Then
                  lbApproval = True
               End If
            End If
         End If
            
         If Not lbUpdate Then
            If .Detail(lnCtr, "nCancelld") > 0 Then lbUpdate = True
         End If
         
         If lbUpdate And lbApproval Then
            Exit For
         End If
      Next
   
      If lbApproval Then
         If (p_oAppDrivr.UserLevel And pxeAllowed) = 0 Then
            If p_sApproval = "" Then
               MsgBox "Invalid Approval Code Detected!" & vbCrLf & _
                     "Seek Manager's Approval for this Stock Request!" & vbCrLf & _
                     "Please verify your entry then try again!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         Else
            .Master("sApproved") = p_oAppDrivr.UserID
         End If
         
         .Master("sAprvCode") = p_sApproval
         lbUpdate = True
      End If
         
      If Not lbUpdate Then
         If MsgBox("Transaction was not Modified!" & vbCrLf & _
               "Saving it will approved all orders!" & vbCrLf & vbCrLf & _
               "Continue Saving Transaction?", vbCritical + vbYesNo, "Confirm") <> vbYes Then
            GoTo endProc
         End If
      End If
      
      ' XerSys - 2014-10-03
      '  update the approved order to the approvers' inventory
      Set loSPTrans = New clsSPInventoryTrans
      With loSPTrans
         Set .AppDriver = p_oAppDrivr
         .Branch = p_sBranchCd
         If .InitTransaction() = False Then GoTo endProc
      End With
      
      For lnCtr = 0 To .ItemCount - 1
         loSPTrans.Detail(lnCtr, "sPartsIDx") = .Detail(lnCtr, "sPartsIDx")
         loSPTrans.Detail(lnCtr, "nQuantity") = .Detail(lnCtr, "nQuantity") - .Detail(lnCtr, "nCancelld")
      Next
      
      If loSPTrans.BranchOrderConfirm(.Master("sTransNox"), p_oAppDrivr.ServerDate, xeModeAddNew) = False Then
         GoTo endProc
      End If
      
      .Master("cTranStat") = xeStateClosed
      .Destination = Left(.Master("sTransNox"), pxeBranchCodeLen)
      
      If Not p_oSPOrder.SaveTransaction() Then GoTo endWithRoll
   End With
   
   SaveTransaction = True

endProc:
   Exit Function
endWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
  ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "OpenTransaction"
   On Error GoTo errProc

   If Not pbInitTran Then
      Call showNotInitMsg
      GoTo endProc
   End If
   
   lsSQL = AddCondition(p_sSQLDetail, "a.sTransNox = " & strParm(sTransNo))
   
   Set p_oDetail = New Recordset
   p_oDetail.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set p_oDetail.ActiveConnection = Nothing
   
   If Not p_oDetail.EOF Then
      pbRecExist = p_oSPOrder.OpenTransaction(sTransNo)
   End If

   OpenTransaction = pbRecExist
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lasSelect() As String
   Dim loRS As Recordset

   lsProcName = "SearchTransaction"
   On Error GoTo errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   lsSQL = "SELECT DISTINCT a.sTransNox" & _
               ", b.sBranchNm" & _
               ", a.dTransact" & _
            " FROM SP_Stock_Request_Master a" & _
               ", Branch b" & _
               ", SP_Stock_Request_Detail c" & _
            " WHERE a.sTransNox LIKE CONCAT(b.sBranchCd, '%')" & _
               " AND a.sTransNox = c.sTransNox" & _
               " AND a.cTranStat = " & strParm(xeStateOpen) & _
               IIf(p_sReqBranchCd <> "", " AND b.sBranchCd = " & strParm(p_sReqBranchCd), "")
               
   If Not IsMissing(sSearch) Then
      If p_oSPOrder.Master("sTransNox") = sSearch Then GoTo endProc
      lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(sSearch & "%"))
   End If
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   Debug.Print lsSQL
   If loRS.EOF Then GoTo endProc
    
   If loRS.RecordCount > 1 Then
      lsSQL = KwikBrowse(p_oAppDrivr, loRS, "sTransNox»sBranchNm»dTransact", _
                        "Transaction #»Branch»Date", "@@@@»@»MMM DD, YYYY")
      If lsSQL = "" Then GoTo endProc
      
      lasSelect = Split(lsSQL, "»")
      
      Call loRS.Find("sTransNox = " & strParm(lasSelect(0)), 0, adSearchForward)
   End If
   
   SearchTransaction = OpenTransaction(loRS("sTransNox"))

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sSearch & " )"
End Function

Function SearchBranch(ByVal lsValue As String, ByVal lbSearch As Boolean) As String
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "getBranch"
   On Error GoTo errProc
   
   SearchBranch = p_sReqBranchNm
   
   If lsValue <> "" Then
      If lsValue = p_sReqBranchNm Then GoTo endProc
      
      If lbSearch Then
         lsSQL = "a.sBranchNm LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsSQL = "a.sBranchNm = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endProc
   End If
   
   lsSQL = "SELECT a.sBranchCd" & _
               ", a.sBranchNm" & _
               ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName) xAddressx" & _
            " FROM Branch a" & _
               ", TownCity b" & _
               ", Province c" & _
            " WHERE a.sTownIDxx = b.sTownIDxx" & _
               " AND b.sProvIDxx = c.sProvIDxx" & _
               " AND a.sBranchCd <> " & strParm(p_sBranchCd) & _
               " AND sBranchCd LIKE " & strParm("M%") & _
               IIf(lsSQL = "", "", " AND " & lsSQL)
   Debug.Print lsSQL
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   If loRS.EOF Then
      If Not lbSearch Then
         p_sReqBranchCd = ""
         p_sReqBranchNm = ""
         p_sAddressx = ""
      End If
      GoTo endProc
   End If
   
   If loRS.RecordCount > 1 Then
      lsSQL = KwikBrowse(p_oAppDrivr, loRS _
                           , "sBranchCd»sBranchNm" _
                           , "ID»Branch")
      If lsSQL = "" Then
         GoTo endProc
      Else
         lasMaster = Split(lsSQL, "»")
         
         loRS.MoveFirst
         loRS.Find "sBranchCd = " & strParm(lasMaster(0)), 0, adSearchForward
      End If
   End If
      
   p_sReqBranchCd = loRS("sBranchCd")
   p_sReqBranchNm = loRS("sBranchNm")
   p_sAddressx = loRS("xAddressx")
   
   SearchBranch = p_sReqBranchNm
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lsValue _
                            & ", " & lbSearch & " )"
End Function

Function GetBranchOrder() As Recordset
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lsProcName As String
   
   lsProcName = "GetBranchOrder"
   On Error GoTo errProc
   
'   If p_sReqBranchCd = "" Then
'      Set GetBranchOrder = Nothing
'      GoTo endProc
'   End If
   
   lsSQL = "SELECT a.sTransNox" & _
               ", a.dTransact" & _
               ", COUNT(b.sPartsIDx) xUnserved" & _
               ", c.sBranchNm" & _
            " FROM SP_Stock_Request_Master a" & _
               ", SP_Stock_Request_Detail b" & _
               ", Branch c" & _
               ", Branch_Others d" & _
            " WHERE a.sTransNox = b.sTransNox" & _
               " AND a.sTransNox LIKE CONCAT(c.sBranchCd, '%')" & _
               " AND c.sBranchCd = d.sBranchCd" & _
               " AND a.cTranStat = " & strParm(xeStateOpen) & _
               IIf(psOrderType = "1", " AND b.nResvOrdr > 0", "") & _
               IIf(psOrderType = "2", " AND b.nResvOrdr = 0", "") & _
            " GROUP BY a.sTransNox" & _
            " ORDER BY a.dTransact ASC, a.sTransNox ASC"
            
   If p_sAreaCode <> "" Then
      lsSQL = AddCondition(lsSQL, "d.sAreaCode = " & strParm(p_sAreaCode))
   End If
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Debug.Print lsSQL
   Set GetBranchOrder = loRS
   
endProc:
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

Private Sub showNotInitMsg()
   MsgBox "Transaction is not yet Initialized!" & vbCrLf & _
            "Please verify your entry then Try Again!", vbCritical, "Warning"
End Sub

Private Sub showNoRecMsg()
   MsgBox "No transaction is currently loaded!" & vbCrLf & _
            "Please verify your entry then Try Again!", vbCritical, "Warning"
End Sub

Private Sub Class_Terminate()
   Set p_oSPOrder = Nothing
   Set p_oDetail = Nothing
End Sub


