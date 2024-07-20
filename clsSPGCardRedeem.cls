VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPGCardRedeem"
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
'     Guanzon Apps Order Redemption Object
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
'  Mac [ 11/02/2018 04:55 pm ]
'     Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME = "clsSPGCardRedeem"
Private Const pxeMasTable = "G_Card_Redemption"
Private Const pxeSourceCd = "GApp"

Private p_oAppDrivr As clsAppDriver

Private p_sBranchCd As String
Private p_nTranStat As Integer
Private pbInitTran As Boolean

Private p_oRSMaster As Recordset
Private p_oRSDetail As Recordset

Private p_sSQLMaster As String
Private p_sSQLDetail As String
Private p_sSQLBrowse As String

Private p_sClientNm As String
Private p_sAddressx As String
Private p_dBirthDte As Date
Private p_nEditMode As Long
Private pnCtr As Integer

Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get EditMode() As Long
   EditMode = p_nEditMode
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oRSDetail.RecordCount
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
   
   p_sSQLMaster = "SELECT" & _
                     "  a.sTransNox" & _
                     ", a.sGCardNox" & _
                     ", a.sCompnyID" & _
                     ", a.dTransact" & _
                     ", a.sSourceNo" & _
                     ", a.sSourceCd" & _
                     ", a.sPromoIDx" & _
                     ", a.nTranAmtx" & _
                     ", a.nPointsxx" & _
                     ", a.sRemarksx" & _
                     ", a.cIsBilled" & _
                     ", b.sCardNmbr" & _
                     ", b.dCardExpr" & _
                     ", b.nAvlPoint" & _
                     ", b.dLastRedm" & _
                     ", c.sPromCode" & _
                     ", d.sCompnyNm" & _
                     ", CONCAT(IFNULL(d.sHouseNox, ''), ' ', d.sAddressx, ', ', e.sTownName, ' ', f.sProvIDxx) xAddressx" & _
                     ", d.dBirthDte"

                     

   p_sSQLMaster = p_sSQLMaster & _
                  " FROM G_Card_Redemption a" & _
                     ", G_Card_Promo_Master c" & _
                     ", G_Card_Master b" & _
                        " LEFT JOIN Client_Master d" & _
                           " ON b.sClientID = d.sClientID" & _
                        " LEFT JOIN TownCity e" & _
                           " ON d.sTownIDxx = e.sTownIDxx" & _
                        " LEFT JOIN Province f" & _
                           " ON e.sProvIDxx = f.sProvIDxx" & _
                  " WHERE a.sGCardNox = b.sGCardNox" & _
                     " AND a.sPromoIDx = c.sTransNox"
                     
   p_sSQLBrowse = "SELECT" & _
                     "  a.sTransNox" & _
                     ", a.dTransact" & _
                     ", a.sPromoIDx" & _
                     ", e.sPromCode" & _
                     ", d.sCompnyNm" & _
                     ", c.sCardNmbr" & _
                  " FROM G_Card_Redemption a" & _
                     ", G_Card_Order_Redeem b" & _
                        " LEFT JOIN G_Card_Promo_Master e" & _
                           " ON b.sPromoIDx = e.sTransNox" & _
                     ", G_Card_Master c" & _
                     ", Client_Master d" & _
                  " WHERE a.sSourceNo = b.sTransNox" & _
                     " AND a.sSourceCd = " & strParm(pxeSourceCd) & _
                     " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                     " AND b.sGCardNox = c.sGCardNox" & _
                     " AND c.sClientID = d.sClientID"
                     
'   If p_nTranStat <> xeStateUnknown Then
'      If p_nTranStat > xeStateUnknown Then
'         lsCondition = "("
'         For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
'            lsCondition = lsCondition & " cTranStat = " & _
'                              strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
'         Next
'         lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
'      Else
'         lsCondition = "cTranStat = " & strParm(p_nTranStat)
'      End If
'   End If
'
'   p_sSQLBrowse = AddCondition(p_sSQLBrowse, lsCondition)

   p_nEditMode = xeModeUnknown

   pbInitTran = True
   InitTransaction = True
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "SaveTransaction"
   'On Error GoTo errProc
   
   If Not pbInitTran Then GoTo endProc
   
   If Not isEntryOK Then GoTo endProc
   
   p_oAppDrivr.BeginTrans
   
   'save master
   lsSQL = ADO2SQL(p_oRSMaster, _
                     pxeMasTable, , _
                     p_oAppDrivr.UserID, _
                     p_oAppDrivr.ServerDate, _
                     "sCardNmbr»dCardExpr»nAvlPoint»dLastRedm»sPromCode»sCompnyNm»xAddressx»dBirthDte")
                   
   Debug.Print lsSQL
   If p_oAppDrivr.Execute(lsSQL, pxeMasTable) <= 0 Then
      MsgBox "Unable to save redemption." & vbCrLf & vbCrLf & _
               "Please inform MIS Department.", vbCritical, "Warning"
      GoTo endWithRoll
   End If
   
   'save ledger
   If Not saveSPTransaction Then GoTo endWithRoll
   
   'update the order status to claimed
   lsSQL = "UPDATE G_Card_Order_Redeem" & _
            " SET cPlcOrder = '2'" & _
            " WHERE sTransNox = " & strParm(p_oRSMaster("sSourceNo"))
   
   Debug.Print lsSQL
   If p_oAppDrivr.Execute(lsSQL, "G_Card_Order_Redeem") <= 0 Then
      MsgBox "Unable to order status." & vbCrLf & vbCrLf & _
               "Please inform MIS Department.", vbCritical, "Warning"
      GoTo endWithRoll
   End If
   
   'deduct the points
   lsSQL = "UPDATE G_Card_Master SET" & _
               "  nTotPoint = nTotPoint - " & p_oRSMaster("nPointsxx") & _
               ", nAvlPoint = nAvlPoint - " & p_oRSMaster("nPointsxx") & _
               ", dLastRedm = " & dateParm(p_oRSMaster("dTransact")) & _
            " WHERE sGCardNox = " & strParm(p_oRSMaster("sGCardNox"))
            
   Debug.Print lsSQL
   If p_oAppDrivr.Execute(lsSQL, "G_Card_Master") <= 0 Then
      MsgBox "Unable to G-Card status." & vbCrLf & vbCrLf & _
               "Please inform MIS Department.", vbCritical, "Warning"
      GoTo endWithRoll
   End If
   
   
   p_oAppDrivr.CommitTrans
   
   p_nEditMode = xeModeReady
   SaveTransaction = True
endProc:
   Exit Function
endWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "CancelTransaction"
   'On Error GoTo errProc
   
   If Not pbInitTran Then GoTo endProc
   
   If Not isEntryOK Then GoTo endProc
   
   p_oAppDrivr.BeginTrans
     
   'save ledger
   If Not saveSPTransaction Then GoTo endWithRoll
   
   'update the order status to claimed
   lsSQL = "UPDATE G_Card_Order_Redeem" & _
            " SET cPlcOrder = '1'" & _
            " WHERE sTransNox = " & strParm(p_oRSMaster("sSourceNo"))
   
   Debug.Print lsSQL
   If p_oAppDrivr.Execute(lsSQL, "G_Card_Order_Redeem") <= 0 Then
      MsgBox "Unable to order status." & vbCrLf & vbCrLf & _
               "Please inform MIS Department.", vbCritical, "Warning"
      GoTo endWithRoll
   End If
   
   'deduct the points
   lsSQL = "UPDATE G_Card_Master SET" & _
               "  nTotPoint = nTotPoint + " & p_oRSMaster("nPointsxx") & _
               ", nAvlPoint = nAvlPoint + " & p_oRSMaster("nPointsxx") & _
               ", dLastRedm = " & dateParm(p_oRSMaster("dTransact")) & _
            " WHERE sGCardNox = " & strParm(p_oRSMaster("sGCardNox"))
            
   Debug.Print lsSQL
   If p_oAppDrivr.Execute(lsSQL, "G_Card_Master") <= 0 Then
      MsgBox "Unable to G-Card status." & vbCrLf & vbCrLf & _
               "Please inform MIS Department.", vbCritical, "Warning"
      GoTo endWithRoll
   End If
   
   
   p_oAppDrivr.CommitTrans
   
   p_nEditMode = xeModeReady
   CancelTransaction = True
endProc:
   Exit Function
endWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function isEntryOK()
   isEntryOK = False
   With p_oRSMaster
      If .Fields("sGCardNox") = "" Then
         MsgBox "Invalid G-Card number detected.", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If .Fields("sSourceNo") = "" Then
         MsgBox "Invalid source number detected", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If .Fields("sPromoIDx") = "" Then
         MsgBox "Invalid promo id detected.", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   isEntryOK = True
endProc:
   Exit Function
End Function

Private Function saveSPTransaction()
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsProcName As String
   Dim lnCtr As Integer
   
   lsProcName = "saveSPTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   ''On Error GoTo errProc
   
   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc
      
      For lnCtr = 0 To p_oRSDetail.RecordCount - 1
         .Detail(pnCtr, "sPartsIDx") = Detail(lnCtr, "sPartsIDx")
         .Detail(pnCtr, "nQuantity") = Detail(lnCtr, "nQuantity")
         pnCtr = pnCtr + 1
      Next
      
      If .GCardRedeem(p_oRSMaster("sTransNox"), _
                        p_oRSMaster("dTransact"), _
                        xeModeAddNew) = False Then
         GoTo endProc
      End If
   End With
   
   saveSPTransaction = True
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function delSPTransaction()
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsProcName As String
   Dim lnCtr As Integer
   
   lsProcName = "saveSPTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   ''On Error GoTo errProc
   
   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc
      
      For lnCtr = 0 To p_oRSDetail.RecordCount - 1
         .Detail(pnCtr, "sPartsIDx") = Detail(lnCtr, "sPartsIDx")
         .Detail(pnCtr, "nQuantity") = Detail(lnCtr, "nQuantity")
         pnCtr = pnCtr + 1
      Next
      
      If .GCardRedeem(p_oRSMaster("sTransNox"), _
                        p_oRSMaster("dTransact"), _
                        xeModeDelete) = False Then
         GoTo endProc
      End If
   End With
   
   saveSPTransaction = True
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Property Let Master(ByVal foIndex As Variant, ByVal foValue As Variant)
   If Not pbInitTran Then Exit Property
   
   If Not IsNumeric(foIndex) Then foIndex = LCase(foIndex)
   
   Select Case foIndex
      Case "sgcardnox", "scompnyid", "dtransact", "ssourceno", "ssourcecd", _
            "spromoidx", "ntranamtx", "npointsxx", "sremarksx", "cisbilled", _
            "dcardexpr", "navlpoint", "dlastredm", "spromcode", "npointsxx"
         p_oRSMaster(foIndex) = foValue
      Case 0 To 10, 12 To 14
         p_oRSMaster(foIndex) = foValue
         RaiseEvent MasterRetrieved(foIndex)
      Case 11, "scardnmbr"
         getClient foValue, False, False
      Case 15, "spromcode"
         getPromo foValue, False, False
      Case "scompnynm", 80
      Case "xaddressx", 81
      Case "dbirthdte", 82
   End Select
End Property

Property Get Master(ByVal foIndex As Variant)
   If Not pbInitTran Then Exit Property
   
   If Not IsNumeric(foIndex) Then foIndex = LCase(foIndex)
   
   Select Case foIndex
      Case "sgcardnox", "scompnyid", "dtransact", "ssourceno", "ssourcecd", _
            "spromoidx", "ntranamtx", "npointsxx", "sremarksx", "cisbilled", _
            "scardnmbr", "dcardExpr", "navlpoint", "dlastredm", "spromcode", "npointsxx"
         Master = p_oRSMaster(foIndex)
      Case 0 To 15
         Master = p_oRSMaster(foIndex)
      Case "scompnynm", 80
         Master = p_sClientNm
      Case "xaddressx", 81
         Master = p_sAddressx
      Case "dbirthdte", 82
         Master = p_dBirthDte
   End Select
End Property

Property Get Detail(ByVal fnIndex As Integer, ByVal foIndex As Variant)
   If Not pbInitTran Then Exit Property
   
   If fnIndex > p_oRSDetail.RecordCount - 1 Then Exit Sub
   p_oRSDetail.Move fnIndex, adBookmarkFirst
   
   If Not IsNumeric(foIndex) Then foIndex = LCase(foIndex)
   
   Select Case foIndex
      Case "spartsidx", "sbarrcode", "sdescript", "nqtyonhnd", "nquantity"
         Detail = p_oRSDetail(foIndex)
      Case 0 To 4
         Detail = p_oRSDetail(foIndex)
   End Select
End Property

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lasSplit() As String
   
   lsProcName = "SearchTransaction"
   'On Error GoTo errProc
   
   If Not pbInitTran Then GoTo endProc
   
   SearchTransaction = False
   If Not IsMissing(bByCode) Then
      If Not IsMissing(sSearch) Then
         If bByCode Then
            lsSQL = "a.sTransNox = " & strParm(sSearch)
         Else
            lsSQL = "c.sCardNmbr LIKE " & strParm("%" & sSearch)
         End If
      End If
   Else
      If Not IsMissing(sSearch) Then
         lsSQL = "d.sCompnyNm LIKE " & strParm(sSearch & "%")
      End If
   End If

   lsSQL = AddCondition(p_sSQLBrowse, lsSQL)
   
   Set loRS = New Recordset
   Debug.Print lsSQL
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   If loRS.EOF Then GoTo endProc
   
   lsSQL = KwikBrowse(p_oAppDrivr, _
                        loRS, _
                        "sTransNox»sCompnyNm»dTransact»sCardNmbr", _
                        "Trans. No»Customer»Date»Card No", _
                        "@»@»Mmm dd, yyyy»@", _
                        "a.sTransNox»d.sCompnyNm»a.dTransact»c.sCardNmbr")
   
   If lsSQL <> "" Then
      lasSplit = Split(lsSQL, "»")
      SearchTransaction = OpenTransaction(lasSplit(0))
   End If

endProc:
   Set loRS = Nothing
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function OpenTransaction(ByVal fsTransNox As String) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   
   OpenTransaction = False
   If fsTransNox = "" Then GoTo endProc
   
   lsSQL = AddCondition(p_sSQLMaster, "a.sTransNox = " & strParm(fsTransNox))
   
   Set p_oRSMaster = New Recordset
   
   Debug.Print lsSQL
   p_oRSMaster.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set p_oRSMaster.ActiveConnection = Nothing
   
   
   If p_oRSMaster.EOF Then GoTo endProc
   
   p_dBirthDte = p_oRSMaster("dBirthDte")
   p_sClientNm = p_oRSMaster("sCompnyNm")
   p_sAddressx = Trim(p_oRSMaster("xAddressx"))
   
   Call loadPromo
   
   OpenTransaction = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   
   If Not pbInitTran Then GoTo endProc
   
   Set p_oRSMaster = New Recordset
   
   p_oRSMaster.Open AddCondition(p_sSQLMaster, "0=1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set p_oRSMaster.ActiveConnection = Nothing
   
   With p_oRSMaster
      .AddNew
      .Fields("sTransNox") = GetNextCode(pxeMasTable, "sTransNox", True, p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
      .Fields("sGCardNox") = ""
      .Fields("sCompnyID") = ""
      .Fields("dTransact") = p_oAppDrivr.ServerDate
      .Fields("sSourceNo") = ""
      .Fields("sSourceCd") = pxeSourceCd
      .Fields("sPromoIDx") = ""
      .Fields("nTranAmtx") = 0
      .Fields("nPointsxx") = 0
      .Fields("sRemarksx") = ""
      .Fields("cIsBilled") = "0"
      .Fields("sCardNmbr") = ""
      .Fields("dCardExpr") = Null
      .Fields("nAvlPoint") = 0
      .Fields("dLastRedm") = Null
      .Fields("sPromCode") = ""
      
      p_sClientNm = ""
      p_sAddressx = ""
      p_dBirthDte = "1900-01-01"
   End With
   
   Call InitDetail
   
   p_nEditMode = xeModeAddNew
   
   NewTransaction = True
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function InitDetail()
   Set p_oRSDetail = New Recordset
   With p_oRSDetail
      .Fields.Append "sPartsIDx", adVarChar, 12
      .Fields.Append "sBarrCode", adVarChar, 20
      .Fields.Append "sDescript", adVarChar, 50
      .Fields.Append "nQtyOnHnd", adInteger, 4
      .Fields.Append "nQuantity", adInteger, 4
      .Open
   End With
End Function

Function SearchMaster(ByVal foIndex As Variant, _
                        ByVal fsValue As Variant, _
                        Optional ByVal fbCode As Boolean = False)
                        
   If Not IsNumeric(foIndex) Then foIndex = LCase(foIndex)
   
   Select Case foIndex
   Case 11, "scardnmbr"
      SearchMaster = getClient(fsValue, fbCode, True)
   Case 15, "spromcode"
      SearchMaster = getPromo(fsValue, fbCode, True)
   End Select
End Function

Private Function getClient(ByVal fsValue As String, _
                           Optional ByVal fbCode As Boolean = False, _
                           Optional ByVal fbSearch As Boolean = False)
                           
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsCondition As String
   Dim lasSplit() As String
   Dim loRS As Recordset
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   
   If fsValue = "" Then
      getClient = False
      GoTo endProc
   End If
   
   If fbCode Then
      If Trim(p_oRSMaster("sGCardNox")) = fsValue Then
         getClient = True
         GoTo endProc
      End If
   
      lsCondition = "a.sGCardNox = " & strParm(fsValue)
   Else
      If Trim(p_oRSMaster("sCardNmbr")) = fsValue Then
         getClient = True
         GoTo endProc
      End If
   
      If fbSearch Then
         lsCondition = "a.sCardNmbr LIKE " & strParm("%" & fsValue)
      Else
         lsCondition = "a.sCardNmbr = " & strParm(fsValue)
      End If
   End If
   
   lsSQL = "SELECT" & _
               "  a.sGCardNox" & _
               ", a.sCardNmbr" & _
               ", a.dCardExpr" & _
               ", a.nAvlPoint" & _
               ", b.sCompnyNm" & _
               ", CONCAT(IFNULL(b.sHouseNox, ''), ' ', b.sAddressx, ', ', c.sTownName, ' ', d.sProvIDxx) xAddressx" & _
               ", b.dBirthDte" & _
            " FROM G_Card_Master a" & _
               ", Client_Master b" & _
                  " LEFT JOIN TownCity c" & _
                     " ON b.sTownIDxx = c.sTownIDxx" & _
                  " LEFT JOIN Province d" & _
                     " ON c.sProvIDxx = d.sProvIDxx" & _
            " WHERE a.sClientID = b.sClientID" & _
               " AND a.cCardStat = '4'" & _
               " AND a.dCardExpr <= " & dateParm(p_oAppDrivr.ServerDate)
               
   lsSQL = AddCondition(lsSQL, lsCondition)
   
   Set loRS = New Recordset
   Debug.Print lsSQL
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   With loRS
      If .EOF Then
         MsgBox "No record found. Please verify your entry.", vbInformation, "Notice"
         initGCardInfo
         getClient = False
         GoTo endProc
      End If
   
      If .RecordCount = 1 Then
         p_oRSMaster("sGCardNox") = loRS("sGCardNox")
         p_oRSMaster("sCardNmbr") = loRS("sCardNmbr")
         p_oRSMaster("dCardExpr") = loRS("dCardExpr")
         p_oRSMaster("nAvlPoint") = loRS("nAvlPoint")
         p_dBirthDte = loRS("dBirthDte")
         p_sClientNm = loRS("sCompnyNm")
         p_sAddressx = Trim(loRS("xAddressx"))
      Else
         lsSQL = KwikSearch(p_oAppDrivr, _
                              lsSQL, _
                              "sCardNmbr»sCompnyNm»dBirthDte", _
                              "Card Number»Name»Birthdate", _
                              "@»@»Mmm dd, yyyy")
                              
         If lsSQL <> "" Then
            lasSplit = Split(lsSQL, "»")
            p_oRSMaster("sGCardNox") = lasSplit(0)
            p_oRSMaster("sCardNmbr") = lasSplit(1)
            p_oRSMaster("dCardExpr") = lasSplit(2)
            p_oRSMaster("nAvlPoint") = lasSplit(3)
            p_oRSMaster("dBirthDte") = lasSplit(6)
            p_sClientNm = lasSplit(4)
            p_sAddressx = Trim(lasSplit(5))
         End If
      End If
   End With
   
   getClient = True
endProc:
   RaiseEvent MasterRetrieved(11)
   RaiseEvent MasterRetrieved(12)
   RaiseEvent MasterRetrieved(13)
   RaiseEvent MasterRetrieved(14)
   RaiseEvent MasterRetrieved(80)
   RaiseEvent MasterRetrieved(81)
   RaiseEvent MasterRetrieved(82)

   Set loRS = Nothing
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function getPromo(ByVal fsValue As String, _
                           Optional ByVal fbCode As Boolean = False, _
                           Optional ByVal fbSearch As Boolean = False)
                           
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsCondition As String
   Dim lasSplit() As String
   Dim loRS As Recordset
   Dim lnAmount As Double
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   
   If Trim(p_oRSMaster("sGCardNox")) = "" Then
      MsgBox "Invalid GCard Holder detected." & vbCrLf & vbCrLf & _
               "Please load a GCard number.", vbInformation, "Notice"
      getPromo = False
      GoTo endProc
   End If
   
   If fbCode Then
      lsCondition = "a.sPromoIDx = " & strParm(fsValue)
   Else
      If fbSearch Then
         lsCondition = "a.sPromCode LIKE " & strParm(fsValue & "%")
      Else
         lsCondition = "a.sPromCode = " & strParm(fsValue)
      End If
   End If
   
   lsSQL = "SELECT" & _
               "  b.sTransNox" & _
               ", a.sTransNox xPromoIDx" & _
               ", a.sPromCode" & _
               ", a.nPointsxx" & _
            " FROM G_Card_Promo_Master a" & _
               ", G_Card_Order_Redeem b" & _
            " WHERE a.sTransNox = b.sPromoIDx" & _
               " AND b.cTranStat = '2'" & _
               " AND b.cPlcOrder = '1'" & _
               " AND b.sGCardNox = " & strParm(p_oRSMaster("sGCardNox"))
               
            
   lsSQL = AddCondition(lsSQL, lsCondition)
   
   Set loRS = New Recordset
   Debug.Print lsSQL
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   With loRS
      If .EOF Then
         MsgBox "Orders are not available for this customer." & vbCrLf & vbCrLf & _
                  "Please check system monitor for order status.", vbInformation, "Notice"
         initGCardPromo
         getPromo = False
         GoTo endProc
      End If
   
      If .RecordCount = 1 Then
         p_oRSMaster("sSourceNo") = loRS("sTransNox")
         p_oRSMaster("sPromoIDx") = loRS("xPromoIDx")
         p_oRSMaster("nTranAmtx") = 0
         p_oRSMaster("nPointsxx") = loRS("nPointsxx")
         p_oRSMaster("sPromCode") = loRS("sPromCode")
      Else
         lsSQL = KwikSearch(p_oAppDrivr, _
                              lsSQL, _
                              "sTransNox»sPromCode»nPointsxx", _
                              "Order No»Code»Points", _
                              "@»@»#,##0")
                              
         If lsSQL <> "" Then
            lasSplit = Split(lsSQL, "»")
            p_oRSMaster("sSourceNo") = lasSplit(0)
            p_oRSMaster("sPromoIDx") = lasSplit(1)
            p_oRSMaster("nTranAmtx") = 0
            p_oRSMaster("nPointsxx") = lasSplit(3)
            p_oRSMaster("sPromCode") = lasSplit(2)
         End If
      End If
      
      Call loadPromo
   End With
   
   getPromo = True
endProc:
   RaiseEvent MasterRetrieved(8)
   RaiseEvent MasterRetrieved(15)
   
   Set loRS = Nothing
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Sub loadPromo()
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lnAmount As Double
   
   'now search the promo items
   lsSQL = "SELECT" & _
               "  a.sPartsIDx" & _
               ", b.sBarrCode" & _
               ", b.sDescript" & _
               ", c.nQtyOnHnd" & _
               ", a.nQuantity" & _
               ", b.nSelPrice" & _
            " FROM G_Card_Promo_Detail a" & _
               ", Spareparts b" & _
               ", SP_Inventory c" & _
            " WHERE a.sPartsIDx = b.sPartsIDx" & _
               " AND b.sPartsIDx = c.sPartsIDx" & _
               " AND c.sBranchCD = " & strParm(p_sBranchCd) & _
               " AND a.sTransNox = " & strParm(p_oRSMaster("sPromoIDx"))
   
   Set loRS = New Recordset
   
   Debug.Print lsSQL
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   InitDetail
   lnAmount = 0
   Do Until loRS.EOF
      p_oRSDetail.AddNew
      p_oRSDetail("sPartsIDx") = loRS("sPartsIDx")
      p_oRSDetail("sBarrCode") = loRS("sBarrCode")
      p_oRSDetail("sDescript") = loRS("sDescript")
      p_oRSDetail("nQtyOnHnd") = IFNull(loRS("nQtyOnHnd"), 0)
      p_oRSDetail("nQuantity") = IFNull(loRS("nQuantity"), 0)
      
      lnAmount = lnAmount + IFNull(loRS("nSelPrice"), 0)
      loRS.MoveNext
   Loop
   
   'set transaction amount
   p_oRSMaster("nTranAmtx") = lnAmount
   
   Set loRS = Nothing
End Sub

Private Sub initGCardInfo()
   p_oRSMaster("sGCardNox") = ""
   p_oRSMaster("sCardNmbr") = ""
   p_oRSMaster("dCardExpr") = Null
   p_oRSMaster("nAvlPoint") = 0
    
   p_sClientNm = ""
   p_sAddressx = ""
   p_dBirthDte = "1900-01-01"
End Sub

Private Sub initGCardPromo()
   p_oRSMaster("sSourceNo") = ""
   p_oRSMaster("sPromoIDx") = ""
   p_oRSMaster("nTranAmtx") = 0
   p_oRSMaster("nPointsxx") = 0
   p_oRSMaster("sPromCode") = ""
End Sub

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub
