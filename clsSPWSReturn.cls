VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPWSReturn"
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
'     Spareparts WholeSale Return Object
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
'  XerSys [ 11/15/2004 03:31 am ]
'     Start creating this object.
'  Kalyptus [ 12/11/2006 10:08 am]
'     Implements the BugTracker
'  Kalyptus [ 04/02/2007 10:31 am]
'     Updated the p_oBase_Clas to conform it with the clsMasterDetail specification
'  XerSys [ 08/14/2007 01:47 pm ]
'     Update this object to adopt the new implementation of SP Inventory Transaction
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME = "clsSPWSReturn"
Private Const pxeSystemCode As String = "SP"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1

Private p_sBranchCd As String
Private p_sFullName As String
Private p_sAddressx As String
Private p_bExcludex As Boolean

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

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Get LocalTrans() As Boolean
   LocalTrans = p_bExcludex
End Property

Property Let LocalTrans(ByVal Value As Boolean)
   p_bExcludex = Value
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
   Case 6, "nquantity"
      CheckReturn Row, Value
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
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
      p_sFullName = Value
   Case 4, "xaddressx"
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

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail

   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "SP_WSO_Return_Master"
      .DetailTable = "SP_WSO_Return_Detail"
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", b.sCompnyNm" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName) xFullname" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName) xAddressx" & _
                        ", a.sRemarksx" & _
                        ", a.nTranTotl" & _
                        ", a.cTranStat" & _
                        ", a.sApproved" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", a.sClientID" & _
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
                        ", d.sTypeName" & _
                        ", e.sModelNme" & _
                        ", a.nUnitPrce" & _
                        ", a.nQuantity" & _
                        ", a.sTransNox" & _
                        ", a.dModified" & _
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
                              " ON b.sModelIdx = e.sModelIdx" & _
                        ", SP_Inventory c" & _
                        ", SP_Type d" & _
                     " WHERE a.sPartsIDx = b.sPartsIDx" & _
                        " AND b.sPartsIDx = c.sPartsIDx" & _
                        " AND b.cPartType = d.cPartType" & _
                        " AND LEFT(a.sTransNox,4) = c.sBranchCd " & _
                     " ORDER BY a.nEntryNox"
      '" AND c.sBranchCd = " & strParm(p_sBranchCd)
      .Reference = "sTransNox|a.sTransNox"
      
      InitTransaction = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID" & _
                        " ORDER BY b.sCompnyNm" & _
                           ", a.dTransact"
      'she 2017-03-09
      '" AND LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = " & strParm(p_sBranchCd)
      .BrowseColumn(0) = "sTransNox»sCompnyNm»dTransact"
      .BrowseCriteria(0) = "a.sTransNox»b.sCompnyNm»a.dTransact"
      .BrowseTitle(0) = "Transaction No»Company Name»Date"
      .BrowseFormat(0) = "@@-@@@@@@»@»MMMM DD, YYYY"

      .BrowseDetailQuery(0) = "SELECT" & _
                                 " DISTINCT a.sBarrCode" & _
                                 ", a.sDescript" & _
                                 ", c.sTypeName" & _
                                 ", d.sModelNme" & _
                                 ", e.nUnitPrce * ((100 - e.nDiscount) / 100) nNetPrice" & _
                                 ", a.sPartsIDx" & _
                                 ", b.nQtyOnHnd" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nReorderx" & _
                                 ", b.nLedgerNo" & _
                              " FROM Spareparts a" & _
                                    " LEFT JOIN SP_Model d" & _
                                       " ON a.sModelIDx = d.sModelIDx" & _
                                 ", SP_Inventory b" & _
                                 ", SP_Type c" & _
                                 ", SP_WSO_Detail e" & _
                                 ", SP_WSO_Master f" & _
                              " WHERE a.sPartsIDx = b.sPartsIDx" & _
                                 " AND a.cPartType = c.cPartType" & _
                                 " AND a.sPartsIDx = e.sPartsIDx" & _
                                 " AND e.sTransNox = f.sTransNox" & _
                                 " AND LEFT(e.sTransNox,4) = b.sBranchCd " & _
                              " ORDER BY a.sDescript"
      '" AND b.sBranchCd = " & strParm(p_sBranchCd)
      .BrowseDetailColumn(0) = "sBarrCode»sDescript»sTypeName»sModelNme»nNetPrice"
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
   
   If Not saveARClient() Then GoTo endProc
   If Not saveClientTransaction() Then GoTo endProc

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
    ShowError lsProcName & "( " & " )"
End Function

Function UpdateTransaction() As Boolean
   p_oBaseClas.EditMode = xeModeUpdate
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Function AddDetail() As Boolean
   AddDetail = p_oBaseClas.AddDetail

   pnCtr = p_oBaseClas.ItemCount - 1
   With p_oBaseClas
      .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      .Detail(pnCtr, "sBarrCode") = ""
      .Detail(pnCtr, "sDescript") = ""
      .Detail(pnCtr, "sTypeName") = ""
      .Detail(pnCtr, "sModelNme") = ""
      .Detail(pnCtr, "nUnitPrce") = 0#
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sPartsIDx") = ""
      .Detail(pnCtr, "nQtyOnHnd") = 0
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
   
   Index = LCase(Index)
   Select Case Index
   Case 2, "scompnynm"
      getMaster Index, Value, True
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

Private Sub getDetail(ByVal lnRow As Integer, _
                      ByVal lnIndex As Integer, _
                      ByVal lsValue As String, _
                      ByVal lbSearch As Boolean)
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsProcName As String

   lsProcName = "GetDetail"
   'On Error GoTo errProc
   
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
               lsDetail = "b.sDescript LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "b.sDescript = " & strParm(Trim(lsValue))
            End If
         End If
         lsDetail = lsDetail & " AND e.sTransNox = f.sTransNox" & _
                     " AND f.sClientID = " & strParm(.Master("sClientID"))
      ElseIf lbSearch = False Then
         GoTo endWithClear
      Else
         lsDetail = "f.sClientId = " & strParm(.Master("sClientID"))
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
      .Detail(lnRow, "sTypeName") = lasDetail(2)
      .Detail(lnRow, "sModelNme") = lasDetail(3)
      .Detail(lnRow, "nUnitPrce") = lasDetail(4)
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "sPartsIDx") = lasDetail(5)
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(6)
      .Detail(lnRow, "nResvOrdr") = lasDetail(7)
      .Detail(lnRow, "nBackOrdr") = lasDetail(8)
'      .Detail(lnRow, "nReorderx") = lasDetail(9)
      .Detail(lnRow, "nLedgerNo") = lasDetail(10)
      Debug.Print .Detail(lnRow, "nQtyOnHnd")
      Debug.Print .Detail(lnRow, "nResvOrdr")
      Debug.Print .Detail(lnRow, "nBackOrdr")
      Debug.Print .Detail(lnRow, "nReorderx")
      Debug.Print .Detail(lnRow, "nLedgerNo")
   End With
   
endProc:
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "sTypeName") = ""
      .Detail(lnRow, "sModelNme") = ""
      .Detail(lnRow, "nUnitPrce") = 0#
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

Private Sub CheckReturn(ByVal lnRow As Integer, ByVal lnValue As Integer)
   Dim lors As Recordset
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "CheckReturn"
   'On Error GoTo errProc
   
   With p_oBaseClas
      If .Detail(lnRow, "sBarrCode") = "" Then GoTo endProc
      If lnValue = 0 Then GoTo endProc
      
      lsSQL = "SELECT SUM(a.nQuantity)" & _
               " FROM SP_WSO_Detail a" & _
                  ", SP_WSO_Master b" & _
               " WHERE a.sTransNox = b.sTransNox" & _
                  " AND a.sPartsIDx = " & strParm(.Detail(lnRow, "sPartsIDx")) & _
                  " AND a.nUnitPrce * ((100 - a.nDiscount) / 100) = " & .Detail(lnRow, "nUnitPrce") & _
                  " AND b.sClientID = " & strParm(.Master("sClientID"))
                  
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If lors.EOF Then
         MsgBox "Unable to Retreive Spareparts Sales Info!!!", vbCritical, "Warning"
         GoTo endWithClear
      End If
      
      If lors(0) < lnValue Then
         MsgBox "Return Item is Greater than the Actual Purchase!!!" & vbCrLf & _
                  "Verify your Entry then Try Again!!!" & lors(0), vbCritical, "Warning"
         GoTo endWithClear
      End If
      .Detail(lnRow, "nQuantity") = lnValue
   End With
   RaiseEvent DetailRetrieved(6)
   
endProc:

   Exit Sub
endWithClear:
   p_oBaseClas.Detail(lnRow, "nQuantity") = 0
   RaiseEvent DetailRetrieved(6)
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lnRow _
                       & ", " & lnValue & " )"
End Sub

Private Sub getMaster(ByVal lnIndex As Integer, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "GetMaster"
   'On Error GoTo errProc
   
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

Private Function delSPTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsProcName As String
   
   lsProcName = "delSPTransaction"
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
      If .WholeSaleReturn(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("dTransact"), _
            xeModeDelete) = False Then
         GoTo endProc
      End If
   End With
   
   If SaveSPCustomerTrans(xeModeDelete) = False Then GoTo endProc
   
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
   Dim lnRow As Long
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_InitMaster"
   'On Error GoTo errProc
   
   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sCompnyNm") = ""
      .Master("sRemarksx") = ""
      .Master("nTranTotl") = 0#
      .Master("cTranStat") = xeStateOpen
      .Master("sApproved") = 0
      .Master("sClientID") = ""

      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "sTypeName") = ""
      .Detail(0, "sModelNme") = ""
      .Detail(0, "nUnitPrce") = 0#
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sPartsIDx") = ""
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nResvOrdr") = 0
      .Detail(0, "nBackOrdr") = 0
      .Detail(0, "nReorderx") = 0
      .Detail(0, "nLedgerNo") = 0
   End With
        
   p_sFullName = ""
   p_sAddressx = ""

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      p_sAddressx = .Master("xAddressx")
      p_sFullName = .Master("xFullName")
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
   Dim loCltTrans As clsClientTrans
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnRow As Long

   lsProcName = "p_oBaseClas_SaveOthers"
   'On Error GoTo errProc
   Cancel = True

   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc

      For pnCtr = 0 To p_oBaseClas.ItemCount - 1
         If p_oBaseClas.Detail(pnCtr, "sPartsIDx") = "" Then Exit For
'         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
'         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         
         .Detail(pnCtr, "sPartsIDx") = p_oBaseClas.Detail(pnCtr, "sPartsIDx")
         .Detail(pnCtr, "nQuantity") = p_oBaseClas.Detail(pnCtr, "nQuantity")
         .Detail(pnCtr, "nQtyOnHnd") = p_oBaseClas.Detail(pnCtr, "nQtyOnHnd")
         .Detail(pnCtr, "nResvOrdr") = p_oBaseClas.Detail(pnCtr, "nResvOrdr")
         .Detail(pnCtr, "nBackOrdr") = p_oBaseClas.Detail(pnCtr, "nBackOrdr")
'         .Detail(pnCtr, "nReorderx") = p_oBaseClas.Detail(pnCtr, "nReorderx")
         .Detail(pnCtr, "nLedgerNo") = p_oBaseClas.Detail(pnCtr, "nLedgerNo")
      Next
      
      ' if a spareparts is not successfully updated, cancel saving.
      If .WholeSaleReturn(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("dTransact"), _
            p_oBaseClas.EditMode) = False Then
         GoTo endProc
      End If
   
      If SaveSPCustomerTrans(p_oBaseClas.EditMode) = False Then GoTo endProc
   
   End With
   
     
   Cancel = False
   
endProc:
   Set loSPTrans = Nothing
   Set loCltTrans = Nothing

   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
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
         If p_oBaseClas.Detail(pnCtr, "sPartsIDx") = "" Then Exit For
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

Private Function SaveSPCustomerTrans(fnEditMode As xeEditMode) As Boolean
   Dim loCltTrans As clsClientTrans
   Dim lsOldProc As String

   lsOldProc = "SaveSPCustomerTrans"
   'On Error GoTo errProc
   
   Set loCltTrans = New clsClientTrans
   With loCltTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .ClientType = xeClientCustomer
      .SourceSystem = pxeSystemCode
      Call .InitTransaction
      
      ' now save the transaction
      If .Sales(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("sClientID"), _
            p_oBaseClas.Master("dTransact"), _
            0#, _
            p_oBaseClas.Master("nTranTotl"), _
            fnEditMode) = False Then
         GoTo endProc
      End If
   End With
   
   SaveSPCustomerTrans = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "(" & fnEditMode & ")"
   GoTo endProc
End Function

Private Function saveARClient() As Boolean
   Dim loClient As clsARClient
   Dim lors As Recordset
   Dim lsProcName As String
   
   lsProcName = "saveARClient"
   Debug.Print pxeMODULENAME & "." & lsProcName
   '''On Error GoTo errProc
   
   Set lors = New Recordset

   lors.Open "SELECT a.*, b.sClientID xClientID" & _
               " FROM SP_Customer a" & _
                  " LEFT JOIN AR_Client_Master b" & _
                     " ON a.sClientID = b.sClientID" & _
               " WHERE a.sClientID = " & strParm(Master("sClientID")) _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If Not IsNull(lors("xClientID")) Then
      saveARClient = True
      GoTo endProc
   End If
   
   Set loClient = New clsARClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .InitRecord
      .NewRecord
            
      .Master("sClientID") = lors("sClientID")
      .Master("sTelNoxxx") = lors("sTelNoxxx")
      .Master("sFaxNoxxx") = lors("sFaxNoxxx")
      .Master("sCPerson1") = lors("sCPerson1")
      .Master("sCPPosit1") = lors("sCPPosit1")
      .Master("sTermIDxx") = lors("sTermIDxx")
      .Master("sRemarksx") = lors("sRemarksx")
      .Master("nCredLimt") = lors("nCredLimt")
      .Master("nABalance") = lors("nABalance")
      .Master("dCltSince") = lors("dCltSince")
      .Master("nLedgerNo") = lors("nLedgerNo")
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

Private Function saveClientTransaction() As Boolean
   Dim loClientsTrans As clsARClientTrans
   Dim lsProcName As String
   
   lsProcName = "saveClientTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   '''On Error GoTo errProc
   
   Set loClientsTrans = New clsARClientTrans
   With loClientsTrans
      Set .AppDriver = p_oAppDrivr
      .ClientType = xeClientCustomer
      .SourceSystem = "SP"
      .Branch = p_sBranchCd
      Call .InitTransaction
   End With
   
   With p_oBaseClas
      If Not loClientsTrans.WholeSaleReturn(.Master("sTransNox") _
                                          , .Master("sClientID") _
                                          , p_oAppDrivr.ServerDate _
                                          , 0# _
                                          , .Master("nTranTotl") _
                                          , xeModeAddNew) Then
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

Function ViewEncoder()
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lsDate As Date
   Dim lsEncoder As String
   Dim lsApprved As String
   
   lsSQL = "SELECT sModified, dModified from xxxReplicationLog" & _
            " WHERE sTransNox LIKE " & strParm(Left(p_oBaseClas.Master("sTransNox"), 6) & "%") & _
               " AND sStatemnt LIKE '%INSERT INTO%sTransNox = %" & p_oBaseClas.Master("sTransNox") & "%'" & _
               " AND sTableNme = 'SP_WSO_Return_Master'"
   
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set lors.ActiveConnection = Nothing
   
   If lors.EOF Then
      MsgBox "Transaction has no history of encoding.", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lsDate = CDate(lors("dModified"))
   lsSQL = "SELECT sUserName" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & strParm(lors("sModified"))

   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If lors.EOF Then
      MsgBox "User does not Exist in the User Table!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lsEncoder = lors("sUserName")
   
   lsSQL = "SELECT sUserName" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & strParm(p_oBaseClas.Master("sApproved"))
   
   Debug.Print lsSQL
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If lors.EOF Then
      lsApprved = "NONE"
   Else
      lsApprved = lors("sUserName")
   End If
   
   MsgBox "    ««« WHOLESALE RETURN TRANSACTION »»»" & vbCrLf & vbCrLf & _
                "Encoded By   : " & Decrypt(lsEncoder, p_oAppDrivr.Machinex) & vbCrLf & _
                "Date Encoded : " & Format(lsDate, "MMMM DD, YYYY HH:NN:SS") _
                , vbInformation, "Notice"
endProc:
    Exit Function
End Function


