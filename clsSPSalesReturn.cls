VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPSalesReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
'
' Copyright 2003 and beyond
' All Rights Reserved
'
'     Spareparts Point-of-Sale Returns Object(Branch Version)
'
'  Note:
'     Be sure to set the transaction no before calling the InitTransaction
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
'     Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsSPSalesReturn"
Private Const pxeSourceCd As String = "SPSR"
Private Const pxeGracePrd As Integer = 180       'Set the default grace period to 14 days requeste
                                                 '2016-12-12 she set the grace period to 180 days requested by support team approved by sir rex
Private p_oAppDrivr As clsAppDriver
Private p_oDetail As Recordset

Private p_sSQLMastr As String
Private p_bInitClas As Boolean

Private p_sTransNox As String
Private p_dTransact As Date
Private p_sClientID As String
Private p_sClientNm As String
Private p_sAddressx As String
Private p_sDeletedx As String
Private p_nTranTotl As Currency

Public Event DetailRetrieved(ByVal Index As Integer)

'Sets the AppDriver to use
Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get TransNo() As String
   TransNo = p_sTransNox
End Property

Property Let TransNo(Value As String)
   p_sTransNox = Value
End Property

Property Get TransactDate() As Date
   TransactDate = p_dTransact
End Property

Property Let TransactDate(Value As Date)
   p_dTransact = Value
End Property

Property Get ClientID() As String
   ClientID = p_sClientID
End Property

Property Let ClientID(Value As String)
   p_sClientID = Value
End Property

Property Get ClientNm() As String
   ClientNm = p_sClientNm
End Property

Property Let ClientNm(Value As String)
   p_sClientNm = Value
End Property

Property Get Address() As String
   Address = p_sAddressx
End Property

Property Let Address(Value As String)
   p_sAddressx = Value
End Property

Property Get Total() As Currency
   Total = p_nTranTotl
End Property

Property Let Total(Value As Currency)
   p_nTranTotl = Value
End Property

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   Dim lsOldProc As String
   
   lsOldProc = "(Get)Detail"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If Not p_bInitClas Then GoTo endProc
   
   If Row - 1 > p_oDetail.RecordCount Then
      GoTo endProc
'   ElseIf Row = 0 Then
   Else
      If p_oDetail.AbsolutePosition - 1 <> Row Then p_oDetail.MoveFirst
'   Else
'      If p_oDetail.AbsolutePosition - 1 <> Row Then p_oDetail.Move Row, 0
   End If
   
   Detail = p_oDetail(Index)

endProc:
   Exit Property
errProc:
   ShowError lsOldProc & "( " & Index & " )"
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   Dim lsOldProc As String
   
   lsOldProc = "(Let)Detail"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If Not p_bInitClas Then GoTo endProc
   
   If Row + 1 > p_oDetail.RecordCount Then
      GoTo endProc
   Else
      If Row = 0 Then
         p_oDetail.MoveFirst
      Else
         p_oDetail.MoveFirst
         p_oDetail.Move Row, 0
      End If
   End If
      
   Select Case Index
   Case 1, "ssalesinv"
      Call getInvoice(Value)
   Case 2, "sbarrcode"
      Call getItem(Value, True)
   Case 3, "sdescript"
      Call getItem(Value, False)
   Case 5
      p_oDetail("nQuantity") = Value
   Case 15, "dmodified"
   End Select

endProc:
   Exit Property
errProc:
   ShowError lsOldProc & "( " & Index _
                       & ", " & Value & " )"
End Property

Property Get ItemCount() As Long
   If p_bInitClas = False Then Exit Property
   On Error Resume Next
   
   ItemCount = p_oDetail.RecordCount
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsCondition As String
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   If p_sTransNox = Empty Then GoTo endProc

   p_sSQLMastr = "SELECT" & _
                         "  a.nEntryNox" & _
                         ", e.sSalesInv" & _
                         ", b.sBarrCode" & _
                         ", b.sDescript" & _
                         ", c.nQtyOnHnd" & _
                         ", a.nQuantity" & _
                         ", a.nUnitPrce" & _
                         ", a.nDiscount" & _
                         ", a.nAddDiscx" & _
                         ", a.sPartsIDx" & _
                         ", a.sReferNox" & _
                         ", c.nResvOrdr" & _
                         ", c.nBackOrdr" & _
                         ", c.nReorderx" & _
                         ", c.nLedgerNo" & _
                         ", a.dModified" & _
                         ", a.nQuantity xQuantity"

   p_sSQLMastr = p_sSQLMastr & _
                 " FROM SP_SO_Return a" & _
                       " LEFT JOIN Spareparts b" & _
                          " ON a.sPartsIDx = b.sPartsIDx" & _
                       " LEFT JOIN SP_Inventory c" & _
                          " ON b.sPartsIDx = c.sPartsIDx AND c.sBranchCd = " & strParm(Left(p_sTransNox, Len(p_oAppDrivr.BranchCode))) & _
                       " LEFT JOIN SP_SO_Master e" & _
                          " ON a.sReferNox = e.sTransNox" & _
                 " ORDER BY a.nEntryNox"
   
   Debug.Print p_sSQLMastr
   
   p_bInitClas = True
   InitTransaction = p_bInitClas
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "NewTransaction"
   'On Error GoTo errProc

   If Not p_bInitClas Then GoTo endProc

   lsSQL = AddCondition(p_sSQLMastr, "0 = 1")
   Set p_oDetail = New ADODB.Recordset
   With p_oDetail
      .Open lsSQL, p_oAppDrivr.Connection, adOpenDynamic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing
   
      .AddNew
   
      Call InitDetail
   End With
   
   p_sDeletedx = ""

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Sub InitDetail()
   p_oDetail("sSalesInv") = ""
   p_oDetail("sPartsIDx") = ""
   p_oDetail("sBarrCode") = ""
   p_oDetail("sDescript") = ""
   p_oDetail("nQtyOnHnd") = 0
   p_oDetail("nQuantity") = 0
   p_oDetail("nUnitPrce") = 0
   p_oDetail("nDiscount") = 0
   p_oDetail("nAddDiscx") = 0
   p_oDetail("sReferNox") = 0
   p_oDetail("nResvOrdr") = 0
   p_oDetail("nBackOrdr") = 0
   p_oDetail("nReorderx") = 0
   p_oDetail("xQuantity") = 0
End Sub

Function SaveTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lnCtr As Integer
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "SaveTransaction"
   'On Error GoTo errProc
   
   If Not p_bInitClas Then GoTo endProc

   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = Left(p_sTransNox, Len(p_oAppDrivr.BranchCode))
      If .InitTransaction() = False Then GoTo endProc
   
      'Delete deleted record
      If p_sDeletedx <> "" Then
         p_sDeletedx = "(" & Mid(p_sDeletedx, 2) & ")"
               
         lsSQL = "DELETE FROM SP_SO_Return" & _
                " WHERE sTransNox = " & strParm(p_sTransNox) & _
                  " AND nEntryNox IN " & p_sDeletedx
         Call p_oAppDrivr.Execute(lsSQL, "SP_SO_Return")
      End If
   
      p_oDetail.MoveFirst
      lnCtr = 0
      Do Until p_oDetail.EOF
         If p_oDetail("sPartsIDx") <> "" Then
            
            If p_oDetail("nQuantity") > p_oDetail("xQuantity") Then
               MsgBox "The quantity released is higher than the quantity return for " & p_oDetail("sBarrCode") & "!"
               GoTo endProc
            End If
                        
            .Detail(lnCtr, "sPartsIDx") = p_oDetail("sPartsIDx")
            .Detail(lnCtr, "nQuantity") = p_oDetail("nQuantity")
            .Detail(lnCtr, "nQtyOnHnd") = p_oDetail("nQtyOnHnd")
            .Detail(lnCtr, "nResvOrdr") = p_oDetail("nResvOrdr")
            .Detail(lnCtr, "nBackOrdr") = p_oDetail("nBackOrdr")
            .Detail(lnCtr, "nLedgerNo") = p_oDetail("nLedgerNo")
             
            If IFNull(p_oDetail("nEntryNox"), 0) <= 0 Then
               lsSQL = "INSERT INTO SP_SO_Return SET" & _
                             "  sTransNox = " & strParm(p_sTransNox) & _
                             ", nEntryNox = " & lnCtr + 1 & _
                             ", sReferNox = " & strParm(p_oDetail("sReferNox")) & _
                             ", sPartsIDx = " & strParm(p_oDetail("sPartsIDx")) & _
                             ", nQuantity = " & p_oDetail("nQuantity") & _
                             ", nUnitPrce = " & p_oDetail("nUnitPrce") & _
                             ", nDiscount = " & p_oDetail("nDiscount") & _
                             ", nAddDiscx = " & p_oDetail("nAddDiscx") & _
                             ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
            Else
               lsSQL = "UPDATE SP_SO_Return SET" & _
                             "  nEntryNox = " & lnCtr + 1 & _
                             ", sReferNox = " & strParm(p_oDetail("sReferNox")) & _
                             ", sPartsIDx = " & strParm(p_oDetail("sPartsIDx")) & _
                             ", nQuantity = " & p_oDetail("nQuantity") & _
                             ", nUnitPrce = " & p_oDetail("nUnitPrce") & _
                             ", nDiscount = " & p_oDetail("nDiscount") & _
                             ", nAddDiscx = " & p_oDetail("nAddDiscx") & _
                             ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                     " WHERE sTransNox = " & strParm(p_sTransNox) & _
                       " AND nEntryNox = " & p_oDetail("nEntryNox")
            End If
         
            p_oAppDrivr.Execute lsSQL, "SP_SO_Return"
            lnCtr = lnCtr + 1
         End If
         p_oDetail.MoveNext
      Loop
   
      If lnCtr > 0 Or p_sDeletedx <> "" Then
         If .SalesReturn(p_sTransNox, p_dTransact, xeModeAddNew) = False Then GoTo endProc
'         If .SalesReturn(p_sTransNox, p_dTransact, xeModeUpdate) = False Then GoTo endProc
      End If
   End With
   
   SaveTransaction = True

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction() As Boolean
   Dim lsSQL  As String
   Dim lsProcName As String
   
   lsProcName = "OpenTransaction"
   'On Error GoTo errProc
   
   If Not p_bInitClas Then GoTo endProc
      
   lsSQL = AddCondition(p_sSQLMastr, "a.sTransNox = " & strParm(p_sTransNox))
   Set p_oDetail = New ADODB.Recordset
   p_oDetail.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_oDetail.ActiveConnection = Nothing
      
   If p_oDetail.RecordCount = 0 Then
      p_oDetail.AddNew
      Call InitDetail
   End If
   
   p_sDeletedx = ""
   OpenTransaction = True
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "()"
End Function

Function DeleteTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "DeleteTransaction"
   'On Error GoTo errProc
   
   If Not p_bInitClas Then GoTo endProc
   
   p_oDetail.MoveFirst
   If p_oDetail("sPartsIDx") = "" Then
      DeleteTransaction = True
      GoTo endProc
   End If
   
   lsSQL = "DELETE FROM SP_SO_Return" & _
          " WHERE sTransNox = " & strParm(p_sTransNox)
   Call p_oAppDrivr.Execute(lsSQL, "SP_SO_Return")

   If CancelTransaction = False Then GoTo endProc

   Set p_oDetail = Nothing

   DeleteTransaction = True
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function PostTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "PostTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If Not p_bInitClas Then GoTo endProc

   PostTransaction = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "()"
End Function

Function CloseTransaction()
   Dim lsProcName As String
   
   lsProcName = "CloseTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If Not p_bInitClas Then GoTo endProc

   CloseTransaction = True

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "()"
End Function

Function CancelTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lnCtr As Integer
   Dim lsProcName As String
   
   lsProcName = "CancelTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If Not p_bInitClas Then GoTo endProc

   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = Left(p_sTransNox, Len(p_oAppDrivr.BranchCode))
      If .InitTransaction() = False Then GoTo endProc
   
      lnCtr = 1
      p_oDetail.MoveFirst
      Do Until p_oDetail.EOF
         If p_oDetail("sPartsIDx") <> "" Then
            .Detail(lnCtr, "sPartsIDx") = p_oDetail("sPartsIDx")
            .Detail(lnCtr, "nQuantity") = p_oDetail("nQuantity")
            .Detail(lnCtr, "nQtyOnHnd") = p_oDetail("nQtyOnHnd")
            .Detail(lnCtr, "nResvOrdr") = p_oDetail("nResvOrdr")
            .Detail(lnCtr, "nBackOrdr") = p_oDetail("nBackOrdr")
            .Detail(lnCtr, "nLedgerNo") = p_oDetail("nLedgerNo")
            lnCtr = lnCtr + 1
         End If
         p_oDetail.MoveNext
      Loop
         
      If .ItemCount > 0 Then
         If .SalesReturn(p_sTransNox, p_dTransact, xeModeDelete) = False Then GoTo endProc
      End If
   End With
   
   CancelTransaction = True
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "()"
End Function

Function SearchDetail(ByVal Row As Long, _
                        ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsProcName As String

   lsProcName = "SearchDetail"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   SearchDetail = False

   If p_bInitClas = False Then GoTo endProc

   Index = LCase(Index)
   Select Case Index
   Case 1, "ssalesinv"
      Call getInvoice(Value)
   Case 2, "sbarrcode"
      Call getItem(Value, True)
   Case 3, "sdescript"
      Call getItem(Value, False)
   End Select
   
   SearchDetail = True
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub getInvoice(ByVal Value As String)
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "getInvoice"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   'test if there is an update in the SalesInvoice
   If Value = p_oDetail("sSalesInv") And Value <> "" Then Exit Sub

   'Load all Sales Invoice
   lsSQL = "SELECT" & _
                  "  a.sSalesInv" & _
                  ", CONCAT(b.sLastname, ', ', b.sFrstName, IF(IFNULL(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sLastName) xClientNm" & _
                  ", a.dTransact" & _
                  ", a.sTransNox" & _
          " FROM SP_SO_Master a" & _
               " LEFT JOIN Client_Master b" & _
                     " ON a.sClientID = b.sClientID" & _
          " WHERE a.sTransNox LIKE " & strParm(Left(p_sTransNox, Len(p_oAppDrivr.BranchCode)) & "%") & _
            " AND a.cTranStat <> '3'" & _
            " AND a.dTransact >= " & dateParm(p_dTransact - pxeGracePrd) & _
            " AND a.sSalesInv LIKE " & strParm(Value & "%")
   Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   
   If loRS.EOF Then
      Call InitDetail
   ElseIf loRS.RecordCount = 1 Then
      Call InitDetail
      p_oDetail("sSalesInv") = loRS("sSalesInv")
      p_oDetail("sReferNox") = loRS("sTransNox")
   ElseIf loRS.RecordCount > 1 Then
      lsSQL = KwikBrowse(p_oAppDrivr, loRS, _
                        "sSalesInv»xClientNm»dTransact", _
                        "Sales Inv#»Client Name»Date")
      If lsSQL = Empty Then GoTo endProc
      InitDetail
      lasMaster = Split(lsSQL, "»")
      p_oDetail("sSalesInv") = lasMaster(0)
      p_oDetail("sReferNox") = lasMaster(3)
   End If

endProc:
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)
   RaiseEvent DetailRetrieved(7)
   
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub getItem(ByVal Value As String, ByVal ByCode As Boolean)
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lasMaster() As String
   Dim lsFieldNm As String
   Dim lsProcName As String

   lsProcName = "getItem"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   'Sales Invoice should be present first
   If p_oDetail("sReferNox") = "" Then Exit Sub
   
   'Determine the Field Name
   lsFieldNm = IIf(ByCode, "sBarrCode", "sDescript")
   
   'test if there is an update in the Detail
   If Value = p_oDetail(lsFieldNm) And Value <> "" Then Exit Sub

   'Load all Sales Invoice
   lsSQL = "SELECT" & _
                  "  a.sPartsIDx" & _
                  ", b.sBarrCode" & _
                  ", b.sDescript" & _
                  ", a.nQuantity" & _
                  ", a.nUnitPrce" & _
                  ", a.nDiscount" & _
                  ", a.nAddDiscx" & _
                  ", c.nLedgerNo" & _
                  ", c.nResvOrdr" & _
                  ", c.nBackOrdr" & _
                  ", Ifnull(c.nReorderx,0) nReorderx" & _
                  ", c.nQtyOnHnd" & _
                  ", a.nQuantity xQuantity" & _
          " FROM SP_SO_Detail a" & _
               " LEFT JOIN Spareparts b" & _
                     " ON a.sPartsIDx = b.sPartsIDx" & _
               " LEFT JOIN SP_Inventory c" & _
                     " ON b.sPartsIDx = c.sPartsIDx AND c.sBranchCd = " & strParm(Left(p_sTransNox, Len(p_oAppDrivr.BranchCode))) & _
          " WHERE a.sTransNox = " & strParm(p_oDetail("sReferNox")) & _
            " AND a.cLaborxxx <> '1'" & _
            " AND b." & lsFieldNm & " LIKE " & strParm(Value & "%")
   
   Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   Debug.Print lsSQL

   If loRS.EOF Then
      GoTo endWithClear
   ElseIf loRS.RecordCount = 1 Then
      p_oDetail("sPartsIDx") = loRS("sPartsIDx")
      p_oDetail("sBarrCode") = loRS("sBarrCode")
      p_oDetail("sDescript") = loRS("sDescript")
      p_oDetail("nQuantity") = loRS("nQuantity")
      p_oDetail("nUnitPrce") = loRS("nUnitPrce")
      p_oDetail("nDiscount") = loRS("nDiscount")
      p_oDetail("nAddDiscx") = loRS("nAddDiscx")
      p_oDetail("nLedgerNo") = loRS("nLedgerNo")
      p_oDetail("nResvOrdr") = loRS("nResvOrdr")
      p_oDetail("nBackOrdr") = loRS("nBackOrdr")
      p_oDetail("nReorderx") = loRS("nReorderx")
      p_oDetail("nQtyOnHnd") = loRS("nQtyOnHnd")
      p_oDetail("xQuantity") = loRS("xQuantity")
   
   ElseIf loRS.RecordCount > 1 Then
      lsSQL = KwikBrowse(p_oAppDrivr, loRS, _
                        "sBarrCode»sDescript»nQuantity»nUnitPrce", _
                        "Bar Code»Description»Qty»Unit Price")
      If lsSQL = Empty Then GoTo endProc
      
      lasMaster = Split(lsSQL, "»")
      p_oDetail("sPartsIDx") = lasMaster(0)
      p_oDetail("sBarrCode") = lasMaster(1)
      p_oDetail("sDescript") = lasMaster(2)
      p_oDetail("nQuantity") = lasMaster(3)
      p_oDetail("nUnitPrce") = lasMaster(4)
      p_oDetail("nDiscount") = lasMaster(5)
      p_oDetail("nAddDiscx") = lasMaster(6)
      p_oDetail("nLedgerNo") = lasMaster(7)
      p_oDetail("nResvOrdr") = lasMaster(8)
      p_oDetail("nBackOrdr") = lasMaster(9)
      p_oDetail("nReorderx") = lasMaster(10)
      p_oDetail("nQtyOnHnd") = lasMaster(11)
      p_oDetail("xQuantity") = lasMaster(12)
   End If

endProc:
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)
   RaiseEvent DetailRetrieved(7)
   
   Exit Sub
endWithClear:
   p_oDetail("sPartsIDx") = ""
   p_oDetail("sBarrCode") = ""
   p_oDetail("sDescript") = ""
   p_oDetail("nQuantity") = 0
   p_oDetail("nUnitPrce") = 0
   p_oDetail("nDiscount") = 0
   p_oDetail("nAddDiscx") = 0
   p_oDetail("nLedgerNo") = 0
   p_oDetail("nResvOrdr") = 0
   p_oDetail("nBackOrdr") = 0
   p_oDetail("nReorderx") = 0
   p_oDetail("nQtyOnHnd") = 0
   p_oDetail("xQuantity") = 0
   
   GoTo endProc
errProc:
   ShowError lsProcName & "()"
End Sub

Sub DeleteDetail(ByVal Row As Integer)
   Dim lsProcName As String

   lsProcName = "DeleteDetail"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If Not p_bInitClas Then GoTo endProc
   
   If Row - 1 > p_oDetail.RecordCount Then
      GoTo endProc
   ElseIf Row = 0 Then
      p_oDetail.MoveFirst
   Else
      p_oDetail.Move Row, 0
   End If

   'Record the deletion
   If IsNull(p_oDetail("nEntryNox")) Then p_sDeletedx = p_sDeletedx & "," & p_oDetail("nEntryNox")
   
   'Delete the detail
   p_oDetail.Delete
   
   'If empty add another record
   If p_oDetail.RecordCount = 0 Then
      p_oDetail.AddNew
      InitDetail
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "()"
End Sub

Function AddDetail() As Boolean
   p_oDetail.AddNew
   InitDetail
   AddDetail = True
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Public Sub ShowReturn()
   Dim loForm As frmSPPOSReturn
   
   Set loForm = New frmSPPOSReturn
   Set loForm.AppDriver = p_oAppDrivr
   Load loForm
   Set loForm.TransObj = Me
   loForm.LoadFields
   loForm.Show 1
   
End Sub
   
