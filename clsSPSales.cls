VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPSales"
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
'     Spareparts Sales Object
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
'  XerSys [ 10/29/2004 01:44 pm ]
'     Start creating this object.
'  XerSys [ 06/12/2006 03:08 pm ]
'     Add the sales invoice and gift coupon as requirement for the branch version of this
'        object
'  XerSys [ 07/24/2006 11:56 am ]
'     Implement the constant value retrieval for special information.
'  XerSys [ 09/21/2006 01:50 pm ]
'     SP Sales must accept Customer Name
'  Kalyptus [ 12/11/2006 10:08 am]
'     Implements the BugTracker
'  Kalyptus [ 04/02/2007 10:24 am]
'     Updated the p_oBase_Clas to conform it with the clsMasterDetail specification
'  XerSys [ 08/11/2007 04:47 pm ]
'     Update this object to adopt the new implementation of SP Inventory Transaction
'  XerSys [ 11/02/2007 05:21 pm ]
'     Reactivate this object for main office
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME = "clsSPSales"
Private Const xeGiftCpn As String = "GiftCpnx"
Private Const xeWHouseID As String = "W2"

Private Type CheckInfo
   CheckNox As String
   AcctNoxx As String
   BankName As String
   CheckDte As Date
   Amountxx As Double
End Type

Private Type ReceiptInfo
   TransNox As String
   ORNoxxxx As String
   ClientID As String
   TranTotl As Double
   CashAmtx As Double
   CheckAmt As Double
   Remarksx As String
End Type

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1

Private p_axChecks(4) As CheckInfo
Private p_axReceipt As ReceiptInfo
Private p_sCashierx As String
Private p_sBranchCd As String
Private p_nGCpnAmtx As Double

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pbWHSearch As Boolean
Private pnCtr As Long

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

Property Get Cashier() As String
   Cashier = p_sCashierx
End Property

Property Let Cashier(ByVal Value As String)
   getCashier Value
End Property

Property Get Checks(ByVal Row As Integer, ByVal Index As Variant) As Variant
   If Row > 4 Then Exit Property

   Select Case Index
      Case 0, "CheckNox"
         Checks = p_axChecks(Row).CheckNox
      Case 1, "sAcctNoxx"
         Checks = p_axChecks(Row).AcctNoxx
      Case 2, "sBankName"
         Checks = p_axChecks(Row).BankName
      Case 3, "dCheckDte"
         Checks = p_axChecks(Row).CheckDte
      Case 4, "nAmountxx"
         Checks = p_axChecks(Row).Amountxx
   End Select
End Property

Property Let Checks(ByVal Row As Integer, ByVal Index As Variant, ByVal Value As Variant)
   If Row > 4 Then Exit Property

   Select Case Index
      Case 0, "CheckNox"
         p_axChecks(Row).CheckNox = Value
      Case 1, "sAcctNoxx"
         p_axChecks(Row).AcctNoxx = Value
      Case 2, "sBankName"
         p_axChecks(Row).BankName = Value
      Case 3, "dCheckDte"
         p_axChecks(Row).CheckDte = Value
      Case 4, "nAmountxx"
         p_axChecks(Row).Amountxx = Value
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
      getDetail Row, Index, Value, False
   Case 2, "sdescript"
      getDetail Row, Index, Value, False
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   Master = p_oBaseClas.Master(Index)
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property

   p_oBaseClas.Master(Index) = Value
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property

   ItemCount = p_oBaseClas.ItemCount
End Property

Property Get Receipt(ByVal Index As Variant) As Variant
   Select Case Index
      Case 0, "sTransNox"
         Receipt = p_axReceipt.TransNox
      Case 1, "sORNoxxxx"
         Receipt = p_axReceipt.ORNoxxxx
      Case 2, "sClientID"
         Receipt = p_axReceipt.ClientID
      Case 3, "nTranTotl"
         Receipt = p_axReceipt.TranTotl
      Case 4, "nCashAmtx"
         Receipt = p_axReceipt.CashAmtx
      Case 5, "nCheckAmt"
         Receipt = p_axReceipt.CheckAmt
      Case 6, "sRemarksx"
         Receipt = p_axReceipt.Remarksx
   End Select
End Property

Property Let Receipt(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
      Case 0, "sTransNox"
      Case 1, "sORNoxxxx"
         p_axReceipt.ORNoxxxx = Value
      Case 2, "sClientID"
         p_axReceipt.ClientID = Value
      Case 3, "nTranTotl"
         p_axReceipt.TranTotl = Value
      Case 4, "nCashAmtx"
         p_axReceipt.CashAmtx = Value
      Case 5, "nCheckAmt"
         p_axReceipt.CheckAmt = Value
      Case 6, "sRemarksx"
         p_axReceipt.Remarksx = Value
   End Select
End Property

Function InitTransaction() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String, lors As Recordset

   lsOldProc = "InitTransaction"
   'On Error Goto errProc
   InitTransaction = False

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc

   Set p_oBaseClas = New clsMasterDetail

   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "SP_SO_Master"
      .DetailTable = "SP_SO_Detail"

      .MasterQuery = "SELECT" & _
                        "  sTransNox" & _
                        ", dTransact" & _
                        ", sSalesInv" & _
                        ", nTranTotl" & _
                        ", nAmtPaidx" & _
                        ", nGiftCpnx" & _
                        ", sCashierx" & _
                        ", sAutoTran" & _
                        ", cTranStat" & _
                        ", sModified" & _
                        ", dModified" & _
                        ", sSalesInv" & _
                        ", sReferNox" & _
                        ", sClientID" & _
                     " FROM " & .MasterTable

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
                        ", a.sPartsIDx" & _
                        ", c.nQtyOnHnd + a.nQuantity xQtyOnHnd" & _
                        ", d.cPartType" & _
                        ", c.nSMaxDisc" & _
                        ", c.nMMaxDisc" & _
                        ", c.nSelPrice" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo"
      .DetailQuery = .DetailQuery & _
                     " FROM " & .DetailTable & " a" & _
                        ", Spareparts b" & _
                        ", SP_Inventory c" & _
                        ", SP_Type d" & _
                     " WHERE a.sPartsIDx = b.sPartsIDx" & _
                        " AND a.sPartsIDx = c.sPartsIDx" & _
                        " AND b.cPartType = d.cPartType" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                     " ORDER BY a.nEntryNox"
      .Reference = "sTransNox|a.sTransNox"

      InitTransaction = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                           "  sTransNox" & _
                           ", dTransact" & _
                           ", nTranTotl" & _
                        " FROM " & .MasterTable & _
                        " WHERE LEFT(sTransNox, 2) = " & strParm(p_sBranchCd) & _
                        " ORDER BY dTransact"
      .BrowseColumn(0) = "sTransNox»dTransact»nTranTotl"
      .BrowseTitle(0) = "Transaction No»Date»Total"
      .BrowseFormat(0) = "@@@@-@@@@@@»MMMM DD, YYYY»#,##0.00"

      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sBarrCode" & _
                                 ", a.sDescript" & _
                                 ", b.nQtyOnHnd" & _
                                 ", b.nSelPrice" & _
                                 ", c.sModelNme" & _
                                 ", a.sPartsIDx" & _
                                 ", a.cPartType" & _
                                 ", b.nSMaxDisc" & _
                                 ", b.nMMaxDisc" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nReorderx" & _
                                 ", b.nLedgerNo" & _
                              " FROM Spareparts a" & _
                                    " LEFT JOIN SP_Model c" & _
                                       " ON a.sModelIDx = c.sModelIDx" & _
                                 ", SP_Inventory b" & _
                              " WHERE a.sPartsIDx = b.sPartsIDx" & _
                                 " AND b.sBranchCd = " & strParm(p_sBranchCd)
      .BrowseDetailColumn(0) = "sBarrCode»sDescript»nQtyOnHnd»nSelPrice»sModelNme"
      .BrowseDetailTitle(0) = "Bar Code»Description»QOH»Price»Model"
      .BrowseDetailFormat(0) = "@»@»#0»#,##0.00»@"
      .BrowseDetailCriteria(0) = "a.sBarrCode»a.sDescript"
      
      ' retrieve the value of Gift Coupon
      lsSQL = "SELECT sValuexxx FROM xxxOtherConfig" & _
               " WHERE sProdctID = " & strParm(p_oAppDrivr.ProductID) & _
                  " AND sConfigID = " & strParm(xeGiftCpn)

      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

      If lors.EOF = False Then
         p_nGCpnAmtx = CDbl(lors("sValuexxx"))
      Else
         p_nGCpnAmtx = 50
      End If
   End With

   pbInitTran = InitTransaction

endProc:
   Set lors = Nothing

   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "NewTransaction"
   'On Error Goto errProc
   
   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "SaveTransaction"
   'On Error Goto errProc
   
   If pbInitTran Then SaveTransaction = p_oBaseClas.SaveTransaction

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "OpenTransaction"
   'On Error Goto errProc
   
   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)

endProc:

   Exit Function
errProc:
    ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "DeleteTransaction"
   'On Error Goto errProc
   
   If pbInitTran Then DeleteTransaction = p_oBaseClas.DeleteTransaction

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsOldProc As String
   Dim lsRecord As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String

   lsOldProc = "SearchTransaction"
   'On Error Goto errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "sTransNox"
      lsDescript = ""
   End If

   SearchTransaction = p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript)

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "CancelTransaction"
   'On Error Goto errProc
   
   If pbInitTran Then CancelTransaction = p_oBaseClas.CancelTransaction()

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
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
      .Detail(pnCtr, "nMMaxDisc") = 0
      .Detail(pnCtr, "nSMaxDisc") = 0
      .Detail(pnCtr, "nSelPrice") = 0#
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
   Dim lsOldProc As String

   lsOldProc = "SearchDetail"
   'On Error Goto errProc
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
   ShowError lsOldProc & "( " & Row _
                       & ", " & Index _
                       & ", " & Value & " )"
End Function

Function SearchFromWarehouse(ByVal Row As Long, _
                              ByVal Index As Variant, _
                              Optional Value As Variant = "") As Boolean
   Dim lsOldQuery As String

   With p_oBaseClas
      lsOldQuery = .BrowseDetailQuery(0)

      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sBarrCode" & _
                                 ", a.sDescript" & _
                                 ", b.nQtyOnHnd" & _
                                 ", b.nSelPrice" & _
                                 ", c.sModelNme" & _
                                 ", a.sPartsIDx" & _
                                 ", a.cPartType" & _
                                 ", b.nSMaxDisc" & _
                                 ", b.nMMaxDisc" & _
                              " FROM Spareparts a" & _
                                    " LEFT JOIN SP_Model c" & _
                                       " ON a.sModelIDx = c.sModelIDx" & _
                                 ", SP_Inventory b" & _
                              " WHERE a.sPartsIDx = b.sPartsIDx" & _
                                 " AND b.sBranchCd = " & strParm(xeWHouseID) & _
                              " ORDER BY a.sDescript"
      SearchFromWarehouse = getDetail(Row, Index, .Detail(Row, Index), True)
      .BrowseDetailQuery(0) = lsOldQuery
      If pbWHSearch = False Then pbWHSearch = SearchFromWarehouse
   End With
End Function

Function UpdateTransaction() As Boolean
   p_oBaseClas.EditMode = xeModeUpdate
   UpdateTransaction = True
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Private Sub ShowError(ByVal lsOldProc As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, "clsSPSales", lsOldProc, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Function getCashier(ByVal lsCashier As String) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "GetCashier"
   'On Error Goto errProc
   
   p_sCashierx = ""
   lsSQL = "SELECT sUserIDxx, sUserName FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & strParm(lsCashier)
   
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   If lors.EOF Then GoTo endProc
   
   p_sCashierx = Decrypt(lors("sUserName"), p_oAppDrivr.Machinex)
   p_oBaseClas.Master("sCashierx") = lors("sUserIDxx")
   
   getCashier = True

endProc:
   Set lors = Nothing

   Exit Function
errProc:
    ShowError lsOldProc & "( " & lsCashier & " )"
End Function

Private Function getDetail(ByVal lnRow As Integer, _
                        ByVal lnIndex As Integer, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsOldProc As String

   lsOldProc = "GetDetail"
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
      .Detail(lnRow, "nDiscount") = 0#
      .Detail(lnRow, "nAddDiscx") = 0#
      .Detail(lnRow, "sPartsIDx") = lasDetail(5)
      .Detail(lnRow, "cPartType") = lasDetail(6)
      .Detail(lnRow, "nSMaxDisc") = lasDetail(7)
      .Detail(lnRow, "nMMaxDisc") = lasDetail(8)
      .Detail(lnRow, "nResvOrdr") = lasDetail(9)
      .Detail(lnRow, "nBakcOrdr") = lasDetail(10)
      .Detail(lnRow, "nReorderx") = lasDetail(11)
      .Detail(lnRow, "nLedgerNo") = lasDetail(12)
      .Detail(lnRow, "nSelPrice") = lasDetail(3)
   End With
   getDetail = True
   
endProc:
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)

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
      .Detail(lnRow, "nMMaxDisc") = 0#
      .Detail(lnRow, "nSMaxDisc") = 0#
      .Detail(lnRow, "nSelPrice") = 0#
      .Detail(lnRow, "nResvOrdr") = 0
      .Detail(lnRow, "nBakcOrdr") = 0
      .Detail(lnRow, "nReorderx") = 0
      .Detail(lnRow, "nLedgerNo") = 0
   End With
   getDetail = False
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & lnRow _
                        & ", " & lnIndex _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function UndoTransaction() As Boolean
   Dim loSPTrans As clsSPTrans
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Long

   lsOldProc = "UndoTransaction"
   'On Error Goto errProc
   UndoTransaction = False

   With p_oBaseClas
      Set loSPTrans = New clsSPTrans
      Set loSPTrans.AppDriver = p_oAppDrivr
      loSPTrans.InitTransaction
      loSPTrans.Branch = p_sBranchCd
      loSPTrans.TransactDate = .MasterOriginal("dTransact")

      lsSQL = AddCondition(.DetailQuery, _
                  "a.sTransNox = " & strParm(.Master("sTransNox")))

      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

      If lors.EOF Then GoTo endProc
      Do
         loSPTrans.PartsID = lors("sPartsIDx")
         loSPTrans.Quantity = lors("nQuantity")
         loSPTrans.EntryNo = lors("nEntryNox")

         ' if a spareparts inventory is not successfully updated, cancel saving.
         If loSPTrans.Sales(.Master("sTransNox"), True) = False Then GoTo endProc
         lors.MoveNext
      Loop Until lors.EOF

      If .MasterOriginal("nAmtPaidx") > 0# Then
         lsSQL = "DELETE FROM Receipt_Master" & _
                  " WHERE sReferNox = " & strParm(.Master("sTransNox"))
         p_oAppDrivr.RegisDelete lsSQL
         p_oAppDrivr.Connection.Execute lsSQL, lnRow, adCmdText

         If lnRow <= 0 Then
            MsgBox "Unable to Delete Receipt Transaction Info!!!", vbCritical, "Warning"
            GoTo endProc
         End If

         lsSQL = "DELETE FROM Checks_Received" & _
                  " WHERE sReferNox = " & strParm(p_axReceipt.TransNox)
         p_oAppDrivr.RegisDelete lsSQL
         p_oAppDrivr.Connection.Execute lsSQL, lnRow, adCmdText
      End If
      
      If .MasterOriginal("sAutoTrans") <> Empty Then
         If DelAutoTransfer = False Then GoTo endProc
      End If
   End With

   UndoTransaction = True

endProc:
   Set loSPTrans = Nothing
   Set lors = Nothing

   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Private Function SaveSalesInvoice() As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim lnRow As Long
   Dim lsRefNo As String
   Dim lnRefNo As Long

   lsOldProc = "SaveSalesInvoice"
   'On Error Goto errProc
   SaveSalesInvoice = False

   With p_oBaseClas
      If .EditMode = xeModeAddNew Then
         If Trim(.Master("sSalesInv")) = Empty Then
            SaveSalesInvoice = True
            GoTo endProc
         End If

         lsSQL = "INSERT INTO SalesInvoice " & _
                     "( sTransNox" & _
                     ", sBranchCd" & _
                     ", dTransact" & _
                     ", sSalesInv" & _
                     ", sSourceCd" & _
                     ", nTranTotl" & _
                     ", sModified" & _
                     ", dModified" & _
                  " ) VALUES ( " & _
                     strParm(.Master("sTransNox")) & _
                     ", " & strParm(Left(.Master("sTransNox"), 2)) & _
                     ", " & dateParm(.Master("dTransact")) & _
                     ", " & strParm(.Master("sSalesInv")) & _
                     ", " & strParm("SPSl") & _
                     ", " & .Master("nTranTotl") & _
                     ", " & strParm(Encrypt(.Master("sModified"))) & _
                     ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
      Else
         If Trim(.MasterOriginal("sSalesInv")) = Empty And _
               Trim(.Master("sSalesInv")) = Empty Then
            SaveSalesInvoice = True
            GoTo endProc
         End If

         lsSQL = "UPDATE SalesInvoice SET" & _
                     "  dTransact = " & dateParm(.Master("dTransact")) & _
                     ", sSalesInv = " & strParm(.Master("sSalesInv")) & _
                     ", nTranTotl = " & .Master("nTranTotl") & _
                     ", sModified = " & strParm(.Master("sModified")) & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                  " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                     " AND sSourceCd = " & strParm("SPSl")
      End If
   
      lnRow = p_oAppDrivr.Execute(lsSQL, "SalesInvoice")
      If lnRow = 0 Then
         MsgBox "Unable to Update Transaction Payment!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   SaveSalesInvoice = True

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Private Function SaveAutoTransfer() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsTransNox As String
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lnCtr As Integer
   Dim lbTransfer As Boolean
   
   SaveAutoTransfer = False
   lsOldProc = "SaveAutoTransfer"
   'On Error Goto errProc
   
   ' if user search from warehouse, check if an item was solicited from warehouse
   With p_oBaseClas
      .Master("sAutoTran") = ""
      If pbWHSearch Then
         Set loSPTrans = New clsSPTrans
         
         With loSPTrans
            Set .AppDriver = p_oAppDrivr
            .Branch = p_sBranchCd
            If .InitTransaction() = False Then GoTo endProc
         End With
         
         lsTransNox = GetNextCode("SP_AutoTransfer_Master", "sTransNox", True, _
                        p_oAppDrivr.Connection, True, xeWHouseID)
         lnCtr = 1
         For pnCtr = 0 To .ItemCount - 1
            lsSQL = "SELECT" & _
                        "  sPartsIDx" & _
                        ", nQtyOnHnd" & _
                        ", nResvOrdr" & _
                        ", nBackOrdr" & _
                        ", nReorderx" & _
                        ", nLedgerNo" & _
                     " FROM SP_Inventory" & _
                     " WHERE sPartsIDx = " & strParm(.Detail(pnCtr, "sPartsIDx")) & _
                        " AND sBranchCd = " & strParm(p_sBranchCd)
            
            lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
            If lors("nQtyOnHnd") < .Detail(pnCtr, "nQuantity") Then
               loSPTrans.PartsID = .Detail(pnCtr, "sPartsIDx")
               loSPTrans.Quantity = .Detail(pnCtr, "nQuantity")
               loSPTrans.EntryNo = pnCtr + 1
               
               If loSPTrans.AcceptAutoDelivery(lsTransNox, xeWHouseID) = False Then GoTo endProc
               
               lsSQL = "INSERT INTO SP_AutoTransfer_Detail " & _
                           "( sTransNox" & _
                           ", nEntryNox" & _
                           ", sPartsIDx" & _
                           ", nQuantity" & _
                           ", dModified" & _
                        " ) VALUES ( " & _
                           strParm(lsTransNox) & _
                           ", " & lnCtr & _
                           ", " & strParm(.Detail(pnCtr, "sPartsIDx")) & _
                           ", " & (.Detail(pnCtr, "nQuantity") - lors("nQtyOnHnd")) & _
                           ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
                           
               p_oAppDrivr.Connection.Execute lsSQL, lnRow, adCmdText
               If lnRow <= 0 Then
                  MsgBox "Unable to Create Spareparts Auto-Transfer Detail Info!!!", vbCritical, "Warning"
                  GoTo endProc
               End If
               lnCtr = lnCtr + 1
            End If
         Next
         
         If lnCtr > 1 Then
            lsSQL = "INSERT INTO SP_AutoTransfer_Master " & _
                        "( sTransNox" & _
                        ", dTransact" & _
                        ", sDestinat" & _
                        ", cTranStat" & _
                        ", sModified" & _
                        ", dModified" & _
                     " ) VALUES ( " & _
                        strParm(lsTransNox) & _
                        ", " & dateParm(.Master("dTransact")) & _
                        ", " & strParm(p_sBranchCd) & _
                        ", " & strParm(xeStateClosed) & _
                        ", " & strParm(Encrypt(p_oAppDrivr.UserID)) & _
                        ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
                        
            p_oAppDrivr.Connection.Execute lsSQL, lnRow, adCmdText
            If lnRow <= 0 Then
               MsgBox "Unable to Create Spareparts Auto-Transfer Detail Info!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
         
         .Master("sAutoTran") = lsTransNox
      End If
   End With
   
   SaveAutoTransfer = True
   
endProc:
   Set loSPTrans = Nothing
   Set lors = Nothing

   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Private Function DelAutoTransfer() As Boolean
   Dim loSPTrans As clsSPTrans
   Dim lors As Recordset
   Dim lsTransNox As String
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Long

   lsOldProc = "DelAutoTransfer"
   'On Error Goto errProc
   
   DelAutoTransfer = False
   
   Set lors = New Recordset
   lsSQL = "SELECT cTranStat FROM SP_AutoTranfer_Master" & _
            " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sAutoTran"))
      
   With p_oAppDrivr
      lors.Open lsSQL, .Connection, , , adCmdText
      
      If lors.EOF Then
         MsgBox "Unable to Retrieve Spareparts Auto-Transfer Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
         
      If lors("cTranStat") = xeStatePosted Then
         MsgBox "Auto-Transfer was Already Posted!!!" & vbCrLf & _
                  "Modification of its Master Transaction is not Allowed!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsSQL = "DELETE FROM SP_AutoTransfer_Master" & _
               " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sAutoTran"))
               
      ' delete the master info
      If .RegisDelete(lsSQL) = False Then GoTo endProc
      .Connection.Execute lsSQL, lnRow, adCmdText
      
      If lnRow <= 0 Then
         MsgBox "Unable to Undo Spareparts Auto-Transfer Master Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsSQL = "SELECT * FROM SP_AutoTransfer_Detail" & _
               " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sAutoTran"))
      lors.Close
      lors.Open lsSQL, .Connection, , , adCmdText
      
      Set loSPTrans = New clsSPTrans
      Set loSPTrans.AppDriver = p_oAppDrivr
      loSPTrans.Branch = p_sBranchCd
      loSPTrans.TransactDate = p_oBaseClas.MasterOriginal("dTransact")
      Do While Not lors.EOF
         loSPTrans.PartsID = lors("sPartsIDx")
         loSPTrans.Quantity = lors("nQuantity")
         loSPTrans.EntryNo = lors("nEntryNox")
         
         If loSPTrans.AcceptAutoDelivery(lsTransNox, Left(lsTransNox, 2), True) = False Then GoTo endProc
      
         lors.MoveNext
      Loop
      
      lsSQL = "DELETE FROM SP_AutoTransfer_Detail" & _
               " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sAutoTran"))
               
      ' then delete the detail info
      If .RegisDelete(lsSQL) = False Then GoTo endProc
      .Connection.Execute lsSQL, lnRow, adCmdText
      
      If lnRow <= 0 Then
         MsgBox "Unable to Undo Spareparts Auto-Transfer Detail Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   DelAutoTransfer = True
   
endProc:
   Set loSPTrans = Nothing
   Set lors = Nothing

   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Private Function isTotalDiscOK() As Boolean
   Dim lnItemPrice As Double
   Dim lnDiscPrice As Double
   Dim lnDiscount As Double
   Dim lnMDiscount As Double
   Dim lnSDiscount As Double
   Dim lnUDiscount As Double
   Dim lsAppvID As String
   Dim lsAppvName As String
   Dim lnAppvRights As Integer
   Dim lbGetApproval As Boolean
                                                                                                          
   isTotalDiscOK = False
                                                                                                          
   lnAppvRights = p_oAppDrivr.UserLevel
   lsAppvID = p_oAppDrivr.UserID
   
   With p_oBaseClas
      lnSDiscount = 0
      lnMDiscount = 0
      lnUDiscount = 0
      lnDiscount = 0
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sPartsIDx") = "" Then Exit For
         
         lnItemPrice = .Detail(pnCtr, "nSelPrice") * .Detail(pnCtr, "nQuantity")
         lnUDiscount = lnUDiscount + _
                  Round(lnItemPrice * 5 / 100, 2)
         lnMDiscount = lnMDiscount + _
                  Round(lnItemPrice * .Detail(pnCtr, "nMMaxDisc") / 100, 2)
         lnSDiscount = lnSDiscount + _
                  Round(lnItemPrice * .Detail(pnCtr, "nSMaxDisc") / 100, 2)
         lnDiscPrice = .Detail(pnCtr, "nUnitPrce") * .Detail(pnCtr, "nQuantity")
         lnDiscPrice = (lnItemPrice - lnDiscPrice) + Round(lnDiscPrice * .Detail(pnCtr, "nDiscount") / 100, 2)
         lnDiscount = lnDiscount + lnDiscPrice
      Next
      
      If lnDiscount = 0# Then
         isTotalDiscOK = True
         GoTo endProc
      End If
      
      lbGetApproval = False
      If p_oAppDrivr.UserLevel = xeEncoder And lnUDiscount < lnDiscount Then
         lbGetApproval = True
      ElseIf p_oAppDrivr.UserLevel = xeSupervisor And lnSDiscount < lnDiscount Then
         lbGetApproval = True
      ElseIf p_oAppDrivr.UserLevel = xeManager And lnMDiscount < lnDiscount Then
         lbGetApproval = True
      End If
      
      If lbGetApproval Then
         If GetApproval(p_oAppDrivr, lnAppvRights, lsAppvID, lsAppvName, p_oAppDrivr.MenuName) = False Then GoTo endProc
         
         If lnAppvRights = xeSupervisor Then
            If lnSDiscount < lnDiscount Then
               MsgBox "Discount is not within the range of the approving officer!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         ElseIf lnAppvRights = xeManager Then
            If lnMDiscount < lnDiscount Then
               MsgBox "Discount is not within the range of the approving officer!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         ElseIf Not (lnAppvRights = xeSysOwner Or lnAppvRights = xeEngineer) Then
            MsgBox "User is not allowed to approved discount!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End With
   isTotalDiscOK = True

endProc:
   Exit Function
End Function

Private Sub Class_Terminate()
   Set p_oBaseClas = Nothing
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lsOldProc As String
   
   lsOldProc = "p_oBaseClas_InitMaster"
   'On Error Goto errProc
   
   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = Format(p_oAppDrivr.ServerDate, "MM/DD/YYYY")
      .Master("nGiftCpnx") = 0
      .Master("nTranTotl") = 0#
      .Master("nAmtPaidx") = 0#
      .Master("cTranStat") = xeStateClosed
      .Master("sCashierx") = ""
      .Master("sAutoTran") = ""
      .Master("sSalesInv") = ""
      .Master("sReferNox") = ""
      .Master("sClientID") = ""
      p_sCashierx = ""

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
      .Detail(0, "nSMaxDisc") = 1
      .Detail(0, "nMMaxDisc") = 1
      .Detail(0, "nSelPrice") = 0#
   End With

   With p_axReceipt
      .TransNox = GetNextCode("Receipt_Master", "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .ClientID = ""
      .TranTotl = 0#
      .CashAmtx = 0#
      .CheckAmt = 0#
      .Remarksx = ""
   End With

   For pnCtr = 0 To UBound(p_axChecks)
      p_axChecks(pnCtr).AcctNoxx = ""
      p_axChecks(pnCtr).BankName = ""
      p_axChecks(pnCtr).CheckNox = ""
      p_axChecks(pnCtr).CheckDte = Date
      p_axChecks(pnCtr).Amountxx = 0
   Next

endProc:
   Exit Sub
errProc:
    ShowError lsOldProc & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      p_sCashierx = p_oAppDrivr.getUserName(.Master("sCashierx"))
      
      For pnCtr = 0 To .ItemCount - 1
         .Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "xQtyOnHnd")
      Next
   End With
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_WillCancel(Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "p_oBaseClas_WillCancel"
   'On Error Goto errProc
   
   Cancel = Not UndoTransaction
   
endProc:
   Exit Sub
errProc:
    ShowError lsOldProc & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillDelete(Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "p_oBaseClas_WillDelete"
   'On Error Goto errProc

   Cancel = Not UndoTransaction

endProc:
   Exit Sub
errProc:
    ShowError lsOldProc & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim loSPTrans As clsSPTrans
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "p_oBaseClass_WillSave"
   'On Error Goto errProc
   Cancel = True

   pbModified = True
   With p_oBaseClas
      If .isMasterModify = False And .isDetailModify = False Then
         pbModified = False
         Cancel = False
         GoTo endProc
      End If

      If p_sCashierx = Empty Then
         MsgBox "Invalid Cashier Detected!!!" & vbCrLf & _
                  "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If

      If .EditMode = xeModeUpdate Then
         If UndoTransaction = False Then GoTo endProc
      End If

      If isTotalDiscOK = False Then GoTo endProc
      
      If .Master("nTranTotl") > .Master("nAmtPaidx") + (p_nGCpnAmtx * .Master("nGiftCpnx")) Then
         MsgBox "Invalid Amount Paid Detected!!!" & vbCrLf & _
                  "Verify Your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      Set loSPTrans = New clsSPTrans
      Set loSPTrans.AppDriver = p_oAppDrivr
      loSPTrans.Branch = p_oAppDrivr.BranchCode
      loSPTrans.TransactDate = .Master("dTransact")

'      p_sCashierx = LCase(p_sCashierx)
'      lsSQL = "SELECT sUserIDxx, sLogNamex FROM xxxSysUser" & _
'               " WHERE sLogNamex = " & strParm(Encrypt(p_sCashierx, p_oAppDrivr.Machinex)) & _
'                  " AND sProdctID = " & strParm(p_oAppDrivr.ProductID)
'
'      Set lors = New Recordset
'      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
'
'      If lors.EOF Then
'         MsgBox "Invalid Cashier Detected!!!" & vbCrLf & _
'                  "Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         GoTo endProc
'      End If
'
'      .Master("sCashierx") = lors("sUserIDxx")
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sPartsIDx") = Empty Then Exit For

         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1

         loSPTrans.PartsID = .Detail(pnCtr, "sPartsIDx")
         loSPTrans.Quantity = .Detail(pnCtr, "nQuantity")
         loSPTrans.EntryNo = pnCtr + 1

         ' if a spareparts inventory is not successfully updated, cancel saving.
         If loSPTrans.Sales(.Master("sTransNox")) = False Then
            MsgBox "Unable to Update Sales Transaction!", vbCritical, "Warning"
            GoTo endProc
         End If
      Next

      If SaveAutoTransfer = False Then GoTo endProc
      
      If .Master("nAmtPaidx") > 0 Then
         Cancel = Not SaveSalesInvoice
         If Cancel Then GoTo endProc
      End If
   End With

   Cancel = False
   
endProc:
   Set loSPTrans = Nothing
   Set lors = Nothing

   Exit Sub
errProc:
    ShowError lsOldProc & "( " & Cancel & " )"
End Sub
