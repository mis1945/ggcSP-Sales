VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPWarranty"
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
'     Spareparts Warranty Release Object
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
'  XerSys [ 05/02/2006 04:17 pm ]
'     Start creating this object.
'  XerSys [ 05/03/2006 09:40 am ]
'     Continue creating this object.
'  XerSys [ 05/05/2006 09:40 am ]
'     I'm back... from a short vacation.
'  XerSys [ 08/10/2006 05:09 am ]
'     Warranty parts don't always come from main office. if an available parts exists in
'        branches it is automatically used as replacement.
'  Kalyptus [ 12/11/2006 10:08 am]
'     Implements the BugTracker
'  Kalyptus [ 04/02/2007 10:26 am]
'     Updated the p_oBase_Clas to conform it with the clsMasterDetail specification
'  XerSys [ 08/14/2007 11:26 am ]
'     Update this object to adopt the new implementation of SP Inventory Transaction
'  XerSys [ 09/20/2008 09:48 am ]
'     Adjust this object to fit to online warranty processing
'  Kalyptus [ 02/11/2011 11:04 am]
'     Updated p_oBaseClas_InitMaster(). Set the initial value of sSalesInv to empty("")
'  Kalyptus [ 06/25/2014 09:12 am]
'     Removed MC_SO_Master/MC_SO_Detail and replace with MC_Registration instead
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME = "clsSPWarranty"
Private Const pxeDealer As String = "DelrCode"
Private Const pxeLaborRate As String = "LaborRte"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1

Private p_cTranStat As String
Private p_bForClose As Boolean
Private p_bAllowCls As Boolean

Private p_sBranchCd As String
Private p_sFullName As String
Private p_sAddressx As String
Private p_sDelrCode As String
Private p_nLaborRte As Double
Private p_sMechName As String

Private p_nTotlLand As Currency
Private p_nOverPart As Currency
Private p_nTotlLabr As Currency
Private p_nDiscount As Single

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

Property Let WStatus(ByVal Value As xeTransactionStatus)
   p_cTranStat = Value
End Property

Property Let ForClose(ByVal Value As Boolean)
   p_bForClose = Value
End Property

Property Get ForClose() As Boolean
   ForClose = p_bForClose
End Property

Property Get AllowClose() As Boolean
   AllowClose = p_bAllowCls
End Property

Property Get DealerCode() As String
   DealerCode = p_sDelrCode
End Property

Property Get LandingCostAllowance() As Currency
   LandingCostAllowance = p_nTotlLand
End Property
Property Get TotalPartCost() As Currency
   TotalPartCost = p_nOverPart
End Property
Property Get TotalLaborCost() As Currency
   TotalLaborCost = p_nTotlLabr
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
   Case 4, 5, 6, "nquantity", "nunitprce", "nflatrate"
      p_oBaseClas.Detail(Row, Index) = Value
      Call ComputeOthers
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "xfullname"
      Master = p_sFullName
   Case 3, "xaddressx"
      Master = p_sAddressx
   Case 18, "xmechname"
      Master = p_sMechName
   Case Else
      Master = p_oBaseClas.Master(Index)
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "dtransact"
       p_oBaseClas.Master(Index) = getValidDate(p_oAppDrivr, Value)
'      If IsDate(Value) Then
'         p_oBaseClas.Master(Index) = Value
'      Else
'         p_oBaseClas.Master(Index) = p_oAppDrivr.ServerDate
'      End If
      
      p_oBaseClas.Master(13) = p_oBaseClas.Master(1)
      p_oBaseClas.Master(29) = p_oBaseClas.Master(1)
      
      RaiseEvent MasterRetrieved(1)
      RaiseEvent MasterRetrieved(13)
      RaiseEvent MasterRetrieved(29)
      
   Case 2, "sfullname"
      getMaster Value, False
   Case 3, "xaddressx"
      p_sAddressx = Value
   Case 12, 13, "dreleased", "dfilingxx"
      If IsDate(Value) Then
         p_oBaseClas.Master(Index) = Value
      Else
         p_oBaseClas.Master(Index) = "01/01/1900"
      End If
      
      RaiseEvent MasterRetrieved(Index)
   Case 14, 15, 16, 17, "nlaborrte", "ntotlfrtx", "ntotlpart", "ntrantotl"
   Case 18, "xmechname"
      getMechanic Value, False
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
   
   lsProcName = "InitTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   p_sDelrCode = p_oAppDrivr.getConfiguration(pxeDealer)
   p_nLaborRte = p_oAppDrivr.getConfiguration(pxeLaborRate)
   
   Set p_oBaseClas = New clsMasterDetail
   
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "SP_Warranty_Master"
      .DetailTable = "SP_Warranty_Detail"

      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", (CASE WHEN LENGTH(a.sClientID) > 4 THEN" & _
                              " CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNULL(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName)" & _
                           " ELSE i.sCompnyNm END) xFullName" & _
                        ", (CASE WHEN LENGTH(a.sClientID) > 4 THEN" & _
                              " CONCAT(IF(IFNull(b.sHouseNox, '') = '', '', CONCAT(b.sHouseNox, ' ')), b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode)" & _
                           " ELSE ' ' END)  xAddressx" & _
                        ", e.sEngineNo" & _
                        ", e.sFrameNox" & _
                        ", f.sModelNme" & _
                        ", j.dPurchase" & _
                        ", a.sWCRNoxxx" & _
                        ", a.sRemarksx" & _
                        ", a.sDiagnosx" & _
                        ", a.sApproved"

      .MasterQuery = .MasterQuery & _
                        ", a.dReleased" & _
                        ", a.dFilingxx" & _
                        ", a.nLaborRte" & _
                        ", a.nTotlFRTx" & _
                        ", a.nTotlPart" & _
                        ", a.nTranTotl" & _
                        ", CONCAT(h.sLastName, ', ', h.sFrstName, ' ', h.sMiddName) xMechName" & _
                        ", a.sKMReadng" & _
                        ", a.cTranStat" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", a.sClientID" & _
                        ", a.sSerialID" & _
                        ", a.sReleased" & _
                        ", a.nEntryNox" & _
                        ", a.sMechanic" & _
                        ", a.sSalesInv" & _
                        ", a.dMalFuncD"

      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                        " LEFT JOIN Client_Master b" & _
                           " LEFT JOIN TownCity c" & _
                              " LEFT JOIN Province d" & _
                                 " ON c.sProvIDxx = d.sProvIDxx" & _
                              " ON b.sTownIDxx = c.sTownIDxx" & _
                           " ON a.sClientID = b.sClientID" & _
                        " LEFT JOIN MC_Serial e" & _
                           " LEFT JOIN MC_Model f" & _
                              " ON e.sModelIDx = f.sModelIDx" & _
                           " ON a.sSerialID = e.sSerialID" & _
                        " LEFT JOIN Company i" & _
                           " ON e.sCompnyID = i.sCompnyID" & _
                        " LEFT JOIN MC_Registration j" & _
                           " ON a.sSerialID = j.sSerialID AND j.sSourceCD = 'SRg'" & _
                        " LEFT JOIN MC_Registration_Expense g" & _
                           " ON j.sTransNox = g.sReferNox AND g.cTranStat <> '3'" & _
                        " LEFT JOIN Mechanic h" & _
                           " ON a.sMechanic = h.sEmployID"
                        
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", c.nQtyOnHnd" & _
                        ", a.nQuantity" & _
                        ", a.nUnitPrce" & _
                        ", a.nFlatRate" & _
                        ", (CASE WHEN d.cTranStat = '2' THEN" & _
                              " c.nQtyOnHnd + a.nQuantity" & _
                           " ELSE c.nQtyOnHnd END) xQtyOnHnd" & _
                        ", a.dModified" & _
                        ", a.sTransNox" & _
                        ", a.sPartsIDx" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo"
      
      .DetailQuery = .DetailQuery & _
                     " FROM " & .DetailTable & " a" & _
                        ", " & .MasterTable & " d" & _
                        ", Spareparts b" & _
                        ", SP_Inventory c" & _
                     " WHERE a.sTransNox = d.sTransNox" & _
                        " AND a.sPartsIDx = b.sPartsIDx" & _
                        " AND b.sPartsIDx = c.sPartsIDx" & _
                     " ORDER BY a.nEntryNox"
                     
      If p_bForClose = False Then
         .DetailQuery = AddCondition(.DetailQuery, "c.sBranchCd = " & strParm(p_sBranchCd))
      Else
         .DetailQuery = AddCondition(.DetailQuery, "c.sBranchCd = LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ")")
      End If
      Debug.Print .DetailQuery
      
      .Reference = "a.sTransNox|a.sTransNox"
      
      pbInitTran = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sWCRNoxxx" & _
                           ", (CASE WHEN LENGTH(a.sClientID) > 4 THEN" & _
                                 " CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNULL(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName)" & _
                              " ELSE d.sCompnyNm END) xFullName" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN Client_Master b" & _
                              " ON a.sClientID = b.sClientID" & _
                           " LEFT JOIN MC_Serial c" & _
                              " LEFT JOIN Company d" & _
                                 " ON c.sCompnyID = d.sCompnyID" & _
                              " ON a.sSerialID = c.sSerialID" & _
                        " ORDER BY xFullName" & _
                           ", a.dTransact"
                           
      If p_bForClose = False Then
         .BrowseQuery(0) = AddCondition(.BrowseQuery(0), "LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = " & strParm(p_sBranchCd))
      End If
      
      If p_cTranStat <> "" Then
         .BrowseQuery(0) = AddCondition(.BrowseQuery(0), "a.cTranStat = " & p_cTranStat)
      End If
      
      .BrowseColumn(0) = "sWCRNoxxx»xFullName»dTransact"
      .BrowseCriteria(0) = "a.sWCRNoxxx»CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNULL(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName)»a.dTransact"
      .BrowseTitle(0) = "WCR No»Customer Name»Date"
      .BrowseFormat(0) = "@@-@@@@@@@@»@»MMMM DD, YYYY"
      
      ' Query statement for the branch destination field
      .BrowseQuery(2) = "SELECT" & _
                           "  a.sClientID" & _
                           ", CONCAT(a.sLastName, ', ', a.sFrstName, IF(IFNULL(a.sSuffixNm, '') = '', ' ', CONCAT(' ', a.sSuffixNm, ' ')), a.sMiddName) xFullName" & _
                           ", CONCAT(IF(IFNull(a.sHouseNox, '') = '', '', CONCAT(a.sHouseNox, ' ')), a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" & _
                           ", d.sEngineNo" & _
                           ", d.sFrameNox" & _
                           ", f.sModelNme" & _
                           ", e.dPurchase" & _
                           ", d.sSerialId" & _
                        " FROM MC_Registration e" & _
                           ", MC_Registration_Expense g" & _
                           ", Client_Master a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                           ", MC_Serial d" & _
                           ", MC_Model f" & _
                        " WHERE e.sTransNox = g.sReferNox" & _
                           " AND e.sSourceCD = 'SRg'" & _
                           " AND g.cTranStat <> '3'" & _
                           " AND e.sClientID = a.sClientID" & _
                           " AND a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND e.sSerialID = d.sSerialID" & _
                           " AND d.sModelIDx = f.sModelIDx" & _
                        " ORDER BY xFullName"
      
      .BrowseColumn(2) = "sClientID»xFullName»sEngineNo»xAddressx"
      .BrowseCriteria(2) = "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName, IF(IFNULL(a.sSuffixNm, '') = '', ' ', CONCAT(' ', a.sSuffixNm, ' ')), a.sMiddName)"
      .BrowseTitle(2) = "Code»Customer Name»Engine No»Address"
      
      .BrowseQuery(27) = "SELECT" & _
                           "  a.sCompnyID sClientID" & _
                           ", b.sCompnyNm xFullName" & _
                           ", '' xAddressx" & _
                           ", a.sEngineNo" & _
                           ", a.sFrameNox" & _
                           ", c.sModelNme" & _
                           ", '' dPurchase" & _
                           ", a.sSerialId" & _
                        " FROM MC_Serial a" & _
                           " LEFT JOIN Company b ON a.sCompnyID = b.sCompnyID" & _
                           " LEFT JOIN MC_Model c ON a.sModelIDx = c.sModelIDx" & _
                        " WHERE a.cSoldStat = '0' AND a.cLocation = '1'" & _
                        " ORDER BY xFullName"
            
      .BrowseColumn(27) = "sClientID»xFullName»sEngineNo»xAddressx"
      .BrowseCriteria(27) = "a.sCompnyID»b.sCompnyNm"
      .BrowseTitle(27) = "Code»Customer Name»Engine No»Address"
      
      .BrowseQuery(18) = "SELECT" & _
                           "  sEmployID" & _
                           ", CONCAT(sLastName, ', ', sFrstName, ' ', sMiddName) xMechName" & _
                        " FROM Mechanic" & _
                        " WHERE sBranchCd = " & strParm(p_sBranchCd) & _
                          " AND cRecdStat = " & strParm(xeRecStateActive) & _
                        " ORDER BY xMechName"
      .BrowseColumn(18) = "sEmployID»xMechName"
      .BrowseCriteria(18) = "sEmployID»CONCAT(sLastName, ', ', sFrstName, ' ', sMiddName)"
      .BrowseTitle(18) = "Mechanic ID»Mechanic Name"
      
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
   Dim lasMaster() As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String
   Dim lsSQL As String
      
   lsProcName = "SearchTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   SearchTransaction = False
   
   If pbInitTran = False Then GoTo endProc
   
   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sWCRNoxxx"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lasMaster = GetSplitedName(sSearch)
            lsValue = Trim(lasMaster(0)) & "»" & Trim(lasMaster(1))
            lsField = "b.sLastName" & "»" & "b.sFrstName"
            lsDescript = "b.sLastName" & "»" & "b.sFrstName"
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
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsSQL As String
   Dim lsProcName As String
   
   lsProcName = "PostTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then
      With p_oBaseClas
         Set loSPTrans = New clsSPInventoryTrans
         With loSPTrans
            Set .AppDriver = p_oAppDrivr
            .Branch = p_sBranchCd
            If .InitTransaction() = False Then GoTo endProc
         End With
         
         For pnCtr = 0 To .ItemCount - 1
            If .Detail(pnCtr, "sPartsIDx") = "" Then Exit For
            
   '         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
   '         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
            If .Detail(pnCtr, "nQtyOnHnd") = 0 Or .Detail(pnCtr, "nQtyOnHnd") < .Detail(pnCtr, "nQuantity") Then
               MsgBox "No Stock is Currently Availble!!!", vbCritical, "Warning"
               Exit Function
            End If
            
            loSPTrans.Detail(pnCtr, "sPartsIDx") = .Detail(pnCtr, "sPartsIDx")
            loSPTrans.Detail(pnCtr, "nQuantity") = .Detail(pnCtr, "nQuantity")
            loSPTrans.Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
            loSPTrans.Detail(pnCtr, "nResvOrdr") = .Detail(pnCtr, "nResvOrdr")
            loSPTrans.Detail(pnCtr, "nBackOrdr") = .Detail(pnCtr, "nBackOrdr")
            loSPTrans.Detail(pnCtr, "nReorderx") = .Detail(pnCtr, "nReorderx")
            loSPTrans.Detail(pnCtr, "nLedgerNo") = .Detail(pnCtr, "nLedgerNo")
         Next
         
         ' if a spareparts is not successfully updated, cancel saving.
         If loSPTrans.WarrantyRelease(.Master("sTransNox"), _
               .Master("dReleased"), _
               xeModeAddNew) = False Then
            GoTo endProc
         End If
         
         ' check if reorder level of some parts has reached
         If loSPTrans.ReOrderParts Then
            MsgBox "Some Parts Included in the Transaction has Reached it Reorder Level Limit!!!" & _
                     vbCrLf & vbCrLf & "Please Inform the Spareparts Manager on this Alert!!!", _
                     vbInformation, "Notice"
         End If
      End With
   
      PostTransaction = p_oBaseClas.PostTransaction(sTransNo)
   
   End If

endProc:
   Exit Function
   Set loSPTrans = Nothing
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
      .Detail(pnCtr, "nQtyOnHnd") = 0
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "nUnitPrce") = 0#
      .Detail(pnCtr, "nFlatRate") = 0#
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sPartsIDx") = ""
      .Detail(pnCtr, "nResvOrdr") = 0
      .Detail(pnCtr, "nBackOrdr") = 0
      .Detail(pnCtr, "nReorderx") = 0
      .Detail(pnCtr, "nLedgerno") = 0
   End With
End Function

Function SearchMaster(ByVal Index As Variant, Optional Value As Variant = "") As Boolean
   If pbInitTran = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "xfullname"
      SearchMaster = getMaster(2 & "»" & Value, True)
   Case 18, "xmechname"
      SearchMaster = getMechanic(Value, True)
   Case 27, "scompnynm"
      SearchMaster = getMaster(27 & "»" & Value, True)
   End Select
   SearchMaster = True
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

Function UpdateTransaction() As Boolean
   With p_oBaseClas
'      If .Master("cReleased") = xeYes Then Exit Function

      .EditMode = xeModeUpdate
   End With
   UpdateTransaction = True
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Function getMaster(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String
   Dim lasValue() As String
   Dim lnValue As Integer
   
   lsProcName = "getMaster(" & lsValue & ", " & lbSearch & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
      
   With p_oBaseClas
      lasValue = Split(lsValue, "»")
      lnValue = lasValue(0)
      Select Case lnValue
      Case 2
         lsValue = lasValue(1)
         If Trim(lsValue) <> "" Then
            If lsValue = p_sFullName Then GoTo endProc
            lasMaster = GetSplitedName(lsValue)
            If lbSearch Then
               lsMaster = "a.sLastName LIKE " & strParm(lasMaster(0) & "%") & _
                           " AND a.sFrstName LIKE " & strParm(lasMaster(1) & "%")
            Else
               lsMaster = "a.sLastName = " & strParm(lasMaster(0)) & _
                           " AND a.sFrstName = " & strParm(lasMaster(1))
            End If
         ElseIf lbSearch = False Then
            GoTo endWithClear
         End If
      Case 27
         lsValue = lasValue(1)
         If Trim(lsValue) <> "" Then
            If lsValue = p_sFullName Then GoTo endProc
            If lbSearch Then
               lsMaster = "b.sCompnyNm LIKE " & strParm(lsValue & "%")
            Else
               lsMaster = "b.sCompnyNm = " & strParm(lsValue & "%")
            End If
            lsMaster = lsMaster & " AND a.sBranchCd = " & strParm(p_sBranchCd)
         ElseIf lbSearch = False Then
            GoTo endWithClear
         End If
      End Select
      Debug.Print lsMaster
      lsMaster = .getMaster(lnValue, lsMaster)
      If lsMaster = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")
      
      .Master(4) = lasMaster(3)
      .Master(5) = lasMaster(4)
      .Master(6) = lasMaster(5)
      .Master(7) = IIf(IsDate(lasMaster(6)), lasMaster(6), "01/01/1900")
      .Master("sClientID") = lasMaster(0)
      .Master("sSerialID") = lasMaster(7)
      
      p_sFullName = lasMaster(1)
      p_sAddressx = lasMaster(2)
      
      RaiseEvent MasterRetrieved(2)
      RaiseEvent MasterRetrieved(3)
      RaiseEvent MasterRetrieved(4)
      RaiseEvent MasterRetrieved(5)
      RaiseEvent MasterRetrieved(6)
      RaiseEvent MasterRetrieved(7)
   End With

   getMaster = True
   
endProc:

   Exit Function
endWithClear:
   With p_oBaseClas
      .Master(4) = ""
      .Master(5) = ""
      .Master(6) = ""
      .Master(7) = "01/01/1900"
      .Master("sClientID") = ""
      .Master("sSerialID") = ""
      p_sFullName = ""
      p_sAddressx = ""
   End With
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   RaiseEvent MasterRetrieved(5)
   RaiseEvent MasterRetrieved(6)
   RaiseEvent MasterRetrieved(7)
   
   getMaster = False
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function
                     
Private Function getMechanic(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String
   
   lsProcName = "getMechanic(" & lsValue & ", " & lbSearch & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
      
   With p_oBaseClas
      If Trim(lsValue) <> "" Then
         If lsValue = p_sMechName Then GoTo endProc
         
         lasMaster = GetSplitedName(lsValue)
         If lbSearch Then
            lsMaster = "sLastName LIKE " & strParm(lasMaster(0) & "%") & _
                        " AND sFrstName LIKE " & strParm(lasMaster(1) & "%")
         Else
            lsMaster = "sLastName = " & strParm(lasMaster(0)) & _
                        " AND sFrstName = " & strParm(lasMaster(1))
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
         
      lsMaster = .getMaster(18, lsMaster)
      If lsMaster = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")
      
      .Master("sMechanic") = lasMaster(0)
      p_sMechName = lasMaster(1)
      
   End With

   getMechanic = True
   
endProc:
   RaiseEvent MasterRetrieved(18)
   Exit Function
endWithClear:
   With p_oBaseClas
      .Master("sMechanic") = ""
      p_sMechName = ""
   End With
   
   getMechanic = False
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function
                     
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
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(3)
      .Detail(lnRow, "sPartsIDx") = lasDetail(4)
      .Detail(lnRow, "nQuantity") = 0
'      .Detail(lnRow, "sWCRNoxxx") = ""
   End With
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   
endProc:

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "sPartsIDx") = ""
   End With
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lnIndex _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Sub
                     
Function LoadDetail() As Boolean
   Dim loRS As Recordset
   Dim lsProcName As String, lsSQL As String
   Dim lnCtr As Integer
   Dim lnRow As Long

   lsProcName = "LoadDetail"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   lsSQL = "SELECT" & _
               "  a.sPartsIDx" & _
               ", b.sBarrCode" & _
               ", b.sDescript" & _
               ", c.nQtyOnHnd" & _
               ", a.nQuantity" & _
               ", a.nUnitPrce" & _
               ", a.nFlatRate" & _
            " FROM SP_Warranty_Detail a" & _
               ", Spareparts b" & _
               ", SP_Inventory c" & _
            " WHERE a.sPartsIDx = b.sPartsIDx" & _
               " AND a.sPartsIDx = c.sPartsIDx" & _
               " AND a.sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
            " ORDER BY b.sBarrCode"
      
   If p_bForClose Then
      lsSQL = AddCondition(lsSQL, "c.sBranchCd = " & strParm(Left(p_oBaseClas.Master("sTransNox"), Len(p_oAppDrivr.BranchCode))))
   Else
      lsSQL = AddCondition(lsSQL, "c.sBranchCd = " & strParm(p_sBranchCd))
   End If
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Debug.Print lsSQL
   With p_oBaseClas
      If loRS.EOF Then GoTo endProc

      lnCtr = 0
      
      Do While loRS.EOF = False
         If lnCtr + 1 > .ItemCount Then .AddDetail
         
         .Detail(lnRow, "sBarrCode") = loRS("sBarrCode")
         .Detail(lnRow, "sDescript") = loRS("sDescript")
         .Detail(lnRow, "nQuantity") = loRS("nQuantity")
         .Detail(lnRow, "nQtyOnHnd") = loRS("nQtyOnHnd")
         .Detail(lnRow, "nUnitPrce") = loRS("nUnitPrce")
         .Detail(lnRow, "nFlatRate") = loRS("nFlatRate")
         
         lnCtr = lnCtr + 1
         loRS.MoveNext
      Loop
   End With
   
   LoadDetail = True
   
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"

   GoTo endProc
End Function

Private Function delSPTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsProcName As String
   
   lsProcName = "delSPTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   'Should test first if already posted
   If p_oBaseClas.Master("cTranStat") = xeStatePosted Then
   
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
         If .WarrantyRelease(p_oBaseClas.Master("sTransNox"), _
               p_oBaseClas.Master("dTransact"), _
               xeModeDelete) = False Then
            GoTo endProc
         End If
      End With
   
   End If
   
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
      .Master("dTransact") = Format(p_oAppDrivr.ServerDate, "mm/dd/yyyy")
      .Master("sEngineNo") = ""
      .Master("sFrameNox") = ""
      .Master("sModelNme") = ""
      .Master("dPurchase") = "01/01/1900"
      .Master("sWCRNoxxx") = ""
      .Master("sRemarksx") = ""
      .Master("sDiagnosx") = ""
      .Master("sApproved") = ""
      .Master("dReleased") = "01/01/1900"
      .Master("dFilingxx") = .Master("dTransact")
      .Master("nLaborRte") = p_nLaborRte
      .Master("nTotlFRTx") = 0
      .Master("nTotlPart") = 0
      .Master("nTranTotl") = 0
      .Master("sKMReadng") = ""
      .Master("sApproved") = ""
      .Master("sClientID") = ""
      .Master("sSerialID") = ""
      .Master("cTranStat") = xeStateOpen
      .Master("sReleased") = ""
      .Master("nEntryNox") = 1
      .Master("sMechanic") = ""
      .Master("dMalFuncD") = Format(p_oAppDrivr.ServerDate, "mm/dd/yyyy")
      .Master("sSalesInv") = ""
      
      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "nQuantity") = 0
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nUnitPrce") = 0#
      .Detail(0, "nFlatRate") = 0#
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sPartsIDx") = ""
      .Detail(0, "nResvOrdr") = 0
      .Detail(0, "nBackOrdr") = 0
      .Detail(0, "nReorderx") = 0
      .Detail(0, "nLedgerNo") = 0
      
   End With

   p_sAddressx = Empty
   p_sMechName = Empty
      
   p_nOverPart = 0
   p_nTotlLabr = 0
   p_nTotlLand = 0
   p_nDiscount = Val(p_oAppDrivr.getConfiguration("sSuppDsc1", p_sBranchCd))

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      p_sAddressx = .Master("xAddressx")
      p_sFullName = IFNull(.Master("xFullName"))
      p_sMechName = IFNull(.Master("xMechName"), "")
      For pnCtr = 0 To .ItemCount - 1
         .Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "xQtyOnHnd")
      Next
   End With
   
   p_sBranchCd = Left(p_oBaseClas.Master("sTransNox"), Len(p_oAppDrivr.BranchCode))
   p_oBaseClas.Branch = p_sBranchCd
   p_nDiscount = Val(p_oAppDrivr.getConfiguration("sSuppDsc1", p_sBranchCd))
   
   Call ComputeOthers
   
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = Not pbModified
End Sub

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   With p_oBaseClas
      If p_bForClose Then
         For pnCtr = 0 To .ItemCount - 1
            If .Detail(pnCtr, "nFlatRate") > 0 Then
               p_bAllowCls = True
            End If
         Next
      End If
   End With
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
   Dim lsSQL As String
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_WillSave"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = True

   pbModified = True
   With p_oBaseClas
  
      If .EditMode = xeModeAddNew Then
         .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
      End If
      Debug.Print "Item Count: " & .ItemCount
      For pnCtr = 0 To .ItemCount - 1
         Debug.Print pnCtr & "»" & .Detail(pnCtr, "sPartsIDx") & "»" & .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         If .Detail(pnCtr, "sPartsIDx") = "" Then Exit For
         
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
         If IsNull(.Detail(pnCtr, "nUnitPrce")) Then .Detail(pnCtr, "nUnitPrce") = 0#
         If IsNull(.Detail(pnCtr, "nFlatRate")) Then .Detail(pnCtr, "nFlatRate") = 0#
      Next
   
      If .isMasterModify = False And .isDetailModify = False Then pbModified = False
   
   End With
   
   Cancel = False

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub ComputeOthers()
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lnFlatRate As Single
   Dim lnUnitPrce As Currency
   Dim lnDiscount As Single
         
   Dim lnTotlLand As Currency
   Dim lnOverPart As Currency
   Dim lnTotlLabr As Currency
      
   For lnCtr = 0 To p_oBaseClas.ItemCount - 1
      lnUnitPrce = lnUnitPrce + (IFNull(p_oBaseClas.Detail(lnCtr, 5), 0) * IFNull(p_oBaseClas.Detail(lnCtr, 4), 0))
      lnFlatRate = lnFlatRate + IFNull(p_oBaseClas.Detail(lnCtr, 6), 0)
   Next
   
   If p_nDiscount > 0 Then
      p_nTotlLand = ((lnUnitPrce / (1 - p_nDiscount)) / 1.12) * 0.15
   Else
      p_nTotlLand = 0
   End If
   
   p_nOverPart = lnUnitPrce + p_nTotlLand
   p_nTotlLabr = lnFlatRate * IFNull(p_oBaseClas.Master("nLaborRte"), 0)
   
   p_oBaseClas.Master("nTotlFRTx") = lnFlatRate
   p_oBaseClas.Master("nTotlPart") = lnUnitPrce
   p_oBaseClas.Master("nTranTotl") = p_nOverPart + p_nTotlLabr
   
   RaiseEvent MasterRetrieved(15) 'Flat Rate
End Sub

