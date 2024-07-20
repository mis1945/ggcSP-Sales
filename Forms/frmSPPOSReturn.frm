VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmSPPOSReturn 
   BorderStyle     =   0  'None
   Caption         =   "SP POS Sales Return"
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5415
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   9551
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   7635
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   570
         Width           =   2235
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1545
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         Width           =   1920
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   5
         Left            =   7455
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "ht0;ft0"
         Top             =   4665
         Width           =   2430
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1545
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   585
         Width           =   5220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   465
         Index           =   4
         Left            =   1545
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   885
         Width           =   5220
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   3210
         Left            =   90
         TabIndex        =   5
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   1395
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   5662
         AllowBigSelection=   -1  'True
         AutoAdd         =   -1  'True
         AutoNumber      =   -1  'True
         BACKCOLOR       =   -2147483643
         BACKCOLORBKG    =   8421504
         BACKCOLORFIXED  =   -2147483633
         BACKCOLORSEL    =   -2147483635
         BORDERSTYLE     =   1
         COLS            =   2
         FILLSTYLE       =   0
         FIXEDCOLS       =   1
         FIXEDROWS       =   1
         FOCUSRECT       =   1
         EDITORBACKCOLOR =   -2147483643
         EDITORFORECOLOR =   -2147483640
         FORECOLOR       =   -2147483640
         FORECOLORFIXED  =   -2147483630
         FORECOLORSEL    =   -2147483634
         FORMATSTRING    =   ""
         Object.HEIGHT          =   3210
         GRIDCOLOR       =   12632256
         GRIDCOLORFIXED  =   0
         BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GRIDLINES       =   1
         GRIDLINESFIXED  =   2
         GRIDLINEWIDTH   =   1
         MOUSEICON       =   "frmSPPOSReturn.frx":0000
         MOUSEPOINTER    =   0
         REDRAW          =   -1  'True
         RIGHTTOLEFT     =   0   'False
         ROWS            =   2
         SCROLLBARS      =   3
         SCROLLTRACK     =   0   'False
         SELECTIONMODE   =   0
         Object.TOOLTIPTEXT     =   ""
         WORDWRAP        =   0   'False
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   1
         Left            =   7230
         TabIndex        =   10
         Top             =   615
         Width           =   375
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   195
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   6705
         TabIndex        =   8
         Top             =   4710
         Width           =   720
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1605
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Address"
         Height          =   195
         Index           =   11
         Left            =   330
         TabIndex        =   7
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   285
         Index           =   3
         Left            =   330
         TabIndex        =   6
         Top             =   615
         Width           =   1125
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10395
      TabIndex        =   11
      Top             =   1845
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Del. Row"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSPPOSReturn.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10395
      TabIndex        =   12
      Top             =   1215
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Searc&h"
      AccessKey       =   "h"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSPPOSReturn.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10395
      TabIndex        =   13
      Top             =   600
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSPPOSReturn.frx":0F10
   End
End
Attribute VB_Name = "frmSPPOSReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmAdvancePayment"

Private p_oSkin As clsFormSkin
Private p_oAppDriver As clsAppDriver
Private WithEvents oTrans As clsSPSalesReturn
Attribute oTrans.VB_VarHelpID = -1
Private pnCtr As Integer
Private pbSearch As Boolean

Property Set TransObj(foObj As clsSPSalesReturn)
   Set oTrans = foObj

   InitForm

End Property

Property Set AppDriver(foObj As clsAppDriver)
   Set p_oAppDriver = foObj
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   With GridEditor1
      Select Case Index
      Case 0
         Unload Me
      Case 1 'Search
         Select Case .Col
         Case 1, 2, 3
            If oTrans.SearchDetail(.Row - 1, .Col) Then .Col = 5
            .SetFocus
            .Refresh
         End Select
      Case 2 'Delete Row
         If .Rows <> 2 Then
            oTrans.DeleteDetail (.Row - 1)
            .DeleteRow
         End If
   
         If .Rows > 13 Then
            .ColWidth(3) = 2800
            .ColWidth(9) = 1000
         Else
            .ColWidth(3) = 2900
            .ColWidth(9) = 1200
         End If
         ComputePOSTotal
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   ''On Error GoTo errProc


   Set p_oSkin = New clsFormSkin
   With p_oSkin
      Set .AppDriver = p_oAppDriver
      Set .Form = Me
      .DisableClose = True
      .ApplySkin xeFormTransDetail
   End With
   InitForm

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub InitForm()
   Dim lnCtr As Integer

   With GridEditor1
      .Cols = 10
      .Rows = 2
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "SI #"
      .TextMatrix(0, 2) = "Barcode"
      .TextMatrix(0, 3) = "Description"
      .TextMatrix(0, 4) = "QOH"
      .TextMatrix(0, 5) = "Qty."
      .TextMatrix(0, 6) = "Unit Price"
      .TextMatrix(0, 7) = "Disc."
      .TextMatrix(0, 8) = "Add. Disc."
      .TextMatrix(0, 9) = "Sub Total"
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'column width
      .ColWidth(0) = 350
      .ColWidth(1) = 570           'Starts here
      .ColWidth(2) = 1900
      .ColWidth(4) = 0
      .ColWidth(6) = 570
      .ColWidth(7) = 1000
      .ColWidth(7) = 650
      .ColWidth(8) = 800

      .ColNumberOnly(4) = True
      .ColNumberOnly(5) = True
      .ColNumberOnly(6) = True
      .ColNumberOnly(8) = True
      .ColNumberOnly(9) = True

      .ColEnabled(4) = False
      .ColEnabled(5) = True
      .ColEnabled(6) = False
      .ColEnabled(7) = False
      .ColEnabled(8) = False
      .ColEnabled(9) = False

      .ColMaxValue(7) = "99"

      .ColDefault(4) = 0
      .ColDefault(5) = 0
      .ColDefault(6) = 0
      .ColDefault(7) = 0 & "%"
      .ColDefault(8) = 0
      .ColDefault(9) = 0

      .ColFormat(6) = "#,##0.00"
      .ColFormat(8) = "#,##0.00"
      .ColFormat(9) = "#,##0.00"

      .ColAlignment(3) = 1
      .ColAlignment(4) = 1

      .WordWrap = True
      
      .EditorBackColor = p_oAppDriver.getColor("HT1")
   End With
   
End Sub

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDriver
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Sub LoadFields()
   Dim lnCtr As Integer
   Dim lnSubTotal As Currency
   Dim lnTotal As Currency
   Dim lsOldProc As String
   
   lsOldProc = "LoadFields"
   ''On Error GoTo errProc
      
   txtField(0).Text = Format(oTrans.TransNo, "@@@@-@@@@@@")
   txtField(1).Text = Format(oTrans.TransactDate, "MMMM DD, YYYY")
   txtField(3).Text = oTrans.ClientNm
   txtField(4).Text = oTrans.Address
   
   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)

      For pnCtr = 0 To oTrans.ItemCount - 1
         For lnCtr = 1 To .Cols - 2
            .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, lnCtr)
         Next
         .TextMatrix(pnCtr + 1, 7) = oTrans.Detail(pnCtr, 7) & "%"
         lnSubTotal = (oTrans.Detail(pnCtr, 5) * oTrans.Detail(pnCtr, 6))
         lnTotal = lnTotal + lnSubTotal
         .TextMatrix(pnCtr + 1, 9) = lnSubTotal - (lnSubTotal * _
                                   (oTrans.Detail(pnCtr, 7) / 100) - oTrans.Detail(pnCtr, 8))
      Next
      
      oTrans.Total = lnTotal
      txtField(5).Text = Format(oTrans.Total, "#,##0.00")
      
      If .Rows > 13 Then
         .ColWidth(3) = 2800
         .ColWidth(9) = 1000
      Else
         .ColWidth(3) = 2900
         .ColWidth(9) = 1200
      End If
      
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 3) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 5) = 0 Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 9) = 0 Then
         Cancel = True
      End If
      
      If Not Cancel Then
         oTrans.AddDetail
'         .Rows = 3
         .Row = .Rows
      End If

      If .Rows > 13 Then
         .ColWidth(3) = 2800
         .ColWidth(9) = 1000
      End If
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lnPercent As Integer
   Dim lnDiscount As Variant
   Dim lnRep As Integer
   Dim lnSubTotal As Currency
   Dim lnTotal As Currency
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_EditorValidate"
   ''On Error GoTo errProc
   
   With GridEditor1
      If .Col = 1 Or .Col = 2 Or .Col = 3 Then
         If pbSearch = False Then oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
      Else
         If .Col = 5 Then
            If CDbl(.TextMatrix(.Row, 5)) > 0 Then
               oTrans.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
               
               lnSubTotal = .TextMatrix(.Row, 5) * .TextMatrix(.Row, 6)
               .TextMatrix(.Row, 9) = lnSubTotal '- (lnSubTotal * _
                                         (oTrans.Detail(.Row, 7) / 100) - oTrans.Detail(.Row, 8))
               Call ComputePOSTotal
               GridEditor1_AddingRow False
            End If
         End If
      End If
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = p_oAppDriver.getColor("HT1")
   End With
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lnRep As Integer
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_KeyDown"
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Then
      With GridEditor1
         Select Case .Col
         Case 1, 2, 3
            pbSearch = True
            If oTrans.SearchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
               'Nothing to do i guess...
            End If
            pbSearch = False
         End Select

         KeyCode = 0
         .SetFocus
         .Refresh
      End With
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )"
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .EditorBackColor = p_oAppDriver.getColor("EB")
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      If Index = 7 Then
         .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index) & "%"
      Else
         .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
      End If
   
      If Index = 7 Then
         ComputePOSSubTotal
         ComputePOSTotal
      End If
      
   End With

End Sub

Private Sub ComputePOSTotal()
   Dim lnCtr As Integer
   Dim lnTotal As Double
   Dim lnNetTl As Double

   With GridEditor1
      For lnCtr = 1 To .Rows - 1
         lnTotal = lnTotal + CDbl(.TextMatrix(lnCtr, 9))
      Next
   End With
   
   oTrans.Total = lnTotal
   
   txtField(5).Text = Format(IIf(lnTotal > 0, lnTotal, "0.00"), "#,##0.00")
End Sub

Private Sub ComputePOSSubTotal()
   With GridEditor1
      If .TextMatrix(.Row, 5) <> 0 Then
         .TextMatrix(.Row, 9) = CDbl(.TextMatrix(.Row, 5)) * CDbl(.TextMatrix(.Row, 6))
         .TextMatrix(.Row, 9) = (.TextMatrix(.Row, 9) * _
                                    (100 - CDbl(Left(.TextMatrix(.Row, 7), _
                                       Len(.TextMatrix(.Row, 7)) - 1))) / 100) - _
                                    CDbl(.TextMatrix(.Row, 8))
      End If
   End With
End Sub



