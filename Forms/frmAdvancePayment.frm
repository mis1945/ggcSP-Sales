VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAdvancePayment 
   BorderStyle     =   0  'None
   Caption         =   "MC Giveaways"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3885
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2085
      Left            =   105
      TabIndex        =   4
      Top             =   1710
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3678
      _Version        =   393216
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   5565
      TabIndex        =   5
      Top             =   555
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
      Picture         =   "frmAdvancePayment.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5565
      TabIndex        =   6
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmAdvancePayment.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1080
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1905
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   540
         Index           =   1
         Left            =   885
         TabIndex        =   3
         Top             =   405
         Width           =   4125
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   885
         TabIndex        =   1
         Top             =   105
         Width           =   4125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   2
         Top             =   405
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   0
         Top             =   105
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmAdvancePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmAdvancePayment"

Private p_oAppDrivr As clsAppDriver

Private p_oSkin As clsFormSkin
Private p_oSource As Recordset
Private p_bCancelxx As Boolean
Private pnCtr As Integer

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Set Source(Value As Recordset)
   Set p_oSource = Value
End Property

Property Let ClientName(Value As String)
   txtField(0).Text = Value
End Property

Property Let ClientAddress(Value As String)
   txtField(1).Text = Value
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsProcName As String
   
   lsProcName = "cmdButton_Click"
   ''On Error GoTo errProc
   
   With MSFlexGrid1
      Select Case Index
      Case 0, 1
         pnCtr = 0
         Me.Hide
         p_bCancelxx = Index = 1
      Case 2
         .Refresh
         .SetFocus
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Index & " )", True
End Sub

Property Get Cancelled() As Integer
   Cancelled = p_bCancelxx
End Property

Private Sub Form_Activate()
   p_bCancelxx = False
   
   Call LoadAdvPaym
End Sub

Private Sub Form_Load()
   Set p_oSkin = New clsFormSkin
   With p_oSkin
      Set .AppDriver = p_oAppDrivr
      Set .Form = Me
      .DisableClose = True
      .ApplySkin xeFormTransDetail
   End With
End Sub

Private Sub LoadAdvPaym()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Cols = 4
      .Rows = 2
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Reference #"
      .TextMatrix(0, 2) = "Amount"
      .TextMatrix(0, 3) = "Date"
      .Row = 0
      
      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next
      
      'column alignment
      .ColAlignment(1) = 1
      .ColAlignment(2) = 6
      .ColAlignment(3) = 1
      
      'column with
      .ColWidth(0) = 300
      .ColWidth(1) = 1300
      .ColWidth(2) = 1500
      .ColWidth(3) = 2000
      
      If p_oSource Is Nothing Then Exit Sub
      
      If p_oSource.RecordCount = 0 Then Exit Sub
      
      p_oSource.MoveFirst
      .Rows = p_oSource.RecordCount + 1
      For pnCtr = 0 To p_oSource.RecordCount - 1
         .TextMatrix(pnCtr + 1, 0) = pnCtr + 1
         .TextMatrix(pnCtr + 1, 1) = p_oSource("sReferNox")
         .TextMatrix(pnCtr + 1, 2) = Format(p_oSource("nTranAmtx"), "#,##0.00")
         .TextMatrix(pnCtr + 1, 3) = Format(p_oSource("dTransact"), "MMMM DD, YYYY")
         
         Call selectRow(pnCtr + 1, p_oSource("cTranStat") = xeStateClosed)
         p_oSource.MoveNext
      Next
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub selectRow(lnRow As Integer, lbSelect As Boolean)
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Row = lnRow
      For lnCtr = 1 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = lbSelect
         .CellBackColor = IIf(lbSelect, p_oAppDrivr.getColor("HT1"), p_oAppDrivr.getColor("EB"))
      Next
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set p_oSkin = Nothing
   Set p_oAppDrivr = Nothing
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      If bEnd Then
         .xShowError
      Else
         With Err
            .Raise .Number, .Source, .Description
         End With
      End If
   End With
End Sub

Private Sub MSFlexGrid1_DblClick()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      If .Row = 0 Then Exit Sub
      
      p_oSource.Move .Row - 1, adBookmarkFirst
      If p_oSource("cTranStat") = xeStateClosed Then
         p_oSource("cTranStat") = xeStateOpen
      Else
         p_oSource("cTranStat") = 1 'xeStateClosed
      End If
      Call selectRow(.Row, p_oSource("cTranStat") = xeStateClosed)
   End With
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeySpace
      Call MSFlexGrid1_DblClick
      KeyCode = 0
   End Select
End Sub
