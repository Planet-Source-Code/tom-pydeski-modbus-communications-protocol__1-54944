VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmMain 
   Caption         =   "Tom Pydeski's Modbus Communications - OffLine"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ModBusMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReadVel 
      Caption         =   "Read Velocity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   6360
      TabIndex        =   35
      Tag             =   "0"
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Auto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7560
      TabIndex        =   34
      Tag             =   "0"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdGetIO 
      Caption         =   "Get I/O"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7560
      TabIndex        =   33
      Tag             =   "0"
      Top             =   50
      Width           =   1095
   End
   Begin VB.CommandButton cmdShowBits 
      Caption         =   "Show Bits"
      Height          =   405
      Left            =   2040
      TabIndex        =   32
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoop 
      Caption         =   "Continuous Loop"
      Height          =   530
      Left            =   2040
      TabIndex        =   31
      Tag             =   "0"
      Top             =   50
      Width           =   1095
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   405
      Left            =   2640
      TabIndex        =   29
      Top             =   1245
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   714
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "txtNode"
      BuddyDispid     =   196626
      OrigLeft        =   3000
      OrigTop         =   1320
      OrigRight       =   3255
      OrigBottom      =   1695
      Max             =   32
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      TabIndex        =   28
      Text            =   "0"
      Top             =   7200
      Width           =   855
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   495
      LargeChange     =   1000
      Left            =   8520
      Max             =   0
      Min             =   32767
      SmallChange     =   100
      TabIndex        =   27
      Top             =   5760
      Width           =   255
   End
   Begin VB.TextBox txtVelocity 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7560
      TabIndex        =   25
      Text            =   "0"
      ToolTipText     =   "Changes the motor velocity for EN Drives"
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      TabIndex        =   22
      Text            =   "1102"
      Top             =   6720
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   8640
      Top             =   120
   End
   Begin VB.TextBox txtDelay 
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Text            =   "1000"
      Top             =   5040
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "---Inputs-------Outputs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   6360
      TabIndex        =   12
      Top             =   960
      Width           =   2895
      Begin VB.CheckBox chkOutput 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   2
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox chkOutput 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   1
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox chkOutput 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   0
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkInput 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CheckBox chkInput 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox chkInput 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   1
         Left            =   120
         Picture         =   "ModBusMain.frx":12FA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chkInput 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   0
         Left            =   120
         Picture         =   "ModBusMain.frx":183C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Image imgInOn 
         Height          =   240
         Left            =   120
         Picture         =   "ModBusMain.frx":1D7E
         Top             =   4200
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgInOff 
         Height          =   240
         Left            =   120
         Picture         =   "ModBusMain.frx":22C0
         Top             =   3840
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgOutOn 
         Height          =   480
         Left            =   1320
         Picture         =   "ModBusMain.frx":2802
         Top             =   4320
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Image imgOutOff 
         Height          =   480
         Left            =   1320
         Picture         =   "ModBusMain.frx":37C4
         Top             =   3840
         Visible         =   0   'False
         Width           =   1350
      End
   End
   Begin MSCommLib.MSComm Comm1 
      Left            =   1440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DTREnable       =   -1  'True
      InputMode       =   1
   End
   Begin VB.TextBox txtNode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   10
      Text            =   "32"
      Top             =   1250
      Width           =   480
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1935
      Left            =   45
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   17
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7455
      Left            =   3240
      TabIndex        =   8
      Top             =   45
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   13150
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbCmd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   50
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1750
      Width           =   3100
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "ModBusMain.frx":4786
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox txtQty 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   "1"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "1"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Get Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   50
      Picture         =   "ModBusMain.frx":478C
      ScaleHeight     =   1575
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   50
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Emerson EN Base Drive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   6480
      TabIndex        =   30
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label lblVelocity 
      Caption         =   "42022"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   26
      ToolTipText     =   "Displays the motor velocity for EN Drives"
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "New Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      Left            =   6360
      TabIndex        =   24
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "New Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      Left            =   6360
      TabIndex        =   23
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Delay"
      Height          =   255
      Left            =   6600
      TabIndex        =   21
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Slave Node #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1920
      TabIndex        =   11
      Top             =   990
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Quantity"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   800
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2175
      Width           =   800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:Tom Pydeski
'BitWise Industrial Automation
'F & L Machinery Design, Inc.
'This project communicates with devices via modbus
'it was originally intended for ascii mode, but is
'being used to talk to Emerson servo drives, which
'communicate via ModBus RTU
'It was tested with an EN-204 drive.
'© BitWise Industrial Automation, 2004
'
Option Explicit
Dim cellX As Integer
Dim cellY As Integer
Dim Unloaded As Byte
Dim gHeight As Integer
Dim ignScroll As Byte
Dim ScrollInit As Byte
Dim newVel As Long

Private Sub cmdShowBits_Click()
bitMax = (MaxReg - 1) * 16
If bitMax > 1000 Then
    eMess$ = "Maximum number of bits is 1000."
    eMess$ = eMess$ & vbCrLf & "You have selected " & bitMax & " bits!"
    eMess$ = eMess$ & vbCrLf & "Please try again."
    MsgBox eMess$, vbCritical + vbMsgBoxSetForeground, "Bit Limit Exceeded!"
    Exit Sub
End If
'SUBROUTINE FOR DECIPHERING DATA
For Bitw = 0 To ((ExpQty / 2) - 1)
    'pick out the registers
    WordIn = (RxIn(AddrIn + Bitw + 1))
    'isolate the bits starting at bitno
    BitNo = (AddrIn * 16) + (Bitw * 16) + 1
    Call DecDat
Next Bitw
ShowBits
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
Text1.Height = (Me.Height - Text1.Top) - 550
Grid1.Height = (Me.Height - Grid1.Top) - 550
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Comm1.PortOpen = False
Unloaded = 1
End
End Sub

Private Sub Form_Load()
On Error GoTo Oops
'
If App.PrevInstance Then
    eMess$ = "ModBus Comm is already running!"
    MsgBox eMess$, vbCritical + vbMsgBoxSetForeground, "Previous Instance Detected"
    Unload Me
    End
End If
For i = 0 To chkInput.Count - 1
    chkInput(i).Picture = imgInOff.Picture
    chkInput(i).DownPicture = imgInOn.Picture
Next i
For i = 0 To chkOutput.Count - 1
    chkOutput(i).Picture = imgOutOff.Picture
    chkOutput(i).DownPicture = imgOutOn.Picture
Next i
InitModBus
'"BBBB,P,D,S"
'Where BBBB is the baud rate, P is the parity, D is the number of data bits,
'and S is the number of stop bits. The default value of value is:
'"9600,N,8,1"
'old qbasic call for initializing the com port ;-)
'Open "COM1:19200,E,7,1,CS0,RS,DS0,CD0,LF,ASC,RB 2048,TB 2048" For Random As #1
PortNo = 1
'emerson uses rtu
Ser$ = "19200,N,8,2"
With Comm1
    .CommPort = PortNo
    .Settings = Ser$
    .InBufferSize = 2048
    .OutBufferSize = 2048
    .PortOpen = True
    .InputLen = 0
End With
cmbCmd.AddItem "Select Function"
cmbCmd.AddItem "1 - Read Coil Status"
cmbCmd.AddItem "2 - Read Input Status"
cmbCmd.AddItem "3 - Read Holding Registers"
cmbCmd.AddItem "4 - Read Input Registers"
cmbCmd.AddItem "5 - Modify Single Coil"
cmbCmd.AddItem "6 - Modify Single Register"
cmbCmd.AddItem "7 - Read Exception Status"
cmbCmd.AddItem "8 - Perform Diagnostic Test"
cmbCmd.AddItem "15 - Force Multiple Coils"
cmbCmd.AddItem "16 - Preset Multiple Registers"
cmbCmd.AddItem "17 - Report Slave ID number"
'cmbCmd.ListIndex = 3
Dim gformat$
gformat = "^Address|^ Value|^ Hex|^Char"
Grid1.FormatString = gformat
gformat = "^Base Addr|^ 1|^ 2|^ 3|^ 4|^ 5|^ 6|^ 7|^ 8|^ 9|^10|^11|^12|^13|^14|^15|^16"
Grid2.FormatString = gformat
'
'initialize exception error array
ExErr$(1) = " Illegal Function "
ExErr$(2) = " Illegal Data Address "
ExErr$(3) = " Illegal Data Value "
ExErr$(4) = " Failure In Associated Slave Device "
ExErr$(5) = " Acknowlege "
ExErr$(6) = " Busy, Rejected Message"
'
txtAddress.Text = GetSetting("ModBusTest", "Settings", "Addr", "1")
txtQty.Text = GetSetting("ModBusTest", "Settings", "Qty", "1")
cmbCmd.ListIndex = GetSetting("ModBusTest", "Settings", "Cmd", 3)
txtNode.Text = GetSetting("ModBusTest", "Settings", "Node", "1")
txtDelay.Text = GetSetting("ModBusTest", "Settings", "Delay", "100")
'
NodeAddr = Val(txtNode.Text)
Debug.Print Comm1.InBufferCount
If Comm1.InBufferCount <> 0 Then EmptyBuffer
NexTemp:
EmptyBuf
GoTo Exit_Form_Load
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Form_Load "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Form_Load"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Form_Load:
gHeight = 7400
Text1.Height = 4575
Text1.Width = 3135
Me.Height = Grid1.Top + gHeight + 550
Grid2.Visible = False
Show
Refresh
Beeep
Text1.Text = ""
'get velocity
ignScroll = 1
GoTo noind
'-------------------------------------------------
'below is for a single indexer, which we skip in this example
ReadInReg "2021", "2" '32021
If ComErr = 1 Then Exit Sub
lblVelocity = RxIn(2)
txtVelocity = lblVelocity
VScroll1.Value = Val(txtVelocity)
DoEvents
ignScroll = 0
'
GetStat
'-----------------------------------------------------
noind:
Me.ZOrder 1
End Sub


Private Sub cmdLoop_LostFocus()
If cmdLoop.Tag = 1 Then SendMessageBynum cmdLoop.hwnd, BM_SETSTATE, 1, 0
End Sub

Private Sub cmdLoop_Click()
Continuous = 1 - Continuous
'thanks to evan toder of psc for this little trick
Call SendMessageBynum(cmdLoop.hwnd, BM_SETSTATE, Continuous, 0)
If Continuous = 1 Then
    cmdLoop.Caption = "Stop Loop"
    DoEvents
    Refresh
    cmdExecute_Click
Else
    cmdLoop.Caption = "Continuous Loop"
End If
cmdLoop.Tag = Continuous
End Sub

Private Sub cmdGetIO_LostFocus()
If cmdGetIO.Tag = 1 Then SendMessageBynum cmdGetIO.hwnd, BM_SETSTATE, 1, 0
End Sub

Private Sub cmdGetIO_Click()
Continuous = 1 - Continuous
'thanks to evan toder of psc for this little trick
Call SendMessageBynum(cmdGetIO.hwnd, BM_SETSTATE, Continuous, 0)
If Continuous = 1 Then
    cmdGetIO.Caption = "Stop I/O"
    DoEvents
    Refresh
    GetStatus
Else
    cmdGetIO.Caption = "Get I/O"
End If
cmdGetIO.Tag = Continuous
End Sub

Private Sub cmdRun_LostFocus()
If cmdRun.Tag = 1 Then SendMessageBynum cmdRun.hwnd, BM_SETSTATE, 1, 0
End Sub

Private Sub cmdRun_Click()
cmdRun.Tag = Not cmdRun.Tag
Call SendMessageBynum(cmdRun.hwnd, BM_SETSTATE, cmdRun.Tag, 0)
If cmdRun.Tag Then
    'write to the enable register
    'this will allow the outputs to be cycled on a timed basis
    'for testing.
    ModBus.WriteReg 104, 32767
    cmdRun.Caption = "Stop Auto"
    DoEvents
    Refresh
Else
    cmdRun.Caption = "Run Auto"
End If
Timer1.Interval = txtDelay.Text
Timer1.Enabled = cmdRun.Tag
GetStat
End Sub

Private Sub cmdReadVel_LostFocus()
If cmdReadVel.Tag = 1 Then SendMessageBynum cmdReadVel.hwnd, BM_SETSTATE, 1, 0
End Sub

Private Sub cmdReadVel_Click()
cmdReadVel.Tag = 1 - cmdReadVel.Tag
Call SendMessageBynum(cmdReadVel.hwnd, BM_SETSTATE, cmdReadVel.Tag, 0)
If cmdReadVel.Tag = 1 Then
    cmdReadVel.Caption = "Stop Read"
    ignScroll = 1
    ReadInReg "2021", "2" '32021
    lblVelocity = RxIn(2)
    txtVelocity = lblVelocity
    VScroll1.Value = Val(txtVelocity)
    DoEvents
    ignScroll = 0
    GetVel
Else
    cmdReadVel.Caption = "Read Velocity"
    ignScroll = 0
    DoEvents
End If
End Sub

Sub GetStatus()
Do
    GetStat
Loop Until Continuous = 0
End Sub

Sub GetStat()
GetIO
For i = 1 To 4
    chkInput(i - 1).Value = Inputs(i)
Next i
For i = 1 To 3
    chkOutput(i - 1).Value = Outputs(i)
Next i
DoEvents
End Sub

Private Sub Text2_Change()
'You can put code in here to allow access for writing new data
End Sub

Private Sub Text3_Change()
'You can put code in here to allow access for writing new data
End Sub

Private Sub Timer1_Timer()
'this will write the outputs on a timed basis for testing.
OutReg = OutReg + 1
If OutReg = 4 Then OutReg = 0
OutData = 2 ^ OutReg
ModBus.WriteReg 103, Str$(OutData)
GetStat
End Sub

Private Sub chkOutput_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Caption = OutData
'toggle the state of the output
OutData = 0
For i = 0 To chkOutput.Count - 1
    OutData = OutData + (chkOutput(i).Value * (2 ^ i))
Next i
Caption = Caption & OutData
'write to the enable register
ModBus.WriteReg 104, 32767
'now write the value
ModBus.WriteReg 103, Str$(OutData)
End Sub

Private Sub cmbCmd_Change()
SaveSetting "ModBusTest", "Settings", "Cmd", cmbCmd.ListIndex
End Sub

Private Sub cmbCmd_Click()
SaveSetting "ModBusTest", "Settings", "Cmd", cmbCmd.ListIndex
Continuous = Val(cmdLoop.Tag)
End Sub

Private Sub cmbCmd_DropDown()
Continuous = 0
End Sub

Sub cmdExecute_Click()
Screen.MousePointer = vbHourglass
Text1.Text = ""
'disable the get data button so we don't try to re-execute
cmdExecute.Enabled = False
Grid1.Visible = False
Grid2.Visible = False
DoEvents
loopS:
Grid1.Redraw = False
Grid2.Redraw = False
'
NodeAddr = Val(txtNode.Text)
nobits = 1
Select Case cmbCmd.ListIndex
    Case Is = 1
        ReadCoils txtAddress.Text, txtQty.Text
        ShowBits
        nobits = 0
    Case Is = 2
        ReadInputs txtAddress.Text, txtQty.Text
        ShowBits
        nobits = 0
    Case Is = 3
        ReadReg txtAddress.Text, txtQty.Text
    Case Is = 4
        ReadInReg txtAddress.Text, txtQty.Text
    Case Is = 5
        SetOutput txtAddress.Text, txtQty.Text
    Case Is = 6
        WriteReg txtAddress.Text, txtQty.Text
    Case Is = 7
        'i can't get this to work with the emerson...
        'it may not be supported
        ReadException
    Case Is > 7
        MsgBox "This fuction has not been programmed yet..."
End Select
If ComErr = 1 Then GoTo ExClick
Grid1.Rows = (QtyIn \ 2) + 1 + (IIf((QtyIn Mod 2) = 0, 0, 1))
Dim HexVal$
For i = 1 To Grid1.Rows - 1
    Grid1.TextMatrix(i, 0) = HRBase + AddrIn + i
    Grid1.TextMatrix(i, 1) = RxIn(i)
    HexVal$ = AddZero(Hex$(RxIn(i)), 4)
    Grid1.TextMatrix(i, 2) = HexVal$
    Grid1.TextMatrix(i, 3) = Chr$(Val("&H" & Left$(HexVal$, 2))) & " " & Chr$(Val("&H" & Right$(HexVal$, 2)))
Next i
If nobits = 1 Then
    'gHeight = 7400
    gHeight = (Me.Height - Grid1.Top) - 550
    'Text1.Height = 4575
    Text1.Height = (Me.Height - Text1.Top) - 550
    Text1.Width = 3135
    If Me.WindowState = vbNormal Then Me.Height = Grid1.Top + gHeight + 550
    Grid2.Visible = False
Else
    gHeight = 4800
    Text1.Height = 2000
    Text1.Width = Grid2.Width
    If Me.WindowState = vbNormal Then Me.Height = Grid2.Top + Grid2.Height + 550
    Grid2.Visible = True
End If
If Grid1.Rows < 25 Then
    gHeight = Grid1.Rows * Grid1.RowHeight(0) + 100
    If Grid1.Top + gHeight > Text1.Top Then
        'resize the text box so it is not obscured
        Text1.Width = 3135
    End If
End If
Grid1.Height = gHeight
Grid1.Visible = True
Grid1.Redraw = True
If Me.WindowState = vbNormal Then
    If Me.Width < 4000 Then
        Me.Visible = False
        Me.Width = 6400
        Refresh
        Me.Top = (Screen.Height - Me.Height) / 2
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Visible = True
    End If
End If
Refresh
DoEvents
If Unloaded = 0 And Continuous = 1 Then GoTo loopS
ExClick:
cmdExecute.Enabled = True
Screen.MousePointer = 0
End Sub

Private Sub Grid1_DblClick()
Dim gRow As Integer
Dim TargetAddr$
Dim NewVal$
'try to write a new value
gRow = Grid1.Row
TargetAddr$ = Grid1.TextMatrix(gRow, 0)
NewVal$ = InputBox$("Please enter a new value for " & TargetAddr$, "Modify Value...")
If Len(NewVal$) > 0 Then
    WriteReg TargetAddr$, NewVal$
End If
DoEvents
cmdExecute_Click
End Sub

Private Sub grid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim r As Integer
Dim c As Integer
cellY = Grid1.MouseCol
cellX = Grid1.MouseRow
'
'below is old way
'cellX = 0
'cellY = 0
'For r = 0 To Grid1.Rows - 1
'    If Grid1.RowPos(r) > Y Then Exit For
'    cellX = r
'Next r
'For c = 0 To Grid1.Cols - 1
'    If Grid1.ColPos(c) > X Then Exit For
'    cellY = c
'Next c
Grid1.ToolTipText = Grid1.TextMatrix(cellX, 0) & " = " & Grid1.TextMatrix(cellX, 1) & " (" & Grid1.TextMatrix(cellX, 2) & "H)"
End Sub

Private Sub Grid2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim r As Integer
Dim c As Integer
cellY = Grid2.MouseCol
cellX = Grid2.MouseRow
'
'cellX = 0
'cellY = 0
'For r = 0 To Grid2.Rows - 1
'    If Grid2.RowPos(r) > Y Then Exit For
'    cellX = r
'Next r
'For c = 0 To Grid2.Cols - 1
'    If Grid2.ColPos(c) > X Then Exit For
'    cellY = c
'Next c
Grid2.ToolTipText = (Val(Grid2.TextMatrix(cellX, 0)) + cellY) & " = " & Grid2.TextMatrix(cellX, cellY)
End Sub

Private Sub txtAddress_Change()
SaveSetting "ModBusTest", "Settings", "Addr", txtAddress.Text
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtQty.SetFocus
End Sub

Private Sub txtDelay_Change()
SaveSetting "ModBusTest", "Settings", "Delay", txtDelay.Text
Timer1.Interval = txtDelay.Text
End Sub

Private Sub txtNode_Change()
If Val(txtNode.Text) < 1 Or Val(txtNode.Text) > 32 Then
    MsgBox "Valid Node Address is 1 to 32!" & vbCrLf & "Please try again.", vbCritical + vbMsgBoxSetForeground, "Node Selection Errror!"
    Exit Sub
End If
SaveSetting "ModBusTest", "Settings", "Node", txtNode.Text
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdExecute_Click
End Sub

Private Sub txtQty_Change()
SaveSetting "ModBusTest", "Settings", "Qty", txtQty.Text
End Sub

Sub ShowBits()
On Error GoTo Oops
Dim BitValue As Integer
'display each bit
'first setup the grid colors
'prevent the grid from showing changes as we make them (flicker-free)
Grid2.Redraw = False
Grid2.Visible = False 'this makes the grid update faster
Grid2.FillStyle = flexFillRepeat
Grid2.Row = 1
Grid2.Col = 1
Grid2.RowSel = Grid2.Rows - 1
Grid2.ColSel = Grid2.Cols - 1
Grid2.CellForeColor = vbBlack
Grid2.CellBackColor = vbWhite
Grid2.Row = 1
Grid2.Col = 1
Grid2.Row = 1
Grid2.ColSel = 1
Grid2.FillStyle = flexFillSingle
Grid2.Rows = (QtyIn \ 2) + 1 + (IIf((QtyIn Mod 2) = 0, 0, 1))
For i = 1 To Grid2.Rows - 1
    If HRBase < 40000 Then
        'emerson drives are not zero based, so the first input is 10001
        Grid2.TextMatrix(i, 0) = HRBase + (AddrIn * 16) + ((i - 1) * 16)
    Else
        'emerson drives are not zero based, so the first input is 10001
        Grid2.TextMatrix(i, 0) = HRBase + i
    End If
    For j = 1 To 16
        'get the value of each bit in the register and display it
        BitNo = (AddrIn * 16) + ((i - 1) * 16) + j
        If BitNo > bitMax Then Exit For
        BitValue = ModIn(BitNo)
        Grid2.TextMatrix(i, j) = BitValue
        If BitValue = 1 Then
            'change the color of bits that are on
            Grid2.Row = i
            Grid2.Col = j
            Grid2.CellForeColor = vbBlue
            Grid2.CellBackColor = vbYellow
        Else
            'grid2.CellForeColor = BLACK
            'grid2.CellBackColor = WHITE
        End If
    Next j
    'check if we have exceeded our intended last bit
    If BitNo > bitMax Then Exit For
Next i
Grid2.Height = Grid2.Rows * Grid2.RowHeight(0) + 100
Grid2.Visible = True
Grid2.Redraw = True
GoTo Exit_ShowBits
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ShowBits "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ShowBits"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ShowBits:
End Sub

Private Sub txtVelocity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdReadVel.Tag = 1 Then
        'update the velocity
        cmdReadVel.Tag = 0
        DoEvents
        ModBus.WriteReg 1102, txtVelocity
        DoEvents
        cmdReadVel.Tag = 1
    Else
        ModBus.WriteReg 1102, txtVelocity
        DoEvents
    End If
End If
End Sub

Sub GetVel()
Do
    ReadInReg "2021", "2" '32021
    lblVelocity = RxIn(2)
    DoEvents
Loop Until cmdReadVel.Tag = 0
End Sub

Private Sub VScroll1_Change()
Dim hun As Long
hun = 100
If ignScroll = 1 Then Exit Sub
If ScrollInit = 0 Then
    ignScroll = 1
    ReadInReg "2021", "2" '32021
    lblVelocity = RxIn(2)
    txtVelocity = lblVelocity
    VScroll1.Value = Val(lblVelocity)
    ignScroll = 0
    ScrollInit = 1
End If
If cmdReadVel.Tag = 0 Then
    DoEvents
    newVel = (VScroll1.Value \ 100)
    newVel = (newVel * hun)
    txtVelocity.Text = newVel
    ModBus.WriteReg 1102, txtVelocity
    DoEvents
Else
    cmdReadVel.Tag = 0
    DoEvents
    ModBus.WriteReg 1102, txtVelocity
    DoEvents
    cmdReadVel.Tag = 1
End If
ReadInReg "2021", "2" '32021
lblVelocity = RxIn(2)
txtVelocity = lblVelocity
End Sub
