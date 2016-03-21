VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBaseAddresses 
   Caption         =   "[Cedalion] Base dll Address Generator."
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ComboIncrement 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   5520
      Width           =   1095
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   2143
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Random Base Address"
      TabPicture(0)   =   "frmBaseAddresses.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdGenerateRandom"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Specify Base Address"
      TabPicture(1)   =   "frmBaseAddresses.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "txtBeginingBaseAddress"
      Tab(1).Control(2)=   "cmdNoRandom"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdNoRandom 
         Caption         =   "Generate"
         Height          =   375
         Left            =   -72000
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtBeginingBaseAddress 
         Height          =   375
         Left            =   -74880
         TabIndex        =   6
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton cmdGenerateRandom 
         Caption         =   "Generate DLL base addresses"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Beginning Base Address"
         Height          =   255
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.TextBox txtNumbertogenerate 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Text            =   "0"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtBaseAddresses 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3360
      Width           =   5295
   End
   Begin VB.Label Label5 
      Caption         =   "increment by "
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Copy and paste into the box on the Project Properties/Compiler page of your VB project."
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   $"frmBaseAddresses.frx":0038
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Number of addresses to generate after base."
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "frmBaseAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Min_Address = 16777216
Const MAX_Address = 2147418112
Const c64K_Bytes = 65536
Const DefaultIncrement = c64K_Bytes * 4

Private m_Increment As Long

Private Property Let Increment(lData As Long)
  If lData Mod c64K_Bytes = 0 Then
    m_Increment = lData
  Else
    MsgBox lData & " is not a multiple of 64K"
  End If
  
End Property

Private Property Get Increment() As Long
  Increment = m_Increment
End Property

Private Function GetBaseAddress(Optional lBaseSet As Long = 0) As String
Dim tempbase As Long
   tempbase = 0
   
   'Make sure the random number is above the minimum 16MB
   Do While Not CheckValid(tempbase)
     tempbase = MAX_Address * Rnd()
   Loop
   
   'Make sure the number is a multiple of 64K
   Do While tempbase Mod c64K_Bytes <> 0
     tempbase = tempbase - 1
   Loop
   
   'Add &H to begining for copying and pasting into VB
   GetBaseAddress = "&H" & Hex$(tempbase) & GenerateFromBase(tempbase, lBaseSet)
   
End Function

Private Function CheckValid(lAddress As Long) As Boolean
  'Check address is > than minimum 16MB
  If lAddress >= Min_Address Then
    CheckValid = True
  End If
End Function

'*****************************************
'  This function generates a series of base
'  addresses from a basenumber.  These are simply
'  increments of 256K

Private Function GenerateFromBase(lBase As Long, lNumber As Long) As String
Dim i As Long
   For i = 1 To lNumber
     'Make the increment large enough to accommodate most dlls.
     lBase = lBase + (Increment)
     If lBase < MAX_Address Then
       GenerateFromBase = GenerateFromBase & vbCrLf & "&H" & Hex$(lBase)
     Else
       Exit For
     End If
   Next i
End Function


Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdGenerateRandom_Click()
  If txtNumbertogenerate.Text <> "" Then
    txtBaseAddresses.Text = GetBaseAddress(CLng(txtNumbertogenerate.Text))
  Else
    txtBaseAddresses.Text = GetBaseAddress()
  End If
End Sub

Private Sub cmdNoRandom_Click()
Dim base As Long
On Error GoTo ErrorTrap
    base = CLng(txtBeginingBaseAddress)
    txtBaseAddresses.Text = GenerateFromBase(base, CLng(txtNumbertogenerate))
Exit Sub
ErrorTrap:
  MsgBox "Unable to convert Base Address to Long" & vbCrLf & Err.Description, vbOKOnly, "Conversion Error"
End Sub

Private Sub Combo1_Change()
  

End Sub


Private Sub ComboIncrement_Click()
  Increment = (ComboIncrement.ListIndex + 1) * c64K_Bytes
End Sub

Private Sub Form_Load()
Dim i As Integer
  Randomize
  With ComboIncrement
    For i = 1 To 20
      .AddItem (CStr(64 * i) & "K")
    Next i
    .ListIndex = 3
  End With
  Increment = DefaultIncrement
End Sub

Private Sub txtNumbertogenerate_Change()
Dim temp As Long
On Error GoTo ErrorTrap
  If txtNumbertogenerate.Text <> "" Then
    temp = CLng(txtNumbertogenerate.Text)
  End If
  Exit Sub
ErrorTrap:
  MsgBox "You can only specify a number"
  txtNumbertogenerate.SelStart = txtNumbertogenerate.SelStart - 1
  txtNumbertogenerate.SelLength = 1
End Sub
