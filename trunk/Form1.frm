VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "POP2OWA"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2580
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   2580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   2265
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
      Begin VB.CheckBox chkFBA 
         Caption         =   "Form-Based-Auth"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox chkSend 
         Caption         =   "Save sent messages"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Apply"
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000A&
         Caption         =   "Exchange Server"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   2265
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Text            =   "25"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkSMTP 
         Caption         =   "SMTP"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Text            =   "110"
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txtIP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "IP"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "POP3"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1440
      Top             =   960
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2655
      Left            =   5
      TabIndex        =   0
      Top             =   5
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4683
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Exchange"
            Key             =   "Exchange"
            Object.ToolTipText     =   "Exchange Config"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "POP/SMTP"
            Key             =   "Protocols"
            Object.ToolTipText     =   "Protocols configuration"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
'Main form, with the program options.

Option Explicit

Private oPOP3 As clsPOP3
Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1
Private bsysTray As Boolean

''
'Enabe/Disable the SMTP features.
Private Sub chkSMTP_Click()
    Me.txtPort(1).Enabled = (chkSMTP.Value = vbChecked)
End Sub

''
'Apply the changes and reset the program
Private Sub cmdOk_Click()
    Timer1.Enabled = False
    oPOP3.Destroy
    Set oPOP3 = Nothing
    writeRegistry
    Reset
    Timer1.Enabled = True
End Sub
Public Sub Init()
    'Get values from registry
    readRegistry
    Reset
End Sub


Private Sub Form_Load()
On Error GoTo ErrHandler
    '
    'Get values from registry
'    readRegistry
       
'    Reset
    Timer1.Enabled = True
Exit Sub
ErrHandler:
    Timer1.Enabled = False
    MsgBox Err.Description, vbCritical, "POP2OWA: " & Err.Source
    WriteLog Err.Source & vbTab & Err.Description
    oPOP3.Destroy
    Set oPOP3 = Nothing
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not (m_frmSysTray Is Nothing) Then
    Unload m_frmSysTray
    Set m_frmSysTray = Nothing
End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized And Not bsysTray Then
        Me.Hide
        bsysTray = True
        SysTray
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    oPOP3.Destroy
    Set oPOP3 = Nothing
End Sub

Private Sub TabStrip1_Click()
    Timer1.Enabled = False
    Frame1(2 - TabStrip1.SelectedItem.Index).Visible = False
    Frame1(TabStrip1.SelectedItem.Index - 1).Visible = True
    Frame1(TabStrip1.SelectedItem.Index - 1).ZOrder 0
    Frame1(TabStrip1.SelectedItem.Index - 1).Refresh
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error GoTo ErrHandler
    '
    Timer1.Enabled = False
    oPOP3.Refresh
    If oPOP3.isActive Then
        Timer1.Interval = 1000
    Else
        If (Timer1.Interval <> 10000 And bsysTray) Then
            Me.Hide
    Me.Refresh
            DoEvents
        End If
        Timer1.Interval = 10000
    End If
    Timer1.Enabled = True
Exit Sub
ErrHandler:
    Timer1.Enabled = False
    MsgBox Err.Description, vbCritical, "POP2OWA: " & Err.Source
    WriteLog Err.Source & vbTab & Err.Description
    oPOP3.Destroy
    Set oPOP3 = Nothing
    Unload Me
End Sub


Public Sub readRegistry()
Dim c As cRegistry
Set c = New cRegistry
'Values are stored in (HKEY_CURRENT_USER\Software\pop2owa)
With c
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Software\pop2owa"
    .ValueType = REG_SZ
    'Exchange server
    .ValueKey = "ExchangeServer"
    Me.txtServer.Text = .Value
    'IP to listen
    .ValueKey = "IP"
    Me.txtIP.Text = .Value
    
    'Port values are integers
    .ValueType = REG_DWORD
    'POP3 Port
    .ValueKey = "POP3"
    Me.txtPort(0).Text = .Value
    'SMTP Port
    .ValueKey = "SMTP"
    Me.txtPort(1).Text = .Value
    'Leave a copy in send folder
    .ValueKey = "Saveinsent"
    Me.chkSend.Value = .Value
    'Form-Based-Authentication on/off
    .ValueKey = "FormBasedAuth"
    Me.chkFBA.Value = .Value
End With
End Sub

Private Sub writeRegistry()
Dim c As cRegistry
Set c = New cRegistry
'Values are stored in (HKEY_CURRENT_USER\Software\pop2owa)
With c
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Software\pop2owa"
    .ValueType = REG_SZ
    'Exchange server
    .ValueKey = "ExchangeServer"
    .Value = Me.txtServer.Text
    'IP to listen
    .ValueKey = "IP"
    .Value = Me.txtIP.Text
    
    'Port values are integers
    .ValueType = REG_DWORD
    'POP3 Port
    .ValueKey = "POP3"
    .Value = Me.txtPort(0).Text
    'SMTP Port
    .ValueKey = "SMTP"
    .Value = Me.txtPort(1).Text
    'Leave a copy in send folder
    .ValueKey = "Saveinsent"
    .Value = chkSend.Value
    'Form-Based-Authentication on/off
    .ValueKey = "FormBasedAuth"
    .Value = Me.chkFBA.Value
End With
End Sub

Public Sub Reset()
    If Not (oPOP3 Is Nothing) Then
        oPOP3.Destroy
        Set oPOP3 = Nothing
    End If
    Set oPOP3 = New clsPOP3
    With oPOP3
        .IP = Me.txtIP.Text
        .Port(0) = Me.txtPort(0).Text
        If chkSMTP.Value = vbChecked Then
            .Port(1) = Me.txtPort(1).Text
        End If
        .Saveinsent = (Me.chkSend.Value = vbChecked)
        .FormBasedAuthentication = (Me.chkFBA.Value = vbChecked)
        .Start
    End With
    strExchSvrName = Me.txtServer.Text
End Sub


Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
   Select Case sKey
   Case "open"
        bsysTray = False
        SysTray
        Me.WindowState = vbNormal
        Me.Show
        Me.ZOrder
   Case "close"
      Unload Me
   Case Else
      MsgBox "Clicked item with key " & sKey, vbInformation
   End Select
    
End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    m_frmSysTray_MenuClick 0, "open"
End Sub

Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If (eButton = vbRightButton) Then
        m_frmSysTray.ShowMenu
    End If
End Sub

Private Sub SetIcon()
 
m_frmSysTray.IconHandle = Me.Icon

End Sub

Private Sub SysTray()
    If bsysTray Then
        Set m_frmSysTray = New frmSysTray
        With m_frmSysTray
            .AddMenuItem "&Open Pop2Owa", "open", True
            .AddMenuItem "-"
            .AddMenuItem "&Close", "close"
            .ToolTip = "Pop2Owa!"
        End With
        SetIcon
    ElseIf Not (m_frmSysTray Is Nothing) Then
            Unload m_frmSysTray
            Set m_frmSysTray = Nothing
    End If
End Sub

