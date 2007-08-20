VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POP2OWA"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2580
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   2580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Apply"
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   1900
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   2415
      Begin VB.CheckBox chkFBA 
         Caption         =   "Form-Based-Auth"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox chkSend 
         Caption         =   "Save sent messages"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2175
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
      Height          =   1900
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Text            =   "25"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkSMTP 
         Caption         =   "SMTP"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Text            =   "110"
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txtIP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   7
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2200
      Left            =   5
      TabIndex        =   0
      Top             =   5
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3889
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

''
'Class to handle POP3/STMP commands.
Private oPOP3 As clsPOP3

''
'Form with systray code.
Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1
''
'Control the state of the SysTray icon.
Private bsysTray As Boolean

''
'Enable/Disable the SMTP features.
Private Sub chkSMTP_Click()
    Me.txtPort(1).Enabled = (chkSMTP.Value = vbChecked)
End Sub

''
'Apply the changes and reset the program
Private Sub cmdOk_Click()
    writeRegistry
    Reset
    FillControls
End Sub


''
'Main function to start the aplication in form mode.
Public Sub Init()
Dim objFrame As Frame

    Me.Caption = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
    'Redraw controls
    For Each objFrame In Me.Frame1
        objFrame.Move Me.TabStrip1.ClientLeft, _
                    Me.TabStrip1.ClientTop, _
                    Me.TabStrip1.ClientWidth, _
                    Me.TabStrip1.ClientHeight
    Next
    Me.Move Me.Left, Me.Top, 2700, 3150
    Set objFrame = Nothing
    'Get values from registry
    FillControls
    Reset
    Me.Show
End Sub

''
'Event used to finish the aplication.
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
    Set oPOP3 = Nothing
End Sub

Private Sub TabStrip1_Click()
    Frame1(2 - TabStrip1.SelectedItem.Index).Visible = False
    Frame1(TabStrip1.SelectedItem.Index - 1).Visible = True
    Frame1(TabStrip1.SelectedItem.Index - 1).ZOrder 0
    Frame1(TabStrip1.SelectedItem.Index - 1).Refresh
End Sub

''
'Fill the control in the form with the registry information.
Private Sub FillControls()
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
''
'Fill the registry with the data in the form.
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
    .Value = CInt(Me.txtPort(0).Text)
    'SMTP Port
    .ValueKey = "SMTP"
    .Value = CInt(Me.txtPort(1).Text)
    'Leave a copy in send folder
    .ValueKey = "Saveinsent"
    .Value = chkSend.Value
    'Form-Based-Authentication on/off
    .ValueKey = "FormBasedAuth"
    .Value = Me.chkFBA.Value
End With
End Sub

''
'Reset the POP/SMTP object to restore the connection with the new values.
Private Sub Reset()
    If Not (oPOP3 Is Nothing) Then
        Set oPOP3 = Nothing
    End If
    Set oPOP3 = New clsPOP3
End Sub

''
'Event to capture one click in the popup menu.
'
'@param lIndex  Number of the mouse button pressed
'@param sKey    Option menu
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
      'MsgBox "Clicked item with key " & sKey, vbInformation
   End Select
    
End Sub

''
'Event to capture double click in the systray icon.
'
'@param eButton Number of the mouse button pressed
Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    m_frmSysTray_MenuClick 0, "open"
End Sub

''
'Event to capture mouse down in the systray icon.
'
'@param eButton Number of the mouse button pressed
Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If (eButton = vbRightButton) Then
        m_frmSysTray.ShowMenu
    End If
End Sub
''
'Put/Remove pop2owa icon from the systray.
Private Sub SysTray()
    If bsysTray Then
        Set m_frmSysTray = New frmSysTray
        With m_frmSysTray
            .AddMenuItem "&Open Pop2Owa", "open", True
            .AddMenuItem "-"
            .AddMenuItem "&Close", "close"
            .ToolTip = "Pop2Owa!"
            .IconHandle = Me.Icon
        End With
    ElseIf Not (m_frmSysTray Is Nothing) Then
            Unload m_frmSysTray
            Set m_frmSysTray = Nothing
    End If
End Sub
