VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXM - ASSISTANT"
   ClientHeight    =   10950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11715
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      Height          =   375
      Left            =   8520
      TabIndex        =   38
      Top             =   7740
      Width           =   675
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   375
      Left            =   8520
      TabIndex        =   37
      Top             =   7200
      Width           =   675
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   375
      Left            =   8520
      TabIndex        =   36
      Top             =   6660
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   375
      Left            =   8520
      TabIndex        =   35
      Top             =   6120
      Width           =   675
   End
   Begin VB.Frame F_Header 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8115
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "By CPS"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   11
         Top             =   660
         Width           =   675
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "EXM - ASSISTANT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   3315
      End
   End
   Begin VB.Frame F_Run_As 
      Height          =   4215
      Left            =   60
      TabIndex        =   20
      Top             =   5340
      Width           =   8235
      Begin VB.CommandButton Command7 
         Caption         =   "Instructions"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   40
         ToolTipText     =   "Best Instructions found that describe how to setup a box to use Exmerge."
         Top             =   3540
         Width           =   1215
      End
      Begin VB.TextBox txtDomain 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1740
         TabIndex        =   2
         Top             =   2460
         Width           =   2325
      End
      Begin VB.TextBox txtExmLocation 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1740
         TabIndex        =   3
         Top             =   2940
         Width           =   4065
      End
      Begin VB.CommandButton Command13 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   5880
         TabIndex        =   4
         Top             =   2940
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd_ra_back 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5100
         TabIndex        =   6
         Top             =   3600
         Width           =   1155
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1740
         TabIndex        =   0
         Top             =   1440
         Width           =   2325
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1740
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1965
         Width           =   2325
      End
      Begin VB.CommandButton cmd_ra_run 
         Caption         =   "Run"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmd_ra_cancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6660
         TabIndex        =   7
         Top             =   3600
         Width           =   1155
      End
      Begin VB.Label lblLabels 
         Caption         =   $"frmMain.frx":0442
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   34
         Top             =   780
         Width           =   7500
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Domain:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Top             =   2475
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Caption         =   "EXM Location:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   3000
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Username:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   1980
         Width           =   1080
      End
      Begin VB.Label Label7 
         Caption         =   "Run EXM As "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame F_Component_Verification 
      Height          =   4215
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   8235
      Begin VB.CheckBox chk_Admin 
         Caption         =   "Windows Server Admin Tools Pack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1980
         TabIndex        =   33
         Top             =   1800
         Width           =   4035
      End
      Begin VB.CheckBox chk_Exmerge 
         Caption         =   "ExMerge Tool"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1980
         TabIndex        =   32
         Top             =   2880
         Width           =   2835
      End
      Begin VB.CheckBox chk_IIS 
         Caption         =   "IIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1980
         TabIndex        =   31
         Top             =   1260
         Width           =   2835
      End
      Begin VB.CheckBox chk_ExchSysTools 
         Caption         =   "Exchange System Management Tools"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1980
         TabIndex        =   30
         Top             =   2340
         Width           =   4815
      End
      Begin VB.TextBox txtAdminToolVersion 
         Height          =   315
         Left            =   6120
         TabIndex        =   29
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Instructions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   27
         ToolTipText     =   "Best Instructions found that describe how to setup a box to use Exmerge."
         Top             =   3540
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cv_verify 
         Caption         =   "Verify"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cv_cancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6660
         TabIndex        =   24
         Top             =   3600
         Width           =   1155
      End
      Begin VB.CommandButton cmd_cv_next 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5100
         TabIndex        =   19
         Top             =   3600
         Width           =   1155
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Read Me"
         Height          =   375
         Left            =   360
         Picture         =   "frmMain.frx":04DB
         TabIndex        =   15
         ToolTipText     =   "Use Windows Add/Remove to install all IIS components.  Click here for information."
         Top             =   1260
         Width           =   975
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Download"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         ToolTipText     =   $"frmMain.frx":091D
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Read Me"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         ToolTipText     =   "From the Exchange 2003 CD. Run setup.exe under the setup\i386 directory"
         Top             =   2340
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Download"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         ToolTipText     =   "Click here Microsoft Download URL for Exmerge.EXE."
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Component Verification"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   18
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Components Required"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1980
         TabIndex        =   17
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "How To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   16
         Top             =   840
         Width           =   675
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Redirect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8580
      TabIndex        =   39
      Top             =   5700
      Width           =   675
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Sub ShowFrame(iFrame)
    
    If iFrame = 0 Then
        F_Component_Verification.ZOrder vbBringToFront
    ElseIf iFrame = 1 Then
        F_Run_As.ZOrder vbBringToFront
    End If
    
End Sub

Private Sub cmd_cv_verify_Click()
    Dim sPath As String, sVal As String, sResult As String

    'verify IIS exists
    chk_IIS.Value = 0
    If ServiceExists("iisadmin") <> "" Then
        chk_IIS.Value = 1
    End If
    
    'verify Admin Tools
    'HKEY_CLASSES_ROOT\Installer\Features\2FC670E5DEFE2A346A32310E6DE27C0E
    chk_ExchSysTools.Value = 0
    sPath = "Installer\Features\0268927B6CAE1D11F878000680E26AE3"
    If CheckRegistryKey(HKEY_CLASSES_ROOT, sPath) Then
        chk_Admin.Value = 1
        txtAdminToolVersion.Text = "2000"
    End If
    sPath = "Installer\Features\2FC670E5DEFE2A346A32310E6DE27C0E"
    If CheckRegistryKey(HKEY_CLASSES_ROOT, sPath) Then
        chk_Admin.Value = 1
        txtAdminToolVersion.Text = "2K3"
    End If
    
    chk_ExchSysTools.Value = 0
    chk_Exmerge.Value = 0
    sResult = ""
    sPath = "SOFTWARE\Microsoft\Exchange\Setup"
    sVal = "ExchangeServerAdmin"
    sResult = Replace(QueryValue(HKEY_LOCAL_MACHINE, sPath, sVal), vbNullChar, "")
    If sResult <> "" Then
        chk_ExchSysTools.Value = 1
        
        'verify that exmerge now exists in the "sresult" path
        If FileExists(sResult & "\bin\exmerge.exe") Then
            chk_Exmerge.Value = 1
        End If
    End If
    
End Sub

Private Sub cmd_ra_run_Click()
    Ra.Username = Trim(txtUserName.Text)
    Ra.Password = Trim(txtPassword.Text)
    Ra.Domain = Trim(txtDomain.Text)
    Ra.ApplicationName = Trim(txtExmLocation.Text)
    If RunAs_VarsReady Then RunAs
End Sub

Function RunAs_VarsReady() As Boolean
    RunAs_VarsReady = False
    If Ra.Username <> "" And Ra.Password <> "" And Ra.Domain <> "" And Ra.ApplicationName <> "" Then
        RunAs_VarsReady = True
    End If
End Function

Private Sub Command1_Click()
    'http://technet.microsoft.com/en-us/exchange/bb288488.aspx
    OpenBrowser "http://technet.microsoft.com/en-us/exchange/bb288488.aspx"
End Sub

Private Sub Command10_Click()
    'if exchange 2007 then
        'http://www.microsoft.com/downloads/details.aspx?FamilyID=6be38633-7248-4532-929b-76e9c677e802&displaylang=en
        
    MsgBox "Exchange Server Management Tools " & vbNewLine & _
            "must be installed from the " & vbNewLine & _
            "Exchange 2003 CD under setup\i386\setup.exe." & vbNewLine & _
            "Then choose the management tools near the bottom "
End Sub

Private Sub Command11_Click()
    ' AdminPak.msi
    'http://www.microsoft.com/downloads/details.aspx?FamilyID=c16ae515-c8f4-47ef-a1e4-a8dcbacff8e3&displaylang=en
    OpenBrowser "http://www.microsoft.com/downloads/details.aspx?FamilyID=c16ae515-c8f4-47ef-a1e4-a8dcbacff8e3&displaylang=en"
End Sub

Private Sub Command12_Click()
    MsgBox "1. From the Add/Remove Windows Components " & vbNewLine & _
            "select Internet Information Services (IIS)" & vbNewLine & _
            "and click on Details." & vbNewLine & _
            vbNewLine & _
            "2. Set the checkbox for the " & _
            "Internet Information Services Snap-In " & vbNewLine & _
            "component and proceed with the installation."
End Sub

Private Sub Command13_Click()
    txtExmLocation.Text = SHFolder(frmMain)
End Sub

Private Sub Command6_Click()
    'http://www.exchangeinbox.com/article.aspx?i=58
    OpenBrowser "http://www.exchangeinbox.com/article.aspx?i=58"
    
End Sub



Private Sub Command7_Click()
    'http://www.exchangeinbox.com/article.aspx?i=58
    OpenBrowser "http://www.exchangeinbox.com/article.aspx?i=58"
End Sub

Private Sub Form_Load()
    Init
End Sub

Sub Init()

    

    With Me
    
        .txtExmLocation = App.Path & "\EXM.exe"
        
        .F_Header.Top = 0
        .F_Header.Left = 0
        .F_Header.Width = 8115
        .F_Header.Height = 1155
        
        .F_Component_Verification.Top = 1080
        .F_Component_Verification.Left = -60
        .F_Component_Verification.Width = 8235
        .F_Component_Verification.Height = 4215
        
        .F_Run_As.Top = 1080
        .F_Run_As.Left = -60
        .F_Run_As.Width = 8235
        .F_Run_As.Height = 4215
    
        .cmd_cv_verify.Top = 3600
        .cmd_cv_verify.Left = 360
        .cmd_cv_next.Top = 3600
        .cmd_cv_next.Left = 5100
        .cmd_cv_cancel.Top = 3600
        .cmd_cv_cancel.Left = 6660
            
        .cmd_ra_run.Top = 3600
        .cmd_ra_run.Left = 360
        .cmd_ra_back.Top = 3600
        .cmd_ra_back.Left = 5100
        .cmd_ra_cancel.Top = 3600
        .cmd_ra_cancel.Left = 6660
    
        .Width = 8190
        .Height = 5775
    
        ShowFrame 0
        .Show
        
    
    End With
    
End Sub



Private Sub cmd_cv_cancel_Click()
    Unload Me
End Sub

Private Sub cmd_cv_next_Click()
    ShowFrame 1
End Sub

Private Sub cmd_ra_back_Click()
    ShowFrame 0
End Sub

Private Sub cmd_ra_cancel_Click()
    Unload Me
End Sub

