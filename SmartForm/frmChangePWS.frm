VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmChangePWS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Your Password"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIn 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2190
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1500
      Width           =   2760
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3225
      TabIndex        =   5
      Top             =   2085
      Width           =   1470
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Change &Password"
      Height          =   375
      Index           =   0
      Left            =   885
      TabIndex        =   4
      Top             =   2115
      Width           =   1470
   End
   Begin VB.TextBox txtIn 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2190
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1065
      Width           =   2760
   End
   Begin VB.TextBox txtIn 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2190
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   630
      Width           =   2760
   End
   Begin VB.TextBox txtIn 
      Height          =   345
      Index           =   0
      Left            =   2190
      TabIndex        =   0
      Top             =   195
      Width           =   2760
   End
   Begin MSComctlLib.ImageList ImgLstIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":0C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":0F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":129C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":15B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":18D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":1BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":1F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHA~1.frx":221E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirm Password: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   285
      TabIndex        =   9
      Top             =   1575
      Width           =   1845
   End
   Begin VB.Image ImgCP 
      Height          =   675
      Left            =   105
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "New Password: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   285
      TabIndex        =   8
      Top             =   1110
      Width           =   1845
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Old Password: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   270
      TabIndex        =   7
      Top             =   690
      Width           =   1845
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "User Id: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   270
      TabIndex        =   6
      Top             =   270
      Width           =   1845
   End
End
Attribute VB_Name = "frmChangePWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strDS As String
Private Sub cmdBtn_Click(Index As Integer)
   
   Select Case Index
    Case 0 '*** Changing Password ***
      
      If txtIn(2).Text = txtIn(3).Text Then
        
        If Change_PWS Then
           MsgBox "'" & txtIn(0).Text & "' Yours Password has changed successfully", vbInformation, "Password Change successful.."
            Unload Me
        End If
      Else
        MsgBox "Please Check your password.." & vbCrLf & _
               "The Password you typed do not match." & vbCrLf & _
               "Type the new password in the text boxes.", vbCritical, "Password Mismatch..."
        txtIn(3).SetFocus
      End If
    Case 1 '*** Termination of process ***
        Unload Me
   End Select
End Sub

Private Sub Form_Activate()
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbDefault
      

frmChangePWS.Icon = ImgLstIcons.ListImages(12).Picture
ImgCP.Picture = ImgLstIcons.ListImages(12).Picture
              
'frmChangePWS.txtIn(0).Text = frmlogin.txtid.Text
cmdBtn(0).Enabled = False
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

'frmLogin.txtid.Text = frmChangePWS.txtIn(0).Text
'frmLogin.txtpws.Text = ""
Unload Me

End Sub

Private Sub txtIn_Change(Index As Integer)
Dim i As Byte, OKBtn As Boolean
   
'If Len(txtIn(0).Text) > 0 And Len(txtIn(1).Text) > 0 And Len(txtIn(2).Text) > 0 Then
If Len(txtIn(0).Text) > 0 And Len(txtIn(1).Text) > 0 And Len(txtIn(2).Text) > 0 And Len(txtIn(3).Text) > 0 Then

     cmdBtn(0).Enabled = True
Else
     cmdBtn(0).Enabled = False
End If


End Sub

Private Sub txtIn_GotFocus(Index As Integer)
  txtIn(Index).SelStart = 0
  txtIn(Index).SelLength = Len(txtIn(Index))
  
  lblTitle(Index).ForeColor = vbGreen
End Sub

Private Sub txtIn_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    
    If cmdBtn(0).Enabled Then
        Call cmdBtn_Click(0)
    Else
    
        Select Case Index
 Case 0
        
        If Len(txtIn(1).Text) > 0 And Len(txtIn(2).Text) > 0 And Len(txtIn(3).Text) = 0 Then
            txtIn(3).SetFocus
        Else
           If Len(txtIn(1).Text) > 0 And Len(txtIn(2).Text) = 0 Then
            txtIn(2).SetFocus
           Else
            txtIn(1).SetFocus
           End If
        End If
        
 Case 1
        If Len(txtIn(2).Text) > 0 And Len(txtIn(3).Text) > 0 And Len(txtIn(0).Text) = 0 Then
            txtIn(0).SetFocus
        Else
           If Len(txtIn(2).Text) > 0 And Len(txtIn(3).Text) = 0 Then
            txtIn(3).SetFocus
           Else
            txtIn(2).SetFocus
           End If
        End If
 
 Case 2
        If Len(txtIn(3).Text) > 0 And Len(txtIn(0).Text) > 0 And Len(txtIn(1).Text) = 0 Then
            txtIn(1).SetFocus
        Else
           If Len(txtIn(3).Text) > 0 And Len(txtIn(0).Text) = 0 Then
            txtIn(0).SetFocus
           Else
            txtIn(3).SetFocus
           End If
        End If
 
 Case 3
        If Len(txtIn(0).Text) > 0 And Len(txtIn(1).Text) > 0 And Len(txtIn(2).Text) = 0 Then
            txtIn(2).SetFocus
        Else
           If Len(txtIn(0).Text) > 0 And Len(txtIn(1).Text) = 0 Then
            txtIn(1).SetFocus
           Else
            txtIn(0).SetFocus
           End If
        End If
        End Select
    End If
    End If
  
End Sub

Private Sub txtIn_LostFocus(Index As Integer)
    
    If txtIn(Index) <> "" Then
        lblTitle(Index).ForeColor = &H80&
    Else
        lblTitle(Index).ForeColor = vbBlack
    End If
    
End Sub

Private Function Change_PWS() As Boolean

MsgBox "Here you need to write a procedure to change the password.", vbInformation, "Password Changing..."

''''    On Error GoTo ErrHnd
''''      Dim oConn As New ADODB.Connection
''''      Dim oRs As New ADODB.Recordset
''''      Dim strConn As String, strChangePWS As String
''''
''''      pLog.SubInName = "Change_PWS"
''''      pLog.SubIn
''''
''''      strConn = "User ID =" & txtIn(0).Text & ";" & _
''''                "Password =" & txtIn(1).Text & ";" & _
''''                "Data Source =" & strDS & ";" & _
''''                "Persist Security Info=True"
''''      strChangePWS = "Alter user " & txtIn(0).Text & _
''''                          " identified by " & txtIn(3).Text
''''      pLog.Debugging strConn
''''      pLog.Activity "making ADO connection"
''''
''''      With oConn
''''        .ConnectionTimeout = 15
''''        .CursorLocation = adUseClient
''''        .Provider = "MSDAORA.1"
''''        .Open (strConn)
''''      End With
''''      pLog.Activity "ADO connection established.."
''''
''''      'oRs.Open "Select * from T_FLOOR", oConn, adOpenDynamic, adLockOptimistic
''''      'MsgBox oRs(0)
''''      oRs.Open strChangePWS, oConn, adOpenDynamic, adLockOptimistic
''''''      pLog.Activity "Password has changed"
''''
''''    Set oRs = Nothing
''''    Set oConn = Nothing
''''    Change_PWS = True
''''    pLog.SubOut
''''    Exit Function
''''ErrHnd:
''''    Change_PWS = False
''''    Debug.Print "Error No: " & CStr(Err.Number) & vbCrLf & "Error Description: " & Err.Description
''''    pLog.Debugging "Error No: " & CStr(Err.Number) & vbCrLf & "Error Description: " & Err.Description
''''
''''   Select Case Err.Number
''''     Case -214767259, -2147217843
''''''     MsgBox "The User Name or Old password is incorrect." & vbCrLf & _
''''                "latters in passwords must typed the correct case." & vbCrLf & _
''''                "Make sure the Caps Lock is not accidentally on", vbCritical, "Please check password"
''''
''''     pLog.Error Err.Description, Err.Number, _
''''                "The User Name or Old password is incorrect." & vbCrLf & _
''''                "latters in passwords must typed the correct case." & vbCrLf & _
''''                "Make sure the Caps Lock is not accidentally on"
''''     pLog.ErrorMsg
''''
''''     Case Else
''''''     MsgBox Err.Description & Err.Number, vbCritical, "Please check the Error"
''''     pLog.Error Err.Description, Err.Number, "Please check the Error"
''''     pLog.ErrorMsg
''''   End Select
''''   pLog.SubOut
End Function
