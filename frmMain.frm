VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Downloader Test Application"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   2880
      TabIndex        =   25
      Text            =   "http://www.allapi.net/"
      Top             =   120
      Width           =   3855
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   24
      Top             =   4080
      Width           =   6375
   End
   Begin VB.Frame Frame1 
      Caption         =   " Proxy Server "
      Height          =   2865
      Left            =   3720
      TabIndex        =   12
      Top             =   600
      Width           =   2895
      Begin VB.ComboBox cmbProxyServer 
         Height          =   315
         ItemData        =   "frmMain.frx":1CFA
         Left            =   600
         List            =   "frmMain.frx":1D0A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Text            =   "Administrator"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   16
         Text            =   "MyPassword"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtProxy 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   15
         Text            =   "10.0.0.1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtProxyPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Text            =   "80"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkRemoteDNS 
         Caption         =   "Use remote DNS"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   270
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proxy:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   750
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   195
         Left            =   1920
         TabIndex        =   19
         Top             =   750
         Width           =   330
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Download Type "
      Height          =   1815
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   3255
      Begin VB.OptionButton optDlType 
         Caption         =   "Download to File"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   480
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   7
         Top             =   480
         Width           =   2655
         Begin VB.OptionButton optDlToFile 
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   10
            Top             =   15
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.TextBox txtFilename 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   9
            Text            =   "C:\myfile.ext"
            Top             =   0
            Width           =   2415
         End
         Begin VB.OptionButton optDlToFile 
            Caption         =   "Use temporary file"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.OptionButton optDlType 
         Caption         =   "Use a download buffer"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
      Begin VB.OptionButton optDlType 
         Caption         =   "Use a download stream"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Value           =   -1  'True
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Progress "
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   3255
      Begin VB.PictureBox picProgress 
         AutoRedraw      =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   100
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblDownloadStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Downloaded x bytes from y bytes"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdFetch 
      Caption         =   "Fetch URL"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the URL of the file to download:"
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   2670
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result:"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   4920
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents RemoteFile As DataConnection
Attribute RemoteFile.VB_VarHelpID = -1
Dim DlType As DOWNLOAD_TYPE
Private Sub Form_Load()
    'Initialize WinSock... (this*must* be done)
    StartWinsock vbNullString
    'Create a new DataConnection class
    Set RemoteFile = New DataConnection
    'Set default proxy to 'No Proxy'
    cmbProxyServer.ListIndex = 0
    'Set default download type to Stream Buffer
    optDlType_Click 2
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Clean up...
    Set RemoteFile = Nothing
    EndWinsock
End Sub
Private Sub optDlType_Click(Index As Integer)
    optDlToFile(0).Enabled = (Index = 0)
    optDlToFile(1).Enabled = (Index = 0)
    txtFilename.Enabled = (Index = 0)
    DlType = Index
End Sub
Private Sub cmbProxyServer_Click()
    chkRemoteDNS.Enabled = (cmbProxyServer.ListIndex = 2 Or cmbProxyServer.ListIndex = 3)
    txtPassword.Enabled = (cmbProxyServer.ListIndex = 3)
    txtUsername.Enabled = chkRemoteDNS.Enabled
    txtProxy.Enabled = (cmbProxyServer.ListIndex <> 0)
    txtProxyPort.Enabled = (cmbProxyServer.ListIndex <> 0)
End Sub
Private Sub cmdFetch_Click()
    txtOutput.Text = ""
    With RemoteFile
        .DownloadType = DlType
        .Filename = txtFilename.Text
        .UseTempFile = optDlToFile(1).Value
        .ProxyType = cmbProxyServer.ListIndex
        .ProxyHostname = txtProxy.Text
        .ProxyPort = Val(txtProxyPort.Text)
        .ProxyUsername = txtUsername.Text
        .ProxyPassword = txtPassword.Text
        .ProxyUseRemoteDNS = (chkRemoteDNS.Value = vbChecked)
        .AutoRedirect = True
        .AllowCache = False
        .UseResume = False
        picProgress.Cls
        .Disconnect
        .MaxDownload = 1972224 '40712
        .FetchURLString txtURL.Text
        Debug.Print CStr(.SocketHandle)
    End With
End Sub
Private Sub RemoteFile_BytesReceived(lByteCount As Long, ID As Long)
    'If the script knows how many bytes that it has to receive then
    If RemoteFile.DownloadLength > 0 Then
        '... draw a progress bar
        lblDownloadStatus.Caption = "Downloaded " + CStr(lByteCount) + " bytes from " + CStr(RemoteFile.DownloadLength) + "."
        picProgress.ScaleWidth = RemoteFile.DownloadLength
        picProgress.Line (0, 0)-(lByteCount, 1), , BF
    'Or else, simply show how many bytes we have received
    Else
        lblDownloadStatus.Caption = "Downloaded " + CStr(lByteCount) + " bytes."
    End If
End Sub
Private Sub RemoteFile_Connected(ID As Long)
    'Successfully connected to the remote host
    Debug.Print "Connected"
End Sub
Private Sub RemoteFile_Disconnected(ID As Long)
    Debug.Print "Disconnected"
    'If we have downloaded everything to a buffer...
    If RemoteFile.DownloadType = dtToBuffer Then
        '... then show it
        txtOutput.Text = RemoteFile.GetBufferAsString
        RemoteFile.ClearBuffer
    End If
    'Tell the user the download is complete
    If RemoteFile.DownloadType = dtToFile And RemoteFile.UseTempFile Then
        MsgBox "Download completed!" + vbCrLf + "Result saved to " + RemoteFile.Filename, vbInformation
    Else
        MsgBox "Download completed!", vbInformation
    End If
End Sub
Private Sub RemoteFile_DownloadFailed(Why As DOWNLOAD_FAILURE, ID As Long)
    'Uhoh... the download failed :(
    Debug.Print "Download Failed"
    MsgBox "The download failed... Error code " + CStr(Why), vbExclamation
End Sub
Private Sub RemoteFile_StreamBytes(lByteCount As Long, bBytes() As Byte, ID As Long)
    'Show the received bytes
    txtOutput.Text = txtOutput.Text + StrConv(bBytes, vbUnicode)
End Sub

