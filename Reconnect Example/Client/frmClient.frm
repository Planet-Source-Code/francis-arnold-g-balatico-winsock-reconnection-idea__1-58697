VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWinSck.ocx"
Begin VB.Form frmClient 
   Caption         =   "Client Side"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Left            =   645
      Top             =   165
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Left            =   195
      Top             =   165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6000
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4860
      TabIndex        =   2
      Top             =   3810
      Width           =   915
   End
   Begin VB.TextBox txtSend 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   3825
      Width           =   4635
   End
   Begin VB.TextBox txtResult 
      Height          =   3615
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   105
      Width           =   5670
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private recon As Boolean

'variable to check if connection is the first time (Listen) or the second time (Connect) from server
Private frstConn As Boolean

Private Sub cmdSend_Click()
    'send message to server
    sckConnect.SendData Trim$(txtSend.Text)
    
    'display message
    txtResult.Text = txtResult.Text & "YOU: " & txtSend.Text & vbCrLf & vbCrLf
    txtResult.SelStart = Len(txtResult.Text)
End Sub

Private Sub Form_Load()
'initialize client winsock

'set variable to indicate that the client has not yet attempted reconnection
    recon = False
'set variable to indicate that client will connect first to listening server winsock
    frstConn = True
    With sckConnect
        .RemoteHost = "localhost"
        .RemotePort = 6000
        .Connect
    
'check connection state
'if closed then enable timer
    If .State <> sckConnected Then
        'client will try to reconnect
        tmrConnect.Interval = 3000 ' + counter1 + counter2 + counter3
        tmrConnect.Enabled = True
        cmdSend.Enabled = False
    Else
        'if connection is open then enable send button and disable timer
        'tmrConnect.Enabled = False
        cmdSend.Enabled = True
    End If
    
    End With
End Sub

Private Sub sckConnect_Close()
    With sckConnect
        .Close
        
        cmdSend.Enabled = False
        
        'display status of connection
        txtResult.Text = txtResult.Text & "Server closed connection" & vbCrLf & vbCrLf
        txtResult.SelStart = Len(txtResult.Text)
        
        If .State <> sckConnected Then
        'client will try to reconnect
        tmrConnect.Interval = 3000
        tmrConnect.Enabled = True
        End If
    End With
End Sub

Private Sub sckConnect_Connect()
    With sckConnect
        tmrConnect.Enabled = False
        
        'display connection status
        
        'upon connecting to server's winsock pool, display status
        If frstConn = False Then
            If recon = False Then
                txtResult.Text = txtResult.Text & "Connected to Server IP Address: " & .RemoteHostIP & vbCrLf & vbCrLf
                recon = True
            Else
                txtResult.Text = txtResult.Text & "Reconnected to Server IP Address: " & .RemoteHostIP & vbCrLf & vbCrLf
            End If
            txtResult.SelStart = Len(txtResult.Text)
            frstConn = True
        Else
            frstConn = False
        End If
        cmdSend.Enabled = True
    End With
End Sub

Private Sub sckConnect_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim msg As String
    Dim prseStr() As String
    
    
    sckConnect.GetData msg
    
    prseStr = Split(msg, "*")
    
    'if msg is signal to connect then close current connection to Listening winsock then reconnect
    'to a winsock in the server's winsock pool/collection
    With sckConnect
        If prseStr(0) = "Connect" Then
            .Close
            .RemoteHost = "Localhost"
            .RemotePort = Val(prseStr(1))
            .Connect
        ElseIf prseStr(0) = "Message" Then
    'else if signal indicates a message display it
            txtResult.Text = txtResult.Text & "SERVER : " & Trim$(prseStr(1)) & vbCrLf & vbCrLf
            txtResult.SelStart = Len(txtResult.Text)
        End If
    End With
End Sub

Private Sub sckConnect_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    With sckConnect
        .Close
                
       'client will try to RECONNECT
        tmrConnect.Interval = 3000
        tmrConnect.Enabled = True
        
        cmdSend.Enabled = False
    End With
End Sub

Private Sub tmrConnect_Timer()
On Error Resume Next
'attempt reconnect every set interval
    With sckConnect
        .RemoteHost = "localhost"
        .RemotePort = 6000
        .Connect
        
'if connected then disable timer
    If .State = sckConnected Then
        tmrConnect.Enabled = False
    End If
    End With
End Sub
