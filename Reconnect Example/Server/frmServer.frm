VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Side"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbComNum 
      Height          =   315
      Left            =   105
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3825
      Width           =   5685
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Index           =   0
      Left            =   720
      Top             =   225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   225
      Top             =   225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6000
   End
   Begin VB.TextBox txtSend 
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   4215
      Width           =   4635
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
      Left            =   4845
      TabIndex        =   1
      Top             =   4200
      Width           =   915
   End
   Begin VB.TextBox txtResult 
      Height          =   3675
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   90
      Width           =   5670
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   4605
      Width           =   45
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private counter As Long
'holds the index used to load the winsock in the winsock pool
Private sckNum As Integer

Private Sub cmdSend_Click()
'verifies and sends message to client
    If Trim$(cmbComNum) = "" Then
        MsgBox "Choose computer number to send to.", vbOKOnly + vbExclamation, "Attention"
        Exit Sub
    Else
        sckConnect(Val(cmbComNum.Text)).SendData "Message*" & txtSend.Text
        txtResult.Text = txtResult.Text & "Server : " & txtSend.Text & vbCrLf & vbCrLf
        txtResult.SelStart = Len(txtResult.Text)
    End If
End Sub

Private Sub Form_Load()
    counter = 0
    lblCount.Caption = counter
'listen for connection attempts
    With sckListen
        .LocalPort = 6000
        .Listen
    End With
End Sub

Private Sub sckConnect_Close(Index As Integer)
'when client closes connection server should close too

    With sckConnect(Index)
        .Close
        
        txtResult.Text = txtResult.Text & "Client Computer # " & Index & " closed @ Port : " & .LocalPort & vbCrLf & vbCrLf
        txtResult.SelStart = Len(txtResult.Text)
        
        If Not Index = 0 Then
            Unload sckConnect(Index)
        End If
        
        cmbComNum.ListIndex = 0
        
'remove index(computer number) from combo box

        Do While Not Val(cmbComNum.Text) = Index
            cmbComNum.ListIndex = cmbComNum.ListIndex + 1
        Loop
        
        cmbComNum.RemoveItem cmbComNum.ListIndex
        
        counter = counter - 1
        lblCount.Caption = counter
    End With
End Sub

Private Sub sckConnect_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'accept connection from client

'connection request to this winsock happens after the client receives the signal
'and port number from the listening server winsock

'this is the second connection request
    With sckConnect(Index)
        If .State <> sckClosed Then
            .Close
        End If
        .Accept requestID
        
        counter = counter + 1
        lblCount.Caption = counter
'display result of connection
        txtResult.Text = txtResult.Text & "Connected to client computer # " & Index & " @ Port : " & .LocalPort & vbCrLf & vbCrLf
        txtResult.SelStart = Len(txtResult.Text)
        
'add computer number to the combobox
        cmbComNum.AddItem Index
    End With
End Sub

Private Sub sckConnect_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim msg As String
'accept message and display it
    With sckConnect(Index)
        .GetData msg
        txtResult.Text = txtResult.Text & "Computer " & Index & " : " & msg & vbCrLf & vbCrLf
        txtResult.SelStart = Len(txtResult.Text)
    End With
End Sub

Private Sub sckConnect_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    With sckConnect(Index)
        .Close
    End With
        Unload sckConnect(Index)
End Sub

Private Sub sckListen_Close()
'listening winsock should close and re-listen when client closes
    sckListen.Close
    sckListen.Listen
End Sub

Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
'when a connection is established to the listening winsock, load a new instance
'of a winsock in the winsock pool to accept the connection
    With sckListen
        If .State <> sckClosed Then
            .Close
        End If
        
        .Accept requestID
        
        sckNum = sckNum + 1
        
        Load sckConnect(sckNum)

'initialize new instance to accept connections
        sckConnect(sckNum).LocalPort = 6000 + sckNum
        sckConnect(sckNum).Listen
'send signal containing the newly created port number
        sckListen.SendData "Connect*" & Str(6000 + sckNum)
    End With
End Sub

Private Sub sckListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    With sckListen
        .Close
        .LocalPort = 6000
        .Listen
    End With
End Sub
