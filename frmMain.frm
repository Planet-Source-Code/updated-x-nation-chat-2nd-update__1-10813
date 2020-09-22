VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6000
   ClientLeft      =   3975
   ClientTop       =   1965
   ClientWidth     =   8310
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8310
   Begin VB.TextBox txtMessage 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   4095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send!"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Winsock"
      ForeColor       =   &H00C0FFFF&
      Height          =   1095
      Left            =   5400
      TabIndex        =   22
      Top             =   4080
      Width           =   2775
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Session"
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "Listen"
         Height          =   375
         Left            =   1440
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Sound Effects"
      ForeColor       =   &H00C0FFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   5175
      Begin VB.CommandButton cmdDoorbell 
         Caption         =   "Doorbell"
         Height          =   375
         Left            =   2520
         TabIndex        =   30
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdHorn 
         Caption         =   "Horn"
         Height          =   375
         Left            =   2520
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSpring 
         Caption         =   "Spring"
         Height          =   375
         Left            =   1320
         TabIndex        =   28
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCough 
         Caption         =   "Cough"
         Height          =   375
         Left            =   1320
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdLaugh 
         Caption         =   "Laugh"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCuckoo 
         Caption         =   "Cuckoo"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Sounds On!"
         ForeColor       =   &H00FF80FF&
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000040&
         Caption         =   "Sounds Off!"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Show  Border (To Move!)"
      ForeColor       =   &H00FF80FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   16
      Top             =   3600
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1275
      Left            =   5520
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   1215
      ScaleWidth      =   2565
      TabIndex        =   15
      Top             =   360
      Width           =   2625
   End
   Begin VB.CommandButton cmdPopUp 
      Caption         =   "Popup"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   14
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtPopup 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   5280
      Width           =   6975
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtLocalIP 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtRemoteIP 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   2640
      Width           =   2655
   End
   Begin RichTextLib.RichTextBox RTFText 
      Height          =   2295
      Left            =   360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4048
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":2CA6
   End
   Begin VB.PictureBox picLocal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   120
      Picture         =   "frmMain.frx":2D54
      ScaleHeight     =   3480
      ScaleWidth      =   5265
      TabIndex        =   4
      Top             =   120
      Width           =   5265
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7680
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picExit 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   7920
      Picture         =   "frmMain.frx":430E
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   285
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   7560
      Picture         =   "frmMain.frx":45D1
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   285
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Port To Listen or Connect To:"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Local IP Address:"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote IP To Connect To:"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status Bar    -    X-Conn Nation by David Bowlin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Width           =   8055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Leave these settings alone to play sounds within a VB program
Private Declare Function mciSendString Lib "winmm.dll" Alias _
        "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
        lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
        hwndCallback As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Check1_Click()
        If Check1.Value = "1" Then
            Me.BorderStyle = 1
            Me.Caption = "X-Conn Nation"
        Else
            Me.Caption = ""
            Me.BorderStyle = 3
        End If
    txtMessage.SetFocus
End Sub

Private Sub cmdConnect_Click()
    On Error Resume Next ' If there's an error, resume the next command.
    If txtPort.Text < "1" Or txtPort.Text > "9999" Then txtPort.Text = "1818"
    Winsock1.Close ' Close any open ports (just in case).
    Winsock1.Connect txtRemoteIP.Text, txtPort.Text ' Try to connect to the computer IP address specified in the txtRemoteIP text box, on the port specified in the txtPort text box.
    lblStatus.Caption = "Connecting to " + txtRemoteIP.Text ' Inform the user we are trying to connect to the specified IP address.
    cmdDisconnect.Enabled = True ' Enable the Disconnect button since we may want to disconnect or stop trying to connect.
    cmdListen.Enabled = False ' Disable the Listen button and the Connect button since we are already trying to connect.
    cmdConnect.Enabled = False ' We are trying to connect, so hide the connect button.
    If err Then lblStatus.Caption = err.Description ' If there are any errors, inform the user by showing it on the lblStatus bar.
End Sub

Private Sub cmdCough_Click()
    On Error Resume Next
    Winsock1.SendData "!"
    If Option1.Value = True Then
        playsound = sndPlaySound("cough.wav", 1)
    End If
    txtMessage.SetFocus
End Sub

Private Sub cmdCuckoo_Click()
    On Error Resume Next
    Winsock1.SendData "@"
    If Option1.Value = True Then
        playsound = sndPlaySound("cuckoo.wav", 1)
    End If
    txtMessage.SetFocus
End Sub

Private Sub cmdDoorbell_Click()
    On Error Resume Next
    Winsock1.SendData "&"
    If Option1.Value = True Then
        playsound = sndPlaySound("doorbell.wav", 1)
    End If
    txtMessage.SetFocus
End Sub

Private Sub cmdHorn_Click()
    On Error Resume Next
    Winsock1.SendData "%"
    If Option1.Value = True Then
        playsound = sndPlaySound("horn.wav", 1)
    End If
    txtMessage.SetFocus
End Sub

Private Sub cmdLaugh_Click()
    On Error Resume Next
    Winsock1.SendData "~"
    If Option1.Value = True Then
        playsound = sndPlaySound("lol.wav", 1)
    End If
    txtMessage.SetFocus
End Sub

Private Sub cmdListen_Click()
    On Error Resume Next ' If there's an error, resume next command.
    cmdConnect.Enabled = False ' We are listening for a connection, so disable the Connect button.
    cmdListen.Enabled = False ' We are already listening for a connection, so disable the Listen button.
    cmdDisconnect.Enabled = True ' Enable the Disconnect button in case you want to stop listening for connection request.
    If txtPort.Text < "1" Or txtPort.Text > "9999" Then
        txtPort.Text = "1818"
    End If
    Winsock1.LocalPort = txtPort.Text ' Set the local port to listen on by getting the value from the txtPort text box.
    Winsock1.Listen ' Listen for the connection request by the other computer.
    lblStatus.Caption = "Listening For Connection Request" ' Inform the user that we are listening for a connection request.
End Sub

Private Sub cmdClear_Click()
    On Error GoTo err
    mbox = MsgBox("Clear the current chat session?", vbOKCancel, "Clear Session?")
    If mbox = vbOK Then
        RTFText.Text = ""
        RTFText.Text = ""
        txtMessage.SetFocus
        Exit Sub
    End If

err: ' If the user pressed Cancel on the message box above, we end up here, since this produces an error in Visual Basic
    txtMessage.SetFocus ' The user pressed Cancel, so we do nothing but reset the focus back to the outgoing message box
    Exit Sub ' Exit the subroutine
End Sub

Private Sub cmdDisconnect_Click()
    On Error Resume Next ' If there's an error, resume next command.
    Dim playsound As Long ' Declare the variable to hold the sound to be played if "Play Sounds" box is checked.
    If Option1.Value = True Then
        playsound = sndPlaySound("xcdiscon.wav", 1) ' If the "Play Sounds" box is checked, play the sound.
    End If
    Winsock1.Close ' We want to disconnect or stop listening for a connection request, so close the connected or listening port.
    cmdConnect.Enabled = True ' Enable the Connect button so we can connect to another computer.
    cmdListen.Enabled = True ' Enable the Listen button so we can listen for a connection request.
    cmdDisconnect.Enabled = False ' We are not connected to anything, so disable the Disconnect button.
    lblStatus.Caption = "Disconnected - Not Listening For Request." ' Show the user we are disconnected, and that we are not listening for a connection request.
End Sub
Private Sub cmdPopUp_Click()
    On Error Resume Next ' If there's an error, continue with next command
    If txtPopup.Text = "" Then Exit Sub ' If the txtPopup text box is empty, don't send any data, exit subroutine.
    Winsock1.SendData ("|" & txtPopup.Text) ' Send the data with a pipe, |, as first character.  The pipe tells our program that this message is for a popup message box.  See DATA_ARRIVAL subroutine.
    txtPopup.Text = "" ' Set the txtPopup box to blank for another popup message.
End Sub

Private Sub cmdSend_Click()
    On Error GoTo err ' If there is an error in this subroutine, go to "err" code at bottom
    If txtMessage.Text = "" Then Exit Sub ' If the message is blank, don't send it
    If Winsock1.State = 0 Then ' If we are not connected, do not attempt to send any messages
        txtMessage.Text = "" ' Clear message box
        txtMessage.SetFocus ' and set the focus back to the message box to send another message
        Exit Sub ' Exit subroutine
    End If
    Winsock1.SendData (txtMessage.Text) ' Send our message to the other computer
    RTFText.SelStart = Len(RTFText.Text) ' Set cursor to end of incoming message box. This keeps the last message on the screen
    RTFText.SelColor = &HC0FFFF    ' Make sure our text is yellow on the incoming message box
    RTFText.SelText = txtMessage.Text & vbCrLf ' Add current message to the incoming message box (RTFText)
    txtMessage.Text = "" ' Reset the outgoing message box
    txtMessage.SetFocus ' Set the focus back on the outgoing message box to send another message
    Exit Sub ' Exit subroutine
err: ' This is our error trap section for this routine
    lblStatus.Caption = err.Description ' Show the error to the user on the status bar
    txtMessage.Text = "" ' Reset the outgoing message box to blank
    txtMessage.SetFocus ' Set the focus to the outgoing message box to send another message
End Sub

Private Sub cmdSpring_Click()
    On Error Resume Next
    Winsock1.SendData "$"
    If Option1.Value = True Then
        playsound = sndPlaySound("spring.wav", 1)
    End If
    txtMessage.SetFocus
End Sub

Private Sub Form_Load()
    txtLocalIP.Text = Winsock1.LocalIP
    Option1.Value = True
End Sub

Private Sub picExit_Click()
    On Error Resume Next
    Winsock1.Close
    Unload Me
End Sub

Private Sub picMin_Click()
    On Error Resume Next
    Me.WindowState = vbMinimized
End Sub

Private Sub Picture1_Click()
    MsgBox "X-Conn Nation TCP Chat Software!" + vbCrLf + vbCrLf + "By David Bowlin" + vbCrLf + "Use at your own risk.", vbInformation, "X-Conn Nation"
    txtMessage.SetFocus
End Sub

Private Sub RTFText_GotFocus()
    txtMessage.SetFocus
End Sub

Private Sub rtftext_KeyPress(KeyAscii As Integer)
    On Error GoTo err
    RTFText.SelStart = Len(RTFText.Text) ' Set cursor to end of outgoing message box. This keeps the last message on the screen.
    RTFText.SelColor = &HC0FFFF
    Winsock1.SendData Chr(KeyAscii) ' Send each character (as it is typed to the other) computer.
    Exit Sub
err:
    lblStatus.Caption = err.Description ' Show the error to the user on the status bar.
    Resume Next ' Resume with next command after showing the error.
End Sub

Private Sub Winsock1_Close()
    Dim playsound As Long
        playsound = sndPlaySound("xcdiscon.wav", 1)
    lblStatus.Caption = "Connection Has Been Closed." ' Show the user that the connection is closed.
    cmdConnect.Enabled = True ' Reset the command buttons.
    cmdListen.Enabled = True  ' Connect and listen need to be enabled.
    cmdDisconnect.Enabled = False ' Disable Disconnect since we're not connected or listening for connection.
    cmdConnect.SetFocus ' Set the focus back to the cmdConnect button.
End Sub

Private Sub Winsock1_Connect()
    On Error Resume Next ' If there's an error, continue with next command.
    Dim playsound As Long ' Declare variable to hold the sound to be played if "Play Sounds" box is checked.
    lblStatus.Caption = "Connection Has Been Established!" ' Show the user we have a connection.
    txtRemoteIP.Text = Winsock1.RemoteHostIP ' Put the remote computer's IP in the remoteIP box.
    cmdConnect.Enabled = False ' Disable the Connect and Listen buttons.
    cmdListen.Enabled = False  ' We don't need these buttons enabled, and it prevents possible errors.
    cmdDisconnect.Enabled = True ' We are connected, so enable the Disconnect button.
        playsound = sndPlaySound("xcestab.wav", 1) '
    txtMessage.SetFocus ' Set the focus on the box to enter messages to send to the other computer.
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next ' Just in case there's an error, continue with next command.
    Dim playsound As Long ' Declare variable to hold the sound to be played if "Play Sounds" box is checked.
    Winsock1.Close ' Close any open socket (just in case).
    Winsock1.Accept requestID ' Accept the other computer's connection request.
    lblStatus.Caption = "Connection Has Been Established!" ' Show the user we have accepted the connection request, and are connected.
    txtRemoteIP.Text = Winsock1.RemoteHostIP ' Show the remote computer's IP in the txtRemoteIP text box.
    cmdConnect.Enabled = False ' We are connected, so disable the Connect and Listen buttons.
    cmdListen.Enabled = False ' This helps to prevent anyone from clicking them and causing errors.
    cmdDisconnect.Enabled = True ' Enable the Disconnect button since we're connected.
    If Option1.Value = True Then
        playsound = sndPlaySound("xcestab.wav", 1) ' If the "Play Sounds" is selected, play the default sound.
    End If
    txtMessage.SetFocus ' Set the focus on the box to enter messages to send to the other computer.
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim currenttext As String ' String to hold contents of rtftext if needed.
    Dim ndata As String ' Declare a variable to hold the incoming data.
    On Error Resume Next ' If there's an error, resume next command.
    Winsock1.GetData ndata ' Get the incoming data and store it in variable "ndata".
                If Mid(ndata, 1, 1) = "|" Then GoTo boxcode ' Check to see if this is a message box, if it is go to subroutine "boxcode".
                If Mid(ndata, 1, 1) = "~" Then GoTo lol
                If Mid(ndata, 1, 1) = "@" Then GoTo cuckoo
                If Mid(ndata, 1, 1) = "!" Then GoTo cough
                If Mid(ndata, 1, 1) = "$" Then GoTo spring
                If Mid(ndata, 1, 1) = "%" Then GoTo horn
                If Mid(ndata, 1, 1) = "&" Then GoTo doorbell
    RTFText.SelStart = Len(RTFText.Text) ' Set the cursor to the end of the text box to hold the incoming messages.
    RTFText.SelColor = &HFF00&
    RTFText.SelText = Mid(ndata, InStr(1, ndata, "^") + 1) & vbCrLf ' Insert the text to the last of the rtftext message box.
    RTFText.SelStart = Len(RTFText.Text) ' Set the cursor to the end of the text box.
    Exit Sub ' Exit the subroutine.
boxcode: ' If the incoming data's first character was a pipe , |, then the program jumps here.
    MsgBox Mid(ndata, 2, Len(ndata) - 1), vbInformation, Winsock1.RemoteHostIP & " says..." ' Display the incoming data as a message box.
    txtMessage.SetFocus ' Put the focus back on the rtftext box to send another message.
    Exit Sub ' Exit the subroutine.
lol:
    If Option1.Value = True Then
        playsound = sndPlaySound("lol.wav", 1)
    End If
    txtMessage.SetFocus
    Exit Sub
cuckoo:
    If Option1.Value = True Then
        playsound = sndPlaySound("cuckoo.wav", 1)
    End If
    txtMessage.SetFocus
    Exit Sub
cough:
    If Option1.Value = True Then
        playsound = sndPlaySound("cough.wav", 1)
    End If
    txtMessage.SetFocus
    Exit Sub
spring:
    If Option1.Value = True Then
        playsound = sndPlaySound("spring.wav", 1)
    End If
    txtMessage.SetFocus
    Exit Sub
horn:
    If Option1.Value = True Then
        playsound = sndPlaySound("horn.wav", 1)
    End If
    txtMessage.SetFocus
    Exit Sub
doorbell:
    If Option1.Value = True Then
        playsound = sndPlaySound("doorbell.wav", 1)
    End If
    txtMessage.SetFocus
    Exit Sub
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblStatus.Caption = Description ' If there was a winsock error, show the user.
    txtMessage.SetFocus ' Set the focus back on the message box to send another message.
End Sub

