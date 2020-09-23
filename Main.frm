VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secure Me (+)"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4800
      Top             =   1680
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2100
      TabIndex        =   2
      Text            =   "You do not have permission to access this service and are being reported to your ISP for malicious activatie."
      Top             =   120
      Width           =   3735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Send Intruder Message:"
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2400
      Top             =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2140
      TabIndex        =   4
      Text            =   "80"
      Top             =   720
      Width           =   520
   End
   Begin VB.TextBox Logtextfile 
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   5655
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Send me a visual alert when a intruder tries to connect to my computer."
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "Port to listen for connection:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
' this is our start button when its clicked
Winsock1.LocalPort = Text3.Text 'telling winsock what the port is... the port is text3.text
Winsock1.Listen 'now that we know which port.. we tell it to listen for a connection on that port!
Check1.Enabled = False 'disabling check box so nothing messes up!
Check2.Enabled = False 'disabling other check box so nothing messes up!
Text2.Enabled = False 'disabling text2 so nothing messes up!
Text3.Enabled = False 'disabling text3 so nothing messes up!
Command1.Enabled = False 'disabling start button so nothing can mess up!
Command2.Enabled = True 'Enabling the stop button so you can stop it!
Logtextfile.Text = "!!! Started " + Format(Now, "General Date") + vbCrLf 'Adding !!! Started then adding the time and date to the log.
Logtextfile.Text = Logtextfile.Text + "!!! Listening for connections on port " + Text3.Text + vbCrLf 'Adding to the log which port it is listening for a connection on!
End Sub

Private Sub Command2_Click()
' this is our stop button when its clicked
Check1.Enabled = True 'Enabling Check1 so that it can be used again
Check2.Enabled = True 'Enabling Check2 so it can be used again
Text2.Enabled = True 'Enabling text2 so it can be used again
Text3.Enabled = True 'Enabling text3 so it can be used again
Command1.Enabled = True 'Enabling the start button so it can be used again!
Command2.Enabled = False 'Disabling the stop button so nothing can mess up!
Winsock1.Close 'Telling winsock to stop listening for a connection!
Logtextfile.Text = Logtextfile.Text + "!!! Stopped " + Format(Now, "General Date") + vbCrLf 'Adding to the log that the program has stopped + the date and time
Logtextfile.Text = Logtextfile.Text + "!!! Closing port " + Text3.Text + vbCrLf 'Adding to the log that we are closing the port!
End Sub

Private Sub Form_Terminate()
'Making shure our program closes and doesnt hang in memory
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Making shure our program closes and doesnt hang in memory
End
End Sub

Private Sub Logtextfile_Change()
'When ever the log text box changes make shure it auto scrolls to the bottom line!
    On Error Resume Next
    Logtextfile.SelLength = 0
    If Len(Logtextfile.Text) > 0 Then
        If Right$(Logtextfile.Text, 1) = vbCrLf Then
            Logtextfile.SelStart = Len(Logtextfile.Text) - 1
            Exit Sub
        End If
        Logtextfile.SelStart = Len(Logtextfile.Text)
    End If
End Sub

Private Sub Text2_Change()
'Checks to see if text2's text is meth0s
If Text2.Text = "meth0s" Then MsgBox "He is elite and so are his friends :)", vbSystemModal, "InSomNic Delta Rez Fraxx & Rebol" 'If text2's text is meth0s then send a msgbox to the user!
End Sub

Private Sub Timer1_Timer()
'Timer1 ones job is to update our label1!
Label1.Caption = Format(Now, "General Date") + " - " + Winsock1.LocalIP + " on " + Winsock1.LocalHostName ' so every 1/1000th's of a second it updates the label to the new time + the IP of your computer and the host name of your computer!
End Sub

Private Sub Timer2_Timer()
'timer2's job is to disconnect the intruder from yur computer 2 seconds after he has connected!
Winsock1.Close 'Disconnecting the intruder from your computer
Winsock1.Listen 'Listening for another intruder to connect to your computer!
If Check2.Value = vbChecked Then MsgBox Winsock1.RemoteHostIP + " has been disconnected from your computer!", vbSystemModal, "ALERT!" 'If check2 is checked then send me a msgbox saying that he is being disconnected!
Timer2.Enabled = False 'Disabling this because there is no more intruder!
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'This is when someone trie's to connect weather we allow them to or not and what happens when they try!
Winsock1.Close 'Under winsock contrl if someone wants to connect you must close the port to alow them!
Winsock1.Accept requestID 'Accepting there request to connect!
Logtextfile.Text = Logtextfile.Text + "** " + Format(Now, "General Date") + " - " + Winsock1.RemoteHostIP + "/" + Text3.Text + " Connection Attempted!" + vbCrLf 'Adding to the log that the intruder *his ip* and *which port* has made a connection attempt!
If Check1.Value = vbChecked Then Winsock1.SendData Text2.Text 'If check1 is checked then send a message to the intruder... the message is text2.text
If Check2.Value = vbChecked Then MsgBox "Computer " + Winsock1.RemoteHostIP + " has connected!" + vbCrLf + "Disconnecting them from your computer!", vbSystemModal, "ALERT!" 'If check2 is checked then send us a msgbox saying that the intruder *IP* has tried to connnect to us!
Timer2.Enabled = True 'Enabling timer2 so we can disconnect him from our computer!
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'If the intruder trie's to send us any data
Dim incomingdata As String
Winsock1.GetData incomingdata 'Where getting the data to prevent any crash but not putting it into anything except a string!
End Sub
