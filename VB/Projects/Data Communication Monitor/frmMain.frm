VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faraz's Data Communication Monitoring Tool"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock SockHTTP 
      Left            =   4320
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5940
      Top             =   2430
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":058A
            Key             =   "HTTP"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B24
            Key             =   "SOCKS"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11F6
            Key             =   "DC"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrCheck 
      Interval        =   100
      Left            =   2790
      Top             =   2880
   End
   Begin MSWinsockLib.Winsock SockOUT 
      Index           =   0
      Left            =   2250
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SockIN 
      Index           =   0
      Left            =   1710
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SockMain 
      Left            =   3780
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Menu"
      Height          =   285
      Left            =   7560
      TabIndex        =   2
      Top             =   2790
      Width           =   915
   End
   Begin ComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   3150
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   15028
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   2265
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   2790
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Connection List:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   1140
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup: Main Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupItems 
         Caption         =   "&Turn Proxy ON"
         Index           =   0
      End
      Begin VB.Menu mnuPopupItems 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPopupItems 
         Caption         =   "&View Log"
         Index           =   2
      End
      Begin VB.Menu mnuPopupItems 
         Caption         =   "&Remove Connection"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lvMB As Integer ' for check double-click on listview

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdMenu_Click()
 If SockMain.State = sckClosed Then
    ' Proxy is not ON yet.
    mnuPopupItems(0).Caption = "Turn Proxy &ON"
    mnuPopupItems(2).Enabled = False
    mnuPopupItems(3).Enabled = False
    PopupMenu mnuPopup, , , , mnuPopupItems(0)
 Else
    ' Proxy is running.
    mnuPopupItems(0).Caption = "Turn Proxy &OFF"
    mnuPopupItems(2).Enabled = True
    mnuPopupItems(3).Enabled = True
    PopupMenu mnuPopup, , , , mnuPopupItems(2)
 End If
End Sub

Private Sub Form_Load()
 With lvList
    .ColumnHeaders.Add , , "Time", 1290
    .ColumnHeaders.Add , , "Source IP", 1215
    .ColumnHeaders.Add , , "S.P.", 675
    .ColumnHeaders.Add , , "Destination IP", 1215
    .ColumnHeaders.Add , , "D.P.", 675
    .ColumnHeaders.Add , , "Bytes Sent", 1320
    .ColumnHeaders.Add , , "Bytes Received", 1320
    .ColumnHeaders.Add , , "Index", 600
 End With

 Status "Proxy Closed."
End Sub

Private Sub Form_Unload(Cancel As Integer)
 ' close all log windows if they're open
 Dim Frm As Form
 For Each Frm In Forms
    Unload Frm
 Next
End Sub

Private Sub lvList_DblClick()
 If lvMB = vbLeftButton Then
    ' View Log
    Call mnuPopupItems_Click(2)
 End If
End Sub

Private Sub lvList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If lvList.HitTest(x, y) Is Nothing Then Set lvList.SelectedItem = Nothing
 lvMB = Button
End Sub

Private Sub mnuPopupItems_Click(Index As Integer)
 Dim Ret    As Long
 Dim vKey   As String

 Select Case Index
    Case 0  ' On/Off
        Select Case Right(mnuPopupItems(0).Caption, 3)
            Case "&ON"
                StartProxy
            Case "OFF"
                StartProxy True
        End Select

    Case 2  ' View Log
        If lvList.SelectedItem Is Nothing Then Exit Sub
        frmLog.Show
        frmLog.vKey = ""
        frmLog.Caption = "View Log - " & lvList.SelectedItem.SubItems(1) & ":" & _
                                         lvList.SelectedItem.SubItems(2)

    Case 3  ' Remove/Terminate Connection
        If lvList.SelectedItem Is Nothing Then Exit Sub
        If lvList.SelectedItem.SmallIcon <> "DC" Then
            Ret& = MsgBox("This will disconnect this socket. Continue?", vbInformation + vbYesNo + vbDefaultButton2)
            If Ret& = vbNo Then Exit Sub
        End If
        ' Determine Key
        vKey = lvList.SelectedItem.SubItems(1) & ":" & lvList.SelectedItem.SubItems(2)
        ' Now do the disconnection working
        SockIN(Conns(vKey).SockINindex).Close
        Unload SockIN(Conns(vKey).SockINindex)
        
        SockOUT(Conns(vKey).SockOUTindex).Close
        Unload SockOUT(Conns(vKey).SockOUTindex)
        ' UI
        lvList.ListItems.Remove vKey
        Conns.Remove vKey
 End Select
End Sub

Public Sub StartProxy(Optional vTurnOff As Boolean)
 If vTurnOff = False Then
    ' Turn ON proxy
    SockMain.Close
    SockMain.LocalPort = SOCKS_LISTEN_PORT
    SockMain.Listen
    SockHTTP.Close
    SockHTTP.LocalPort = HTTP_LISTEN_PORT
    SockHTTP.Listen
    Status "Proxy started. HTTP Port: " & HTTP_LISTEN_PORT & _
           " and SOCKS Port: " & SOCKS_LISTEN_PORT
 Else
    ' Close proxy
    SockMain.Close
    SockHTTP.Close
    Status "Proxy Closed."
 End If
End Sub

Private Sub SockHTTP_ConnectionRequest(ByVal requestID As Long)
 Dim i As Integer, iTx As MSComctlLib.ListItem
 ' Technical stuff
 i = GetFreeSocket
 SockIN(i).Accept requestID
 Conns.Add SockIN(i).RemoteHostIP, SockIN(i).RemotePort, True, cHTTP
 ' For UI
 Set iTx = lvList.ListItems.Add
 With iTx
    .Key = SockIN(i).RemoteHostIP & ":" & SockIN(i).RemotePort
    .Text = Time
    .SmallIcon = "HTTP"
    .SubItems(1) = SockIN(i).RemoteHostIP
    .SubItems(2) = SockIN(i).RemotePort
    .SubItems(7) = SockIN(i).Index
 End With
 Status "HTTP Connection request from " & SockIN(i).RemoteHostIP & ":" & SockIN(i).RemotePort
End Sub

Private Sub SockIN_DataArrival(Index As Integer, ByVal bytesTotal As Long)
 Dim dGetS      As String
 Dim dGetB()    As Byte
 Dim vBytes     As Long
 Dim vKey       As String
 Dim iTx        As MSComctlLib.ListItem
 Dim vSendBack  As String
 Dim i          As Integer
 Dim dumData    As String
 Dim dumData2   As String

 vKey = SockIN(Index).RemoteHostIP & ":" & SockIN(Index).RemotePort

 ' Set bytes received.
 Set iTx = lvList.ListItems(vKey)
 vBytes = CLng("0" & Replace(iTx.SubItems(5), ",", "")) + bytesTotal
 iTx.SubItems(5) = Format(vBytes, "###,###,###0")

 ' Process received data
 SockIN(Index).GetData dGetS
 dGetB = dGetS

 If Conns(vKey).First Then
    If Conns(vKey).cType = cSOCKS Then      ' SOCKS Connection
        ' This is the first time data received at this socket. This data
        ' contains the destination IP and Port. Strip it!
        Conns(vKey).Send = dGetS
        Conns(vKey).DestPort = 256 * dGetB(4) + dGetB(6)
        Conns(vKey).DestIP = dGetB(8) & "." & dGetB(10) & "." & dGetB(12) & "." & dGetB(14)
        iTx.SubItems(3) = Conns(vKey).DestIP
        iTx.SubItems(4) = Conns(vKey).DestPort
    
        vSendBack = Mid(dGetS, 3, 6)
    
        AddToLog vKey, "SOCKS Connecion Started at " & iTx.Text & vbCrLf & _
                       "Source IP                  " & Conns(vKey).SourceIP & vbCrLf & _
                       "Source Port                " & Conns(vKey).SourcePort & vbCrLf & _
                       "Dest IP                    " & Conns(vKey).DestIP & vbCrLf & _
                       "Dest Port                  " & Conns(vKey).DestPort & vbCrLf & _
                       IIf(vSendBack <> "", "To Send Back               " & vSendBack & vbCrLf, "") & _
                       "---------------------"
        i = GetFreeSocketOUT
        Conns(vKey).SockOUTindex = i
        Conns(vKey).SockINindex = Index
     
        SockOUT(i).Connect PROXY_SOCKS_IP, PROXY_SOCKS_PORT
    
    Else        ' HTTP Connection
    
        Conns(vKey).Send = dGetS
        Conns(vKey).DestPort = 80
        dumData = Mid(dGetS, InStr(dGetS, " ") + 1)
        Conns(vKey).DestIP = Left(dumData, InStr(dumData, " ") - 1)
        iTx.SubItems(3) = Conns(vKey).DestIP
        iTx.SubItems(4) = Conns(vKey).DestPort
    
        AddToLog vKey, "HTTP Connecion Started at " & iTx.Text & vbCrLf & _
                       "Source IP                 " & Conns(vKey).SourceIP & vbCrLf & _
                       "Source Port               " & Conns(vKey).SourcePort & vbCrLf & _
                       "Dest IP                   " & Conns(vKey).DestIP & vbCrLf & _
                       "Dest Port                 " & Conns(vKey).DestPort & vbCrLf & _
                       "---------------------"
        AddToLog vKey, "SENT    > " & dGetS

        i = GetFreeSocketOUT
        Conns(vKey).SockOUTindex = i
        Conns(vKey).SockINindex = Index
     
        SockOUT(i).Connect PROXY_HTTP_IP, PROXY_HTTP_PORT
    End If
 
 Else
    ' Normal communication.
    dumData = StrConv(dGetS, vbUnicode)
    For i = 1 To Len(dumData) Step 2
        If i <> 0 Then
            dumData2 = dumData2 & Mid(dumData, i, 1)
        End If
    Next

    AddToLog vKey, "SENT    > " & dumData2
    SockOUT(Conns(vKey).SockOUTindex).SendData dGetS
 End If
End Sub

Private Sub SockMain_ConnectionRequest(ByVal requestID As Long)
 Dim i As Integer, iTx As MSComctlLib.ListItem
 ' Technical stuff
 i = GetFreeSocket
 SockIN(i).Accept requestID
 Conns.Add SockIN(i).RemoteHostIP, SockIN(i).RemotePort, True, cSOCKS
 ' For UI
 Set iTx = lvList.ListItems.Add
 With iTx
    .Key = SockIN(i).RemoteHostIP & ":" & SockIN(i).RemotePort
    .Text = Time
    .SmallIcon = "SOCKS"
    .SubItems(1) = SockIN(i).RemoteHostIP
    .SubItems(2) = SockIN(i).RemotePort
    .SubItems(7) = SockIN(i).Index
 End With
 Status "SOCKS Connection request from " & SockIN(i).RemoteHostIP & ":" & SockIN(i).RemotePort
End Sub

Private Function GetFreeSocket() As Integer
 ' NOTE: We do not use the Zero-Index of the socket.
 Dim i As Integer
 i = SockIN.UBound + 1
 Load SockIN(i)
 SockIN(i).Close
 GetFreeSocket = i
End Function

Private Function GetFreeSocketOUT() As Integer
 ' NOTE: We do not use the Zero-Index of the socket.
 Dim i As Integer
 i = SockOUT.UBound + 1
 Load SockOUT(i)
 SockOUT(i).Close
 GetFreeSocketOUT = i
End Function

Private Sub SockOUT_Connect(Index As Integer)
 Dim i As Integer
 ' Check which SockIN asked SockOUT to connect
 For i = 1 To Conns.Count
    If Conns(i).SockOUTindex = Index Then Exit For
 Next
 SockOUT(Index).SendData Conns(i).Send  ' we send the data 'as-is', we do not mess with it.
End Sub

Private Sub SockOUT_DataArrival(Index As Integer, ByVal bytesTotal As Long)
 Dim dGetS      As String
 Dim i          As Integer
 Dim vKey       As String
 Dim vBytes     As Long
 Dim iTx        As MSComctlLib.ListItem
 Dim dumData    As String
 Dim dumData2   As String
 
 SockOUT(Index).GetData dGetS
 
 ' Check which SockIN asked SockOUT to connect
 For i = 1 To Conns.Count
    If Conns(i).SockOUTindex = Index Then Exit For
 Next

 vKey = Conns(i).SourceIP & ":" & Conns(i).SourcePort

 ' Set bytes receieved from original proxy
 Set iTx = lvList.ListItems(vKey)
 vBytes = CLng("0" & Replace(iTx.SubItems(6), ",", "")) + bytesTotal
 iTx.SubItems(6) = Format(vBytes, "###,###,###0")

 If Conns(vKey).cType = cSOCKS Then
    ' SOCKS Connection
    If Conns(vKey).First Then
       ' this is the reply to the connection request.
       dumData2 = StrConv(dGetS, vbUnicode)
    Else
       ' make sure binary data is shown properly.
       dumData = StrConv(dGetS, vbUnicode)
       For i = 1 To Len(dumData) Step 2
           If i <> 0 Then
               dumData2 = dumData2 & Mid(dumData, i, 1)
           End If
       Next
    End If
 Else
    ' HTTP Connection
    dumData2 = dGetS
 End If
 Conns(vKey).First = False
 
 AddToLog vKey, "RECIEVED> " & dumData2

 If SockIN(Conns(vKey).SockINindex).State = sckConnected Then
    ' the SockIN 'may' get closed if the client program (requester)
    ' disconnects the socket itself before the proxy could return
    ' the data. we just made sure that the client is still connected.
    SockIN(Conns(vKey).SockINindex).SendData dGetS
 End If
End Sub

Private Sub tmrCheck_Timer()
 Dim iTx As MSComctlLib.ListItem

 On Error Resume Next   ' An error may occur here when you remove a connection
                        ' from the listview. when a connection is removed, its
                        ' socket is also unloaded. So below the sockets are checked
                        ' from 0 to ubound... so not necessarily all the sockets
                        ' in this count will be loaded. some may be unloaded.
                        '
                        ' NOTE: below the count (For i = 1 to ...) starts from
                        ' '1' because i dont like using 0-indexed sockets
                        ' because they cannot be unloaded. So me and my apps
                        ' work only loadable and unloadable sockets. :D ;D

 For i = 1 To SockIN.UBound
    If SockIN(i).State >= sckClosing Then
        ' This socket is either Closing or has occured an Error.
        For i2 = 1 To lvList.ListItems.Count
            Set iTx = lvList.ListItems(i2)
            If CLng(iTx.SubItems(7)) = i Then
                If iTx.SmallIcon <> "DC" Then
                    ' Show Disconnected icon
                    iTx.SmallIcon = "DC"
                End If
            End If
        Next
    End If
 Next
End Sub
