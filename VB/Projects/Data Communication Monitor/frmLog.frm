VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmLog 
   Caption         =   "View Log"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CmnDlg 
      Left            =   2610
      Top             =   3150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Log"
      Height          =   285
      Left            =   6300
      TabIndex        =   2
      Tag             =   "lt"
      Top             =   3690
      Width           =   1275
   End
   Begin VB.Timer tmrLog 
      Interval        =   10
      Left            =   6840
      Top             =   3150
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Tag             =   "wh"
      Top             =   90
      Width           =   7485
   End
   Begin VB.Label lblLength 
      AutoSize        =   -1  'True
      Caption         =   "Log Length:"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Tag             =   "t"
      Top             =   3690
      Width           =   855
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vKey As String
Dim PrevH As Long, PrevW As Long

Private Sub Command1_Click()
 On Error GoTo Hell
 With CmnDlg
    .DialogTitle = "Save As..."
    .Filter = "Text Files|*.txt|All Files|*.*"
    .ShowSave
    Open .FileName For Binary As #1
        Put #1, , txtLog
    Close
 End With

Hell:
 If Err.Number = cdlCancel Then Exit Sub
 MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
 PrevH = 0
 PrevW = 0
 ResizeAll Me
End Sub

Private Sub Form_Load()
 PrevH = 0
 PrevW = 0
 ResizeAll Me
End Sub

Private Sub Form_Resize()
 ResizeAll Me
End Sub

Private Sub tmrLog_Timer()
 If vKey = "" Then
    If InStr(Caption, "-") = 0 Then Exit Sub
    vKey = Trim(Mid(Caption, InStr(Caption, "-") + 2))
 End If

 If Len(Conns(vKey).Log) <> Len(txtLog) Then
    txtLog = Replace(Conns(vKey).Log, Chr(0), " ")
    txtLog.SelStart = Len(txtLog)
 End If

 lblLength = "Log Length: " & Format(Len(Conns(vKey).Log), "###,###,###0")
End Sub

Private Function ResizeAll(Frm As Form) ' MODIFIED. DO NOT USE THIS.
 '
 ' === Novacrome Form-Resizer v2.3 ===
 '
 ' Keys Used: t l w h g p

 ' Grp - Group: used for adjusting widths & left usually
 ' Prt - Partners: same as Grp, but used for adjusting Top & Heights
 Dim Ctl As Control, Tg As String, Grp As Boolean, Prt As Boolean
 Dim ChngH    As Long, ChngW As Long
 'Static PrevH As Long, PrevW As Long

 If Frm.WindowState = vbMinimized Then Exit Function
 
 If PrevH <> 0 And PrevW <> 0 Then
    ChngH = Frm.Height - PrevH
    ChngW = Frm.Width - PrevW
 
    For Each Ctl In Frm.Controls
        Tg = LCase(Ctl.Tag)
        
        If Tg <> "" Then
            Grp = False ' defines a group
            Prt = False ' defines a group
            If InStr(Tg, "g") Then Grp = True
            If InStr(Tg, "p") Then Prt = True

            Do Until Tg = ""
                Select Case Left(Tg, 1)
                    Case "l"    ' Left
                        If Grp Then
                            Ctl.Left = Ctl.Left + (ChngW / 2)
                        Else
                            Ctl.Left = Ctl.Left + ChngW
                        End If
                    Case "t"    ' Top
                        If Prt Then
                            Ctl.Top = Ctl.Top + (ChngH / 2)
                        Else
                            Ctl.Top = Ctl.Top + ChngH
                        End If
                    Case "h"    ' Height
                        If Prt Then
                            Ctl.Height = Ctl.Height + (ChngH / 2)
                        Else
                            Ctl.Height = Ctl.Height + ChngH
                        End If
                    Case "w"    ' Width
                        If Grp Then
                            Ctl.Width = Ctl.Width + (ChngW / 2)
                        Else
                            Ctl.Width = Ctl.Width + ChngW
                        End If
                End Select
                Tg = Mid(Tg, 2)
            Loop

        End If
    Next
  End If
 
 PrevW = Frm.Width
 PrevH = Frm.Height
End Function

