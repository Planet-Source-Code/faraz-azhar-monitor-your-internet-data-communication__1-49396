Attribute VB_Name = "modFunctions"

' PLEASE CHANGE THESE VALUES ACCORDING TO YOUR NEEDS
Public Const HTTP_LISTEN_PORT = 8000        ' these are for listening.
Public Const SOCKS_LISTEN_PORT = 1100

Public Const PROXY_HTTP_IP = "192.168.0.1"  ' these are for further connecting to.
Public Const PROXY_HTTP_PORT = 8080
Public Const PROXY_SOCKS_IP = "192.168.0.1"
Public Const PROXY_SOCKS_PORT = 1080
'

Public Enum eProtocols
    cSOCKS = 0
    cHTTP = 1
End Enum

Public Conns    As New clsConns

Public Sub Status(Txt)
 frmMain.stBar.SimpleText = Txt
End Sub

Public Sub AddToLog(vKey As String, vData As String)
 Conns(vKey).Log = Conns(vKey).Log & vData & vbCrLf
End Sub
