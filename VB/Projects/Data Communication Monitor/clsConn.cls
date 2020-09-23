VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SourceIP     As String
Public SourcePort   As String
Public DestIP       As String
Public DestPort     As String
Public First        As Boolean
Public cType        As eProtocols

Public Log          As String

Public SockINindex  As Integer
Public SockOUTindex As Integer

Public Send         As String
Public Recieved     As String
