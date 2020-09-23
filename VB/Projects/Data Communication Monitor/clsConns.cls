VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mcolConns   As Collection

Private Sub Class_Initialize()
 Set mcolConns = New Collection
End Sub

' Read-only Count property.
Public Property Get Count() As Integer
   Count = mcolConns.Count
End Property

' Add method for creating new Spoke objects.
Public Function Add(SourceIP As String, SourcePort As String, First As Boolean, cType As eProtocols) As clsConn

   Dim NewConn As New clsConn
   With NewConn
        .SourceIP = SourceIP
        .SourcePort = SourcePort
        .First = First
        .cType = cType
   End With

   mcolConns.Add NewConn, (SourceIP & ":" & SourcePort)
   Set Add = NewConn
End Function

Public Function Clear()
 Set mcolConns = New Collection
End Function

Public Property Get Item(Index) As clsConn
Attribute Item.VB_UserMemId = 0
  Set Item = mcolConns(Index)
End Property

Public Function Remove(Index)
  On Error Resume Next          ' It caused a problem here. I dont know why!! ?? !
  mcolConns.Remove Index
End Function

Public Function CheckIfExist(IPandPort) As Boolean
 Dim dumbConn As clsConn
 For Each dumbConn In mcolConns
    If IPandPort = dumbConn.SourceIP & ":" & dumbConn.SourcePort Then
        CheckIfExist = True
        Exit Function
    End If
 Next
End Function
