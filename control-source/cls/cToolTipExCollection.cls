VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolTipExCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCollection As New Collection

Public Sub Remove(nHwnd As Long)
    mCollection.Remove CStr(nHwnd)
End Sub

Public Sub Add(nObject As Object, nHwnd As Long)
    mCollection.Add nObject, CStr(nHwnd)
End Sub

Public Function GetCollection() As Collection
    Set GetCollection = mCollection
End Function
