VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAddInInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'This is the interface class which defines the interface for Your addin,
'You can add remove as many procedures as you like providing you change the classes
'that implement this interface
Public Sub OnConnection(pInst As ICommClass)
    '
End Sub
Public Sub ShowSelf()
    '
End Sub
Public Sub HideSelf()
    '
End Sub
