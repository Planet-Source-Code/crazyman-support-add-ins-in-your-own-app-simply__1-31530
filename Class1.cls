VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IAddInInterface ' Must Implement the expected interface for ALL Addins
Private m_Inst As ICommClass 'Hold private scoped instance of the caller object
                     
Dim f As New Form1

Private Sub IAddInInterface_HideSelf()
    '
    f.Hide
End Sub

Private Sub IAddInInterface_OnConnection(pInst As vbCustomAddin.ICommClass)
    '
    pInst.AddinStatusMessage Me, "Add 1 Loaded Okay"
    pInst.RequestOperation Me, "ENABLE_BUTTONS"
    Set m_Inst = pInst 'Hold our caller object instance for later use
End Sub

Private Sub IAddInInterface_ShowSelf()
    Set f.OwnerInterface = m_Inst 'pass on our caller instance
    Set f.OwnerClass = Me 'pass ourself
    f.Show
End Sub
