VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "AddIn Test APP"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Hide Addin 2 Form"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hide Addin 1 Form"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   5175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show Addin 2 Form"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show Addin 1 Form"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Addin2"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Addin1"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ICommClass 'Implement the ICommClass Interface
                      'This Allows you to pass Any object to the OnConnection
                      'Since this object will Support the expected interface [ICommClass]
                      
Dim oInterface1 As IAddInInterface
Dim oInterface2 As IAddInInterface
Private Sub Command1_Click()
    Set oInterface1 = CreateObject("TestAddin.Class1")
    oInterface1.OnConnection Me 'If you dont like exposing your form
                               'Create another class supporting ICommClass
                               'and pass the events through that class - and raise them to your form if required
End Sub

Private Sub Command2_Click()
    Set oInterface2 = CreateObject("TestAddin.Class2")
    oInterface2.OnConnection Me 'If you dont like exposing your form
                               'Create another class supporting ICommClass
                               'and pass the events through that class - and raise them to your form if required
End Sub




Private Sub Command3_Click()
    oInterface1.ShowSelf
    Command5.Enabled = True
End Sub

Private Sub Command4_Click()
    oInterface2.ShowSelf
    Command6.Enabled = True
End Sub

Private Sub Command5_Click()
    oInterface1.HideSelf
End Sub

Private Sub Command6_Click()
    oInterface2.HideSelf
End Sub

Private Sub ICommClass_AddinStatusMessage(pInst As vbCustomAddin.IAddInInterface, sMessage As String)
    Label1.Caption = sMessage
End Sub

Private Sub ICommClass_RequestOperation(pInst As vbCustomAddin.IAddInInterface, Operation As String)
    Text1 = Text1 & "Operation " & Operation & " Requested" & vbCrLf
    If pInst Is oInterface1 Then
        Command3.Enabled = True
    Else
        Command4.Enabled = True
    End If
End Sub
