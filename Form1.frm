VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Addin 1"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Text            =   "ADDIN 1 : TYPE REQUEST HERE"
      Top             =   600
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Request to Main App"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "THIS IS A FORM FROM ADDIN 1"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OwnerInterface As ICommClass
Public OwnerClass As IAddInInterface
Private Sub Command1_Click()
    OwnerInterface.RequestOperation OwnerClass, Text1.Text
End Sub
