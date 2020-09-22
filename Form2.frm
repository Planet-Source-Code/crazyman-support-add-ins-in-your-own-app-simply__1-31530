VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Addin 2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "Addin 2 : TYPE REQUEST HERE"
      Top             =   960
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Request to Main App"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "THIS IS A FORM FROM ADDIN 2"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Form2"
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

