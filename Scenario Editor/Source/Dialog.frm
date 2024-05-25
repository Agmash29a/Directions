VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scenario Details"
   ClientHeight    =   4170
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4815
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Initial Question"
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4575
      Begin VB.TextBox FirstQuestion 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scenario Description"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
      Begin VB.TextBox ScenarioTitle 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()

    Dialog.Visible = False

End Sub

Private Sub OKButton_Click()

    FirstQuestionVar = Dialog.FirstQuestion.Text
    ScenarioTitleVar = Dialog.ScenarioTitle.Text
    
    Dialog.Visible = False

End Sub
