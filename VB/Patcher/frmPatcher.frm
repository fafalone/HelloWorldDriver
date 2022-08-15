VERSION 5.00
Begin VB.Form frmPatcher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete msvbvm by the trick"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3255
   Icon            =   "frmPatcher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPatch 
      Caption         =   "Patch"
      Height          =   390
      Left            =   1620
      TabIndex        =   2
      Top             =   525
      Width           =   1185
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   390
      Left            =   465
      TabIndex        =   1
      Top             =   525
      Width           =   1140
   End
   Begin VB.TextBox txtFile 
      Height          =   345
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   2970
   End
End
Attribute VB_Name = "frmPatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // frmPatcher.frm - main form of Patcher application
' // © Krivous Anatoly Anatolevich (The trick), 2014

Option Explicit

Private Sub cmdBrowse_Click()
    Dim fName As String
    
    fName = GetFile(Me.hwnd)
    If Len(fName) Then txtFile.Text = fName
    
End Sub

Private Sub cmdPatch_Click()

    If Not RemoveRuntimeFromIAT(txtFile.Text) Then
        MsgBox "Error occurs", vbCritical
    End If
    
End Sub

