VERSION 5.00
Begin VB.Form frmThread 
   Caption         =   "Thread"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1890
   LinkTopic       =   "Form1"
   ScaleHeight     =   960
   ScaleWidth      =   1890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents mThread As Iterator
Attribute mThread.VB_VarHelpID = -1

'Unload this form since the thread has finished
'
'Once again if you did not unload this form then the ThreadServer would stay open
'with a reference to the Iterator object pointed to by (mThread)
'
'We donot have to set the mThread to nothing here since this will happen
'when the Form is unloaded
Private Sub mThread_Done()
    Unload Me
End Sub

'Update text box
Private Sub mThread_NumberIncrement(x As Integer)
    Text1 = CStr(x)
End Sub

