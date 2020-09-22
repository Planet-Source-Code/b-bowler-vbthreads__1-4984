VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Thread Client"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Thread"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Iterator object in ThreadServer that implements IThreadable
Private WithEvents mThread As Iterator
Attribute mThread.VB_VarHelpID = -1

'Begin a Thread and display details in a frmThread form
Private Sub Command1_Click()
    Dim frmT As frmThread
    
    Set frmT = New frmThread
    
    Set mThread = New Iterator
    Set frmT.mThread = mThread
    frmT.Show
    
    mThread.StartThread
End Sub

'Halt the current thread pointed to by mThread!
Private Sub Command2_Click()
    If Not (mThread Is Nothing) Then mThread.HaltThread
End Sub

'Current thread running in frmMain has finished.
'
'IF YOU DONOT set mthread = nothing then the THREADSERVER will stay open
'even when all threads have finished execution!
'
'Object references are the most common causes of a process being open even
'when all appears to have finised
Private Sub mThread_Done()
    MsgBox "Done"
    Set mThread = Nothing
End Sub

'increment the text box
Private Sub mThread_NumberIncrement(x As Integer)
    Text1 = CStr(x)
End Sub
