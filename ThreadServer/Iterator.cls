VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Iterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Must implement IThreadable in order to call the Method (Start)
'within modThreads
Implements IThreadable

'Example events....You don't have to use events but it is good to
'some how tell the client when your thread object has finished
'....I will leave it upto on how you want to implement this!
Event Done()
Event NumberIncrement(x As Integer)

'This is used to inform the object that the client is
'requesting it to stop...Again you do not have to do it in this
'manner but this is an example
Private mStop As Boolean


'These public methods donot have to be called StartThread/HaltThread
'You can have any methods to start and stop threads
Public Sub StartThread()
    Start Me 'Calls modThread.Start with IThreadable object
End Sub

Public Sub HaltThread()
    mStop = True
End Sub

'The following methods are called from within modThreads to start/stop
'the thread
Private Sub IThreadable_Halt()
    mStop = True
End Sub

Private Sub IThreadable_Start()
    'You would place whatever code here that you want
    Dim iCount As Integer
    
    For iCount = 1 To 4000
        DoEvents
        RaiseEvent NumberIncrement(iCount)
        If (mStop) Then Exit For
    Next
    
    RaiseEvent Done
End Sub
