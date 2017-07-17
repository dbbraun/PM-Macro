VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PMUserForm 
   Caption         =   "Paramedic Method Tool V7.0"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9915
   OleObjectBlob   =   "PMUserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PMUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AbortButton_Click()
'
    PMUserForm_Abort_Pressed = True
    
End Sub

Private Sub ProgressBarLabel_Click()

End Sub

Private Sub ProgressLabel1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    ' Initialize UserForm Caption
    Me.Caption = "Paramedic Method Tool V7.0"
    
    ' Initialize caption below progress bar
    Me.ProgressLabel1.Caption = "Paramedic Method Analysis Starting"
    
    ' Resize the label
    Me.ProgressBarLabel.Height = 25
    Me.ProgressBarLabel.Caption = ""
    Me.ProgressBarLabel.Width = 0
    Me.ProgressBarLabel.BackColor = wdColorBlue
    
End Sub

Public Function IncrementProgressBar(sPercentComplete As Single, _
    sDescription As String)
    On Error Resume Next
    Me.ProgressLabel1.Caption = sDescription
        
    Dim lMaxProgressBarWidth As Long
    Dim sBarWidth As Single
    
    lMaxProgressBarWidth = 400
    sBarWidth = sPercentComplete * lMaxProgressBarWidth / 100
    Me.ProgressBarLabel.Width = Format(sBarWidth, "#.##")
    
    Me.Repaint
    DoEvents
    
End Function 'IncrementProgressBar

