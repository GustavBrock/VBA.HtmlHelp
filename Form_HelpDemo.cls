VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_HelpDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CloseHelp_Click()

    HelpControl CloseAll

End Sub

Private Sub Form_Close()

    ' Make sure the Help Viewer is closed, or Access may crash when closing.
    HelpControl CloseAll

End Sub

Private Sub OpenHelp_Click()

    ' Name of help file.
    Const FileName  As String = "HelpStudioSample.chm"

    Dim Command     As hhCommand
    Dim Context     As Long
    Dim HelpFile    As String
    
    ' Position the help file in the folder of the application.
    HelpFile = CurrentProject.Path & "\" & FileName
    
    ' Open the Help Viewer.
    Select Case Me!HelpGroup.Value
        Case 1
            Command = DisplayContents
        Case 2
            Command = DisplayIndex
        Case 3
            Command = DisplaySearch
        Case 4
            Command = OpenContext
            Context = Val(Nz(Me!ID.Value))
    End Select
    HelpControl Command, HelpFile, Context
    
End Sub
