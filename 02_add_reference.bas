Option Explicit

Sub add_reference()

Dim vbProject as Object

Set vbProject = ActiveDocument.VBProject
vbProject.References.AddFromFile "C:\temp\reference_use.dot"

End Sub
