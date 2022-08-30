Option Explicit

Sub remove_reference()

Dim vbaProject As Object ' Current VBA project.
Dim objReference As Object ' Counting variable for each reference.
Dim strReference As String

Set vbaProject = ActiveDocument.VBProject ' Refer to the current document's VBA project.
strReference = "Sample_Reference" ' Set the reference's name

' Loop over "References" dialog box and delete the previously defined reference
For Each objReference In vbaProject.References
   If objReference.Name = strReference Then vbaProject.References.Remove objReference
Next

End Sub
