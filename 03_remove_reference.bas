Option Explicit

Sub remove_reference()

Dim vbaProject As Object ' Current VBA project.
Dim objReference As Object ' Counting variable for each reference.
Dim strReference As String

  ' Refer to the current document's VBA project.
Set vbaProject = ActiveDocument.VBProject
strReference = "Sample_Reference"

' Loop over "References" dialog box.
For Each objReference In vbaProject.References
   ' Delete the previously defined reference.
   If objReference.Name = strReference Then vbaProject.References.Remove objReference
Next

End Sub
