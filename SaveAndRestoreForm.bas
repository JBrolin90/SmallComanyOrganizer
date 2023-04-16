Attribute VB_Name = "SaveAndRestoreForms"
'@Folder "HEMS"

public sub SaveForms()
    Foreach form in Application.CurrentProject.AllForms
        Debug.Print "Saving form " & form.Name
        Application.SaveAsText acForm form.Name form.Name & ".form"
    Next
End Sub

