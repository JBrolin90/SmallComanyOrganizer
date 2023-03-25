Attribute VB_Name = "mForm_Schedule"
'@Folder "Schedule"
Option Compare Database
Option Explicit

'https://www.utteraccess.com/topics/2039985

Public Sub Schedules_Form_Afterupdate()
  Form_Schedule.Form_SelectionChange
End Sub

