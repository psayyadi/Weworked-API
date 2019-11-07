Attribute VB_Name = "Module2"
Option Explicit

Sub add_tasks()
'This subroutine automates adding the tasks to a Weworked project.
'Input: excel file containing all the tasks, their codes and their budget

Dim ie As Object
Dim webpage As Object
Dim i As Integer

Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True 'Optional if you want to make the window visible

'open login window
ie.navigate ("https://www.weworked.com/app/login_new")
Do While ie.readyState = 4: DoEvents: Loop
Do Until ie.readyState = 4: DoEvents: Loop
While ie.Busy
    DoEvents
Wend
Set webpage = ie.document

On Error Resume Next
'logging into weworked.com
webpage.getelementbyID("txtEmail").Value = "your email" 'username
webpage.getelementbyID("txtPassword").Value = "your password" 'password
webpage.getelementbyID("btnLogIn").Click

'click settings
ie.navigate ("https://www.weworked.com/app/usertimesheet.php")

Do While ie.readyState = 4: DoEvents: Loop
Do Until ie.readyState = 4: DoEvents: Loop
While ie.Busy
    DoEvents
Wend
Set webpage = ie.document
webpage.getelementbyID("btnAccountSettings").Click

'navigating to test project: this URL should be copied from the project page.
ie.navigate ("https://www.weworked.com/app/manage.php?p=manage_project.php&projectid=181473")

Do While ie.readyState = 4: DoEvents: Loop
Do Until ie.readyState = 4: DoEvents: Loop
While ie.Busy
    DoEvents
Wend
Set webpage = ie.document

'looping through tasks
For i = 2 To 100 'or how many tasks you have
    
    webpage.getelementbyID("btnAddNewTask").Click
    'adding task name
    webpage.getelementbyID("txtTaskName").Value = Range("A" & i) 'column A contains the task names
    'adding task code
    webpage.getelementbyID("txtTaskCode").Value = Range("B" & i) 'column B contains the task codes
    'adding budget in hours
    If VarType(Range("J" & i)) <> 8 Then
        webpage.getelementbyID("txtTaskBudgetHours").Value = Range("C" & i) 'column C contains the budgets
    End If
    webpage.getelementbyID("btnAddTask").Click

Next i

MsgBox "Finished adding the tasks"

ie.Quit
Set ie = Nothing
End Sub
