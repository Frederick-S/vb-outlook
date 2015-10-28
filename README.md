# vb-outlook
## Get recipients in selected item
```
Sub GetRecipientsInSelectedItem()
    Dim Explorer As Outlook.Explorer
    Dim Selection As Outlook.Selection
    Dim Recipients As Outlook.Recipients
    Dim Index As Integer
    
    Set Explorer = Application.ActiveExplorer
    Set Selection = Explorer.Selection
    Set Recipients = Selection.Item(1).Recipients
    
    For Index = 1 To Recipients.Count
        MsgBox (Recipients.Item(Index).Name)
    Next Index
End Sub
```

## Send a task
```
Sub SendTask(name)
    Dim TaskItem As Outlook.TaskItem
    Dim TaskItemRecipients As Outlook.Recipients
    Dim TaskItemRecipient As Outlook.recipient
    
    Set TaskItem = Application.CreateItem(Outlook.OlItemType.olTaskItem)
    TaskItem.Assign
    Set TaskItemRecipients = TaskItem.Recipients
    Set TaskItemRecipient = TaskItemRecipients.Add(name)
    TaskItemRecipient.Resolve
    
    If (TaskItemRecipient.Resolved) Then
        TaskItem.Subject = "Test"
        TaskItem.Body = "Body"
        TaskItem.StartDate = Now
        TaskItem.DueDate = Now + 1
        TaskItem.Display
        TaskItem.Send
    End If
End Sub
```
