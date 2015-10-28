# vba-outlook
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

## Create a custom context menu
```
Private Sub Application_ItemContextMenuDisplay(ByVal CommandBar As Office.CommandBar, ByVal Selection As Selection)
    Const msoControlButton = 1
    Const msoButtonIconAndCaption = 3
    Dim objButton As CommandBarButton
        If Selection.count = 1 Then
            If Selection.Item(1).Class = olMail Then
                Set objButton = CommandBar.Controls.Add(msoControlButton)
                With objButton
                    .Style = msoButtonIconAndCaption
                    .Caption = "Reply Special with Attachments"
                    .Parameter = Selection.Item(1).EntryID
                    'List of face IDs here: http://www.kebabshopblues.co.uk/2007/01/04/visual-studio-2005-tools-for-office-commandbarbutton-faceid-property/'
                    .FaceId = 355
                    .OnAction = "ReplySpecial"
                End With
            End If
        End If
End Sub
```

Put the code in `ThisOutlookSession`.
