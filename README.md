# vb-outlook
## Get recipients in selected item
```
Sub GetRecipientsInSelectedItem()
    Dim Explorer As Outlook.Explorer
    Dim Selection As Outlook.Selection
    Dim Recipients As Outlook.Recipients
    Dim Index As Integer
    
    Set Explorer = application.ActiveExplorer
    Set Selection = Explorer.Selection
    Set Recipients = Selection.Item(1).Recipients
    
    For Index = 1 To Recipients.Count
        MsgBox (Recipients.Item(Index).Name)
    Next Index
End Sub
```
