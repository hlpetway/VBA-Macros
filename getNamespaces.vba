    ' Loop through available namespaces
    For Each namespace In outlookApp.GetNamespace("MAPI").Folders
        ' Print the name of each namespace
        Debug.Print "Namespace: " & namespace.Name
    Next namespace
