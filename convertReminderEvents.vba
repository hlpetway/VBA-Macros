Sub CreateCalendarEvents()
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olFolder As Object
    Dim regexPattern As String
    Dim regEx As Object
    Dim olItems As Object
    Dim olItem As Object
    Dim olCalendar As Object
    Dim olAppointment As Object
       ' Create Outlook application
    Set olApp = CreateObject("Outlook.Application")
        ' Get the default calendar folder
    Set olCalendar = olNamespace.GetDefaultFolder(9)  ' 9 corresponds to the Calendar
    Set olNamespace = olApp.GetNamespace("MAPI")
       ' Specify the folder to search (e.g., Inbox)
    Set olFolder = olNamespace.GetDefaultFolder(6)  ' 6 corresponds to the Inbox
    ' Get all items in the folder
    Set olItems = olFolder.Items
    ' Loop through each email
    For Each olItem In olItems
        ' first find emails from Ajera
        If InStr(1, olItem.Subject, "Ajera Alert - Project Assignment", vbTextCompare) > 0 Then
        ' Check if the email contains a date
        If InStr(1, olItem.Body, "End Date", vbTextCompare) > 0 Then
        For Each line In olItem.Body
            
            ' Extract the date from the email body (adjust the logic based on your email format)
            ' For example, you might use Regular Expressions to extract the date.
               ' Regex pattern for matching dates in the format MM/DD/YYYY
            regexPattern = "(\b\d{1,2}/\d{1,2}/\d{4}\b)"
    
            ' Create a RegExp object
            Set regEx = CreateObject("VBScript.RegExp")
            
            ' Set the pattern
            regEx.Pattern = regexPattern
            ' Create a new appointment item
            Set olAppointment = olCalendar.Items.Add
            ' Customize the appointment details
            With olAppointment
                .Subject = "Meeting Title"
                .Start = Now + TimeValue("09:00:00")  ' Set the start time
                .End = Now + TimeValue("10:00:00")    ' Set the end time
                .Location = "Meeting Room 123"
                .Body = "Discussion on important matters."
                ' You can add more properties or customize as needed
            End With
            ' Save and display the appointment
            olAppointment.Save
            olAppointment.Display
        End If
    Next olItem
    ' Clean up
    Set olApp = Nothing
    Set olNamespace = Nothing
    Set olFolder = Nothing
    Set olItems = Nothing
    Set olItem = Nothing
End Sub
