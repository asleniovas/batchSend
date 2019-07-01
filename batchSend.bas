Attribute VB_Name = "Module2"
Sub batchSend()

'variable declarations
Dim YourRecipientName As String
Dim YourRecipientEmail As String
Dim YourContent As String
Dim Outlook As Object
Dim EmailObject As Outlook.MailItem
Dim rw As Range

'start looping through your column, in this case it's A
For Each rw In ActiveSheet.Range("A:A")

    'check if cell in column A is empty then exit the program
    If rw.Value = "" Then
        
        MsgBox "Loop Finished"
        
        Exit Sub
    End If
 
    'create email object
    Set EmailObject = CreateItem(olMailItem)
 
    'pull info from cells, rw.Offset(0, 1) moves 1 cell to the right from the current cell in column A
    YourRecipientEmail = rw.Offset(0, 1).Value
    YourRecipientName = rw.Value
    
    'define your content here, or in a cell if you prefer (then use something like rw.Offset(0,2).Value to fetch it)
    YourContent = "Hello " & YourRecipientName & "," & vbCrLf & vbCrLf & "Your email is " & YourRecipientEmail
        
    'work with the email object by defining the recipient, subject, attachments etc.
    With EmailObject
    
        .to = YourRecipientEmail
        .Subject = "Your Subject"
        .body = YourContent
        .display
        .Attachments.Add ("C:\Users\asleniovas\Desktop\test.txt")
        
        'you can use .Send instead of .display straight away if you're brave
        
    
    End With
 
    'next row
    Next rw


End Sub


