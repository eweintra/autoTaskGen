VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim strBody As String
    Dim strSubject As String
    Dim arrLines As Variant
    Dim varLine As Variant
    Dim intPos As Integer
    Dim objTask As Outlook.TaskItem
    Dim strDayOfWeek As String
    Dim strWW As String
    Dim intWW As Integer
    Dim dueDate As Date
    Dim blnARFound As Boolean
    Dim strNewText As String
    Dim strOldText As String
    Dim intStartOfNewText As Integer

    If TypeOf Item Is Outlook.MailItem Then
        strBody = Item.Body
        strSubject = Item.Subject
        blnARFound = False
        
        ' Identify the start of the new text in the email body
        intStartOfNewText = InStr(1, strBody, "From: ") ' Adjust this based on the email client's thread marker
        If intStartOfNewText > 0 Then
            strNewText = Left(strBody, intStartOfNewText - 1)
            strOldText = Mid(strBody, intStartOfNewText)
        Else
            strNewText = strBody ' If no thread marker is found, assume the entire body is new
            strOldText = ""
        End If
    
        ' Check if "AR" is in the subject or body
        If InStr(1, strSubject, "AR") > 0 Or InStr(1, strNewText, "AR") > 0 Then
            blnARFound = True
            arrLines = Split(strNewText, vbCrLf) ' Split the body into lines
            For Each varLine In arrLines
                intPos = InStr(1, varLine, "AR") ' Find position of "AR" in the line
                If intPos > 0 Then
                    ' Extract the work week number from the line
                    strWW = Mid(varLine, InStr(UCase(varLine), "WW") + 2, 2)
                    If IsNumeric(strWW) Then
                        intWW = CInt(strWW)
                        ' Calculate the due date as the Tuesday of the given work week
                        ' dueDate = GetTuesdayOfWW(intWW)
                        ' Extract the day of the week from the line
                        strDayOfWeek = Trim(Mid(varLine, InStr(varLine, ".") + 1))
                        ' Calculate the due date based on the work week and day of the week
                        dueDate = CalculateDateFromWWAndDay(strDayOfWeek, intWW)
                        Set objTask = Application.CreateItem(olTaskItem)
                        With objTask
                            .Subject = "AR:" & Mid(varLine, intPos + 2) & vbCrLf & " Subject:" & strSubject ' Set subject to string following "AR"
                            .Body = "Subject: " & strSubject & vbCrLf & "AR Details: " & Mid(varLine, intPos + 2) & vbCrLf & "Body: " & strNewText ' Include email subject and details following "AR"
                            .dueDate = dueDate ' Set due date to the calculated Tuesday
                            .Save
                        End With
                     End If
                End If
            Next varLine
        End If
        If blnARFound Then
            Item.Body = strNewText & "*** Automatic tasks were generated to follow the ARs ***" & vbCrLf & strOldText & vbCrLf & vbCrLf
        End If
        ' Continue sending the mail
        Cancel = False
    End If
End Sub


' Function to calculate the date based on the work week and day of the week
Function CalculateDateFromWWAndDay(strDayOfWeek As String, intWW As Integer) As Date
    ' Dim intWW As Integer
    Dim yearStart As Date
    Dim firstDayOfYear As Date
    Dim dayOfWeekDate As Date
    Dim daysToAdd As Integer
    
    ' Extract the work week number from the string
    ' intWW = CInt(Trim(Split(strDayOfWeek, ".")(0)))
    ' Get the first day of the year
    yearStart = DateSerial(Year(Now), 1, 1)
    ' Find the first day of the specified work week
    firstDayOfYear = yearStart + (intWW - 1) * 7 - Weekday(yearStart, vbSunday)
    
    ' Determine the number of days to add based on the day of the week
    Select Case UCase(strDayOfWeek) ' UCase(Trim(Split(strDayOfWeek, ".")(1)))
        Case "MONDAY"
            daysToAdd = 2
        Case "TUESDAY"
            daysToAdd = 3
        Case "WEDNESDAY"
            daysToAdd = 4
        Case "THURSDAY"
            daysToAdd = 5
        Case "FRIDAY"
            daysToAdd = 6
        Case "SATURDAY"
            daysToAdd = 7
        Case "SUNDAY"
            daysToAdd = 1
    End Select
    
    ' Calculate the date for the specified day of the week
    dayOfWeekDate = firstDayOfYear + daysToAdd ' - 1
    CalculateDateFromWWAndDay = dayOfWeekDate
End Function

