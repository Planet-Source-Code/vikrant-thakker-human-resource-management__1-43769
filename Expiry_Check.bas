Attribute VB_Name = "Expiry_Check"
Dim Disk As String
Public Sub ExpiryCheck()
Disk = "C:\"
'*****************************************************************************************************************
' This will set the expiry date of the software when it is installed for 1st time
' At the time of Installation
If rsExpiry.RecordCount = 0 Then
    rsExpiry.AddNew
    If rsExpiry!usage = 0 Then
        rsExpiry!InstallDate = Date
        rsExpiry!expirydate = DateAdd("d", 30, Date)
        rsExpiry!Paid = "N"
        rsExpiry!CurrentYear = Date
        rsExpiry!YearChanged = "N"
        rsExpiry!HDSerial = VolumeSerialNumber(Disk)
    End If
    rsExpiry.Update
End If


If rsExpiry!HDSerial <> VolumeSerialNumber(Disk) Then
    MsgBox "Contact Developer for Re-installing the Software !                                             Developed by: AnaSys Softwares (Erode), Geetha Thakkar,                                      Ph : 7911226, 7911833, Email : geethathakker@hotmail.com", vbCritical, "Developer"
    End
End If

    rsExpiry!usage = rsExpiry!usage + 1
    rsExpiry.Update


'******************************************************************************************************************
' If user has not paid then This will check for the expiry date and usage, everytime the program is runned...
' After installation
If rsExpiry!Paid = "N" Then
    If rsExpiry!expirydate = Date Or rsExpiry!expirydate < Date Or rsExpiry!usage > 100 Then
        rsExpiry!accexpired = True
        rsExpiry.Update
    End If
End If

'******************************************************************************************************************
' After Trial version has expired
If rsExpiry!Paid = "N" Then

    If rsExpiry!accexpired = True Then
             MsgBox "Kindly Pay the negotiated amount for further usage of this Software !                          Developed by: AnaSys Softwares (Erode), Geetha Thakkar, Ph : 256631, 256031, Email : geethathakker@hotmail.com", vbCritical, "Developer"
        End
    End If
End If

If Year(Now) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "OASYS"
    rsExpiry!YearChanged = "Y"
    rsExpiry.Update
End
ElseIf Year(Now) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "OASYS"
    rsExpiry!YearChanged = "Y"
    rsExpiry.Update
End
End If

' Once the year gets changed then even if the user changes the year of his computer
'manually, he should not be allowed to use the software
' The software can be used only after the developer writes the status of the YearChanged="N" in the database
' Also the developer needs to write the new currentyear

If rsExpiry!YearChanged = "Y" Then
    MsgBox "You cannot work in changed working year ! Contact Developer !", vbCritical, "OASYS"
    End
End If

'If Year(Now) <> rsExpiry!CurrentYear Then
'    MsgBox "Please call developer to run the software in Changed Year"
'    End
'End If


'rsExpiry!expirydate = DateAdd("d", 30, Date)

End Sub
