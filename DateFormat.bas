Attribute VB_Name = "DateFormat"
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long


'Call this function to check the date format and change it if necessary
Public Sub ChangeDateFormat()
    Dim sReturn As String
    Dim r As Long
    Dim LCID As Long
    LCID = GetSystemDefaultLCID()
    r = GetLocaleInfo(LCID, LOCALE_SSHORTDATE, sReturn, Len(sReturn))
    If r Then
     'pad the buffer with spaces to create the size of memory buffer
      sReturn = Space$(r)
     'and call again passing the buffer
      r = GetLocaleInfo(LCID, LOCALE_SSHORTDATE, sReturn, Len(sReturn))
     'if successful (r > 0)
      If r Then
        'r holds the size of the string
        'including the terminating null
        If Left$(sReturn, r - 1) <> "dd/MM/yyyy" Then
            Call SetLocaleInfo(LCID, LOCALE_SSHORTDATE, "dd/MM/yyyy")
        End If
      End If
   End If
End Sub
