Attribute VB_Name = "Test"
' This sub sends e-mail to awesome statistic team that execution is over:
' (http://www.rondebruin.nl/win/winmail/Outlook/tips.htm=
Sub SendEmailWithLink()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim EmailBody As String
    Dim EmailSubject As String
    Dim PathToSwedWin As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    EmailSubject = "Statistic reports (" & Right(WEEK_SUFIX, 5) & ") are ready"
    PathToSwedWin = "\\kihome3.rnd.ki.sw.ericsson.se\volh16001_1\eantest\"
    
    'file:///W:/Statistika/

    EmailBody = "Wassup, statistic team. <br><br>" & _
              "<B><I>Generation of data for week: " & Right(WEEK_SUFIX, 4) & " is done and data is ready for publishing.</I></B><br>" & _
              "<A HREF=""file://" & PathToSwedWin & """>Results</A>" & _
              "<br><br>Sincerely yours,<br>" & _
              "MondayStatisticsTool."
              'ThisWorkbook.Path & "\" & "Results\"
    On Error Resume Next
    With OutMail
        .to = "antea.stojic@ericsson.com"
        .CC = ""
        .BCC = ""
        .Subject = EmailSubject
        .HtmlBody = EmailBody
        .Display
        .Importance = 2
        .Send   'or use .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

