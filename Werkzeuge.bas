Attribute VB_Name = "Werkzeuge"

Public Function getColumn(headline As Integer, key As String) As Integer
    ' *****************************************************************
    ' Diese Funktion nimmt als Input einen Spaltennamen (key) und
    ' eine Suchzeilennummer (headline) und gibt die Spalte zurück
    ' in der der Name zu finden ist. Die Operation wird im ActiveSheet
    ' ausgeführt.
    ' *****************************************************************

    Dim Ende As Integer

    Ende = ActiveSheet.UsedRange.Columns.Count

    For spalte = 1 To Ende
        If Cells(headline, spalte).Value = key Then
            getColumn = spalte
        End If
    Next spalte

End Function

Public Function tableGetNextRow(searchColumn As Integer) As Integer
    ' ******************************************************************
    ' Diese Funktion sucht ein Sheet von unten durch und gibt die erste
    ' freie Zeile in der searchColumn zurück. Operation wird im
    ' ActiveSheet ausgeführt.
    ' ******************************************************************
    
    Dim row As Integer
    
    For row = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
        If Not IsEmpty(ActiveSheet.Cells(row, searchColumn)) Then
            tableGetNextRow = row + 1
            Exit Function
        End If
    Next row
End Function

Public Sub copyLine(wksSource As Worksheet, wksTarget As Worksheet, rowSource As Integer, rowTarget As Integer)
    
    ' ******************************************************************
    ' Dieser Sub kopiert die angegebene Source Row in die angegebene
    ' Target Row. Das aktive WorkSheet wird dabei u.U. beeinträchtigt.
    ' ******************************************************************
    Application.ScreenUpdating = False 'sonst flackert der Bildschirm
    
    wksSource.Activate
    wksSource.Range(Cells(rowSource, 1), Cells(rowSource, wksSource.UsedRange.Columns.Count)).Copy
    wksTarget.Activate
    wksTarget.Cells(rowTarget, 1).Select
    wksTarget.Paste
    
End Sub

Public Sub addNewElem(liste As Variant, elem As Variant)
    If Not (liste.Contains(elem)) Then
        liste.Add elem
    End If
End Sub

Public Function checkoutDict(key As Variant, val As Variant, _
                             dict As Object) As Boolean

    ' ******************************************************
    ' Dieser Sub untersucht ein Dict mit List values auf die
    ' Anwesenheit eines Wertes
    ' ******************************************************
    
    If dict.Exists(key) Then
        checkoutDict = dict(key).Contains(val)
      Else
        checkoutDict = False
    End If

End Function

Public Function findWs(name As String, wb As Workbook) As Integer

    ' ******************************************************
    ' Dieser Sub sucht im übergebenen Workbook nach dem
    ' Sheet mit dem übergebenen Namen und gibt dessen Position
    ' zurück. Sollte es nicht existieren: return 0.
    ' ******************************************************
    Dim i As Integer
    
    findWs = 0
    
    For i = 1 To wb.Sheets.Count
        If wb.Sheets(i).name = name Then
            findWs = i
        End If
    Next i
    
End Function

Public Sub sendMail(empfaenger As String, betreff As String, textkoerper As String, sendDirectly As Boolean, objOutlook As Object)

' **************************************************************
' Dieser Sub dient zum Versenden von Emails.
' Der Parameter sendDirectly entscheidet, ob die Mail
' direkt oder manuell versendet wird.
' **************************************************************

Dim objMail As Object

Set objMail = objOutlook.CreateItem(0)

With objMail
    .To = empfaenger
    .Subject = betreff
    .body = textkoerper
    If Not sendDirectly Then
        .Display
      Else
        .Send
    End If
    
End With
    
End Sub

Public Sub sendMailHTML(empfaenger As String, betreff As String, textkoerper As String, sendDirectly As Boolean, objOutlook As Object)

' **************************************************************
' Dieser Sub dient zum Versenden von Emails.
' Der Parameter sendDirectly entscheidet, ob die Mail
' direkt oder manuell versendet wird.
' **************************************************************

Dim objMail As Object

Set objMail = objOutlook.CreateItem(0)

With objMail
    .To = empfaenger
    .Subject = betreff
    .HTMLBody = textkoerper
    If Not sendDirectly Then
        .Display
      Else
        .Send
    End If
    
End With
    
End Sub

Public Function getValueByName(nameKey As String, column As Integer, ws As Worksheet) As Variant

Dim Vorname As String, Nachname As String
Dim spalteNachname As Integer, spalteVorname As Integer

spalteVorname = getColumn(1, "Vorname")
spalteNachname = getColumn(1, "Nachname")

Nachname = Split(nameKey, ";")(0)
Vorname = Split(nameKey, ";")(1)

Dim zeile As Integer

For zeile = 2 To ws.UsedRange.Rows.Count
    If ((ws.Cells(zeile, spalteNachname).Value = Nachname) And (ws.Cells(zeile, spalteVorname).Value = Vorname)) Then
        getValueByName = Cells(zeile, column).Value
        Exit Function
    End If
Next zeile

MsgBox (Vorname + " " + Nachname + " wurde unter den MailAdressen nicht gefunden.")
    
getValueByName = False

End Function
