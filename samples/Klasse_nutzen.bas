Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim lr As LineareRegression
Dim sqlAbfrage As String

sqlAbfrage = "SELECT XWert, YWert FROM TestDaten"
Set db = CurrentDb
Set rs = db.OpenRecordset(sqlAbfrage, dbOpenSnapshot)

' Initialisiere die LineareRegression-Klasse
Set lr = New LineareRegression

lr.XFeldName = "XWert"
lr.YFeldName = "YWert"

lr.Berechne rs


' Zeige die Ergebnisse an
MsgBox "Ergebnisse der Linearen Regression:" & vbCrLf & _
       "Steigung: " & Format(lr.Steigung, "0.0000") & vbCrLf & _
       "Achsenabschnitt: " & Format(lr.Achsenabschnitt, "0.0000") & vbCrLf & _
       "R²: " & Format(lr.RSquared, "0.0000"), vbInformation, "Regressionsergebnisse"

Bereinigung:
If Not rs Is Nothing Then
    rs.Close
    Set rs = Nothing
End If
Set db = Nothing
Set lr = Nothing
