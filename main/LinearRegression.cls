
Class LineareRegression
' Private Variablen zur Speicherung der Feldnamen und laufenden Summen
Private pXFeldName As String
Private pYFeldName As String

' Laufende Summen für die Berechnung
Private pAnzahl As Long
Private pSummeX As Double
Private pSummeY As Double
Private pSummeXY As Double
Private pSummeXQuadrat As Double
Private pSummeYQuadrat As Double

' Ergebnisse der Regression
Private pSteigung As Double
Private pAchsenabschnitt As Double
Private pRSquared As Double

' Property zum Festlegen des X-Feldnamens
Public Property Let XFeldName(ByVal Wert As String)
    pXFeldName = Wert
End Property

Public Property Get XFeldName() As String
    XFeldName = pXFeldName
End Property

' Property zum Festlegen des Y-Feldnamens
Public Property Let YFeldName(ByVal Wert As String)
    pYFeldName = Wert
End Property

Public Property Get YFeldName() As String
    YFeldName = pYFeldName
End Property

' Properties zum Abrufen der Ergebnisse
Public Property Get Steigung() As Double
    Steigung = pSteigung
End Property

Public Property Get Achsenabschnitt() As Double
    Achsenabschnitt = pAchsenabschnitt
End Property

Public Property Get RSquared() As Double
    RSquared = pRSquared
End Property

Public Sub Initialisieren()
    pAnzahl = 0
    pSummeX = 0
    pSummeY = 0
    pSummeXY = 0
    pSummeXQuadrat = 0
    pSummeYQuadrat = 0
    pSteigung = 0
    pAchsenabschnitt = 0
    pRSquared = 0
End Sub

Public Sub Berechne(rs As DAO.Recordset)
    Dim xWert As Variant
    Dim yWert As Variant

    Initialisieren

    If rs.EOF And rs.BOF Then
        Debug.Print "Das Recordset ist leer.", vbExclamation
        Exit Sub
    End If

    rs.MoveFirst
    Do While Not rs.EOF
        xWert = rs.Fields(pXFeldName).Value
        yWert = rs.Fields(pYFeldName).Value
    
        ' Überspringe Datensätze mit Nullwerten
        If Not IsNull(xWert) And Not IsNull(yWert) Then
            pAnzahl = pAnzahl + 1
            pSummeX = pSummeX + xWert
            pSummeY = pSummeY + yWert
            pSummeXY = pSummeXY + (xWert * yWert)
            pSummeXQuadrat = pSummeXQuadrat + (xWert ^ 2)
            pSummeYQuadrat = pSummeYQuadrat + (yWert ^ 2)
        End If
    
        rs.MoveNext
    Loop

    If pAnzahl < 2 Then
        Debug.Print "Nicht genügend Datenpunkte für die Regression.", vbExclamation
        Exit Sub
    End If

    ' Berechne Steigung und Achsenabschnitt
    pSteigung = (pAnzahl * pSummeXY - pSummeX * pSummeY) / (pAnzahl * pSummeXQuadrat - pSummeX ^ 2)
    pAchsenabschnitt = (pSummeY - pSteigung * pSummeX) / pAnzahl

    Dim zaehler As Double
    Dim nenner As Double
    zaehler = (pAnzahl * pSummeXY - pSummeX * pSummeY) ^ 2
    nenner = (pAnzahl * pSummeXQuadrat - pSummeX ^ 2) * (pAnzahl * pSummeYQuadrat - pSummeY ^ 2)
    If nenner <> 0 Then
        pRSquared = zaehler / nenner
    Else
        pRSquared = 0
    End If
End Sub
    
End Class
