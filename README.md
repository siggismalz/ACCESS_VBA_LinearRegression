# ACCESS_VBA_LinearRegression
Die Bereitstellung einer neuen Klasse in VBA Access zum durchführen von Linearen-Regressionen unter Verwendung von Recordsets.

Den Code der Klasse stelle ich hier zur Verfügung. Durch die DLL, kann man die Klasse herunterladen, den Verweis "VBA_ML" im VBA Editor aktivieren und die Klasse in Access nutzen.

# Funktionen
## Die LineareRegression-Klasse bietet folgende Hauptfunktionen:

- **Properties:**
  - **XFeldName** und **YFeldName**: Ermöglichen das Festlegen der Feldnamen für die unabhängige und abhängige Variable.
  - **Steigung**, **Achsenabschnitt**, **RSquared**: Geben die berechneten Ergebnisse der linearen Regression zurück.
  - **Anzahl**: Zeigt die Anzahl der verarbeiteten Datenpunkte an.

- **Methoden:**
  - **Initialisieren**: Setzt alle internen Summen und Ergebnisse zurück.
  - **Berechne**: Führt die Regression durch, indem es ein übergebenes Recordset verarbeitet.
