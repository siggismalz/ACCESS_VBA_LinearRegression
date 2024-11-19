# ACCESS_VBA_LinearRegression

Die Bereitstellung einer neuen Klasse in VBA Access zum Durchführen von linearen Regressionen unter Verwendung von Recordsets.

Den Code der Klasse stelle ich hier zur Verfügung. Durch die DLL kann man die Klasse herunterladen, den Verweis "VBA_ML" im VBA-Editor aktivieren und die Klasse in Access nutzen.

## Funktionen

### Die LineareRegression-Klasse bietet folgende Hauptfunktionen:

#### Properties:

- **XFeldName (Property)**
  - **Beschreibung:** Legt den Namen des X-Feldes (unabhängige Variable) fest.
  - **Typ:** `String`
  - **Beispiel:**
    ```vba
    lr.XFeldName = "XWert"
    ```

- **YFeldName (Property)**
  - **Beschreibung:** Legt den Namen des Y-Feldes (abhängige Variable) fest.
  - **Typ:** `String`
  - **Beispiel:**
    ```vba
    lr.YFeldName = "YWert"
    ```

- **Steigung (Property)**
  - **Beschreibung:** Gibt die berechnete Steigung der Regressionslinie zurück.
  - **Typ:** `Double`
  - **Beispiel:**
    ```vba
    Dim slope As Double
    slope = lr.Steigung
    ```

- **Achsenabschnitt (Property)**
  - **Beschreibung:** Gibt den berechneten Achsenabschnitt der Regressionslinie zurück.
  - **Typ:** `Double`
  - **Beispiel:**
    ```vba
    Dim intercept As Double
    intercept = lr.Achsenabschnitt
    ```

- **RSquared (Property)**
  - **Beschreibung:** Gibt das Bestimmtheitsmaß \( R² \) der Regression zurück.
  - **Typ:** `Double`
  - **Beispiel:**
    ```vba
    Dim rSquared As Double
    rSquared = lr.RSquared
    ```

- **Anzahl (Property)**
  - **Beschreibung:** Gibt die Anzahl der verarbeiteten Datenpunkte zurück.
  - **Typ:** `Long`
  - **Beispiel:**
    ```vba
    Dim count As Long
    count = lr.Anzahl
    ```

#### Methoden:

- **Initialisieren (Methode)**
  - **Beschreibung:** Setzt alle internen Summen und Ergebnisse auf Null. Sollte vor der Durchführung einer neuen Regression aufgerufen werden.
  - **Parameter:** Keine
  - **Beispiel:**
    ```vba
    lr.Initialisieren
    ```

- **Berechne (Methode)**
  - **Beschreibung:** Führt die lineare Regression anhand eines übergebenen Recordsets durch. Berechnet Steigung, Achsenabschnitt und \( R² \).
  - **Parameter:** `rs` (DAO.Recordset) – Das Recordset mit den Datenpunkten.
  - **Beispiel:**
    ```vba
    lr.Berechne rs
    ```
