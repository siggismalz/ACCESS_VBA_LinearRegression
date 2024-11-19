Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim fld As DAO.Field
Dim SQL As String

Set db = CurrentDb

On Error Resume Next
db.TableDefs.Delete "TestDaten"
db.TableDefs.Refresh
On Error GoTo 0

Set tdf = db.CreateTableDef("TestDaten")

With tdf
    .Fields.Append .CreateField("XWert", dbDouble)
    .Fields.Append .CreateField("YWert", dbDouble)
End With

db.TableDefs.Append tdf

SQL = "INSERT INTO TestDaten (XWert, YWert) VALUES (1, 2)"
db.Execute SQL, dbFailOnError
SQL = "INSERT INTO TestDaten (XWert, YWert) VALUES (2, 4)"
db.Execute SQL, dbFailOnError
SQL = "INSERT INTO TestDaten (XWert, YWert) VALUES (3, 6)"
db.Execute SQL, dbFailOnError
SQL = "INSERT INTO TestDaten (XWert, YWert) VALUES (4, 8)"
db.Execute SQL, dbFailOnError
SQL = "INSERT INTO TestDaten (XWert, YWert) VALUES (5, 10)"
db.Execute SQL, dbFailOnError

MsgBox "Testtabelle 'TestDaten' wurde erstellt und mit Beispieldaten gefüllt.", vbInformation, "Erfolg"

Set fld = Nothing
Set tdf = Nothing
Set db = Nothing
