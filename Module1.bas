Attribute VB_Name = "ConectaBD"
Public cnnSegura As New ADODB.Connection
Public cnnCodbar As New ADODB.Connection

Public Sub ConectDB()
Path = App.Path & "\agenda.mdb"
'db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Path & ";Persist Security Info=False"
'globDBConnect.Open strVar



Dim oConn As ADODB.Connection

    Set oConn = New ADODB.Connection
    With oConn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Mode = adModeReadWrite
        .Open "Data source=" & Path & ""
    End With


    Dim oRec As ADODB.Recordset
    Set oRec = New ADODB.Recordset
    oRec.Open "SELECT * FROM agenda", oConn, 3, 3

End Sub

