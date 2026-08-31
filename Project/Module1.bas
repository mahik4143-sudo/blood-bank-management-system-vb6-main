Attribute VB_Name = "Module1"

Public C As New ADODB.Connection ' New keyword add karein
Public R As New ADODB.Recordset
Public sql As String



Public Function CONN()
Set C = New ADODB.Connection
C.Open "PROVIDER=MSDAORA.1;USER ID=PRJ2531G/PRJ2531G;PERSIST SECURITY INFO =FALSE"
Set R = New ADODB.Recordset
End Function

