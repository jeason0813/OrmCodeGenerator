'
' Copyright Diego Lunadei (d.lunadei--at--gmail.com)
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'


Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Xml.Serialization

Module mdlConvertitori
    Public ConnAttiva As System.Data.Common.DbConnection

    Public LinguaggioScelto As enLinguaggio = enLinguaggio.VbNet

    Public Function CreaDaStrutturaSqlCompact(ByVal PathFile As String, ByVal pwd As String, Optional ByVal LasciaConnAttiva As Boolean = False) As cDatabase

        Dim DB As New cDatabase
        DB.TipoOrigine = "MS SQL Compact - Percorso: " & PathFile
        DB.TipoSorgente = enTipoSorgente.SQLServerCompact
        DB.OriginalTipoSorgente = DB.TipoSorgente
        Return DB

    End Function

    Public Function CreaDaStrutturaMSSql(ByVal Server As String, ByVal DbName As String, ByVal Login As String, ByVal Pwd As String, Optional ByVal AutenticazioneIntegrata As Boolean = False, Optional ByVal LasciaConnAttiva As Boolean = False, Optional ByVal LoadView As Boolean = False) As cDatabase
        Dim DB As New cDatabase
        DB.TipoOrigine = "MS SQL Server - " & Server & " Database: " & DbName & " Login: " & Login & " Password: " & Pwd & " Integrated security: " & IIf(AutenticazioneIntegrata, "Yes", "No")
        DB.TipoSorgente = enTipoSorgente.SQLServer
        DB.OriginalTipoSorgente = DB.TipoSorgente
        Try

            Cursor.Current = Cursors.WaitCursor

            Dim ConnectionString As String
            If AutenticazioneIntegrata Then
                ConnectionString = "Integrated Security=SSPI;MultipleActiveResultsets=true;Persist Security Info=False;User ID=dbsql;Initial Catalog=" & DbName & ";Server=" & Server
            Else
                ConnectionString = "Server=" & Server & ";Database=" & DbName & ";MultipleActiveResultsets=true;Uid=" & Login & ";Pwd=" & Pwd & ";"
            End If

            Dim Cn As SqlConnection = New SqlConnection(ConnectionString)

            Dim mytb As New DataTable

            Dim myCommand As SqlCommand = Cn.CreateCommand()
            'LOAD FROM TABLE AND VIEW
            Dim Sql As String = ""

            Sql = "SELECT TABLE_NAME,TABLE_TYPE FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE IN( 'BASE TABLE'"
            If LoadView Then Sql &= ",'VIEW'"
            Sql &= ") ORDER BY TABLE_TYPE,TABLE_NAME"

            myCommand.CommandText = Sql

            Cn.Open()

            Dim myReader As SqlDataReader = myCommand.ExecuteReader()

            mytb.Load(myReader)

            myReader.Close()
            myCommand.Dispose()

            Dim Dr As DataRow

            For Each Dr In mytb.Rows

                Dim tb As New cTabella
                tb.NomeTabella = Dr("TABLE_NAME").ToString
                tb.NomeTabella = tb.NomeTabella.Substring(0, 1).ToUpper & tb.NomeTabella.Substring(1).ToLower
                tb.NomeClasse = tb.NomeTabella
                tb.NomeClasseDAO = tb.NomeClasse & "DAO"
                tb.StrutturaRiferimento = DB
                If Dr("TABLE_TYPE").ToString = "VIEW" Then tb.TipoTabella = enTipoTabella.View

                'retrive dei campi

                Dim mycmdTab As SqlCommand = Cn.CreateCommand
                mycmdTab.CommandText = "SELECT top 0 * FROM [" & tb.NomeTabella & "]"

                Dim myreadCampi As SqlDataReader = mycmdTab.ExecuteReader

                Dim x As New System.Data.DataTable, col As DataColumn

                x.Load(myreadCampi)

                'creo la struttura

                Dim Indice As Integer = 0
                Dim TrovataChiave As Boolean = False
                For Each col In x.Columns

                    Dim Campo As New cCampoDb
                    Campo.Ordinal = Indice
                    Campo.Nome = col.Caption
                    Campo.AllowDBNull = col.AllowDBNull
                    Campo.Contatore = col.AutoIncrement
                    Campo.Tipo = col.DataType.ToString
                    Campo.MaxLength = col.MaxLength
                    Campo.DefaultValue = col.DefaultValue.ToString
                    If col.AutoIncrement = True And TrovataChiave = False Then
                        Campo.Contatore = True
                        Campo.CampoChiave = True
                        tb.CampoChiave = Campo
                        TrovataChiave = True
                    End If

                    tb.Campi.Add(Campo)

                    Campo = Nothing
                    Indice += 1
                Next

                If TrovataChiave = False Then
                    Dim campo As cCampoDb = tb.Campi(0)
                    campo.CampoChiave = True
                    tb.CampoChiave = campo
                End If

                DB.Tabelle.Add(tb)
                tb = Nothing

            Next


            'PER LE RELAZIONI SONO ARRIVATO QUI:
            '*****************************
            'SELECT * FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS

            'Dim restrictions() As String

            'mytb = Cn.GetOleDbSchemaTable(OleDbSchemaGuid.Foreign_Keys, restrictions)

            'For Each Dr In mytb.Rows
            '    'qui carico la struttura delle relazioni
            '    Dim rel As New cRelazioneTabella
            '    rel.TabellaOrigine = Dr("PK_TABLE_NAME").ToString
            '    rel.TabellaDestinazione = Dr("FK_TABLE_NAME").ToString
            '    rel.CampoOrigine = Dr("PK_COLUMN_NAME").ToString
            '    rel.CampoDestinazione = Dr("FK_COLUMN_NAME").ToString

            '    DB.Relazioni.Add(rel)

            'Next


            Cn.Close()
            Cn.Dispose()

            Cursor.Current = Cursors.Default


        Catch ex As Exception
            GestisciErrore(ex)
        End Try

        Return DB

    End Function

    Public Function CreaDaStrutturaLunaProject(ByVal Path As String) As cDatabase
        Dim Db As cDatabase

        Try
            Cursor.Current = Cursors.WaitCursor

            Dim serialize As XmlSerializer = New XmlSerializer(GetType(cDatabase))
            Dim deSerialize As IO.FileStream = New IO.FileStream(Path, IO.FileMode.Open)
            Db = serialize.Deserialize(deSerialize)

            If Db.OriginalTipoSorgente = 0 Then Db.OriginalTipoSorgente = Db.TipoSorgente

            Db.TipoOrigine = "Luna Data Schema - File: " & Path
            Db.TipoSorgente = enTipoSorgente.LunaDataSchema

            For Each t As cTabella In Db.Tabelle
                t.StrutturaRiferimento = Db
            Next

            Cursor.Current = Cursors.Default

        Catch ex As Exception
            GestisciErrore(ex)
        End Try

        Return Db
    End Function

    Public Function CreaDaStrutturaMSAccess(ByVal Path As String, Optional ByVal LasciaConnAttiva As Boolean = False) As cDatabase
        Dim DB As New cDatabase
        DB.TipoOrigine = "MS Access - File: " & Path
        DB.TipoSorgente = enTipoSorgente.Access
        DB.OriginalTipoSorgente = DB.TipoSorgente
        Try

            Cursor.Current = Cursors.WaitCursor

            Dim ConnectionString As String = ""
            If Path.EndsWith("accdb") Then
                ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Path & ";Persist Security Info=False;"
            Else
                ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Path & ";Persist Security Info=False;"
            End If
            'Provider=Microsoft.ACE.OLEDB.12.0;
            Dim Cn As OleDbConnection = New OleDbConnection(ConnectionString)

            Dim mytb As New DataTable

            Cn.Open()
            Dim Res() As String = New String(3) {}
            Res(3) = "TABLE"
            mytb = Cn.GetSchema("TABLES", Res)

            Dim Dr As DataRow

            For Each Dr In mytb.Rows

                Dim tb As New cTabella
                tb.NomeTabella = Dr("Table_Name").ToString
                tb.NomeTabella = tb.NomeTabella.Substring(0, 1).ToUpper & tb.NomeTabella.Substring(1).ToLower
                tb.NomeClasse = tb.NomeTabella
                tb.NomeClasseDAO = tb.NomeClasse & "DAO"
                tb.StrutturaRiferimento = DB
                'retrive dei campi

                Dim mycmdTab As OleDbCommand = Cn.CreateCommand
                mycmdTab.CommandText = "SELECT top 1 * FROM [" & tb.NomeTabella & "]"

                Dim myreadCampi As OleDbDataReader = mycmdTab.ExecuteReader(CommandBehavior.KeyInfo)

                Dim x As New System.Data.DataTable

                x = myreadCampi.GetSchemaTable()

                'For each field in the table...

                Dim Indice As Integer = 0
                Dim DrTab As DataRow
                Dim TrovataChiave As Boolean = False
                For Each DrTab In x.Rows

                    Dim Campo As New cCampoDb
                    Campo.Ordinal = Indice
                    Campo.Nome = DrTab("ColumnName")
                    Campo.AllowDBNull = DrTab("AllowDBNull")
                    Campo.Tipo = DrTab("DataType").ToString
                    Campo.MaxLength = DrTab("ColumnSize")
                    Campo.DefaultValue = ""
                    If DrTab("Isautoincrement") = True And TrovataChiave = False Then
                        Campo.Contatore = True
                        Campo.CampoChiave = True
                        tb.CampoChiave = Campo
                        TrovataChiave = True
                    End If

                    tb.Campi.Add(Campo)

                    Campo = Nothing
                    Indice += 1

                Next

                If TrovataChiave = False Then
                    Dim campo As cCampoDb = tb.Campi(0)
                    campo.CampoChiave = True
                    tb.CampoChiave = campo
                End If

                DB.Tabelle.Add(tb)
                tb = Nothing

            Next

            Dim restrictions() As String

            mytb = Cn.GetOleDbSchemaTable(OleDbSchemaGuid.Foreign_Keys, restrictions)

            For Each Dr In mytb.Rows
                'qui carico la struttura delle relazioni
                Dim rel As New cRelazioneTabella
                rel.TabellaOrigine = Dr("FK_TABLE_NAME").ToString
                rel.TabellaDestinazione = Dr("PK_TABLE_NAME").ToString
                rel.CampoOrigine = Dr("FK_COLUMN_NAME").ToString
                rel.CampoDestinazione = Dr("PK_COLUMN_NAME").ToString

                DB.Relazioni.Add(rel)

            Next

            Cn.Close()
            Cn.Dispose()

            Cursor.Current = Cursors.Default

        Catch ex As Exception
            GestisciErrore(ex)
        End Try

        Return DB
    End Function

End Module
