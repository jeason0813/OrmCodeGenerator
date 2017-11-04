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
' prova auto diego e camillo  ariprova 


Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Xml.Serialization

Namespace LunaEngine

    Public Class LunaInterpreter

        Public Shared Function CreaDaStrutturaSqlCompact(ByVal PathFile As String, ByVal pwd As String, Optional ByVal LasciaConnAttiva As Boolean = False) As cDatabase

            Dim DB As New cDatabase
            DB.TipoOrigine = "MS SQL Compact - Percorso: " & PathFile
            DB.TipoSorgente = enDatasourceType.SQLServerCompact
            DB.OriginalTipoSorgente = DB.TipoSorgente
            Return DB

        End Function

        Public Shared Function CreaDaStrutturaMSSql(ByVal Server As String, ByVal DbName As String, ByVal Login As String, ByVal Pwd As String, Optional ByVal AutenticazioneIntegrata As Boolean = False, Optional ByVal LasciaConnAttiva As Boolean = False, Optional ByVal LoadView As Boolean = False) As cDatabase
            Dim DB As New cDatabase
            DB.TipoOrigine = "MS SQL Server - " & Server & " Database: " & DbName & " Login: " & Login & " Password: " & Pwd & " Integrated security: " & IIf(AutenticazioneIntegrata, "Yes", "No")
            DB.TipoSorgente = enDatasourceType.SQLServer
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
                    tb.IsChanged = True
                    If Dr("TABLE_TYPE").ToString = "VIEW" Then tb.TipoTabella = enTableType.View

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
                        'Campo.Ordinal = Indice
                        Campo.Nome = col.Caption
                        Campo.AllowDBNull = col.AllowDBNull
                        Campo.Contatore = col.AutoIncrement
                        Campo.Tipo = col.DataType.ToString
                        Campo.MaxLength = col.MaxLength
                        Campo.DefaultValue = col.DefaultValue.ToString
                        If col.AutoIncrement Then
                            Campo.Contatore = True
                        End If
                        If Array.IndexOf(x.PrimaryKey, col) >= 0 And TrovataChiave = False Then
                            Campo.CampoChiave = True
                            tb.CampoChiave = Campo
                            tb.TipoCampoChiaveString = Campo.TipoStringa
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
        Private Shared Function GetOrigineDati(db As cDatabase) As cDatabase
            Dim DbOrig As cDatabase = Nothing
            ErroreDaGestire = False
            Try
                Select Case db.OriginalTipoSorgente

                    Case enDatasourceType.Access
                        DbOrig = LunaInterpreter.CreaDaStrutturaMSAccess(db.PathDboServer)
                    Case enDatasourceType.SQLServer
                        DbOrig = LunaInterpreter.CreaDaStrutturaMSSql(db.PathDboServer, db.NomeDb, db.NomeUt, Cripta(db.Dbpwd), db.AutenticazioneIntegrata, 0)
                    Case enDatasourceType.SQLServerCompact
                        DbOrig = LunaInterpreter.CreaDaStrutturaSqlCompact(db.PathDboServer, "")
                    Case enDatasourceType.Oracle
                End Select
                Return DbOrig
            Catch ex As Exception
                Return Nothing
            End Try
            ErroreDaGestire = True

        End Function
        Public Shared Function CreaDaStrutturaLunaProject(ByVal Path As String) As cDatabase
            Dim DbFile As cDatabase = Nothing
            Dim DbOrig As cDatabase = Nothing
            Dim tabOr As cTabella
            Dim tabFile As cTabella
            Dim CampoDb As cCampoDb
            Dim CampoOrig As cCampoDb
            Dim OrigineValida As Boolean = False
            Dim oldNomeClasse As String = ""
            Dim oldNomeClasseDAO As String = ""
            'Dim t As cTabella
            Try
                Cursor.Current = Cursors.WaitCursor

                Dim serialize As XmlSerializer = New XmlSerializer(GetType(cDatabase))
                Dim deSerialize As IO.FileStream = New IO.FileStream(Path, IO.FileMode.Open)

                DbFile = serialize.Deserialize(deSerialize)

                DbOrig = GetOrigineDati(DbFile)

                If DbOrig.Tabelle.Count > 0 Then
                    OrigineValida = True
                Else
                    MessageBox.Show("Unable to load original database, File not found!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If

                DbFile.TipoOrigine = "Luna Data Schema - File: " & Path
                DbFile.TipoSorgente = DbFile.OriginalTipoSorgente
                'DbOrig = DaaBase
                'db = file Luna

                If OrigineValida Then
                    For Each tabOr In DbOrig.Tabelle

                        tabFile = DbFile.Tabelle.Find(Function(item) item.NomeTabella = tabOr.NomeTabella)

                        If tabFile Is Nothing Then
                            tabFile = tabOr.Clone
                            tabFile.IsChanged = True
                            tabFile.StrutturaRiferimento = DbFile
                            DbFile.Tabelle.Add(tabOr)
                        Else
                            Dim Posizione As Integer = 0
                            For Each CampoOrig In tabOr.Campi
                                CampoDb = tabFile.Campi.Find(Function(item) item.Nome = CampoOrig.Nome)
                                If CampoDb Is Nothing Then
                                    tabFile.IsChanged = True
                                    tabFile.Campi.Insert(Posizione, CampoOrig)
                                ElseIf CampoDb.TipoStringa <> CampoOrig.TipoStringa Then
                                    'qui controllo il tipo se e' cambiato
                                    tabFile.IsChanged = True
                                    tabFile.Campi(Posizione) = CampoOrig
                                End If
                                Posizione += 1
                            Next
                            'controllo se qualche campo è stato eliminato

                            Dim l As List(Of cCampoDb) = tabFile.Campi.FindAll(Function(x) x.Nome.Length <> 0)

                            For Each CampoDb In l
                                CampoOrig = tabOr.Campi.Find(Function(item) item.Nome = CampoDb.Nome)
                                If CampoOrig Is Nothing Then
                                    tabFile.Campi.Remove(CampoDb)
                                    tabFile.IsChanged = True
                                End If

                            Next

                        End If
                    Next
                End If

                For Each r As cRelazioneTabella In DbFile.Relazioni
                    r.StrutturaRiferimento = DbFile
                Next

                Cursor.Current = Cursors.Default

            Catch ex As Exception
                GestisciErrore(ex)
            End Try

            Return DbFile
        End Function

        Public Shared Function CreaDaStrutturaMSAccess(ByVal Path As String, Optional ByVal LasciaConnAttiva As Boolean = False) As cDatabase
            Dim DB As New cDatabase
            DB.TipoOrigine = "MS Access - File: " & Path
            DB.TipoSorgente = enDatasourceType.Access
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
                    tb.IsChanged = True
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
                    'For Each DrTab In x.Rows

                    '    For Each c As DataColumn In DrTab.Table.Columns
                    '        MessageBox.Show(c.ColumnName)
                    '    Next
                    '    Exit For
                    'Next
                    For Each DrTab In x.Rows

                        Dim Campo As New cCampoDb
                        'Campo.Ordinal = Indice
                        Campo.Nome = DrTab("ColumnName")
                        Campo.AllowDBNull = DrTab("AllowDBNull")
                        Campo.Tipo = DrTab("DataType").ToString
                        Campo.MaxLength = DrTab("ColumnSize")
                        Campo.DefaultValue = ""
                        If DrTab("Isautoincrement") = True Then
                            Campo.Contatore = True
                        End If
                        If DrTab("IsKey") = True And TrovataChiave = False Then
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
                    rel.StrutturaRiferimento = DB
                    If rel.TabellaOrigine.StartsWith("MSys") = False And rel.TabellaDestinazione.StartsWith("MSys") = False Then
                        'qui devo controllare che la tabella esista nella sorgente

                        DB.Relazioni.Add(rel)


                    End If

                Next

                Cn.Close()
                Cn.Dispose()

                Cursor.Current = Cursors.Default

            Catch ex As Exception
                GestisciErrore(ex)
            End Try

            Return DB
        End Function

    End Class

    Public Class LunaConverter

        Public Shared Property OverWriteAllTables As Boolean = True

        Private Shared Function GeneraClassBase(source As cDatabase) As String

            Dim Buffer As String = ""
            Dim x As ILunaTemplate

            x = New My.Templates.LunaBaseClassTvb(source)

            Buffer = x.TransformTextEX()
            Return Buffer

        End Function

        Private Shared Function GeneraContext(source As cDatabase) As String

            Dim Buffer As String = ""
            Dim x As ILunaTemplate

            x = New My.Templates.LunaContextClassTvb()

            Buffer = x.TransformTextEX()
            Return Buffer

        End Function

        Private Shared Function GeneraClassEntity(source As cDatabase, T As cTabella) As String

            Dim Buffer As String = ""
            Dim x As ILunaTemplate

            x = New My.Templates.LunaEntityClassTvb(source, T)

            Buffer = x.TransformTextEX()
            Return Buffer

        End Function

        Private Shared Function GeneraTableClass(source As cDatabase, T As cTabella) As String

            Dim Buffer As String = ""
            Dim x As ILunaTemplate

            x = New My.Templates.LunaTableClassTvb(source, T)

            Buffer = x.TransformTextEX()
            Return Buffer

        End Function

        Private Shared Function GeneraClassDAO(source As cDatabase, T As cTabella) As String

            Dim Buffer As String = ""
            Dim x As ILunaTemplate

            x = New My.Templates.LunaDAOClassTvb(source, T)

            Buffer = x.TransformTextEX()
            Return Buffer


        End Function

        Private Shared Function GeneraClassDAOUser(source As cDatabase, T As cTabella) As String

            Dim Buffer As String = ""
            Dim x As ILunaTemplate

            x = New My.Templates.LunaDAOClassUserTvb(source, T)

            Buffer = x.TransformTextEX()
            Return Buffer


        End Function

        Private Shared Function GeneraInterface(source As cDatabase, T As cTabella) As String

            Dim Buffer As String = ""
            Dim x As ILunaTemplate

            x = New My.Templates.LunaEntityInterfaceTvb(source, T)

            Buffer = x.TransformTextEX()
            Return Buffer

        End Function

        Public Shared Function GeneraConfig(source As cDatabase) As String
            Dim Buffer As String = String.Empty
            Dim x As ILunaTemplate
            x = New My.Templates.LunaConfig(source)
            Buffer = x.TransformTextEX
            Return Buffer
        End Function

        Public Shared Function ElaboraStruttura(source As cDatabase) As LunaResult

            Dim Ris As New LunaResult

            Ris.BufferSQL = ElaboraSQL(source)
            Ris.BufferClassiBase = GeneraClassBase(source)
            Ris.BufferContext = GeneraContext(source)
            Ris.BufferConfig = GeneraConfig(source)

            For Each Tb As cTabella In source.Tabelle

                Dim LavoraTabella As Boolean = False
                If OverWriteAllTables Then
                    LavoraTabella = Tb.Selezionata
                Else
                    If Tb.Selezionata Then
                        LavoraTabella = Tb.IsChanged
                    End If
                End If

                If LavoraTabella Then
                    Dim TbRes As New LunaResultTableCode
                    TbRes.NomeClasse = Tb.NomeClasse
                    TbRes.NomeClasseDAO = Tb.NomeClasseDAO

                    If Tb.TipoTabella = enTableType.Table Then
                        TbRes.CodiceTableClass = GeneraTableClass(source, Tb)
                        TbRes.CodiceClasseEntity = GeneraClassEntity(source, Tb)
                        TbRes.CodiceClasseEntity &= vbNewLine & GeneraInterface(source, Tb)
                    End If

                    TbRes.CodiceClasseDAO = GeneraClassDAO(source, Tb)
                    TbRes.CodiceClasseDAOUser = GeneraClassDAOUser(source, Tb)
                    Ris.ListTablesCode.Add(TbRes)
                End If
            Next

            Return Ris

        End Function

        'Private Shared Function GeneraModulo() As String
        '    Dim Buffer As String = ""
        '    Dim x As ILunaModuleT
        '    If LinguaggioScelto = enLinguaggio.VbNet Then
        '        x = New My.Templates.LunaModuleTvb
        '    Else
        '        x = New My.Templates.LunaModuleTcs
        '    End If
        '    Buffer = x.TransformTextEX()

        '    Return Buffer
        'End Function

#Region "SQL Function"

        Private Shared Function ElaboraSQL(source As cDatabase) As String
            Dim _BufferSQL As String = String.Empty
            Dim Tb As cTabella

            For Each Tb In source.Tabelle

                Dim NomeTab As String = Tb.NomeTabella

                If Tb.Selezionata Then
                    If Tb.TipoTabella = enTableType.Table Then
                        _BufferSQL &= CreaTabellaSql(NomeTab)

                        Dim Campo As cCampoDb, NomeCampoChiave As String = ""

                        For Each Campo In Tb.Campi
                            If Campo.CampoChiave Then
                                NomeCampoChiave = Campo.Nome
                            End If

                            _BufferSQL &= CreaCampoSql(NomeTab, Campo)
                        Next

                        If NomeCampoChiave.Length Then _BufferSQL &= CreaChiaveSql(NomeTab, NomeCampoChiave)


                        '_BufferSQL &= CreaInsertDati(NomeTab)


                        _BufferSQL &= vbNewLine

                    Else
                        'TODO: qui estrapolare il codice della vista


                    End If
                End If

            Next

            Return _BufferSQL
        End Function

        Private Shared Function CreaTabellaSql(ByVal NomeTab As String) As String

            Dim buffer As String

            buffer = "DROP TABLE " & NomeTab & ";" & vbNewLine
            buffer &= "CREATE TABLE " & NomeTab & ";" & vbNewLine

            Return buffer

        End Function

        Private Shared Function CreaChiaveSql(ByVal NomeTab As String, ByVal NomeCampoChiave As String) As String

            Dim buffer As String

            buffer = "CREATE INDEX idx" & NomeCampoChiave & " ON " & NomeTab & "(" & NomeCampoChiave & ") WITH PRIMARY;" & vbNewLine

            Return buffer

        End Function

        Private Shared Function CreaCampoSql(ByVal NomeTab As String, ByRef Campo As cCampoDb) As String

            '"campo: " & col.Caption & " - allownull: " & col.AllowDBNull & " -  autoincrement: " & col.AutoIncrement & " - tipo: " & col.DataType.ToString & " - size: " & col.MaxLength & " - defaultvalue: " & col.DefaultValue & " - readonly:" & col.ReadOnly
            Dim buffer As String

            buffer = Campo.Nome

            Select Case Campo.TipoStringa.ToLower

                Case Is = "int32"
                    If Campo.Contatore Then
                        buffer &= " COUNTER "
                    Else
                        buffer &= " INT"
                    End If

                Case Is = "int16"
                    If Campo.Contatore Then
                        buffer &= " COUNTER "
                    Else
                        buffer &= " INT"
                    End If

                Case Is = "integer"
                    If Campo.Contatore Then
                        buffer &= " COUNTER "
                    Else
                        buffer &= " INT"
                    End If


                Case Is = "string"

                    If Campo.MaxLength < 256 Then
                        buffer &= " TEXT"
                    Else
                        buffer &= " MEMO"
                    End If

                Case Is = "decimal"

                    buffer &= " SINGLE"

                Case Is = "boolean"
                    buffer &= " YESNO "

                Case Else
                    buffer &= " " & Campo.TipoStringa.ToUpper

            End Select

            '  buffer &= " " & Campo.DataType.Name

            If Campo.TipoStringa = "String" And Campo.MaxLength < 256 Then buffer &= " (" & Campo.MaxLength & ")"

            If Campo.DefaultValue.ToString.Length Then buffer &= " DEFAULT " & Campo.DefaultValue

            buffer &= " " & IIf(Campo.AllowDBNull, "NULL", "NOT NULL")

            buffer = "ALTER TABLE " & NomeTab & " ADD COLUMN " & buffer & ";" & vbNewLine
            'buffer &= ";" & vbNewLine

            Return buffer

        End Function

#End Region

    End Class

    Public Class LunaResult

        Public Property BufferConfig As String = String.Empty

        Private _BufferClassiBase As String
        Public Property BufferClassiBase() As String
            Get
                Return _BufferClassiBase
            End Get
            Set(value As String)
                _BufferClassiBase = value
            End Set
        End Property

        Private _BufferSQL As String = ""
        Public Property BufferSQL() As String
            Get
                Return _BufferSQL
            End Get
            Set(value As String)
                _BufferSQL = value
            End Set
        End Property

        Public Property BufferContext As String

        Public ListTablesCode As New List(Of LunaResultTableCode)

    End Class

    Public Class LunaResultTableCode

        Public Property NomeClasse As String
        Public Property CodiceClasseEntity As String
        Public Property CodiceTableClass As String

        Public Property NomeClasseDAO As String
        Public Property CodiceClasseDAO As String
        Public Property CodiceClasseDAOUser As String

    End Class

End Namespace

