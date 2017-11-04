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

Imports System.IO
Imports System.Xml.Serialization
Imports Microsoft.VisualStudio.TextTemplating
Imports Luna.LunaEngine

Public Class frmMain
    Private OriginalPathLunaPrj As String = ""
    Private StrutturaDati As cDatabase = Nothing
    Private RichCode As List(Of RichTextBox)

    Private Sub AssociaLDS()
        Try
            My.Computer.Registry.ClassesRoot.CreateSubKey(".lds").SetValue("", "LunaDataSchema", Microsoft.Win32.RegistryValueKind.String)
            My.Computer.Registry.ClassesRoot.CreateSubKey("LunaDataSchema\shell\open\command").SetValue("", Application.ExecutablePath & " ""%l"" ", Microsoft.Win32.RegistryValueKind.String)

        Catch ex As Exception


        End Try


    End Sub

    Private Sub ImpostaInfo()

        If My.Application.Info.Title <> "" Then
            ApplicationTitle.Text = My.Application.Info.Title
        Else
            'If the application title is missing, use the application name, without the extension
            ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If

        Version.Text = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision

    End Sub

    Private Sub lnkAvanti_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkAvanti.LinkClicked

        If MessageBox.Show("Do you confirm selected data source for class creation?", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            StrutturaDati = Nothing
            ResettaControlli()
            'qui prendo la sorgente e la do in pasto all'interprete che crea il database in formato unico

            If rdoMsAccess.Checked Then
                If lblFile.Text.Length Then
                    StrutturaDati = LunaInterpreter.CreaDaStrutturaMSAccess(lblFile.Text)
                    StrutturaDati.PathDboServer = lblFile.Text
                Else
                    MessageBox.Show("Select Access File!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            ElseIf rdoMSSql.Checked Then
                If rdoSQL.Checked And (txtSQLLogin.Text.Length = 0 Or txtSQLPwd.Text.Length = 0) Then
                    MessageBox.Show("Insert credentials!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    StrutturaDati = LunaInterpreter.CreaDaStrutturaMSSql(txtSQLServer.Text, txtSQLDb.Text, txtSQLLogin.Text, txtSQLPwd.Text, rdoNT.Checked, chkLoadView.Checked)
                    StrutturaDati.PathDboServer = txtSQLServer.Text
                    StrutturaDati.NomeDb = txtSQLDb.Text.ToString
                    StrutturaDati.NomeUt = txtSQLLogin.Text.ToString
                    StrutturaDati.Dbpwd = Cripta(txtSQLPwd.Text.ToString)
                    StrutturaDati.AutenticazioneIntegrata = rdoNT.Checked
                End If
            ElseIf rdoSqlCompact.Checked Then
                If lblFile.Text.Length Then
                    StrutturaDati = LunaInterpreter.CreaDaStrutturaSqlCompact(lblFileSql.Text, "")
                    StrutturaDati.PathDboServer = lblFileSql.Text
                Else
                    MessageBox.Show("Select SQL Compact File!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            ElseIf rdoOracle.Checked Then
                MessageBox.Show("Tipo origine non ancora supportata!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf rdoLuna.Checked Then
                If lblFilePrj.Text.Length Then
                    OriginalPathLunaPrj = lblFilePrj.Text
                    StrutturaDati = LunaInterpreter.CreaDaStrutturaLunaProject(lblFilePrj.Text)


                Else
                    MessageBox.Show("Select Luna Data Schema File!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            End If
            CaricaStrutturaDati()
        End If

    End Sub

    Private Sub CaricaStrutturaDati()
        If Not StrutturaDati Is Nothing Then
            lblSorgenteInUso.Text = StrutturaDati.TipoOrigine
            AggiornaListaTabelle()
            'se e' stata creata una struttura vado al passaggio successivo
            tabMain.SelectedIndex += 1
            If rdoLuna.Checked Then
                SelezionaTabelle(StrutturaDati)
            End If
        End If
    End Sub

    Private Sub ResettaControlli()

        lstTabelle.Items.Clear()

    End Sub


    Private Sub lnkAggiornaTab_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkAggiornaTab.LinkClicked

        ChangeNameObject()

    End Sub

    Private Sub CambiaNomeClasse()

        Dim Node As TreeNode = tvwTabelle.SelectedNode, Trovato As Boolean = False


        If Not Node Is Nothing Then

            If Node.Name.ToLower.StartsWith("c") Then
                Dim NomeTabella As String = Node.Parent.Name.Substring(2)
                Trovato = True
                Dim NuovoNome As String = InputBox("Insert generated name for this class:", "Luna", Node.Tag)
                Dim ErrGiaPres As Boolean = False
                If NuovoNome.Trim.ToString.Length Then
                    For Each nodeSing As TreeNode In tvwTabelle.Nodes.Find("c_" & NuovoNome, True)
                        If nodeSing.Tag.ToString = NuovoNome And Node.Name <> nodeSing.Name Then ErrGiaPres = True
                    Next

                    For Each nodeSing As TreeNode In tvwTabelle.Nodes.Find("cd_" & NuovoNome, True)
                        If nodeSing.Tag.ToString = NuovoNome And Node.Name <> nodeSing.Name Then ErrGiaPres = True
                    Next

                    If ErrGiaPres Then
                        MessageBox.Show("A class named " & NuovoNome & " already exist in this project. Choose another name please!", "Luna", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        Node.Tag = NuovoNome
                        'qui devo andare a cercare la tabella in questione e mettere ischanged = true 
                        Dim t As cTabella = StrutturaDati.Tabelle.Find(Function(x) x.NomeTabella = NomeTabella)
                        t.IsChanged = True

                        If Node.Name.ToLower.StartsWith("cd_") Then
                            Node.Text = "Generated DAO Class : " & NuovoNome
                            Node.Name = "cd_" & NuovoNome
                        Else
                            Node.Text = "Generated Class : " & NuovoNome
                            Node.Name = "c_" & NuovoNome
                        End If

                    End If

                End If

            End If

        End If

        If Not Trovato Then
            MessageBox.Show("Select a Class Node (Class or DAOClass)or a Field Node to change generated name", "Luna", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub CambiaNomeCampo()

        Dim Node As TreeNode = tvwTabelle.SelectedNode, Trovato As Boolean = False

        If Not Node Is Nothing Then

            If Node.Name.ToLower.StartsWith("fs_") Then

                'qui mi prendo il nodo padre che e' per forza il nome della tabella 
                Dim nodeF As TreeNode = Node.Parent.Parent
                Dim NomeTabella As String = Node.Parent.Parent.Name.Substring(2)
                Trovato = True
                Dim NuovoNome As String = InputBox("Insert generated name for this field:", "Luna", Node.Tag)
                Dim ErrGiaPres As Boolean = False
                If NuovoNome.Trim.ToString.Length Then
                    For Each nodeSing As TreeNode In nodeF.Nodes(2).Nodes
                        If nodeSing.Tag.ToString = NuovoNome And Node.Name <> nodeSing.Name Then ErrGiaPres = True
                    Next

                    If ErrGiaPres Then
                        MessageBox.Show("A field named " & NuovoNome & " already exist for this table. Choose another name please!", "Luna", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        Node.Tag = NuovoNome
                        Node.Text = NuovoNome
                        'Node.Name = "fs_" & NomeTabella & NuovoNome
                        Dim t As cTabella = StrutturaDati.Tabelle.Find(Function(x) x.NomeTabella = NomeTabella)
                        t.IsChanged = True

                    End If

                End If

            End If

        End If

        If Not Trovato Then
            MessageBox.Show("Select a Class Node (Class or DAOClass) or a Field Node to change generated name", "Luna", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub AggiornaListaTabelle()

        If Not StrutturaDati Is Nothing Then

            lstTabelle.Items.Clear()

            tvwTabelle.Nodes.Clear()

            Dim tb As cTabella

            For Each tb In StrutturaDati.Tabelle
                Dim Checkato As Boolean = False

                If StrutturaDati.TipoSorgente = enDatasourceType.LunaDataSchema Then Checkato = True
                lstTabelle.Items.Add("(" & tb.TipoTabellaString & ") " & tb.NomeTabella, Checkato)
                'Node = tvwTabelle.Nodes.Add("t_" & tb.NomeTabella, tb.NomeTabella, 1, 1)
                'Node.BackColor = Color.Red

                'Dim NodeC As TreeNode = Node.Nodes.Add("c_" & tb.NomeTabella, "Generated Class: " & tb.NomeClasse, 2, 2)
                'NodeC.Tag = tb.NomeClasse
                'Dim NodecD As TreeNode = Node.Nodes.Add("cd_" & tb.NomeTabella, "Generated DAO Class: " & tb.NomeClasseDAO, 2, 2)
                'NodecD.Tag = tb.NomeClasseDAO
                'Node.Nodes.Add("f" & tb.Nome, "Field")

                ' qui se la tabella e' selezionata la devo mettere checckata


            Next

            tvwTabelle.ExpandAll()

        End If

    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        SalvaSetting()

    End Sub

    Private Sub frmStart_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Text = Application.ProductName & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision

        AssociaLDS()

        CaricaSetting()

        CaricaTutorial()

        'controllo se doppio click su qualche file lds
        If My.Application.CommandLineArgs.Count Then
            Dim Param As String = My.Application.CommandLineArgs(0).ToString
            If Param.EndsWith(".lds") Then
                rdoLuna.Checked = True
                'chkSovrascriviLds.Checked = True
                OriginalPathLunaPrj = Param
                StrutturaDati = LunaInterpreter.CreaDaStrutturaLunaProject(Param)

                'If Not StrutturaDati Is Nothing Then
                '    chkSovrascriviLds.Checked = True
                '    lblSorgenteInUso.Text = StrutturaDati.TipoOrigine

                '    AggiornaListaTabelle()
                '    'se e' stata creata una struttura vado al passaggio successivo
                '    tabMain.SelectedIndex += 1

                'End If
                CaricaStrutturaDati()
            End If
        End If

    End Sub

    Private Sub CaricaTutorial(Optional ByVal UrlWebPage As String = "http://www.diegolunadei.it/luna")

        WebPage.Navigate(UrlWebPage)

    End Sub

    Private Sub CaricaSetting()

        txtSQLServer.Text = My.Settings.SqlServer
        txtSQLDb.Text = My.Settings.SqlDB
        txtSQLLogin.Text = My.Settings.SqlLogin
        txtSQLPwd.Text = My.Settings.SqlPwd

        lblFile.Text = My.Settings.AccessPath

        cmbLanguage.Items.Add(New With {.Descrizione = "Vb.Net", .valore = enLanguage.VbNet})
        'cmbLanguage.Items.Add(New With {.Descrizione = "C#", .valore = enLinguaggio.CSharp})
        cmbLanguage.SelectedIndex = 0

    End Sub

    Private Sub SalvaSetting()

        My.Settings.Item("SQLServer") = txtSQLServer.Text
        My.Settings.Item("SQLDB") = txtSQLDb.Text
        My.Settings.Item("SQLLogin") = txtSQLLogin.Text
        My.Settings.Item("SQLPwd") = txtSQLPwd.Text
        My.Settings.Item("AccessPath") = lblFile.Text

        My.Settings.Save()

    End Sub

    Private Sub btnScegliFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnScegliFile.Click
        rdoMsAccess.Checked = True
        OpenMSAccessDialog.ShowDialog()

        If OpenMSAccessDialog.FileName.Length And (OpenMSAccessDialog.FileName.EndsWith("mdb") Or OpenMSAccessDialog.FileName.EndsWith("accdb")) Then

            lblFile.Text = OpenMSAccessDialog.FileName

        End If

    End Sub

    Private Sub EsciToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        End
    End Sub

    Private Sub lnkSel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkSel.LinkClicked

        Dim I As Integer = 0

        For I = 0 To lstTabelle.Items.Count - 1
            lstTabelle.SetItemChecked(I, True)
        Next

    End Sub

    Private Sub lnkDesel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkDesel.LinkClicked
        Dim I As Integer = 0

        For I = 0 To lstTabelle.Items.Count - 1
            lstTabelle.SetItemChecked(I, False)
        Next

    End Sub

    Private Sub lnkElabora_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkElabora.LinkClicked

        Cursor.Current = Cursors.WaitCursor
        tabClassi.TabPages.Clear()
        txtRisSql.Text = ""
        RichCode = New List(Of RichTextBox)

        If Not StrutturaDati Is Nothing Then

            If cmbLanguage.SelectedIndex = 0 Then
                LinguaggioScelto = enLanguage.VbNet
            Else
                LinguaggioScelto = enLanguage.CSharp
            End If
            'azzero tutte le tabelle selezionate
            For Each Tab As cTabella In StrutturaDati.Tabelle
                Tab.Selezionata = False
            Next

            Dim I As Integer = 0

            For I = 0 To lstTabelle.CheckedItems.Count - 1

                'TODO: migliorare il metodo di ricerca con il find della lista

                Dim Tabella As String = lstTabelle.CheckedItems(I).ToString

                Dim Punto As Integer = Tabella.IndexOf(")")
                Tabella = Tabella.Substring(Punto + 2)
                Dim x As cTabella = StrutturaDati.Tabelle.Find(Function(item) item.NomeTabella = Tabella)
                x.Selezionata = True
                If chkOverWriteTable.Checked = True Or x.IsChanged = True Then

                    'qui cerco dentro le classi contenute all'interno del nome della tabella 
                    Dim Node As TreeNode = tvwTabelle.Nodes("t_" & x.NomeTabella)
                    x.NomeClasse = Node.Nodes(0).Tag
                    x.NomeClasseDAO = Node.Nodes(1).Tag

                    'qui devo caricare anche tutti i fields
                    For Each c As cCampoDb In x.Campi
                        Dim Key As String = "fs_" & x.NomeTabella & "_" & c.Nome
                        Dim NodeC As TreeNode = Node.Nodes(2).Nodes(Key)
                        c.NomeLogico = NodeC.Tag
                    Next


                End If
            Next

            'CODE CREATION
            LunaConverter.OverWriteAllTables = chkOverWriteTable.Checked
            Dim Ris As LunaResult = LunaConverter.ElaboraStruttura(StrutturaDati)

            txtRisSql.Text = Ris.BufferSQL

            Dim Chiavi As String = ""
            Dim Estensione As String = ""

            Select Case LinguaggioScelto
                Case enLanguage.VbNet
                    Chiavi = My.Resources.keywordColored.KeywordVB
                    Estensione = "vb"
                Case enLanguage.CSharp
                    Chiavi = My.Resources.keywordColored.KeywordCSharp
                    Estensione = "cs"
            End Select

            Dim tpBase As New TabPage
            Dim rtBase As New RichTextBox
            tpBase.Name = "LunaBaseClasses"
            tpBase.Text = "LunaBaseClasses"
            tabClassi.TabPages.Add(tpBase)
            rtBase.BorderStyle = BorderStyle.None
            rtBase.ReadOnly = True
            rtBase.Dock = DockStyle.Fill
            rtBase.ForeColor = Color.White
            rtBase.BackColor = Color.FromArgb(35, 35, 35)
            rtBase.Tag = "LunaEngine\ClassBase\LunaBaseClasses." & Estensione
            rtBase.Text = Ris.BufferClassiBase
            tpBase.Controls.Add(rtBase)
            RichCode.Add(rtBase)

            ColoraKey(rtBase, Chiavi, "'")

            Dim tpContext As New TabPage
            Dim rtContext As New RichTextBox
            tpContext.Name = "LunaContext"
            tpContext.Text = "LunaContext"
            tabClassi.TabPages.Add(tpContext)
            rtContext.BorderStyle = BorderStyle.None
            rtContext.ReadOnly = True
            rtContext.Dock = DockStyle.Fill
            rtContext.ForeColor = Color.White
            rtContext.BackColor = Color.FromArgb(35, 35, 35)
            rtContext.Tag = "LunaUserClass\ClassBase\LunaContext." & Estensione
            rtContext.Text = Ris.BufferContext
            tpContext.Controls.Add(rtContext)
            RichCode.Add(rtContext)

            ColoraKey(rtContext, Chiavi, "'")

            Dim tpConfig As New TabPage
            Dim rtConfig As New RichTextBox
            tpConfig.Name = "LunaConfig"
            tpConfig.Text = "LunaConfig"
            tabClassi.TabPages.Add(tpConfig)
            rtConfig.BorderStyle = BorderStyle.None
            rtConfig.ReadOnly = True
            rtConfig.Dock = DockStyle.Fill
            rtConfig.ForeColor = Color.White
            rtConfig.BackColor = Color.FromArgb(35, 35, 35)
            rtConfig.Tag = "Luna.config"
            rtConfig.Text = Ris.BufferConfig
            tpConfig.Controls.Add(rtConfig)
            RichCode.Add(rtConfig)

            For Each tbRis As LunaResultTableCode In Ris.ListTablesCode

                Dim tp As New TabPage
                Dim rt As New RichTextBox
                Dim tpDao As New TabPage
                Dim rtDao As New RichTextBox
                Dim tpEnt As New TabPage
                Dim rtEnt As New RichTextBox
                Dim tpInt As New TabPage
                Dim rtInt As New RichTextBox

                If tbRis.CodiceClasseEntity.Length Then
                    tp.Name = tbRis.NomeClasse
                    tp.Text = "Class Entity - " & tbRis.NomeClasse
                    tabClassi.TabPages.Add(tp)
                    rt.BorderStyle = BorderStyle.None
                    rt.ReadOnly = True
                    rt.Dock = DockStyle.Fill
                    rt.BackColor = Color.FromArgb(35, 35, 35)
                    rt.ForeColor = Color.White
                    rt.Tag = "LunaUserClass\ClassEntity\" & tbRis.NomeClasse & "." & Estensione
                    rt.Text = tbRis.CodiceClasseEntity

                    tp.Controls.Add(rt)
                    RichCode.Add(rt)

                    ColoraKey(rt, Chiavi, "'")

                    Dim tpTable As New TabPage
                    Dim rtTable As New RichTextBox
                    tpTable.Name = "_" & tbRis.NomeClasse
                    tpTable.Text = "TableClass - _" & tbRis.NomeClasse
                    tabClassi.TabPages.Add(tpTable)
                    rtTable.BorderStyle = BorderStyle.None
                    rtTable.ReadOnly = True
                    rtTable.Dock = DockStyle.Fill
                    rtTable.BackColor = Color.FromArgb(35, 35, 35)
                    rtTable.ForeColor = Color.White
                    rtTable.Tag = "LunaEngine\ClassDAO\_" & tbRis.NomeClasse & "." & Estensione
                    rtTable.Text = tbRis.CodiceTableClass
                    tpTable.Controls.Add(rtTable)
                    RichCode.Add(rtTable)
                    ColoraKey(rtTable, Chiavi, "'")
                End If

                tpDao.Name = "_" & tbRis.NomeClasseDAO
                tpDao.Text = "DAOClass - _" & tbRis.NomeClasseDAO
                tabClassi.TabPages.Add(tpDao)
                rtDao.BorderStyle = BorderStyle.None
                rtDao.ReadOnly = True
                rtDao.Dock = DockStyle.Fill
                rtDao.BackColor = Color.FromArgb(35, 35, 35)
                rtDao.ForeColor = Color.White
                rtDao.Tag = "LunaEngine\ClassDAO\_" & tbRis.NomeClasseDAO & "." & Estensione
                rtDao.Text = tbRis.CodiceClasseDAO
                tpDao.Controls.Add(rtDao)
                RichCode.Add(rtDao)
                ColoraKey(rtDao, Chiavi, "'")

                Dim tpDaoUser As New TabPage
                Dim rtDaoUser As New RichTextBox
                tpDaoUser.Name = tbRis.NomeClasseDAO
                tpDaoUser.Text = "DAOClassUser - " & tbRis.NomeClasseDAO
                tabClassi.TabPages.Add(tpDaoUser)
                rtDaoUser.BorderStyle = BorderStyle.None
                rtDaoUser.ReadOnly = True
                rtDaoUser.Dock = DockStyle.Fill
                rtDaoUser.BackColor = Color.FromArgb(35, 35, 35)
                rtDaoUser.ForeColor = Color.White
                rtDaoUser.Tag = "LunaUserClass\ClassDAO\" & tbRis.NomeClasseDAO & "." & Estensione
                rtDaoUser.Text = tbRis.CodiceClasseDAOUser
                tpDaoUser.Controls.Add(rtDaoUser)
                RichCode.Add(rtDaoUser)
                ColoraKey(rtDaoUser, Chiavi, "'")

            Next

            Chiavi = My.Resources.keywordColored.KeywordSQL
            ColoraKey(txtRisSql, Chiavi)

            TabRisCodice.SelectedIndex = 0
            tabMain.SelectedIndex += 1

        End If

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub ColoraKey(ByVal txtRif As RichTextBox, ByVal Chiavi As String, Optional ByVal Commenti As String = "")

        txtRif.AutoWordSelection = True
        Dim PuntoParola As Integer = 0, tipoRicerca As RichTextBoxFinds
        tipoRicerca = RichTextBoxFinds.WholeWord

        Dim Key() As String = Chiavi.Split(",")

        Dim SingParola As String
        For Each SingParola In Key
            PuntoParola = txtRif.Find(SingParola, tipoRicerca)
            While PuntoParola <> -1
                txtRif.SelectionColor = Color.Orange
                PuntoParola = txtRif.Find(SingParola, PuntoParola + 1, tipoRicerca)
            End While
        Next

        If Commenti.Length Then
            PuntoParola = txtRif.Find(Commenti)
            While PuntoParola <> -1

                Dim FineParola As Integer = txtRif.Find(vbCr, PuntoParola + 1, RichTextBoxFinds.None)
                If FineParola <> -1 Then
                    txtRif.SelectionStart = PuntoParola
                    txtRif.SelectionLength = FineParola - PuntoParola
                    txtRif.SelectionColor = Color.LightGreen
                End If
                PuntoParola = txtRif.Find(Commenti, (PuntoParola + 1), RichTextBoxFinds.MatchCase)
            End While
        End If

    End Sub

    Private Sub rdoNT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoNT.CheckedChanged, rdoSQL.CheckedChanged

        AggiornaStatoCampi()

    End Sub

    Private Sub AggiornaStatoCampi()

        txtSQLLogin.Enabled = Not rdoNT.Checked
        txtSQLPwd.Enabled = Not rdoNT.Checked

        lblLogin.Enabled = Not rdoNT.Checked
        lblPwd.Enabled = Not rdoNT.Checked

    End Sub

    Public Sub New()

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        ImpostaInfo()

    End Sub

    Private Sub lnkWeb_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)

        Dim x As New Process

        x.StartInfo.FileName = "http://www.diegolunadei.it"
        x.Start()

    End Sub

    Private Sub lnkMail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        Dim x As New Process

        x.StartInfo.FileName = "mailto:d.lunadei@gmail.com"
        x.Start()

    End Sub

    Private Sub lnkSave_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkSave.LinkClicked
        'salvo tutti i file in una cartella selezionata dall'utente

        If Not RichCode Is Nothing Then
            FolderClassDialog.Description = "Select folder to save generated class file and app.config"

            If FolderClassDialog.ShowDialog = Windows.Forms.DialogResult.OK Then

                Dim PathScelto As String = FolderClassDialog.SelectedPath

                Dim NomeCartella As String = "-" & Now.Year & Now.Month.ToString("00") & Now.Day.ToString("00") & "-" & Now.Hour.ToString("00") & "." & Now.Minute.ToString("00") & "." & Now.Second.ToString("00")
                PathScelto &= "\LunaCode" & NomeCartella
                CreateLongDir(PathScelto)
                'MkDir(PathScelto & "\Class")
                CreateLongDir(PathScelto & "\LunaEngine\ClassBase\")
                CreateLongDir(PathScelto & "\LunaEngine\ClassDAO\")
                CreateLongDir(PathScelto & "\LunaUserClass\ClassBase\")
                CreateLongDir(PathScelto & "\LunaUserClass\ClassDAO\")
                CreateLongDir(PathScelto & "\LunaUserClass\ClassEntity\")
                'MkDir(PathScelto & "\ClassPartial")

                'qui ho il path dove salvare tutti i file class
                Dim Pagina As RichTextBox

                For Each Pagina In RichCode

                    Dim Buffer As String = Pagina.Text
                    If Buffer.Length Then
                        Dim NomeFile As String = PathScelto & "\" & Pagina.Tag
                        Dim scrit As New StreamWriter(NomeFile)
                        scrit.Write(Buffer)
                        scrit.Close()
                    End If

                Next

                'SALVO IL DATASCHEMA

                Dim PathXml As String = PathScelto & "\DataSchema.lds"
                If chkSovrascriviLds.Checked = True And OriginalPathLunaPrj.Length > 0 Then
                    PathXml = OriginalPathLunaPrj
                End If

                'reinposto TipoOrigine a LunaDataSchema, perch?in lettura viene modificato
                StrutturaDati.TipoOrigine = enDatasourceType.LunaDataSchema
                For Each T As cTabella In StrutturaDati.Tabelle
                    T.IsChanged = False
                Next

                'Dim serial As XmlSerializer = New XmlSerializer(GetType(Luna.cDatabase), New Type() {GetType(cCampoDb), GetType(cTabella), GetType(cRelazioneTabella)})
                Dim serial As XmlSerializer = New XmlSerializer(GetType(cDatabase))

                Dim Writer As New System.IO.StreamWriter(PathXml)
                serial.Serialize(Writer, StrutturaDati)

                Writer.Close()

                'qui salvo anche il file sql
                Dim PathSql As String = PathScelto & "\DbScript.txt"

                If chkSovrascriviLds.Checked = True And OriginalPathLunaPrj.Length > 0 Then
                    PathSql = GetFolder(OriginalPathLunaPrj) & "\DbScript.txt"
                End If

                Using W As New StreamWriter(PathSql, False)
                    W.Write(txtRisSql.Text)
                End Using

                'MessageBox.Show("Generated class saved in: " & vbNewLine & PathScelto, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)

                ShellEx(PathScelto)


            End If


        End If

    End Sub

    Private Sub ShellEx(ByVal Path As String)
        Try

            Dim x As New Process

            x.StartInfo.FileName = Path
            x.Start()

        Catch ex As Exception

        End Try

    End Sub
    Private Sub lnkCopy_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkCopy.LinkClicked

        If Not tabClassi.SelectedTab Is Nothing Then

            If tabClassi.SelectedTab.Controls.Count Then
                Dim rt As RichTextBox = tabClassi.SelectedTab.Controls(0)

                Clipboard.SetText(rt.Text)

            End If

        End If

    End Sub
    Private Sub SelezionaTabelle(db As cDatabase)
        Dim x As cTabella
        Dim Node As TreeNode
        Dim a
        For Each x In db.Tabelle
            If x.Selezionata = True Then

                For i As Integer = 0 To lstTabelle.Items.Count - 1
                    a = lstTabelle.Items(i)

                    If lstTabelle.Items(i) = "(TABLE) " & x.NomeTabella Then

                        lstTabelle.SetItemChecked(i, True)
                        'tb.TipoTabellaString
                        'lstTabelle.Items.Add(
                    End If
                Next


                'Node = tvwTabelle.Nodes.Add("t_" & x.NomeTabella, x.NomeTabella, 1, 1)
                'Node.BackColor = Color.Green
                'Node.Tag = x.NomeTabella

                'Dim NodeC As TreeNode = Node.Nodes.Add("c_" & x.NomeTabella, "Generated Class: " & x.NomeClasse, 2, 2)
                'NodeC.Tag = x.NomeClasse
                'Dim NodecD As TreeNode = Node.Nodes.Add("cd_" & x.NomeTabella, "Generated DAO Class: " & x.NomeClasseDAO, 2, 2)
                'NodecD.Tag = x.NomeClasseDAO

                'Dim NodeF As TreeNode = Node.Nodes.Add("f_" & x.NomeTabella, "Fields", 3, 3)
                'For Each c As cCampoDb In x.Campi
                '    Dim NodeSf As TreeNode = NodeF.Nodes.Add("fs_" & x.NomeTabella & "_" & c.Nome, c.NomeLogico, 3, 3)
                '    NodeSf.Tag = c.NomeLogico
                'Next
                'Node.Expand()
            End If
        Next


    End Sub
    Private Sub lstTabelle_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lstTabelle.ItemCheck

        Dim tbl As String = lstTabelle.Items(e.Index)

        Dim Punto As Integer = tbl.IndexOf(")")
        tbl = tbl.Substring(Punto + 2)
        Dim x As cTabella = StrutturaDati.Tabelle.Find(Function(item) item.NomeTabella = tbl)

        Dim Node As TreeNode

        If e.NewValue = CheckState.Checked Then
            'NewColor = Color.Green
            'aggiungo il nodo per la tabella
            Node = tvwTabelle.Nodes.Add("t_" & x.NomeTabella, x.NomeTabella, 1, 1)
            Node.BackColor = Color.Green
            Node.Tag = x.NomeTabella

            Dim NodeC As TreeNode = Node.Nodes.Add("c_" & x.NomeTabella, "Generated Class: " & x.NomeClasse, 2, 2)
            NodeC.Tag = x.NomeClasse
            Dim NodecD As TreeNode = Node.Nodes.Add("cd_" & x.NomeTabella, "Generated DAO Class: " & x.NomeClasseDAO, 2, 2)
            NodecD.Tag = x.NomeClasseDAO

            Dim NodeF As TreeNode = Node.Nodes.Add("f_" & x.NomeTabella, "Fields", 3, 3)
            For Each c As cCampoDb In x.Campi
                Dim NodeSf As TreeNode = NodeF.Nodes.Add("fs_" & x.NomeTabella & "_" & c.Nome, c.NomeLogico, 3, 3)
                NodeSf.Tag = c.NomeLogico
            Next
            Node.Expand()
        Else
            'NewColor = Color.Red
            'tolgo il nodo per la tabella
            Dim NodeDel() As TreeNode = tvwTabelle.Nodes.Find("t_" & tbl, True)
            'Elimino il nodo
            tvwTabelle.Nodes.Remove(NodeDel(0))

        End If
        'tvwTabelle.ExpandAll()

        'Dim Node() As TreeNode = tvwTabelle.Nodes.Find("t_" & tbl, True)
        'Node(0).BackColor = NewColor

    End Sub

    Private Sub tvwTabelle_NodeMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvwTabelle.NodeMouseDoubleClick

        ChangeNameObject()

    End Sub

    Private Sub ChangeNameObject()
        Dim Node As TreeNode = tvwTabelle.SelectedNode
        If Not Node Is Nothing Then
            If Node.Name.StartsWith("c") Then
                CambiaNomeClasse()
            ElseIf Node.Name.StartsWith("fs_") Then
                'field name change
                CambiaNomeCampo()
            End If
        End If

    End Sub

    Private Sub lstTabelle_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstTabelle.SelectedIndexChanged

    End Sub

    Private Sub btnScegliPrj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnScegliPrj.Click
        rdoLuna.Checked = True
        OpenLunaPrj.ShowDialog()

        If OpenLunaPrj.FileName.Length And OpenLunaPrj.FileName.EndsWith("lds") Then

            lblFilePrj.Text = OpenLunaPrj.FileName

        End If

    End Sub

    Private Sub rdoLuna_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoLuna.CheckedChanged

    End Sub

    Private Sub rdoTutorial_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoTutorial.CheckedChanged
        CaricaTutorial()

    End Sub

    Private Sub rdoTwitter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoTwitter.CheckedChanged
        CaricaTutorial("http://twitter.com/#!/dlunadei")
    End Sub

    Private Sub rdoOfficial_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoOfficial.CheckedChanged
        CaricaTutorial("http://www.diegolunadei.it")
    End Sub

    Private Sub btnScegliCompact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnScegliCompact.Click

    End Sub

    Private Sub tvwTabelle_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvwTabelle.AfterSelect

    End Sub

    Private Sub lnkWeb_LinkClicked_1(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnkWeb.LinkClicked

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        rdoApache.Checked = True
        tabMain.SelectedTab = tpTutorial
    End Sub

    Private Sub tpTutorial_Click(sender As Object, e As EventArgs) Handles tpTutorial.Click

    End Sub

    Private Sub rdoApache_CheckedChanged(sender As Object, e As EventArgs) Handles rdoApache.CheckedChanged
        CaricaTutorial("http://www.apache.org/licenses/LICENSE-2.0.html")
    End Sub

End Class