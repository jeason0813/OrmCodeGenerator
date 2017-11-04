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

Module mdlLuna
    Public ErroreDaGestire As Boolean = True
    Public LinguaggioScelto As enLanguage = enLanguage.VbNet

    Public Class cFormSopra
        Inherits Windows.Forms.Form

        Private OggettoPadre As Windows.Forms.Form

        Public Sub New(ByVal Oggetto As Windows.Forms.Form)
            Me.BackColor = System.Drawing.Color.Black
            'Me.ClientSize = New System.Drawing.Size(292, 266)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
            'Me.Name = "cFormSopra"
            Me.Opacity = 0.5
            Me.ShowIcon = False
            Me.ShowInTaskbar = False
            Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
            Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable

            Me.CausesValidation = False

            Me.ControlBox = False

            Me.MaximizeBox = False
            Me.MinimizeBox = False
            'Me.TopMost = True
            Me.ResumeLayout(False)
            OggettoPadre = Oggetto

            If OggettoPadre.FormBorderStyle = Windows.Forms.FormBorderStyle.None Then
                Left = OggettoPadre.Left '+ 8 'OggettoPadre.ClientRectangle.X
                Top = OggettoPadre.Top + (OggettoPadre.Height - OggettoPadre.ClientRectangle.Height) '- 4 'ClientRectangle.Top ' OggettoPadre.Top + 3

            Else
                Left = OggettoPadre.Left + 8 'OggettoPadre.ClientRectangle.X
                Top = OggettoPadre.Top + (OggettoPadre.Height - OggettoPadre.ClientRectangle.Height) - 8 'ClientRectangle.Top ' OggettoPadre.Top + 3

            End If

            'Me.Parent = OggettoPadre.Region

            'Me.WindowState = FormWindowState.Maximized

            Width = OggettoPadre.ClientRectangle.Width
            Height = OggettoPadre.ClientRectangle.Height

        End Sub

        Private Sub InitializeComponent()
            Me.SuspendLayout()
            '
            'cFormSopra
            '
            Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
            Me.BackColor = System.Drawing.Color.Black
            Me.CausesValidation = False
            Me.ClientSize = New System.Drawing.Size(292, 266)
            Me.ControlBox = False
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "cFormSopra"
            Me.Opacity = 0.5
            Me.ShowIcon = False
            Me.ShowInTaskbar = False
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.ResumeLayout(False)

        End Sub
    End Class

    Public Postazione As New cPostazione

    Public Cn As Data.Common.DbConnection

    'VARIE
    Public Function GestisciErrore(ByRef ex As Exception, Optional ByVal dettaglio As String = "") As Integer
        'gestione errori generale
        If ErroreDaGestire = False Then Exit Function

        Dim x As New Random

        Dim NomeFile As String = Postazione.PathApplicazione & "Err" & x.Next(0, 100000).ToString & ".txt"

        If dettaglio.Length = 0 Then
            If Not ex.InnerException Is Nothing Then
                If ex.InnerException.Message.Length Then dettaglio = ex.InnerException.Message
            End If
        End If

        Dim Messaggioerrore As String = "Inviare il seguente messaggio di posta elettronica all' indirizzo: d.lunadei@gmail.com" & ControlChars.NewLine & ControlChars.NewLine
        Messaggioerrore &= "OGGETTO EMAIL: Errore Luna" & ControlChars.NewLine & ControlChars.NewLine

        Messaggioerrore &= "TESTO EMAIL: " & ControlChars.NewLine & ControlChars.NewLine
        Messaggioerrore &= "Postazione: Luna Versione: " & Postazione.Version & ControlChars.NewLine
        Messaggioerrore &= "Data e ora: " & Date.Now & ControlChars.NewLine
        Messaggioerrore &= "Dettaglio: " & dettaglio & ControlChars.NewLine & ControlChars.NewLine
        Messaggioerrore &= "Errore: " & ex.Message & ControlChars.NewLine & ControlChars.NewLine
        Messaggioerrore &= "Sorgente: " & ex.Source & ControlChars.NewLine & ControlChars.NewLine
        Messaggioerrore &= "Stack: " & ControlChars.NewLine & ControlChars.NewLine & ex.StackTrace & ControlChars.NewLine
        Postazione.ScriviLogFile(NomeFile, Messaggioerrore)

        Dim ErrAvv As String = " Si è verificato un errore, si vuole inviare una email all'assistenza? " & ControlChars.NewLine & ControlChars.NewLine & "- premendo Si, verrà visualizzato il testo da inviare all'assistenza;" & ControlChars.NewLine & "- premendo No, il programma tenterà di continuare;" & ControlChars.NewLine & "- premendo Annulla, il programma verrà terminato;" & ControlChars.NewLine & ControlChars.NewLine & "In ogni caso il dettaglio dell'errore è stato salvato nel file " & NomeFile

        Dim res As System.Windows.Forms.DialogResult = MessageBox.Show(ErrAvv, Application.ProductName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Error)

        Select Case res
            Case DialogResult.Yes
                Dim p As New Process

                p.StartInfo.FileName = NomeFile
                p.Start()
                End
            Case DialogResult.No
                'qui devo cercare di capire se c'e' aperta la form di sottofondo per toglierla
            Case DialogResult.Cancel
                End

        End Select

        Return ex.GetHashCode

    End Function

    Function PulisciStringa(ByVal in_put As String) As String
        Dim clean_and_clear As String
        Dim car As Char
        clean_and_clear = " ,%/'();éòàùè@!+^&\?][{}*ç°§"
        For Each car In clean_and_clear
            in_put = Replace(in_put, car, "")
        Next
        Return Trim$(in_put)
    End Function

    Public Sub Sottofondo(ByVal Oggetto)
        If Oggetto.visible Then
            If Oggetto.formsopra Is Nothing Then

                Dim x As New cFormSopra(Oggetto)
                Oggetto.formsopra = x
                'x.StartPosition = FormStartPosition.CenterParent

                x.Show(Oggetto)

            Else

                Oggetto.Focus()
                Oggetto.formsopra.hide()
                Oggetto.formsopra.dispose()
                Oggetto.formsopra = Nothing

            End If
        End If
    End Sub

    Public Sub RoundBorder(ByRef FormRif As Windows.Forms.Form)


    End Sub


    Public Function SelectIndexCombo(ByVal combo As ComboBox, ByVal Valore As Integer) As Integer


        If IsDBNull(Valore) OrElse Valore = 0 Then SelectIndexCombo = -1

        combo.SelectedValue = Valore

        Return -1

    End Function

    Public Sub AvviaFile(ByVal PathMod As String)

        Try

            Dim x As New Process

            x.StartInfo.FileName = PathMod
            x.Start()

        Catch ex As Exception

            MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Public Function SelectIndexComboValore(ByVal combo As ComboBox, ByVal valore As String)

        If IsDBNull(valore) OrElse valore.Length = 0 Then SelectIndexComboValore = -1

        combo.SelectedValue = valore

        Return -1

    End Function


    Public Function SelectIndexCombo(ByVal combo As ComboBox, ByVal Valore As String) As Integer

        If IsDBNull(Valore) OrElse Valore.Trim = "" Then SelectIndexCombo = -1
        Dim i As Integer
        For i = 0 To combo.Items.Count - 1
            combo.SelectedIndex = i
            If combo.Items(i).ToString = Valore Then
                Return i
                Exit For
            End If
        Next
        Return -1
    End Function

  

    'STRINGHE

    Public Function Ap(ByVal Testo) As String

        Dim str As String

        str = Testo.ToString

        str = str.Replace("'", "''")
        str = "'" & str & "'"
        Return str

    End Function

    Public Function ApCancelletto(ByVal Testo) As String

        Dim str As String

        str = "#" & Testo.ToString & "#"

        Return str

    End Function

    Public Function ApLike(ByVal testo)
        Dim str As String
        str = testo.ToString

        str = str.Replace("'", "''")
        str = "like '%" & str & "%'"

        Return str

    End Function

    Public Function ApLikeRight(ByVal testo)
        Dim str As String
        str = testo.ToString

        str = str.Replace("'", "''")
        str = "like '" & str & "%'"

        Return str

    End Function

    Public Function FormattaPrezzo(ByVal Val As Double) As String

        Try
            Return Format(Val, "#,##0.00")
            'Return String.Format("{0:F2}", Val)
        Catch ex As Exception

        End Try
    End Function
    Public Function Cripta(ByVal pwd As String) As String
        Dim lunghezzaPwd As Integer
        Dim i As Integer
        Dim codAscii As Long
        Dim pwdCriptata As String

        'Let pwd = StrConv(pwd, vbUpperCase)
        lunghezzaPwd = Len(pwd)
        pwdCriptata = ""

        For i = 1 To lunghezzaPwd
            codAscii = Asc(Mid(pwd, i, 1))
            codAscii = 255 - codAscii
            pwdCriptata = pwdCriptata & Chr(codAscii)
        Next i

        Cripta = pwdCriptata
    End Function
    Public Function FormattaDouble(ByVal Val As Double) As String
        Try
            Return Format(Val, "0.00")
        Catch ex As Exception

        End Try
    End Function


    'FILE SYSTEM
    Public Sub CreateLongDir(ByVal sDir As String)
        Dim sBuild As String = ""

        Dim startpos As Integer = 2

        If sDir.StartsWith("\\") Then
            startpos = 3
        End If

        While InStr(startpos, sDir, "\") > 1
            sBuild = sBuild & Left(sDir, InStr(startpos, sDir, "\") - 1)
            sDir = Mid(sDir, InStr(startpos, sDir, "\"))
            If Dir(sBuild, 16) = "" Then
                System.IO.Directory.CreateDirectory(sBuild)
            End If
        End While
    End Sub

    Public Function GetNomeFile(ByVal NomeFile As String, Optional ByVal IdOrd As Integer = 0) As String

        Dim NuovoNome As String = ""

        Dim posizione As Integer = NomeFile.LastIndexOf("\")
        NuovoNome = NomeFile.Substring(posizione + 1)
        If IdOrd Then
            NuovoNome = IdOrd & "_" & NuovoNome
        End If
        Return NuovoNome

    End Function
    Public Function GetNomeFileFtp(ByVal NomeFile As String) As String

        Dim NuovoNome As String = ""

        Dim posizione As Integer = NomeFile.LastIndexOf("/")
        NuovoNome = NomeFile.Substring(posizione + 1)

        Return NuovoNome

    End Function

    Public Function GetNomeFileTemp(Optional ByVal extension As String = ".jpg") As String

        Dim Numero As New Random

        Randomize()

        Dim NomeFile As String = Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & Now.Millisecond & Numero.Next(0, 1000) & extension

        Return NomeFile

    End Function

    Public Function GetFolder(ByVal Path As String) As String

        Return Path.Substring(0, Path.LastIndexOf("\"))

    End Function

    Public Sub CreaCartella(ByVal nome As String)

        'crea una cartella nel percorso dell'applicazione

        Dim Path As String = Postazione.PathApplicazione & nome

        If System.IO.Directory.Exists(Path) = False Then

            System.IO.Directory.CreateDirectory(Path)

        End If

    End Sub

    'GRAFICA
    Public Sub ResizeImgPublic(ByVal PathOld As String, ByVal PathNew As String, Optional ByVal NumPixLong As Integer = 800, Optional ByVal NumPixShort As Integer = 600, Optional ByVal Watermark As Boolean = False)

        Dim Img As New Bitmap(PathOld)

        Dim width As Integer = 0, height As Integer = 0

        If Img.Width > Img.Height Then
            width = NumPixLong
            height = NumPixShort
        ElseIf Img.Width < Img.Height Then
            width = NumPixShort
            height = NumPixLong
        Else
            width = NumPixLong
            height = NumPixLong
        End If

        Dim ImgNew As Bitmap = New Bitmap(Img, New Size(width, height))

        If Watermark Then
            ImgNew = DrawWatermark(ImgNew, 10, 10)
        End If

        Dim ms As New MemoryStream()
        ImgNew.Save(ms, Imaging.ImageFormat.Jpeg)

        Dim imgData(ms.Length - 1) As Byte
        ms.Position = 0
        ms.Read(imgData, 0, ms.Length)
        Dim fs As New FileStream(PathNew, FileMode.Create, FileAccess.Write)
        fs.Write(imgData, 0, UBound(imgData))
        fs.Close()

    End Sub

    Friend Function DrawWatermark(ByVal result_bm As Bitmap, ByVal x As Integer, ByVal y As Integer) As Bitmap

        Const ALPHA As Byte = 128
        Dim watermark_bm As Bitmap = My.Resources.waterM
        ' Set the watermark's pixels' Alpha components.
        Dim clr As Color
        For py As Integer = 0 To watermark_bm.Height - 1
            For px As Integer = 0 To watermark_bm.Width - 1
                clr = watermark_bm.GetPixel(px, py)
                watermark_bm.SetPixel(px, py, _
                    Color.FromArgb(ALPHA, clr.R, clr.G, clr.B))
            Next px
        Next py

        ' Set the watermark's transparent color.
        watermark_bm.MakeTransparent(watermark_bm.GetPixel(0, _
            0))

        ' Copy onto the result image.
        Dim gr As Graphics = Graphics.FromImage(result_bm)
        gr.DrawImage(watermark_bm, x, y)

        Return result_bm
    End Function


End Module
