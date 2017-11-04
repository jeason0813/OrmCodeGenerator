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


<Serializable()>
Public Class cDatabase

    Private _Nome As String = ""
    Public Property Nome As String
        Get
            Return _Nome
        End Get
        Set(ByVal value As String)
            _Nome = value
        End Set
    End Property

    Private _TipoOrigine As String
    Public Property TipoOrigine() As String
        Get
            Return _TipoOrigine
        End Get
        Set(ByVal value As String)
            _TipoOrigine = value
        End Set
    End Property

    Private _TipoSorgente As Integer
    Public Property TipoSorgente() As Integer
        Get
            Return _TipoSorgente
        End Get
        Set(ByVal value As Integer)
            _TipoSorgente = value
        End Set
    End Property

    Private _OriginalTipoSorgente As Integer
    Public Property OriginalTipoSorgente() As Integer
        Get
            Return _OriginalTipoSorgente
        End Get
        Set(ByVal value As Integer)
            _OriginalTipoSorgente = value
        End Set
    End Property
    Private _PathDboServer As String = ""
    Public Property PathDboServer() As String
        Get
            Return _PathDboServer
        End Get
        Set(ByVal value As String)
            _PathDboServer = value
        End Set
    End Property
    Private _NomeDb As String = ""
    Public Property NomeDb() As String
        Get
            Return _NomeDb
        End Get
        Set(ByVal value As String)
            _NomeDb = value
        End Set
    End Property
    Private _AutenticazioneIntegrata As Boolean = False
    Public Property AutenticazioneIntegrata() As Integer
        Get
            Return _AutenticazioneIntegrata
        End Get
        Set(ByVal value As Integer)
            _AutenticazioneIntegrata = value
        End Set
    End Property

    Private _NomeUt As String = ""
    Public Property NomeUt() As String
        Get
            Return _NomeUt
        End Get
        Set(ByVal value As String)
            _NomeUt = value
        End Set
    End Property

    Private _Dbpwd As String = ""
    Public Property Dbpwd() As String
        Get
            Return _Dbpwd
        End Get
        Set(ByVal value As String)
            _Dbpwd = value
        End Set
    End Property

    Public Relazioni As New List(Of cRelazioneTabella)

    Public Tabelle As New List(Of cTabella)

End Class


