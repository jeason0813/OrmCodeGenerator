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
Public Class cRelazioneTabella

    Private _TabellaOrigine As String = ""
    Public Property TabellaOrigine As String
        Get
            Return _TabellaOrigine
        End Get
        Set(ByVal value As String)
            _TabellaOrigine = value
        End Set
    End Property

    Private _TabellaDestinazione As String = ""
    Public Property TabellaDestinazione As String
        Get
            Return _TabellaDestinazione
        End Get
        Set(ByVal value As String)
            _TabellaDestinazione = value
        End Set
    End Property

    Private _CampoOrigine As String = ""
    Public Property CampoOrigine As String
        Get
            Return _CampoOrigine
        End Get
        Set(ByVal value As String)
            _CampoOrigine = value
        End Set
    End Property

    Private _CampoDestinazione As String = ""
    Public Property CampoDestinazione As String
        Get
            Return _CampoDestinazione
        End Get
        Set(ByVal value As String)
            _CampoDestinazione = value
        End Set
    End Property

    Public ReadOnly Property TbOrig As cTabella
        Get
            Dim _tbOrig As cTabella = Nothing
            If Not StrutturaRiferimento Is Nothing Then
                _tbOrig = StrutturaRiferimento.Tabelle.Find(Function(item) item.NomeTabella = TabellaOrigine)
            End If
            Return _tbOrig
        End Get
    End Property

    Public ReadOnly Property TbDest As cTabella
        Get
            Dim _tbDest As cTabella = Nothing
            If Not StrutturaRiferimento Is Nothing Then
                _tbDest = StrutturaRiferimento.Tabelle.Find(Function(item) item.NomeTabella = TabellaDestinazione)
            End If
            Return _tbDest
        End Get
    End Property

    Friend StrutturaRiferimento As cDatabase

End Class
