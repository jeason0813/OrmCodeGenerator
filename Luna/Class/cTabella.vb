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
Public Class cTabella
    Implements ICloneable

    Private _CampoChiave As cCampoDb
    Public Property CampoChiave As cCampoDb
        Get
            Return _CampoChiave
        End Get
        Set(value As cCampoDb)
            _CampoChiave = value
        End Set
    End Property

    Private _Campi As New List(Of cCampoDb)
    Public Property Campi As List(Of cCampoDb)
        Get
            Return _Campi
        End Get
        Set(ByVal value As List(Of cCampoDb))
            _Campi = value
        End Set
    End Property

    Private _NomeTabella As String
    Public Property NomeTabella() As String
        Get
            Return _NomeTabella
        End Get
        Set(ByVal value As String)
            _NomeTabella = value
        End Set
    End Property

    Private _TipoTabella As enTableType = enTableType.Table
    Friend Property TipoTabella() As enTableType
        Get
            Return _TipoTabella
        End Get
        Set(ByVal value As enTableType)
            _TipoTabella = value
        End Set
    End Property

    Public Property TipoCampoChiaveString As String

    Public ReadOnly Property TipoTabellaString() As String
        Get
            If _TipoTabella = enTableType.Table Then
                Return "TABLE"
            Else
                Return "VIEW"
            End If
        End Get
    End Property

    Private _NomeClasse As String
    Public Property NomeClasse() As String
        Get
            Return _NomeClasse
        End Get
        Set(ByVal value As String)
            _NomeClasse = value
        End Set
    End Property

    Private _NomeClasseDAO As String
    Public Property NomeClasseDAO() As String
        Get
            Return _NomeClasseDAO
        End Get
        Set(ByVal value As String)
            _NomeClasseDAO = value
        End Set
    End Property

    Public ReadOnly Property NomeInterfaccia As String
        Get
            Return "I" & _NomeClasse
        End Get
    End Property


    Private _Selezionata As Boolean = False
    Public Property Selezionata As Boolean
        Get
            Return _Selezionata

        End Get
        Set(ByVal value As Boolean)
            _Selezionata = value
        End Set
    End Property

    Private _StrutturaRiferimento As cDatabase
    Friend Property StrutturaRiferimento As cDatabase
        Get
            Return _StrutturaRiferimento
        End Get
        Set(value As cDatabase)
            _StrutturaRiferimento = value
        End Set
    End Property

    Private _IsChanged As Boolean = False
    Public Property IsChanged As Boolean
        Get
            Return _IsChanged

        End Get
        Set(ByVal value As Boolean)
            _IsChanged = value
        End Set
    End Property

    Public Function Clone() As Object Implements ICloneable.Clone
        Return Me.MemberwiseClone
    End Function
End Class
