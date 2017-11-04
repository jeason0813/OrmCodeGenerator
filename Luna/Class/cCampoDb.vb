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
Public Class cCampoDb

    Private _Nome As String = String.Empty
    Public Property Nome() As String
        Get
            Return _Nome
        End Get
        Set(ByVal value As String)
            _Nome = value
        End Set
    End Property

    Private _NomeLogico As String = String.Empty
    Public Property NomeLogico() As String
        Get
            If _NomeLogico.Length Then
                Return _NomeLogico
            Else
                Return _Nome
            End If
        End Get
        Set(ByVal value As String)
            _NomeLogico = value
        End Set
    End Property

    Private _CampoChiave As Boolean = False
    Public Property CampoChiave As Boolean
        Get
            Return _CampoChiave
        End Get
        Set(ByVal value As Boolean)
            _CampoChiave = value
        End Set
    End Property

    Private _Contatore As Boolean = False
    Public Property Contatore() As Boolean
        Get
            Return _Contatore
        End Get
        Set(ByVal value As Boolean)
            _Contatore = value
        End Set
    End Property

    Private _AllowDBNull As Boolean = False
    Public Property AllowDBNull() As Boolean
        Get
            Return _AllowDBNull
        End Get
        Set(ByVal value As Boolean)
            _AllowDBNull = value
        End Set
    End Property

    Private _Tipo As String = String.Empty
    Public Property Tipo() As String
        Get
            Return _Tipo
        End Get
        Set(ByVal value As String)
            _Tipo = value
        End Set
    End Property

    Public ReadOnly Property ValoreDefault As String

        Get
            Dim ValDef As String = ""
            If LinguaggioScelto = enLanguage.VbNet Then
                Select Case TipoStringa.ToLower
                    Case "string"
                        ValDef &= " = """" "
                    Case "integer"
                        ValDef &= " = 0 "
                    Case "boolean"
                        ValDef &= " = False "
                    Case "single"
                        ValDef &= " = 0 "
                    Case "double"
                        ValDef &= " = 0 "
                    Case "decimal"
                        ValDef &= " = 0 "
                    Case Else
                        ValDef = " = Nothing "
                End Select
            Else
                Select Case TipoStringa.ToLower
                    Case "string"
                        ValDef &= " = """" "
                    Case "int"
                        ValDef &= " = 0 "
                    Case "bool"
                        ValDef &= " = false "
                    Case "single"
                        ValDef &= " = 0 "
                    Case "double"
                        ValDef &= " = 0 "
                    Case "decimal"
                        ValDef &= " = 0 "
                    Case "datetime"
                        ValDef &= " = DateTime.MinValue "
                    Case Else
                        ValDef = " = null "
                End Select
            End If

            Return ValDef
        End Get
    End Property

    Friend ReadOnly Property TipoStringa() As String

        Get
            Dim _TipoCampo As String = ""
            Dim TipoInterno As String = _Tipo
            If TipoInterno.StartsWith("System.") Then TipoInterno = TipoInterno.Substring(7)
            If LinguaggioScelto = enLanguage.VbNet Then
                Select Case TipoInterno.ToLower

                    Case "int32", "int16"
                        _TipoCampo = "integer"

                    Case "string"
                        _TipoCampo = "string"
                    Case "decimal"
                        _TipoCampo = "single"
                    Case Else
                        _TipoCampo = TipoInterno.Replace("[", "").Replace("]", "")

                End Select
            Else
                Select Case TipoInterno.ToLower

                    Case "int32", "int16"
                        _TipoCampo = "int"

                    Case "string"
                        _TipoCampo = "string"
                    Case "datetime"
                        _TipoCampo = "DateTime"

                    Case Else
                        _TipoCampo = TipoInterno.Replace("[", "").Replace("]", "")

                End Select
            End If

            Return _TipoCampo
        End Get
    End Property

    Private _DefaultValue As String = String.Empty
    Public Property DefaultValue() As String
        Get
            Return _DefaultValue
        End Get
        Set(ByVal value As String)
            _DefaultValue = value
        End Set
    End Property

    Private _MaxLength As Integer = 0
    Public Property MaxLength() As String
        Get
            Return _MaxLength
        End Get
        Set(ByVal value As String)
            _MaxLength = value
        End Set
    End Property

    'Public Property Attivo As Boolean = True

    'Private _Ordinal As Integer = 0
    'Public Property Ordinal() As String
    '    Get
    '        Return _Ordinal
    '    End Get
    '    Set(ByVal value As String)
    '        _Ordinal = value
    '    End Set
    'End Property

End Class
