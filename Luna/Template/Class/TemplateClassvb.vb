Namespace My.Templates

    Partial Public Class LunaBaseClassTvb
        Implements ILunaTemplate

        Private _Sorgente As cDatabase
        Public ReadOnly Property Sorgente As cDatabase
            Get
                Return _Sorgente
            End Get
        End Property

        Public Sub New(Sorgente As cDatabase)
            _Sorgente = Sorgente
        End Sub

        Public Function TransformTextEX() As String Implements ILunaTemplate.TransformTextEX
            Return TransformText()
        End Function

    End Class

    Partial Public Class LunaContextClassTvb
        Implements ILunaTemplate

        Public Function TransformTextEX() As String Implements ILunaTemplate.TransformTextEX
            Return TransformText()
        End Function

    End Class

    Partial Public Class LunaTableClassTvb
        Implements ILunaTemplate
        Private _Tabella As cTabella
        Private _Sorgente As cDatabase
        Public ReadOnly Property Tabella As cTabella
            Get
                Return _Tabella
            End Get
        End Property

        Public ReadOnly Property Sorgente As cDatabase
            Get
                Return _Sorgente
            End Get
        End Property

        Public Sub New(Sorgente As cDatabase, Tabella As cTabella)
            _Tabella = Tabella
            _Sorgente = Sorgente
        End Sub

        Public Function TransformTextEX() As String Implements ILunaTemplate.TransformTextEX
            Return TransformText()
        End Function

    End Class

    Partial Public Class LunaDAOClassTvb
        Implements ILunaTemplate
        Private _Tabella As cTabella
        Private _Sorgente As cDatabase
        Public ReadOnly Property Tabella As cTabella
            Get
                Return _Tabella
            End Get
        End Property

        Public ReadOnly Property Sorgente As cDatabase
            Get
                Return _Sorgente
            End Get
        End Property

        Public Sub New(Sorgente As cDatabase, Tabella As cTabella)
            _Tabella = Tabella
            _Sorgente = Sorgente
        End Sub

        Public Function TransformTextEX() As String Implements ILunaTemplate.TransformTextEX
            Return TransformText()
        End Function

    End Class

    Partial Public Class LunaDAOClassUserTvb
        Implements ILunaTemplate
        Private _Tabella As cTabella
        Private _Sorgente As cDatabase
        Public ReadOnly Property Sorgente As cDatabase
            Get
                Return _Sorgente
            End Get
        End Property

        Public Sub New(Sorgente As cDatabase, Tabella As cTabella)
            _Tabella = Tabella
            _Sorgente = Sorgente
        End Sub

        Public ReadOnly Property Tabella As cTabella
            Get
                Return _Tabella
            End Get
        End Property

        Public Function TransformTextEX() As String Implements ILunaTemplate.TransformTextEX
            Return TransformText()
        End Function

    End Class

    Partial Public Class LunaEntityClassTvb
        Implements ILunaTemplate
        Private _Tabella As cTabella
        Private _Sorgente As cDatabase
        Public ReadOnly Property Tabella As cTabella
            Get
                Return _Tabella
            End Get
        End Property

        Public ReadOnly Property Sorgente As cDatabase
            Get
                Return _Sorgente
            End Get
        End Property

        Public Sub New(Sorgente As cDatabase, Tabella As cTabella)
            _Tabella = Tabella
            _Sorgente = Sorgente
        End Sub

        Public Function TransformTextEX() As String Implements ILunaTemplate.TransformTextEX
            Return TransformText()
        End Function

    End Class

    Partial Public Class LunaModuleTvb
        Implements ILunaTemplate
        Public Function TransformTextEX() As String Implements ILunaTemplate.TransformTextEX
            Return TransformText()
        End Function
    End Class

    Partial Public Class LunaConfig
        Implements ILunaTemplate
        Private _Sorgente As cDatabase
        Public ReadOnly Property Sorgente As cDatabase
            Get
                Return _Sorgente
            End Get
        End Property
        Public Sub New(Sorgente As cDatabase)
            _Sorgente = Sorgente
        End Sub
        Public Function TransformTextEX() As String Implements ILunaTemplate.TransformTextEX
            Return TransformText()
        End Function
    End Class

    Partial Public Class LunaEntityInterfaceTvb
        Implements ILunaTemplate


        Private _Tabella As cTabella
        Private _Sorgente As cDatabase
        Public ReadOnly Property Tabella As cTabella
            Get
                Return _Tabella
            End Get
        End Property

        Public ReadOnly Property Sorgente As cDatabase
            Get
                Return _Sorgente
            End Get
        End Property

        Public Sub New(Sorgente As cDatabase, Tabella As cTabella)
            _Tabella = Tabella
            _Sorgente = Sorgente
        End Sub

        Public Function TransformTextEX() As String Implements ILunaTemplate.TransformTextEX
            Return TransformText()
        End Function

    End Class

End Namespace