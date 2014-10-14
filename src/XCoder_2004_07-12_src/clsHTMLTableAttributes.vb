Public Class clsHTMLTableAttributes
    Private _border As String
    Private _width As String
    Private _firstColWidth As String
    Private _cellPadding As String
    Private _cellSpacing As String
    Private _bgColor As String
    Private _foreColor As String
    Private _cols As String
    Private _rows As String


    Property border() As String
        Get
            Return Me._border
        End Get
        Set(ByVal Value As String)
            Me._border = Value
        End Set
    End Property

    Property width() As String
        Get
            Return Me._width
        End Get
        Set(ByVal Value As String)
            Me._width = Value
        End Set
    End Property

    Property cellPadding() As String
        Get
            Return Me._cellPadding
        End Get
        Set(ByVal Value As String)
            Me._cellPadding = Value
        End Set
    End Property

    Property cellSpacing() As String
        Get
            Return Me._cellSpacing
        End Get
        Set(ByVal Value As String)
            Me._cellSpacing = Value
        End Set
    End Property

    Property bgColor() As String
        Get
            Return Me._bgColor
        End Get
        Set(ByVal Value As String)
            Me._bgColor = Value
        End Set
    End Property

    Property foreColor() As String
        Get
            Return Me._foreColor
        End Get
        Set(ByVal Value As String)
            Me._foreColor = Value
        End Set
    End Property

    Property firstColWidth() As String
        Get
            Return Me._firstColWidth
        End Get
        Set(ByVal Value As String)
            Me._firstColWidth = Value
        End Set
    End Property

    Property rows() As String
        Get
            Return Me._rows
        End Get
        Set(ByVal Value As String)
            Me._rows = Value
        End Set
    End Property

    Property cols() As String
        Get
            Return Me._cols
        End Get
        Set(ByVal Value As String)
            Me._cols = Value
        End Set
    End Property

End Class
