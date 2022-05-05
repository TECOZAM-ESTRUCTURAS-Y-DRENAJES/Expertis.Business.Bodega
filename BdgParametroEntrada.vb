Public Class BdgParametroEntrada

    Private mBdgDecimalesEntradas As Integer?

    Public ReadOnly Property BdgDecimalesEntradas() As Integer
        Get
            If mBdgDecimalesEntradas Is Nothing Then mBdgDecimalesEntradas = New BdgParametro().BdgDecimalesEntradas
            Return mBdgDecimalesEntradas
        End Get
    End Property

End Class

