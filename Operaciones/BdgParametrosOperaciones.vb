Public Class BdgParametrosOperaciones

    Private mLotePorDefecto As String
    Private mLoteExplicitoEnBotellero As Boolean?
    Private mComprobarFechaOperacion As Boolean?
    Private mUnidadesCampoLitros As String

    Public ReadOnly Property LotePorDefecto() As String
        Get
            If mLotePorDefecto Is Nothing Then mLotePorDefecto = New BdgParametro().LotePorDefecto
            Return mLotePorDefecto
        End Get
    End Property


    Public ReadOnly Property LoteExplicitoEnBotellero() As Boolean
        Get
            If mLoteExplicitoEnBotellero Is Nothing Then mLoteExplicitoEnBotellero = New BdgParametro().LoteExplicitoEnBotellero
            Return mLoteExplicitoEnBotellero
        End Get
    End Property

    Public ReadOnly Property ComprobarFechaOperacion() As Boolean
        Get
            If mComprobarFechaOperacion Is Nothing Then mComprobarFechaOperacion = New BdgParametro().BdgComprobarFechaOperacion
            Return mComprobarFechaOperacion
        End Get
    End Property

    Public ReadOnly Property UnidadesCampoLitros() As String
        Get
            If mUnidadesCampoLitros Is Nothing Then mUnidadesCampoLitros = New BdgParametro().UnidadesCampoLitros
            Return mUnidadesCampoLitros
        End Get
    End Property

End Class

