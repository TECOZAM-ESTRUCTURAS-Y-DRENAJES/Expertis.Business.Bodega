Public Class BdgEntradaVinoVariedad

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaVinoVariedad"

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("Porcentaje", AddressOf ProcesoEntradaVino.CambioPorcentaje)
        oBrl.Add("Cantidad", AddressOf ProcesoEntradaVino.CambioCantidad)
        Return oBrl
    End Function

#End Region

End Class