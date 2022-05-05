Public Class BdgOperacionVinoPlanMaterial

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgOperacionVinoPlanMaterial"

#End Region


#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim OBrl As New BusinessRules
        OBrl.Add("IDArticulo", AddressOf BdgGeneral.CambioMaterialVinoMaterial)
        OBrl.Add("Cantidad", AddressOf BdgGeneral.CambioCantidadVinoMaterial)
        OBrl.Add("Merma", AddressOf BdgGeneral.CambioMermaVinoMaterial)
        OBrl.Add("IDAlmacen", AddressOf BdgGeneral.CambioAlmacenVinoMaterial)
        Return OBrl
    End Function

#End Region


End Class

<Serializable()> _
Public Class _OVPMAT
    Public Const IDOperacionVinoPlanMaterial As String = "IDOperacionVinoPlanMaterial"
    Public Const IDLineaOperacionVinoPlan As String = "IDLineaOperacionVinoPlan"
    Public Const IDOperacionPlanMaterial As String = "IDOperacionPlanMaterial"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const Cantidad As String = "Cantidad"
    Public Const Merma As String = "Merma"
    Public Const RecalcularMaterial As String = "RecalcularMaterial"
End Class