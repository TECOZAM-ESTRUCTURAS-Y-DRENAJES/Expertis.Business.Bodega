Public Class BdgVendimiaVto

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVendimiaVto"

#End Region

#Region "Eventos Entidad"

    Public Overloads Function GetItemRow(ByVal IDVendimiaVto As Guid) As DataRow
        Dim dt As DataTable = New BdgVendimiaVto().SelOnPrimaryKey(IDVendimiaVto)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe el vencimiento de Vendimia")
        Else : Return dt.Rows(0)
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _Vvto
    Public Const IDVendimiaVto As String = "IDVendimiaVto"
    Public Const Vendimia As String = "Vendimia"
    Public Const Fecha As String = "Fecha"
    Public Const DescVto As String = "DescVto"
    Public Const TipoVto As String = "TipoVto"
    Public Const TipoCalculo As String = "TipoCalculo"
    Public Const PrecioKiloT As String = "PrecioKiloT"
    Public Const PrecioKiloB As String = "PrecioKiloB"
    Public Const IDTarifaT As String = "IDTarifaT"
    Public Const IDTarifaB As String = "IDTarifaB"
    Public Const Facturado As String = "Facturado"
End Class

Public Enum BdgTipoVto
    Liquidacion = 0
    Anticipo
    Bonificacion
End Enum