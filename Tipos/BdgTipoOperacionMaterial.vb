Public Class BdgTipoOperacionMaterial

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgTipoOperacionMaterial"

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function SelOnIDTipoOperacion(ByVal IDTipoOperacion As String, ByVal services As ServiceProvider) As DataTable
        Return New BdgTipoOperacionMaterial().Filter(New StringFilterItem(_TOM.IDTipoOperacion, FilterOperator.Equal, IDTipoOperacion))
    End Function

#End Region

End Class

<Serializable()> _
Public Class _TOM
    Public Const IDTipoOperacion As String = "IDTipoOperacion"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const Cantidad As String = "Cantidad"
End Class