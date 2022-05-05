Public Class BdgTipoOperacionVarios

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgTipoOperacionVarios"

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function SelOnIDTipoOperacion(ByVal IDTipoOperacion As String, ByVal services As ServiceProvider) As DataTable
        Return New BdgTipoOperacionVarios().Filter(New StringFilterItem(_TOV.IDTipoOperacion, FilterOperator.Equal, IDTipoOperacion))
    End Function

#End Region

End Class

<Serializable()> _
Public Class _TOV
    Public Const IDTipoOperacion As String = "IDTipoOperacion"
    Public Const IDVarios As String = "IDVarios"
    Public Const DescVarios As String = "DescVarios"
    Public Const Cantidad As String = "Cantidad"
End Class