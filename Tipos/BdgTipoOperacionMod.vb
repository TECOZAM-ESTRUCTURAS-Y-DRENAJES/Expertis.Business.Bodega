Public Class BdgTipoOperacionMod

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgTipoOperacionMod"

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function SelOnIDTipoOperacion(ByVal IDTipoOperacion As String, ByVal services As ServiceProvider) As DataTable
        Return New BdgTipoOperacionMod().Filter(New StringFilterItem(_TOMD.IDTipoOperacion, FilterOperator.Equal, IDTipoOperacion))
    End Function

#End Region

End Class

<Serializable()> _
Public Class _TOMD
    Public Const IDTipoOperacion As String = "IDTipoOperacion"
    Public Const IDOperario As String = "IDOperario"
    Public Const Tiempo As String = "Tiempo"
End Class