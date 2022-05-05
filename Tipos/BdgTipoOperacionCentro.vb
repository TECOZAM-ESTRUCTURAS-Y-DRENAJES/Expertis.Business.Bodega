Public Class BdgTipoOperacionCentro

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgTipoOperacionCentro"

#End Region

#Region "Funciones Públicas"

    Public Function SelOnIDTipoOperacion(ByVal IDTipoOperacion As String) As DataTable
        Return New BdgTipoOperacionCentro().Filter(New StringFilterItem(_TOC.IDTipoOperacion, FilterOperator.Equal, IDTipoOperacion))
    End Function

#End Region

End Class

<Serializable()> _
Public Class _TOC
    Public Const IDTipoOperacion As String = "IDTipoOperacion"
    Public Const IDCentro As String = "IDCentro"
    Public Const Tiempo As String = "Tiempo"
    Public Const PorCantidad As String = "PorCantidad"
End Class