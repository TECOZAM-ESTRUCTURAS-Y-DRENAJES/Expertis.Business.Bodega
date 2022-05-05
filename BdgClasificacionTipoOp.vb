Public Class BdgClasificacionTipoOp

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgClasificacionTipoOp"

#End Region

#Region "Eventos Entidad"

    Public Overloads Function GetItemRow(ByVal IDClasificacion As Integer) As DataRow
        Dim dt As DataTable = New BdgClasificacionTipoOp().SelOnPrimaryKey(IDClasificacion)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe la Clasificación del Tipo de Operación")
        Else : Return dt.Rows(0)
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _CTO
    Public Const IdClasificacion As String = "IdClasificacion"
    Public Const DescClasificacion As String = "DescClasificacion"
End Class