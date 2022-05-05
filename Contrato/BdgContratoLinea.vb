Public Class BdgContratoLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgContratoLinea"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarQRecibidaEstado)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data(_BdgContratoLinea.IDContratoLinea)) = 0 OrElse CType(data(_BdgContratoLinea.IDContratoLinea), Guid).Equals(Guid.Empty) Then
                data(_BdgContratoLinea.IDContratoLinea) = Guid.NewGuid
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarQRecibidaEstado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If data(_BdgContratoLinea.QRecibida) < 0 Then data(_BdgContratoLinea.QRecibida) = 0
            If data(_BdgContratoLinea.QRecibida) = 0 Then
                data(_BdgContratoLinea.Estado) = EstadoLineaContrato.Abierto
            Else
                If data(_BdgContratoLinea.QRecibida) < data(_BdgContratoLinea.Cantidad) Then
                    data(_BdgContratoLinea.Estado) = EstadoLineaContrato.Abierto
                Else : data(_BdgContratoLinea.Estado) = EstadoLineaContrato.Cerrado
                End If
            End If
        End If
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _BdgContratoLinea
    Public Const IDContratoLinea As String = "IDContratoLinea"
    Public Const IDContrato As String = "IDContrato"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const DescArticulo As String = "DescArticulo"
    Public Const Cantidad As String = "Cantidad"
    Public Const Precio As String = "Precio"
    Public Const TipoPrecio As String = "TipoPrecio"
    Public Const PrecioPorte As String = "PrecioPorte"
    Public Const Estado As String = "Estado"
    Public Const QRecibida As String = "QRecibida"
End Class

Public Enum EstadoLineaContrato
    Abierto
    Cerrado
End Enum