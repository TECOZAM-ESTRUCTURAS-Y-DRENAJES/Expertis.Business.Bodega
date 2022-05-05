Public Class BdgContrato

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgContrato"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DataCont As New Contador.DatosDefaultCounterValue(data, "BdgContrato", _BdgContrato.NContrato)
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, DataCont, services)
        data(_BdgContrato.IDContrato) = Guid.NewGuid
        data(_BdgContrato.Fecha) = Date.Today
        data(_BdgContrato.TipoPrecio) = TipoPrecioContrato.PorLitro
        data(_BdgContrato.PrecioPorte) = 0
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarContador)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(_BdgContrato.IDContrato) = 0 OrElse CType(data(_BdgContrato.IDContrato), Guid).Equals(Guid.Empty) Then
                data(_BdgContrato.IDContrato) = Guid.NewGuid
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data(_BdgContrato.IDContador)) > 0 Then
                data(_BdgContrato.NContrato) = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data(_BdgContrato.IDContador), services)
            End If
        End If
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _BdgContrato
    Public Const IDContrato As String = "IDContrato"
    Public Const NContrato As String = "NContrato"
    Public Const IDContador As String = "IDContador"
    Public Const IDProveedor As String = "IDProveedor"
    Public Const Fecha As String = "Fecha"
    Public Const IDFormaPago As String = "IDFormaPago"
    Public Const IDCondicionPago As String = "IDCondicionPago"
    Public Const TipoPrecio As String = "TipoPrecio"
    Public Const PrecioPorte As String = "PrecioPorte"
End Class

Public Enum TipoPrecioContrato
    PorLitro
    PorHectoGrado
End Enum