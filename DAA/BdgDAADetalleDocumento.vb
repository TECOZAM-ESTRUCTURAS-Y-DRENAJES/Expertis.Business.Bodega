Imports Solmicro.Expertis.Business.Bodega.BdgDAACabecera

Public Class BdgDAADetalleDocumento

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub


    Private Const cnEntidad As String = "tbDaaDetalleDocumento"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDDaaDetalleDocumento") = Guid.NewGuid
    End Sub

    <Task()> Public Shared Function CrearDAADetalleDocumento(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        Dim bsnDetalleDocumento As BusinessHelper = BusinessHelper.CreateBusinessObject("BdgDAADetalleDocumento")
        Dim detalleDocumentoDefault As DataTable = bsnDetalleDocumento.Filter(New GuidFilterItem("IDDaa", data.DefaultDAA.Rows(0)("IDDaa")))
        If Not detalleDocumentoDefault Is Nothing AndAlso detalleDocumentoDefault.Rows.Count > 0 Then
            data.DetalleDocumento = detalleDocumentoDefault.Clone()
            For Each dtr As DataRow In detalleDocumentoDefault.Rows
                data.DetalleDocumento.Rows.Add(dtr.ItemArray)
            Next
        End If
        Return data
    End Function

#End Region

End Class