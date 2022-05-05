Imports Solmicro.Expertis.Business.Bodega.BdgDAACabecera

Public Class BdgDAADetalleTransporte

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub


    Private Const cnEntidad As String = "tbDaaDetalleTransporte"

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
        If data.RowState = DataRowState.Added Then data("IDDaaDetalleTransporte") = Guid.NewGuid
    End Sub

    <Task()> Public Shared Function CrearDAADetalleTransporte(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        Dim bsnDetalleTransporte As BusinessHelper = BusinessHelper.CreateBusinessObject("BdgDAADetalleTransporte")

        Dim stISOPais As String = String.Empty
        If Length(data.Cabecera.Rows(0)("IDPaisTransportista")) > 0 Then
            Dim clsPais As New Pais
            Dim dtPais As DataTable = clsPais.SelOnPrimaryKey(data.Cabecera.Rows(0)("IDPaisTransportista") & String.Empty)
            If dtPais.Rows.Count = 1 Then
                stISOPais = dtPais.Rows(0)("CodigoISO") & String.Empty
            End If
        End If

        'Contenedor y Precinto
        If Length(data.Contenedor) > 0 Then
            If (data.DetalleTransporte Is Nothing) Then data.DetalleTransporte = bsnDetalleTransporte.AddNew
            Dim newRow As DataRow = data.DetalleTransporte.NewRow
            newRow("IDDaaDetalleTransporte") = Guid.NewGuid
            newRow("IDDaa") = data.Cabecera.Rows(0)("IDDaa")

            newRow("MedioTransporteInterno") = 1 'MedioTransporteInterno.Camion
            newRow("IDPaisTransporteInterno") = data.Cabecera.Rows(0)("IDPaisTransportista") & String.Empty
            newRow("ISOPaisTransporteInterno") = stISOPais
            newRow("UnidadTransporte") = 1 'unidadestransporte.contenedor
            newRow("IdentidadUnidadTransporte") = data.Contenedor
            If Length(data.Precinto) > 0 Then
                newRow("IdentidadPrecintoComercial") = data.Precinto
            End If
            data.DetalleTransporte.Rows.Add(newRow)
        End If

        'Matricula
        If Length(data.Matricula) > 0 Then
            If (data.DetalleTransporte Is Nothing) Then data.DetalleTransporte = bsnDetalleTransporte.AddNew
            Dim newRow As DataRow = data.DetalleTransporte.NewRow
            newRow("IDDaaDetalleTransporte") = Guid.NewGuid
            newRow("IDDaa") = data.Cabecera.Rows(0)("IDDaa")
            newRow("IDPaisTransporteInterno") = data.Cabecera.Rows(0)("IDPaisTransportista") & String.Empty
            newRow("ISOPaisTransporteInterno") = stISOPais
            newRow("MedioTransporteInterno") = 1 'MedioTransporteInterno.Camion
            newRow("UnidadTransporte") = 2 'unidadestransporte.vehículo
            newRow("IdentidadUnidadTransporte") = data.Matricula
            data.DetalleTransporte.Rows.Add(newRow)
        End If

        'Remolque
        If Length(data.Remolque) > 0 Then
            If (data.DetalleTransporte Is Nothing) Then data.DetalleTransporte = bsnDetalleTransporte.AddNew
            Dim newRow As DataRow = data.DetalleTransporte.NewRow
            newRow("IDDaaDetalleTransporte") = Guid.NewGuid
            newRow("IDDaa") = data.Cabecera.Rows(0)("IDDaa")
            newRow("MedioTransporteInterno") = 1 'MedioTransporteInterno.Camion
            newRow("IDPaisTransporteInterno") = data.Cabecera.Rows(0)("IDPaisTransportista") & String.Empty
            newRow("ISOPaisTransporteInterno") = stISOPais
            newRow("UnidadTransporte") = 3 'unidadestransporte.remolque
            newRow("IdentidadUnidadTransporte") = data.Remolque
            data.DetalleTransporte.Rows.Add(newRow)
        End If

        Return data
    End Function

#End Region

End Class