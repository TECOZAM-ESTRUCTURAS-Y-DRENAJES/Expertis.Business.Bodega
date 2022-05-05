Imports Solmicro.Expertis.Business.Bodega.BdgDAACabecera

Public Class BdgDAALineaOperacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub


    Private Const cnEntidad As String = "tbDAALineaOperacion"

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
        If data.RowState = DataRowState.Added Then data("IDDaaLineaOperacion") = Guid.NewGuid
    End Sub

    <Serializable()> _
    Public Class stCrearDAALineaOperacionInfo
        Public IDDaaLinea As String
        Public dtDefaultDAAOperacion As DataTable

        Public Sub New(ByVal IDDaaLinea As String, ByVal dtDefaultDAAOperacion As DataTable)
            Me.IDDaaLinea = IDDaaLinea
            Me.dtDefaultDAAOperacion = dtDefaultDAAOperacion
        End Sub
    End Class

    <Task()> Public Shared Function CrearDaaLineasOperacion(ByVal data As stCrearDAALineaOperacionInfo, ByVal services As ServiceProvider) As DataTable
        Dim dttLinea As DataTable = New BdgDAALineaOperacion().AddNew
        If Not data Is Nothing AndAlso Not data.dtDefaultDAAOperacion Is Nothing AndAlso data.dtDefaultDAAOperacion.Rows.Count > 0 Then
            For Each dtr As DataRow In data.dtDefaultDAAOperacion.Rows
                Dim dtrNewRow As DataRow = dttLinea.NewRow
                dtrNewRow("IDDaaLinea") = data.IDDaaLinea
                dtrNewRow("CodigoManipulacionVitivinicola") = dtr("CodigoManipulacionVitivinicola")
                dttLinea.Rows.Add(dtrNewRow.ItemArray)
            Next
        End If
        Return dttLinea
    End Function

#End Region

End Class