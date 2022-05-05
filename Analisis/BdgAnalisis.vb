Public Class BdgAnalisis

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgAnalisis"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarContador)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarOrden)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(_A.IDContador) = 0 Then data(_A.IDAnalisis) = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data(_A.IDContador), services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarOrden(ByVal Data As DataRow, ByVal services As ServiceProvider)
        Dim DtAnalisis As DataTable = New BdgAnalisisVariable().Filter(New FilterItem("IDAnalisis", FilterOperator.Equal, Data("IDAnalisis")), "Orden")
        If Not DtAnalisis Is Nothing AndAlso DtAnalisis.Rows.Count > 0 Then
            Dim i As Integer = 1
            For Each DrVar As DataRow In DtAnalisis.Select
                DrVar("Orden") = i
                i += 1
            Next
            BusinessHelper.UpdateTable(DtAnalisis)
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub ValidatePrimaryKey(ByVal IDAnalisis As String, ByVal Services As ServiceProvider)
        Dim dtAux As DataTable = New BdgAnalisis().SelOnPrimaryKey(IDAnalisis)
        If dtAux.Rows.Count = 0 Then
            ApplicationService.GenerateError("Actualización en conflicto con el valor de la clave | de la tabla |.", IDAnalisis, cnEntidad)
        End If
    End Sub

    <Task()> Public Shared Sub ValidateDuplicateKey(ByVal IDAnalisis As String, ByVal services As ServiceProvider)
        Dim dtAux As DataTable = New BdgAnalisis().SelOnPrimaryKey(IDAnalisis)
        If dtAux.Rows.Count > 0 Then ApplicationService.GenerateError("No se permite insertar una clave duplicada en la tabla |.", cnEntidad)
    End Sub

    <Serializable()> _
    Public Class StCrearAnalisis
        Public IDEntrada As Integer
        Public IDAnalisis As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDEntrada As Integer, ByVal IDAnalisis As String)
            Me.IDEntrada = IDEntrada
            Me.IDAnalisis = IDAnalisis
        End Sub
    End Class

    <Task()> Public Shared Function CrearAnalisis(ByVal data As StCrearAnalisis, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDAnalisis) > 0 Then
            Dim dtAV As DataTable = New BdgAnalisisVariable().Filter(New StringFilterItem(_AV.IDAnalisis, data.IDAnalisis))
            If dtAV.Rows.Count > 0 And Not dtAV Is Nothing Then
                Dim dtEA As DataTable = New Bodega.BdgEntradaAnalisis().AddNew
                For Each rwAV As DataRow In dtAV.Rows
                    Dim rwEA As DataRow = dtEA.NewRow
                    rwEA(_EA.IDEntrada) = data.IDEntrada
                    rwEA(_EA.IDVariable) = rwAV(_AV.IDVariable)
                    dtEA.Rows.Add(rwEA)
                Next
                Return dtEA
            End If
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _A
    Public Const IDAnalisis As String = "IDAnalisis"
    Public Const DescAnalisis As String = "DescAnalisis"
    Public Const IDContador As String = "IDContador"
End Class