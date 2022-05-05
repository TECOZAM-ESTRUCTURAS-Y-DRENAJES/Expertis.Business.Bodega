Public Class ObraParteAgrupadoCentro
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbObraParteAgrupadoCentro"

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarObraCentroControl)
    End Sub

    <Task()> Public Shared Sub BorrarObraCentroControl(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim OCC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCentroControl")
        Dim dt As DataTable = OCC.Filter(New GuidFilterItem("IDParteAgrupadoCentro", data("IDParteAgrupadoCentro")))
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                dr("IDParteAgrupadoCentro") = DBNull.Value
            Next
            OCC.Delete(dt)
        End If
    End Sub

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDCentro", AddressOf CambioCentro)
        Obrl.Add("QHoras", AddressOf CalcularImporte)
        Obrl.Add("TasaA", AddressOf CalcularImporte)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioCentro(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim drCentro As DataRow = New Centro().GetItemRow(data.Value)
            data.Current("DescCentro") = drCentro("DescCentro")
            data.Current("TasaA") = drCentro("TasaEjecucionA")
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalcularImporte, data, services)
        Else
            data.Current("DescCentro") = DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CalcularImporte(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current(data.ColumnName) = data.Value
            data.Current("ImporteA") = Nz(data.Current("TasaA"), 0) * Nz(data.Current("QHoras"), 0)
        End If
    End Sub

#End Region


#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDParteAgrupadoCentro")) = 0 Then data("IDParteAgrupadoCentro") = Guid.NewGuid 'AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

End Class