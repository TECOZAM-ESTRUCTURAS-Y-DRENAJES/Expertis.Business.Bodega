Public Class ObraParteAgrupadoVarios
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbObraParteAgrupadoVarios"

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarObraVariosControl)
    End Sub

    <Task()> Public Shared Sub BorrarObraVariosControl(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim OVC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraVariosControl")
        Dim dt As DataTable = OVC.Filter(New GuidFilterItem("IDParteAgrupadoVarios", data("IDParteAgrupadoVarios")))
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                dr("IDParteAgrupadoVarios") = DBNull.Value
            Next
            OVC.Delete(dt)
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
            If Length(data("IDParteAgrupadoVarios")) = 0 Then data("IDParteAgrupadoVarios") = Guid.NewGuid 'AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

End Class