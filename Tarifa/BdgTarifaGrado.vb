Public Class BdgTarifaGrado

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgTarifaGrado"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data(_TG.IDTarifaGrado)) = 0 OrElse data(_TG.IDTarifaGrado) = 0 Then
                data(_TG.IDTarifaGrado) = AdminData.GetAutoNumeric
            End If
        End If
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _TG
    Public Const IDTarifaGrado As String = "IDTarifaGrado"
    Public Const IDTarifa As String = "IDTarifa"
    Public Const GradoDesde As String = "GradoDesde"
    Public Const Precio As String = "Precio"
End Class