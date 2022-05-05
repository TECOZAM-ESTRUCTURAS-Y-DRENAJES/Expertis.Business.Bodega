Public Class BdgFraDesglose

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgFraDesglose"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_FD.IDProveedor)) = 0 OrElse Length(data(_FD.IDProveedorFra)) = 0 Then
            ApplicationService.GenerateError("No se puede insertar un registro con identificadores de proveedor nulos.")
        ElseIf data(_FD.IDProveedor) = data(_FD.IDProveedorFra) Then
            ApplicationService.GenerateError("No se puede insertar un registro en el que los dos identificadores de proveedor son iguales.")
        End If
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _FD
    Public Const IDProveedor As String = "IDProveedor"
    Public Const TipoVariedad As String = "TipoVariedad"
    Public Const IDProveedorFra As String = "IDProveedorFra"
    Public Const Porcentaje As String = "Porcentaje"
End Class