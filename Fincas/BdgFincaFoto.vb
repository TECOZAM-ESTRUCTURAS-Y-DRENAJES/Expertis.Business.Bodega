Public Class BdgFincaFoto

#Region "Constuctor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgFincaFoto"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFinca")) = 0 Then ApplicationService.GenerateError("El código de la finca es obligatorio.")
        If Length(data("DescFoto")) = 0 Then ApplicationService.GenerateError("La descripción es obligatoria.")
    End Sub


#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function SelOnIDFinca(ByVal IDFinca As Guid, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter(cnEntidad, New GuidFilterItem("IDFinca", IDFinca), "IDFoto,IDFinca,DescFoto")
    End Function

#End Region

End Class