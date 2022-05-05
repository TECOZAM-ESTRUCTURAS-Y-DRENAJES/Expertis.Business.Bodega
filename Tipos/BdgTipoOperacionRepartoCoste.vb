Public Class BdgTipoOperacionRepartoCoste

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgTipoOperacionRepartoCoste"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatos)
    End Sub

    <Task()> Public Shared Sub ValidarDatos(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipo")) = 0 Then ApplicationService.GenerateError("El Tipo es un dato obligatorio.")
        If Length(data("IDFamilia")) = 0 Then ApplicationService.GenerateError("La Familia es un dato obligatorio.")
        If Length(data("Porcentaje")) = 0 Then ApplicationService.GenerateError("El Porcentaje es un dato obligatorio.")
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function SelOnIDTipoOperacion(ByVal IDTipoOperacion As String, ByVal services As ServiceProvider) As DataTable
        Return New BdgTipoOperacionRepartoCoste().Filter(New StringFilterItem(_TORC.IDTipoOperacion, FilterOperator.Equal, IDTipoOperacion))
    End Function

#End Region

End Class

<Serializable()> _
Public Class _TORC
    Public Const IDTipoOperacion As String = "IDTipoOperacion"
    Public Const IDTipo As String = "IDTipo"
    Public Const IDFamilia As String = "IDFamilia"
    Public Const Porcentaje As String = "Porcentaje"
End Class