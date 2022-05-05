Public Class BdgOperacionPlanMaterial

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgOperacionPlanMaterial"

#End Region

#Region "Eventos Entidad"

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDArticulo", AddressOf BdgGeneral.CambioMaterialGlobal)
        Obrl.Add("Cantidad", AddressOf BdgGeneral.CambioCantidadGlobal)
        Obrl.Add("Merma", AddressOf BdgGeneral.CambioMermaGlobal)
        Obrl.Add("IDAlmacen", AddressOf BdgGeneral.CambioAlmacenGlobal)
        Return Obrl
    End Function

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
    End Sub

#End Region

#End Region

End Class