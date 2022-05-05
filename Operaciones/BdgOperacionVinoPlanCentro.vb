Public Class BdgOperacionVinoPlanCentro

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgOperacionVinoPlanCentro"

#End Region

#Region "Eventos Entidad"

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDCentro", AddressOf BdgGeneral.CambioCentro)
        Obrl.Add("IDIncidencia", AddressOf BdgGeneral.CambioIncidencia)
        Return Obrl
    End Function

#End Region


#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCentro")) = 0 Then ApplicationService.GenerateError("El Centro es un dato obligatorio.")
    End Sub

#End Region

#End Region

End Class

<Serializable()> _
Public Class _OVPC
    Public Const IDOperacionVinoPlanCentro As String = "IDOperacionVinoPlanCentro"
    Public Const IDLineaOperacionVinoPlan As String = "IDLineaOperacionVinoPlan"
    Public Const IDOperacionPlanCentro As String = "IDOperacionPlanCentro"
    Public Const IDCentro As String = "IDCentro"
    Public Const Tiempo As String = "Tiempo"
    Public Const UDTiempo As String = "UDTiempo"
    Public Const PorCantidad As String = "PorCantidad"
End Class