Public Class BdgCartillistaMunicipio

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCartillistaMunicipio"

#End Region

#Region "Eventos Entidad"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("HaT", AddressOf CambiosHa)
        Obrl.Add("HaB", AddressOf CambiosHa)
        Obrl.Add("RdtoT", AddressOf CambiosRdto)
        Obrl.Add("RdtoB", AddressOf CambiosRdto)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambiosHa(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsDBNull(data.Current("IdCartillista")) Then
            If Not IsNumeric(data.Value) Then data.Value = 0
        End If
    End Sub

    <Task()> Public Shared Sub CambiosRdto(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsDBNull(data.Current("IdCartillista")) Then
            If Not IsNumeric(data.Value) Then data.Value = 100
        End If
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarVendimia)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCartillista")) = 0 Then ApplicationService.GenerateError("El Cartillista es obligatorio.")
        If Length(data("IDMunicipio")) = 0 Then ApplicationService.GenerateError("El municipio es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarVendimia(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If Length(data("Vendimia")) = 0 OrElse data("Vendimia") = 0 Then
                ApplicationService.GenerateError("El valor asignado a la vendimia no es válido.")
            End If
        End If
    End Sub

#End Region

End Class