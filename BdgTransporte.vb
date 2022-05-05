Public Class BdgTransporte
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgTransporte"

#Region "RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosRepetidos)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTransporte")) = 0 Then ApplicationService.GenerateError("El Transporte es obligatorio.")
        If Length(data("DescTransporte")) = 0 Then ApplicationService.GenerateError("La descripción es obligatoria.")
    End Sub

    <Task()> Public Shared Sub ValidarDatosRepetidos(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            ProcessServer.ExecuteTask(Of String)(AddressOf ValidateDuplicateKey, data("IDTransporte"), services)
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub ValidatePrimaryKey(ByVal IDTransporte As String, ByVal services As ServiceProvider)
        Dim DtAux As DataTable = New BdgTransporte().SelOnPrimaryKey(IDTransporte)
        If Not DtAux Is Nothing AndAlso DtAux.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Transporte: | no existe en la tabla: |.", IDTransporte, cnEntidad)
        End If
    End Sub

    <Task()> Public Shared Sub ValidateDuplicateKey(ByVal IDTransporte As String, ByVal services As ServiceProvider)
        Dim DtAux As DataTable = New BdgTransporte().SelOnPrimaryKey(IDTransporte)
        If Not DtAux Is Nothing AndAlso DtAux.Rows.Count > 0 Then
            ApplicationService.GenerateError("El Transporte: | ya existe en la tabla: |", IDTransporte, cnEntidad)
        End If
    End Sub

End Class

Public Enum BdgTipoTransporte
    Transportista = 0
    Tractor = 1
    Remolque = 2
End Enum