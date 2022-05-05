Public Class BdgUsoSigPac
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgMaestroUsoSigPac"

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIdentificadorExistente)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDUsoSigpac")) = 0 Then ApplicationService.GenerateError("Debe indicar el Identificador")
        If Length(data("DescUsoSigpac")) = 0 Then ApplicationService.GenerateError("Debe indicar la descripción")
    End Sub

    <Task()> Public Shared Sub ValidarIdentificadorExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New BdgUsoSigPac().SelOnPrimaryKey(data("IDUsoSigpac"))
            If dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Uso SigPag {0} ya existe en el sistema.", Quoted(data("IDUsoSigpac")))
            End If
        End If
    End Sub

#End Region


End Class
