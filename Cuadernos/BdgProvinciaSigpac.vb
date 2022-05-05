Public Class BdgProvinciaSigpac
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub


    Private Const cnEntidad As String = "tbBdgMaestroProvinciaSigpac"



#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIdentificadorNumerico)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIdentificadorExistente)
    End Sub

    <Task()> Public Shared Sub ValidarIdentificadorNumerico(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDProvinciaSigpac")) > 0 AndAlso Not IsNumeric(data("IDProvinciaSigpac")) Then
            ApplicationService.GenerateError("El Identificador debe ser numérico.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("IDProvinciaSigpac"), 0) = 0 Then ApplicationService.GenerateError("Debe indicar el Identificador")
        If Length(data("DescProvinciaSigpac")) = 0 Then ApplicationService.GenerateError("Debe indicar la descripción")
    End Sub

    <Task()> Public Shared Sub ValidarIdentificadorExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New BdgProvinciaSigpac().SelOnPrimaryKey(data("IDProvinciaSigpac"))
            If dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("La Provincia {0} ya existe en el sistema.", Quoted(data("IDProvinciaSigpac")))
            End If
        End If
    End Sub

#End Region



End Class
