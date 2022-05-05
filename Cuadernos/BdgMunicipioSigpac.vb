Public Class BdgMunicipioSigpac
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub


    Private Const cnEntidad As String = "tbBdgMaestroMunicipioSigpac"



#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIdentificadorExistente)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDProvinciaSigPac")) = 0 Then ApplicationService.GenerateError("Debe indicar el Identificador de la Provincia")
        If Length(data("IDMunicipioSigPac")) = 0 Then ApplicationService.GenerateError("Debe indicar el Identificador del Municipio")
        If Length(data("DescMunicipioSigPac")) = 0 Then ApplicationService.GenerateError("Debe indicar la descripción del Municipio")
    End Sub

    <Task()> Public Shared Sub ValidarIdentificadorExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New BdgMunicipioSigpac().SelOnPrimaryKey(data("IDProvinciaSigPac"), data("IDMunicipioSigPac"))
            If dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Municipio SigPag {0} ya existe en el sistema asociado a la Provincia {1}.", Quoted(data("IDMunicipioSigPac")), Quoted(data("IDProvinciaSigPac")))
            End If
        End If
    End Sub

#End Region



End Class
