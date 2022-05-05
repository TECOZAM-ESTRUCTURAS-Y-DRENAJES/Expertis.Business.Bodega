Public Class BdgCuadernoCampoEstatalAnalisisFitosanitario

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    ''' <summary>
    ''' Establezca el nombre de la tabla que corresponde con esta Entidad
    ''' </summary>
    ''' <remarks>NO SE OLVIDE DE CAMBIAR EL NOMBRE DE TABLA</remarks>
    Private Const cnEntidad As String = "tbBdgCuadernoCampoEstatalAnalisisFitosanitario"

#End Region

#Region "Eventos Entidad"

    ''' <summary>
    ''' Evento para establecer las tareas necesarias para valores por defecto para un nuevo registro de la entidad
    ''' </summary>
    ''' <param name="addnewProcess"></param>
    ''' <remarks>BORRAR PROCESO SI NO SE INSERTA NINGUNA TAREA</remarks>
    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
    End Sub

    ''' <summary>
    ''' Evento para establecer las tareas necesarias para validar los datos previo a su inserción o modificación de datos la entidad
    ''' </summary>
    ''' <param name="validateProcess"></param>
    ''' <remarks>BORRAR PROCESO SI NO SE INSERTA NINGUNA TAREA</remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    ''' <summary>
    ''' Evento para establecer las tareas necesarias para llevar a cabo establecimiento de datos, grabado en otras tablas.
    ''' Previo a la inserción o modificación de datos de la entidad
    ''' </summary>
    ''' <param name="updateProcess"></param>
    ''' <remarks>BORRAR PROCESO SI NO SE INSERTA NINGUNA TAREA</remarks>
    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

	
	''' <summary>
    ''' Evento para establecer las reglas de negocio necesarias para esta entidad
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>BORRAR PROCESO SI NO SE INSERTA NINGUNA TAREA</remarks>
    Public Overrides Function GetBusinessRules() As Solmicro.Expertis.Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDAnalisisCabecera", AddressOf CambioIDAnalisisCabecera)
        oBrl.Add("NAnalisisCabecera", AddressOf CambioNAnalisisCabecera)

        oBrl.Add("IDFinca", AddressOf CambioIDFinca)
        oBrl.Add("CFinca", AddressOf CambioCFinca)
        Return oBrl
    End Function
	
#End Region

#Region "Funciones Públicas"
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Fecha")) = 0 Then ApplicationService.GenerateError("La Fecha es un dato obligatorio.")
        If Length(data("IDMaterialAnalizado")) = 0 Then ApplicationService.GenerateError("El Material Analizado es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDCuadernoAnalisisFitosanitario") = Guid.NewGuid
    End Sub

    <Task()> Public Shared Sub CambioIDAnalisisCabecera(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalAnalisisFitosanitario
            Dim dt As DataTable = New BdgAnalisisCabecera().SelOnPrimaryKey(data.Current("IDAnalisisCabecera"))
            If (Not dt Is Nothing AndAlso dt.Rows.Count > 0) Then
                data.Current("NAnalisisCabecera") = dt.Rows(0)("NAnalisisCabecera")
                data.Current("Fecha") = dt.Rows(0)("Fecha")
                'Material Analizado
                If Length(dt.Rows(0)("IDAnalisis")) > 0 Then
                    Dim dtAnalisis As DataTable = New BdgAnalisis().SelOnPrimaryKey(dt.Rows(0)("IDAnalisis"))
                    If (Not dtAnalisis Is Nothing AndAlso dtAnalisis.Rows.Count > 0) Then
                        If Length(dtAnalisis.Rows(0)("IDMaterialAnalizado")) > 0 Then
                            data.Current("IDMaterialAnalizado") = dtAnalisis.Rows(0)("IDMaterialAnalizado")
                        End If
                    End If
                End If
                'Finca
                If dt.Rows(0)("TipoAnalisis") = 0 Then
                    data.Current("IDFinca") = dt.Rows(0)("Valor")
                    data.Current = ClsBdg.ApplyBusinessRule("IDFinca", data.Current("IDFinca"), data.Current)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioNAnalisisCabecera(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalAnalisisFitosanitario
            Dim f As New StringFilterItem("NAnalisisCabecera", FilterOperator.Equal, data.Value)
            Dim dt As DataTable = New BdgAnalisisCabecera().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("IDAnalisisCabecera") = dt.Rows(0)("IDAnalisisCabecera")
                data.Current = ClsBdg.ApplyBusinessRule("IDAnalisisCabecera", data.Current("IDAnalisisCabecera"), data.Current)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim dt As DataTable = New BdgFinca().SelOnPrimaryKey(data.Current("IDFinca"))
            If (Not dt Is Nothing AndAlso dt.Rows.Count > 0) Then
                data.Current("CFinca") = dt.Rows(0)("CFinca")
                data.Current("DescFinca") = dt.Rows(0)("DescFinca")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalAnalisisFitosanitario
            Dim f As New StringFilterItem("CFinca", FilterOperator.Equal, data.Value)
            Dim dt As DataTable = New BdgFinca().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("IDFinca") = dt.Rows(0)("IDFinca")
                data.Current = ClsBdg.ApplyBusinessRule("IDFinca", data.Current("IDFinca"), data.Current)
            End If
        End If
    End Sub
#End Region

End Class