Public Class BdgCuadernoCampoEstatalFertilizacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    ''' <summary>
    ''' Establezca el nombre de la tabla que corresponde con esta Entidad
    ''' </summary>
    ''' <remarks>NO SE OLVIDE DE CAMBIAR EL NOMBRE DE TABLA</remarks>
    Private Const cnEntidad As String = "tbBdgCuadernoCampoEstatalFertilizacion"

#End Region

#Region "Eventos Entidad"

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
        oBrl.Add("IDFinca", AddressOf CambioIDFinca)
        oBrl.Add("CFinca", AddressOf CambioCFinca)
        oBrl.Add("IDVariedad", AddressOf CambioVariedad)
        oBrl.Add("IDArticulo", AddressOf CambioArticulo)
        Return oBrl
    End Function
#End Region

#Region "Funciones Públicas"
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFinca")) = 0 Then ApplicationService.GenerateError("La Finca es un dato obligatorio.")
        Dim dt As DataTable = New BdgFinca().SelOnPrimaryKey(data("IDFinca"))
        If Length(dt) = 0 Then ApplicationService.GenerateError("La Finca no existe.")
        If Length(data("Fecha")) = 0 Then ApplicationService.GenerateError("La Fecha es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDCuadernoFertilizacion") = Guid.NewGuid
    End Sub

    <Task()> Public Shared Sub CambioIDFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalFertilizacion
            Dim dttFinca As DataTable = New BdgFinca().SelOnPrimaryKey(data.Current("IDFinca"))
            If (Not dttFinca Is Nothing AndAlso dttFinca.Rows.Count > 0) Then
                With dttFinca
                    data.Current("CFinca") = .Rows(0)("CFinca")
                    data.Current = ClsBdg.ApplyBusinessRule("CFinca", data.Current("CFinca"), data.Current)
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalFertilizacion
            Dim f As New StringFilterItem("CFinca", FilterOperator.Equal, data.Value)
            Dim dt As DataTable = New BdgFinca().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("IDFinca") = dt.Rows(0)("IDFinca")
                data.Current("IDVariedad") = dt.Rows(0)("IDVariedad")
                data.Current = ClsBdg.ApplyBusinessRule("IDVariedad", data.Current("IDVariedad"), data.Current)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioVariedad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim f As New StringFilterItem("IDVariedad", FilterOperator.Equal, data.Value)
            Dim dt As DataTable = New BdgVariedad().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("IDEspecie") = dt.Rows(0)("IDEspecie")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim dtArt As DataTable = New Articulo().SelOnPrimaryKey(data.Value)
            If (Not dtArt Is Nothing AndAlso dtArt.Rows.Count > 0) Then
                With dtArt
                    data.Current("DescArticulo") = .Rows(0)("DescArticulo")
                    data.Current("IDTipoAbono") = .Rows(0)("IDTipoAbono")
                    data.Current("RiquezaNPK") = .Rows(0)("RiquezaNPK")
                End With
            End If
        End If
    End Sub
#End Region

End Class