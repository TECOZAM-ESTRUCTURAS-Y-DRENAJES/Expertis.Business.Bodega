Public Class ObraParteAgrupadoMaterial
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbObraParteAgrupadoMaterial"

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarObraMaterialControl)
    End Sub

    <Task()> Public Shared Sub BorrarObraMaterialControl(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim OMC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraMaterialControl")
        Dim dt As DataTable = OMC.Filter(New GuidFilterItem("IDParteAgrupadoMat", data("IDParteAgrupadoMat")))
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                dr("IDParteAgrupadoMat") = DBNull.Value
            Next
            OMC.Delete(dt)
        End If
    End Sub

#End Region

    'Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
    '    Dim oBrl As New BusinessRules
    '    oBrl.Add("IDArticulo", AddressOf CambioIDArticulo)
    '    '    'oBrl.Add("Fecha", AddressOf CambioFecha)
    '    '    'oBrl.Add("IDAlmacen", AddressOf CambioIDAlmacen)
    '    '    'oBrl.Add("PrecioRealMatA", AddressOf CalcularImporte)
    '    '    'oBrl.Add("PrecioVentaA", AddressOf CalcularImporte)
    '    '    oBrl.Add("Cantidad", AddressOf CalcularImporte)
    '    '    'oBrl.Add("UDValoracion", AddressOf CalcularImporte)
    '    '    'oBrl.Add("Dto1", AddressOf CalcularImporte)
    '    '    'oBrl.Add("Dto2", AddressOf CalcularImporte)
    '    '    'oBrl.Add("Dto3", AddressOf CalcularImporte)
    '    '    'oBrl.Add("IDLineaMaterial", AddressOf CambioIDLineaMaterial)
    '    Return oBrl
    'End Function

    '<Task()> Public Shared Sub CambioIDArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
    '    If data.Value > 0 Then
    '        data.Current(data.ColumnName) = data.Value

    '    End If
    'End Sub

    ''<Task()> Public Shared Sub GetTarifaCosteArticulo(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
    ''    If Length(data("Cantidad")) > 0 AndAlso Length(data("IDMaterial")) > 0 Then
    ''        Dim dataCalculoTarifa As New DataCalculoTarifaComercial(data("IDMaterial"), data.Context("IDCliente"), CDbl(data("Cantidad")), Nz(data("Fecha"), Today))
    ''        Dim d As DataTarifaComercial = ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial, DataTarifaComercial)(AddressOf ProcesoComercial.TarifaComercial, dataCalculoTarifa, services)
    ''        If Not d Is Nothing AndAlso d.Precio > 0 Then
    ''            data("PrecioA") = d.Precio
    ''        End If
    ''    End If
    ''End Sub

#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDArticulo", AddressOf CambioArticulo)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) = 0 Then Return
        Dim strAlmacen As String
        If (Length(data.Current("IDParteAgrupadoMat")) = 0) Then data.Current("IDParteAgrupadoMat") = Guid.NewGuid
        If New Parametro().GestionBodegas AndAlso data.Context.ContainsKey("IDZonaFinca") Then
            Dim dtr As DataRow = New BdgZonaFinca().GetItemRow(data.Context("IDZonaFinca"))
            If Not dtr Is Nothing Then
                If Length(dtr("IDAlmacen")) > 0 Then
                    strAlmacen = dtr("IDAlmacen")
                ElseIf Length(dtr("IDExplotacionFinca")) > 0 Then
                    dtr = New BdgExplotacionFinca().GetItemRow(dtr("IDExplotacionFinca"))
                    If Not dtr Is Nothing AndAlso Length(dtr("IDAlmacen")) > 0 Then
                        strAlmacen = dtr("IDAlmacen")
                    End If
                End If
            End If
        End If
        If (Length(strAlmacen) = 0) Then
            Dim StDatos As New DataArtAlm
            StDatos.IDArticulo = data.Value
            strAlmacen = ProcessServer.ExecuteTask(Of DataArtAlm, String)(AddressOf ArticuloAlmacen.AlmacenPredeterminadoArticulo, StDatos, services)
        End If

        data.Current("IDAlmacen") = strAlmacen

    End Sub

#End Region

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDParteAgrupadoMat")) = 0 Then data("IDParteAgrupadoMat") = Guid.NewGuid 'AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region " RegisterValidateTask "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAlmacen")) = 0 Then ApplicationService.GenerateError("El almacén es un dato obligatorio.")
    End Sub

#End Region

End Class