Public Class BdgCuadernoCampoEstatalVenta

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    ''' <summary>
    ''' Establezca el nombre de la tabla que corresponde con esta Entidad
    ''' </summary>
    ''' <remarks>NO SE OLVIDE DE CAMBIAR EL NOMBRE DE TABLA</remarks>
    Private Const cnEntidad As String = "tbBdgCuadernoCampoEstatalVenta"

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
        oBrl.Add("IDArticulo", AddressOf CambioArticulo)
        oBrl.Add("IDAlbaran", AddressOf CambioIDAlbaran)
        oBrl.Add("NAlbaran", AddressOf CambioNAlbaran)
        oBrl.Add("IDCliente", AddressOf CambioIDCliente)
        oBrl.Add("IDFactura", AddressOf CambioIDFactura)
        oBrl.Add("NFactura", AddressOf CambioNFactura)
        oBrl.Add("IDLineaAlbaran", AddressOf CambioIDLineaAlbaran)
        oBrl.Add("IDLineaFactura", AddressOf CambioIDLineaFactura)
        Return oBrl
    End Function

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFinca")) = 0 Then ApplicationService.GenerateError("La Finca es un dato obligatorio.")
        Dim dt As DataTable = New BdgFinca().SelOnPrimaryKey(data("IDFinca"))
        If Length(dt) = 0 Then ApplicationService.GenerateError("La Finca no existe.")
        If Length(data("Fecha")) = 0 Then ApplicationService.GenerateError("La Fecha es un dato obligatorio.")
        If Length(data("IDCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDCuadernoVenta") = Guid.NewGuid
    End Sub

    <Task()> Public Shared Sub CambioCFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalVenta
            Dim f As New StringFilterItem("CFinca", FilterOperator.Equal, data.Value)
            Dim dt As DataTable = New BdgFinca().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("IDFinca") = dt.Rows(0)("IDFinca")
                data.Current = ClsBdg.ApplyBusinessRule("IDFinca", data.Current("IDFinca"), data.Current)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim dttFinca As DataTable = New BdgFinca().SelOnPrimaryKey(data.Current("IDFinca"))
            If (Not dttFinca Is Nothing AndAlso dttFinca.Rows.Count > 0) Then
                With dttFinca
                    data.Current("CFinca") = .Rows(0)("CFinca")
                    data.Current("DescFinca") = .Rows(0)("DescFinca")
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim dtArt As DataTable = New Articulo().SelOnPrimaryKey(data.Value)
            If (Not dtArt Is Nothing AndAlso dtArt.Rows.Count > 0) Then
                With dtArt
                    data.Current("DescArticulo") = .Rows(0)("DescArticulo")
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDAlbaran(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalVenta
            Dim dt As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(data.Value)
            If (Not dt Is Nothing AndAlso dt.Rows.Count > 0) Then
                With dt
                    data.Current("Fecha") = .Rows(0)("FechaAlbaran")
                    data.Current("NAlbaran") = .Rows(0)("NAlbaran")
                    data.Current("IDCliente") = .Rows(0)("IDCliente")
                    data.Current = ClsBdg.ApplyBusinessRule("IDCliente", data.Current("IDCliente"), data.Current)
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioNAlbaran(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalVenta
            Dim f As New StringFilterItem("NAlbaran", FilterOperator.Equal, data.Value)
            Dim dt As DataTable = New AlbaranVentaCabecera().Filter(f)
            If (Not dt Is Nothing AndAlso dt.Rows.Count > 0) Then
                With dt
                    data.Current("IDAlbaran") = .Rows(0)("IDAlbaran")
                    data.Current = ClsBdg.ApplyBusinessRule("IDAlbaran", data.Current("IDAlbaran"), data.Current)
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalVenta
            Dim dt As DataTable = New Cliente().SelOnPrimaryKey(data.Value)
            If (Not dt Is Nothing AndAlso dt.Rows.Count > 0) Then
                With dt
                    data.Current("RazonSocial") = .Rows(0)("RazonSocial")
                    data.Current("CifCliente") = .Rows(0)("CifCliente")
                    data.Current("Direccion") = .Rows(0)("Direccion")
                    data.Current("CodPostal") = .Rows(0)("CodPostal")
                    data.Current("Poblacion") = .Rows(0)("Poblacion")
                    data.Current("Provincia") = .Rows(0)("Provincia")
                    data.Current("IDPais") = .Rows(0)("IDPais")
                    data.Current("IDRGSEAA") = .Rows(0)("IDRGSEAA")
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDFactura(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalVenta
            Dim dt As DataTable = New FacturaVentaCabecera().SelOnPrimaryKey(data.Value)
            If (Not dt Is Nothing AndAlso dt.Rows.Count > 0) Then
                With dt
                    data.Current("NFactura") = .Rows(0)("NFactura")
                    data.Current("IDCliente") = .Rows(0)("IDCliente")
                    data.Current = ClsBdg.ApplyBusinessRule("IDCliente", data.Current("IDCliente"), data.Current)
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioNFactura(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalVenta
            Dim f As New StringFilterItem("NFactura", FilterOperator.Equal, data.Value)
            Dim dt As DataTable = New FacturaVentaCabecera().Filter(f)
            If (Not dt Is Nothing AndAlso dt.Rows.Count > 0) Then
                With dt
                    data.Current("IDFactura") = .Rows(0)("IDFactura")
                    data.Current = ClsBdg.ApplyBusinessRule("IDFactura", data.Current("IDFactura"), data.Current)
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDLineaAlbaran(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalVenta
            Dim dtLineas As DataTable = New AlbaranVentaLinea().SelOnPrimaryKey(data.Value)
            If (Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0) Then
                With dtLineas
                    data.Current("IDArticulo") = .Rows(0)("IDArticulo")
                    data.Current("DescArticulo") = .Rows(0)("DescArticulo")
                    data.Current("Cantidad") = .Rows(0)("QInterna")
                    data.Current("IDAlbaran") = .Rows(0)("IDAlbaran")
                    data.Current = ClsBdg.ApplyBusinessRule("IDAlbaran", data.Current("IDAlbaran"), data.Current)
                    'Lotes
                    Dim strLotes As String = LotesIDLineaAlbaran(data.Value)
                    If Length(strLotes) > 0 Then data.Current("Lote") = strLotes

                    'Finca
                    If Length(.Rows(0)("IDObra")) > 0 Then
                        Dim f As New StringFilterItem("IDObraCampaña", FilterOperator.Equal, .Rows(0)("IDObra"))
                        Dim dtFinca As DataTable = New BdgFinca().Filter(f)
                        If (Not dtFinca Is Nothing AndAlso dtFinca.Rows.Count > 0) Then
                            data.Current("IDFinca") = dtFinca.Rows(0)("IDFinca")
                            data.Current = ClsBdg.ApplyBusinessRule("IDFinca", data.Current("IDFinca"), data.Current)
                        End If
                    End If

                    'Factura Venta
                    Dim ff As New StringFilterItem("IDLineaAlbaran", FilterOperator.Equal, data.Value)
                    Dim dtLineasFV As DataTable = New FacturaVentaLinea().Filter(ff)
                    If (Not dtLineasFV Is Nothing AndAlso dtLineasFV.Rows.Count > 0) Then
                        data.Current("IDLineaFactura") = dtLineasFV.Rows(0)("IDLineaFactura")
                        data.Current = ClsBdg.ApplyBusinessRule("IDLineaFactura", data.Current("IDLineaFactura"), data.Current)
                    End If
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDLineaFactura(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalVenta
            Dim dtLineas As DataTable = New FacturaVentaLinea().SelOnPrimaryKey(data.Value)
            If (Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0) Then
                With dtLineas
                    If Length(data.Current("IDLineaAlbaran")) > 0 Then
                        data.Current("IDArticulo") = .Rows(0)("IDArticulo")
                        data.Current("DescArticulo") = .Rows(0)("DescArticulo")
                        data.Current("Cantidad") = .Rows(0)("QInterna")
                    Else
                        data.Current("IDLineaAlbaran") = .Rows(0)("IDLineaAlbaran")
                        data.Current = ClsBdg.ApplyBusinessRule("IDLineaAlbaran", data.Current("IDLineaAlbaran"), data.Current)
                    End If
                    data.Current("IDFactura") = .Rows(0)("IDFactura")
                    data.Current = ClsBdg.ApplyBusinessRule("IDFactura", data.Current("IDFactura"), data.Current)
                End With
            End If
        End If
    End Sub

    Public Shared Function LotesIDLineaAlbaran(ByVal IDLineaAlbaran As Long) As String
        Dim strLotes As String = String.Empty
        Dim f As New StringFilterItem("IDLineaAlbaran", FilterOperator.Equal, IDLineaAlbaran)
        Dim dtLotes As DataTable = New AlbaranVentaLote().Filter(f)
        If (Not dtLotes Is Nothing AndAlso dtLotes.Rows.Count > 0) Then
            Dim lstLotes As List(Of String) = (From c In dtLotes Select CStr(c("Lote")) Distinct).ToList
            strLotes = String.Join(", ", lstLotes.ToArray)
        End If
        Return strLotes
    End Function

#End Region

End Class