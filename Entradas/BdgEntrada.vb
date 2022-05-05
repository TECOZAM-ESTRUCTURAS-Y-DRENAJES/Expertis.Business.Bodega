Public Class BdgEntrada
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntrada"

    Public Overloads Function GetItemRow(ByVal IDEntrada As Integer) As DataRow
        Dim dt As DataTable = New BdgEntrada().SelOnPrimaryKey(IDEntrada)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe la entrada")
        Else : Return dt.Rows(0)
        End If
    End Function

#Region " RegisterAddnewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dataCont As New Contador.DatosDefaultCounterValue(data, "BdgEntrada", _E.NEntrada)
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, dataCont, services)
        data("IDEntrada") = AdminData.GetAutoNumeric

        data("Fecha") = Date.Today
        data("Hora") = Date.Now.ToShortTimeString

        Dim DescBodega As String
        Dim fwnParam As New BdgParametro
        DescBodega = fwnParam.BodegaPredetEntradaUva()
        Dim v As System.Guid
        Dim oBodega As New BdgBodega
        Dim dtBodega As DataTable
        data("CosteKgUvaCalculada") = 0
        dtBodega = oBodega.Filter("IdBodega", "DescBodega = '" & DescBodega & "'")
        If dtBodega Is Nothing Or dtBodega.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe la Bodega Predeterminada")
        Else
            v = dtBodega.Rows(0)("IdBodega")
            If Length(v.ToString) <> 0 Then
                data("IdBodega") = v
            End If
        End If

        Dim ClsEntrada As New BdgEntrada
        data("IDArticuloCaja") = fwnParam.BdgArticuloCajaEU
        ClsEntrada.ApplyBusinessRule("IDArticuloCaja", data("IDArticuloCaja"), data)
        data("IDArticuloPalet") = fwnParam.BdgArticuloPaletEU
        ClsEntrada.ApplyBusinessRule("IDArticuloPalet", data("IDArticuloPalet"), data)

        Dim strTiporecogida As String = New BdgParametro().TipoRecogidaPredefinidoEntradas & String.Empty
        If Length(strTiporecogida) > 0 Then
            Dim dtrTipoRecogida As DataRow = New BdgTipoRecogida().GetItemRow(strTiporecogida)
            If Not dtrTipoRecogida Is Nothing Then
                data(_E.IDTipoRecogida) = strTiporecogida
                data(_E.IncrementoTipoRecogida) = dtrTipoRecogida("IncrementoTipoRecogida")
            End If
        End If
    End Sub

#End Region

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        'deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarContador)
    End Sub

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("Bruto", AddressOf CambioPesos)
        Obrl.Add("Tara", AddressOf CambioPesos)
        Obrl.Add("Neto", AddressOf CambioPesos)
        Obrl.Add("QDestrio", AddressOf CambioPesos)
        Obrl.Add("BrutoB", AddressOf CambioBrutoB)
        Obrl.Add("Declarado", AddressOf CambioDeclarado)
        Obrl.Add("IDTipoRecogida", AddressOf CambioTipoRecogida)

        Obrl.Add("IDArticuloCaja", AddressOf CambioIDArticuloCaja)
        Obrl.Add("IDArticuloPalet", AddressOf CambioIDArticuloPalet)
        Obrl.Add("PesoCaja", AddressOf CambioCajasPalets)
        Obrl.Add("NumeroCajas", AddressOf CambioCajasPalets)
        Obrl.Add("PesoPalet", AddressOf CambioCajasPalets)
        Obrl.Add("NumeroPalets", AddressOf CambioCajasPalets)
        Obrl.Add("PesoTotalEmbalaje", AddressOf CambioPesos)

        Obrl.Add("TieneAlbaranOrigen", AddressOf CambioPesos)
        Obrl.Add("PesoOrigen", AddressOf CambioPesos)

        Obrl.Add("CosteKgUva", AddressOf CambioCosteKgUva)
        Obrl.Add("IDTractor", AddressOf CambioTractor)
        Obrl.Add("IDCartillista", AddressOf CambioCartillista)
        Return Obrl
    End Function


    <Task()> Public Shared Sub CambioIDArticuloCaja(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        If Length(data.Current("NEntrada")) > 0 Then
            If Length(data.Current("IDArticuloCaja")) > 0 Then
                Dim dr As DataRow = New Articulo().GetItemRow(data.Current("IDArticuloCaja"))
                data.Current("PesoCaja") = dr("PesoNeto")
                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioCajasPalets, data, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDArticuloPalet(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        If Length(data.Current("NEntrada")) > 0 Then
            If Length(data.Current("IDArticuloPalet")) > 0 Then
                Dim dr As DataRow = New Articulo().GetItemRow(data.Current("IDArticuloPalet"))
                data.Current("PesoPalet") = dr("PesoNeto")
                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioCajasPalets, data, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCajasPalets(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim AppParams As BdgParametroEntrada = services.GetService(Of BdgParametroEntrada)()
        data.Current(data.ColumnName) = data.Value

        data.Current("PesoTotalCajas") = Math.Round(Nz(data.Current("PesoCaja"), 0) * Nz(data.Current("NumeroCajas"), 0), AppParams.BdgDecimalesEntradas)
        data.Current("PesoTotalPalets") = Math.Round(Nz(data.Current("PesoPalet"), 0) * Nz(data.Current("NumeroPalets"), 0), AppParams.BdgDecimalesEntradas)
        data.Current("PesoTotalEmbalaje") = Math.Round(Nz(data.Current("PesoTotalCajas"), 0) + Nz(data.Current("PesoTotalPalets"), 0), AppParams.BdgDecimalesEntradas)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioPesos, data, services)
    End Sub

    <Task()> Public Shared Sub CambioTipoRecogida(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        If Length(data.Current("IDTipoRecogida")) > 0 Then
            Dim dr As DataRow = New BdgTipoRecogida().GetItemRow(data.Current("IDTipoRecogida"))
            data.Current("IncrementoTipoRecogida") = dr("IncrementoTipoRecogida")
        Else
            data.Current("IncrementoTipoRecogida") = 0
        End If
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioPesos, data, services)
    End Sub

    <Task()> Public Shared Sub CambioBrutoB(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        If Length(data.Current("NEntrada")) > 0 Then
            If Not IsNumeric(data.Current("BrutoB")) Then
                ApplicationService.GenerateError("El campo {0} debe ser numérico.", Quoted(data.ColumnName))
            Else
                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioPesos, data, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioDeclarado(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        If Length(data.Current("NEntrada")) > 0 Then
            If Not IsNumeric(data.Current("Declarado")) Then
                ApplicationService.GenerateError("El campo {0} debe ser numérico.", Quoted(data.ColumnName))
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioPesos(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        If Length(data.Current("NEntrada")) > 0 Then
            If data.ColumnName = "Neto" Then
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf CalcularBruto, data.Current, services)
            Else
                'Es importante respetar el orden: 1º Peso, 2º Neto, 3º Declarado
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf CalcularPeso, data.Current, services)
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf CalcularNeto, data.Current, services)
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf CalcularDeclarado, data.Current, services)
            End If

            If Nz(data.Current("NumeroCajas"), 0) = 0 Then
                data.Current("KgPorCaja") = 0
            Else
                data.Current("KgPorCaja") = Nz(data.Current("Neto"), 0) / data.Current("NumeroCajas")
            End If

            If Nz(data.Current("NumeroPalets"), 0) = 0 Then
                data.Current("KgPorPalet") = 0
            Else
                data.Current("KgPorPalet") = Nz(data.Current("Neto"), 0) / data.Current("NumeroPalets")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCartillista(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If (Length(data.Value) > 0) Then
            Dim dtrC As DataRow = New BdgCartillista().GetItemRow(data.Value)
            If (Not dtrC Is Nothing AndAlso Length(dtrC("IDExplotadorFincas")) > 0) Then
                data.Current("IDExplotadorFincas") = dtrC("IDExplotadorFincas")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioTractor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If (Length(data.Value) > 0) Then
            Dim dtrTrans As DataRow = New BdgTransporte().GetItemRow(data.Value)
            If (Not dtrTrans Is Nothing) Then
                data.Current("Matricula") = dtrTrans("Matricula")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCosteKgUva(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("NEntrada")) > 0 Then data.Current("CosteKgUvaCalculada") = 0
    End Sub


#Region "  Cálculo de Pesos "

    <Task()> Public Shared Sub CalcularBruto(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim AppParams As BdgParametroEntrada = services.GetService(Of BdgParametroEntrada)()

        data("Bruto") = Math.Round(data("Neto") + Nz(data("Tara"), 0) + Nz(data("PesoTotalEmbalaje"), 0), AppParams.BdgDecimalesEntradas)
    End Sub

    <Task()> Public Shared Sub CalcularPeso(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim AppParams As BdgParametroEntrada = services.GetService(Of BdgParametroEntrada)()

        data("Peso") = Math.Round(Nz(data("Bruto"), 0) - Nz(data("Tara"), 0), AppParams.BdgDecimalesEntradas)
    End Sub

    <Task()> Public Shared Sub CalcularNeto(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim AppParams As BdgParametroEntrada = services.GetService(Of BdgParametroEntrada)()

        Dim DblNeto As Double
        If (Nz(data("TieneAlbaranOrigen"), False)) Then
            DblNeto = Math.Round(Nz(data("PesoOrigen"), 0), AppParams.BdgDecimalesEntradas)
        Else
            DblNeto = Math.Round(Nz(data("Bruto"), 0) - Nz(data("Tara"), 0) - Nz(data("PesoTotalEmbalaje"), 0) - Nz(data("QDestrio"), 0), AppParams.BdgDecimalesEntradas)
        End If
        If DblNeto < 0 Then
            ApplicationService.GenerateError("El campo Neto no puede ser inferior a cero.")
        End If
        data("Neto") = DblNeto
    End Sub

    <Task()> Public Shared Sub CalcularDeclarado(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim AppParams As BdgParametroEntrada = services.GetService(Of BdgParametroEntrada)()

        Dim DblDeclarado As Double
        If (Nz(data("TieneAlbaranOrigen"), False)) Then
            DblDeclarado = Math.Round(Nz(data("PesoOrigen"), 0), AppParams.BdgDecimalesEntradas)
        Else
            Dim dblBruto As Double
            If Nz(data("BrutoB"), 0) = 0 Then
                dblBruto = Nz(data("Bruto"), 0)
            Else
                dblBruto = Nz(data("BrutoB"), 0)
            End If
            DblDeclarado = Math.Round((dblBruto - Nz(data("Tara"), 0) - Nz(data("PesoTotalEmbalaje"), 0) - Nz(data("QDestrio"), 0)) * (1 + Nz(data("IncrementoTipoRecogida"), 0) / 100), AppParams.BdgDecimalesEntradas)
        End If
        data("Declarado") = DblDeclarado
    End Sub
#End Region


#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarEntidadesSecundarias)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If (New BdgParametro().GestionExplotadorObligatoria) Then
            If Length(data(_E.IDExplotadorFincas)) = 0 Then ApplicationService.GenerateError("El Explotador es obligatorio.")
        End If
        If (New BdgParametro().TipoRecogidaObligatorioEntradas) Then
            If Length(data(_E.IDTipoRecogida)) = 0 Then ApplicationService.GenerateError("El tipo de recogida es obligatorio.")
        End If
        If Length(data(_E.Vendimia)) = 0 Then ApplicationService.GenerateError("El valor asignado a la Vendimia no es válido.")
        If Length(data(_E.IDCartillista)) = 0 Then ApplicationService.GenerateError("El Cartillista es obligatorio.")
        If Length(data(_E.IDVariedad)) = 0 Then ApplicationService.GenerateError("La Variedad de vino es obligatoria.")
        If Length(data(_E.Neto)) = 0 OrElse data(_E.Neto) < 0 Then ApplicationService.GenerateError("El peso neto no es válido.")
        If Length(data(_E.Declarado)) = 0 OrElse data(_E.Declarado) < 0 Then ApplicationService.GenerateError("El peso declarado no es válido.")
        If Length(data(_E.Fecha)) = 0 Then
            If data.RowState <> DataRowState.Added Then ApplicationService.GenerateError("La Fecha no es válida.")
        End If
        If Length(data(_E.Hora)) = 0 Then
            If data.RowState <> DataRowState.Added Then ApplicationService.GenerateError("La hora no es válida.")
        End If
        If Length(data(_E.NEntrada)) = 0 Then ApplicationService.GenerateError("El Nº de Entrada no es válido.")
    End Sub

    <Task()> Public Shared Sub ValidarEntidadesSecundarias(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_E.IDFinca)) > 0 Then ProcessServer.ExecuteTask(Of String)(AddressOf BdgFinca.ValidatePrimaryKey, data(_E.IDFinca).ToString, services)
        If Length(data(_E.IdBodega)) > 0 Then ProcessServer.ExecuteTask(Of String)(AddressOf BdgBodega.ValidatePrimaryKey, data(_E.IdBodega).ToString, services)
        ProcessServer.ExecuteTask(Of String)(AddressOf BdgVariedad.ValidatePrimaryKey, data(_E.IDVariedad), services)
        If Length(data(_E.IDMunicipio)) > 0 Then ProcessServer.ExecuteTask(Of String)(AddressOf BdgMunicipio.ValidatePrimaryKey, data(_E.IDMunicipio), services)
        If Length(data(_E.IDTractor)) > 0 OrElse Length(data(_E.IDRemolque)) > 0 Then
            If Length(data(_E.IDTractor)) > 0 Then ProcessServer.ExecuteTask(Of String)(AddressOf BdgTransporte.ValidatePrimaryKey, data(_E.IDTractor), services)
            If Length(data(_E.IDRemolque)) > 0 Then ProcessServer.ExecuteTask(Of String)(AddressOf BdgTransporte.ValidatePrimaryKey, data(_E.IDRemolque), services)
        End If
        If Length(New BdgParametro().AnalisisEntrada) > 0 Then
            ProcessServer.ExecuteTask(Of String)(AddressOf BdgAnalisis.ValidatePrimaryKey, New BdgParametro().AnalisisEntrada, services)
        End If
    End Sub
#End Region

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of UpdatePackage, BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.CrearDocumento)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.ComprobarTotales)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.AsignarValores)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.TratarVariedades)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.TratarProveedores)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.TratarFincas)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.TratarCartillista)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.TratarAnalisis)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.TratarFacturacion)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.TratarDepositos)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.Recalculos)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.RecalcularEntradaFacturacion)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf Comunes.UpdateDocument)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of BdgDocumentoEntrada)(AddressOf BdgProcesoEntrada.CambioFechaOperacionMovimiento)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data(_E.IDEntrada)) = 0 Then data(_E.IDEntrada) = AdminData.GetAutoNumeric()
        End If
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data(_E.IDContador)) > 0 Then
                data(_E.NEntrada) = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data(_E.IDContador), services)
            End If
        End If
    End Sub

#End Region

#Region " Funciones Públicas "

    <Serializable()> _
    Public Class clsEstadoFacturacion
        Public IDEntrada As Integer
        Public blPermitirEdicionFacturacion As Boolean
        Public strEstado As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDEntrada As Integer)
            Me.IDEntrada = IDEntrada
        End Sub
    End Class

    <Task()> Public Shared Function PermitirEdicionFacturacion(ByVal data As clsEstadoFacturacion, ByVal services As ServiceProvider) As clsEstadoFacturacion
        data.strEstado = "Sin Estado"
        data.blPermitirEdicionFacturacion = False

        Dim strVista As String = "select ev.IDEntrada, vv.TipoVto, pv.IDFactura"
        strVista &= " from tbBdgProveedorVto pv"
        strVista &= " inner join tbBdgEntradaVto ev on pv.IDVendimiaVto = ev.IDVendimiaVto and pv.IDProveedor = ev.IDProveedor and pv.TipoVariedad = ev.TipoVariedad"
        strVista &= " inner join tbBdgVendimiaVto vv on pv.IDVendimiaVto = vv.IDVendimiaVto"

        Dim f As New Filter
        f.Add("ev.IDEntrada", data.IDEntrada)
        Dim strWhere As String = f.Compose(New AdoFilterComposer)
        Dim dtEntradasVto As DataTable = AdminData.Execute(strVista & " where " & strWhere, ExecuteCommand.ExecuteReader, False)
        If dtEntradasVto Is Nothing OrElse dtEntradasVto.Rows.Count = 0 Then
            data.strEstado = "Sin Facturar"
            data.blPermitirEdicionFacturacion = True
        Else
            Dim dblFacturas As Long = (Aggregate c In dtEntradasVto Where Not c.IsNull("IDFactura") And (Not c.IsNull("TipoVto") AndAlso c("TipoVto") <> BdgTipoVto.Anticipo) Into Count(c("IDEntrada")))
            If dblFacturas > 0 Then
                data.strEstado = "Facturada (Liquidación o Bonificación)"
                data.blPermitirEdicionFacturacion = False
            Else
                Dim dblNoAnticipo As Long = (Aggregate c In dtEntradasVto Where c.IsNull("IDFactura") And (Not c.IsNull("TipoVto") AndAlso c("TipoVto") <> BdgTipoVto.Anticipo) Into Count(c("IDEntrada")))
                If dblNoAnticipo > 0 Then
                    data.strEstado = "Liquidación o Bonificación Sin Facturar"
                    data.blPermitirEdicionFacturacion = False
                Else
                    Dim dblAnticipo As Long = (Aggregate c In dtEntradasVto Where Not c.IsNull("IDFactura") And (Not c.IsNull("TipoVto") AndAlso c("TipoVto") = BdgTipoVto.Anticipo) Into Count(c("IDEntrada")))
                    If dblAnticipo > 0 Then
                        data.strEstado = "Facturada (Anticipo)"
                        data.blPermitirEdicionFacturacion = True
                    Else
                        data.strEstado = "Anticipo Sin Facturar"
                        data.blPermitirEdicionFacturacion = True
                    End If
                End If
            End If
        End If
        Return data
    End Function

    <Task()> Public Shared Function TipoVariedad(ByVal IDVariedad As String, ByVal services As ServiceProvider) As Integer
        Return New BdgVariedad().GetItemRow(IDVariedad)("TipoVariedad")
    End Function

    <Task()> Public Shared Sub ValidatePrimaryKey(ByVal IDEntrada As Integer, ByVal services As ServiceProvider)
        Dim DtAux As DataTable = New BdgEntrada().SelOnPrimaryKey(IDEntrada)
        If DtAux.Rows.Count = 0 Then ApplicationService.GenerateError("Actualización en conflicto con el valor de la clave | de la tabla |.", IDEntrada, cnEntidad)
    End Sub

    <Task()> Public Shared Sub ValidateDuplicateKey(ByVal IDEntrada As Integer, ByVal services As ServiceProvider)
        Dim DtAux As DataTable = New BdgEntrada().SelOnPrimaryKey(IDEntrada)
        If DtAux.Rows.Count > 0 Then ApplicationService.GenerateError("No se permite insertar una clave duplicada en la tabla |.", cnEntidad)
    End Sub

    <Serializable()> _
    Public Class StObtenerDatos
        Public NCartilla As String
        Public Vendimia As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal NCartilla As String, ByVal Vendimia As Integer)
            Me.NCartilla = NCartilla
            Me.Vendimia = Vendimia
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerDatosConNCartilla(ByVal data As StObtenerDatos, ByVal services As ServiceProvider) As DataTable
        Dim e As New Filter
        e.Add("NCartilla", FilterOperator.Equal, data.NCartilla, FilterType.String)
        e.Add("Vendimia", FilterOperator.Equal, data.Vendimia, FilterType.Numeric)
        Return New BE.DataEngine().Filter("frmBdgCartillistaVendimiaCartillista", e)
    End Function

    <Task()> Public Shared Sub AccioncalcularCosteKgMasivo(ByVal dtObra As DataTable, ByVal services As ServiceProvider)
        For Each drObra As DataRow In dtObra.Rows
            Dim dteInicio As Date = System.Data.SqlTypes.SqlDateTime.MinValue
            Dim dteFin As Date = System.Data.SqlTypes.SqlDateTime.MaxValue

            If Length(drObra("FechaInicio")) > 0 Then dteInicio = drObra("FechaInicio")
            If Length(drObra("FechaFin")) > 0 Then dteFin = drObra("FechaFin")

            Dim StAccion As New StAccionCalcCosteKg(drObra("IDObra"), dteInicio, dteFin)
            ProcessServer.ExecuteTask(Of StAccionCalcCosteKg)(AddressOf AccioncalcularCosteKg, StAccion, services)
        Next
    End Sub

    <Serializable()> _
    Public Class StAccionCalcCosteKg
        Public IDObra As Integer
        Public FechaInicio As DateTime
        Public FechaFin As DateTime

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDObra As Integer, ByVal FechaInicio As DateTime, ByVal FechaFin As DateTime)
            Me.IDObra = IDObra
            Me.FechaInicio = FechaInicio
            Me.FechaFin = FechaFin
        End Sub
    End Class

    <Task()> Public Shared Sub AccioncalcularCosteKg(ByVal data As StAccionCalcCosteKg, ByVal services As ServiceProvider)
        Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")

        'Buscar la obra padre de la obra que me llega.
        Dim dtObraPadre As DataTable = OC.SelOnPrimaryKey(data.IDObra)
        If Not IsNothing(dtObraPadre) AndAlso dtObraPadre.Rows.Count > 0 AndAlso Length(dtObraPadre.Rows(0)("IDObraPadre")) > 0 Then
            Dim StCalculo As New StAccionCalcCosteKg(dtObraPadre.Rows(0)("IDObraPadre"), data.FechaInicio, data.FechaFin)
            Dim dtK As DataTable = ProcessServer.ExecuteTask(Of StAccionCalcCosteKg, DataTable)(AddressOf CalculoCosteKg, StCalculo, services)

            If Not IsNothing(dtK) AndAlso dtK.Rows.Count > 0 Then
                Dim Kilos As Double = dtK.Compute("SUM (KilosNeto)", Nothing)
                Dim f As New Filter
                f.Clear()
                f.Add("IDObra", FilterOperator.Equal, data.IDObra)
                Dim dtObra As DataTable = OC.Filter(f)
                If Not IsNothing(dtObra) AndAlso dtObra.Rows.Count > 0 Then
                    dtObra.Rows(0)("KgsEntradaUva") = Kilos
                    If Length(Kilos) > 0 AndAlso Kilos > 0 Then
                        dtObra.Rows(0)("CosteKgUva") = Nz(dtObra.Rows(0)("ImpRealA"), 0) / Kilos
                    Else
                        dtObra.Rows(0)("CosteKgUva") = 0
                    End If
                    'Entrada
                    For Each drEntrada As DataRow In dtK.Rows
                        Dim StActua As New StActualizarEntradasUva(drEntrada("IDEntrada"), dtObra.Rows(0)("CosteKgUva"), 1)
                        ProcessServer.ExecuteTask(Of StActualizarEntradasUva)(AddressOf ActualizarEntradasUva, StActua, services)
                    Next
                    'Fin Entrada

                    dtObra.Rows(0)("FechaCosteKgUva") = Date.Today
                    'dtObra.Rows(0)("Vendimia") = Vendimia                
                    BusinessHelper.UpdateTable(dtObra)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Function CalculoCosteKg(ByVal data As StAccionCalcCosteKg, ByVal services As ServiceProvider) As DataTable
        Dim f As New Filter
        f.Add("IDObra", FilterOperator.Equal, data.IDObra)
        f.Add("Fecha", FilterOperator.GreaterThanOrEqual, data.FechaInicio)
        f.Add("Fecha", FilterOperator.LessThanOrEqual, data.FechaFin)

        Dim dtKilos As DataTable = New BE.DataEngine().Filter("frmBdgKilosEntradaUva", f)
        If Not IsNothing(dtKilos) AndAlso dtKilos.Rows.Count > 0 Then
            Return dtKilos
        Else : Return Nothing
        End If
    End Function

    <Serializable()> _
    Public Class StActualizarEntradasUva
        Public IDEntrada As String
        Public CosteKgUva As Double
        Public Calculada As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDEntrada As String, ByVal CosteKgUva As Double, ByVal Calculada As Integer)
            Me.IDEntrada = IDEntrada
            Me.CosteKgUva = CosteKgUva
            Me.Calculada = Calculada
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarEntradasUva(ByVal data As StActualizarEntradasUva, ByVal services As ServiceProvider)
        Dim ClsEnt As New BdgEntrada
        Dim f As New Filter
        f.Add("IDEntrada", FilterOperator.Equal, data.IDEntrada)
        Dim dtEntrada As DataTable = ClsEnt.Filter(f)
        If Not IsNothing(dtEntrada) AndAlso dtEntrada.Rows.Count > 0 Then
            For Each dr As DataRow In dtEntrada.Rows
                If dr("CosteKgUvaManual") = False Then
                    dr("CosteKgUva") = data.CosteKgUva
                    dr("CosteKgUvaCalculada") = data.Calculada
                End If
            Next
            ClsEnt.Update(dtEntrada)
        End If
    End Sub

    <Task()> Public Shared Sub TratarEntradaUvaManual(ByVal dtEntradas As DataTable, ByVal services As ServiceProvider)
        If Not IsNothing(dtEntradas) AndAlso dtEntradas.Rows.Count > 0 Then
            For Each dr As DataRow In dtEntradas.Rows
                If dr("CosteKgUvaCalculada") = False Then 'No bloqueado
                    Dim StActua As New StActualizarEntradasUva(dr("IDEntrada"), Nz(dr("CantidadMarca1"), 0), 0)
                    ProcessServer.ExecuteTask(Of StActualizarEntradasUva)(AddressOf ActualizarEntradasUva, StActua, services)
                End If
            Next
        End If
    End Sub

    <Serializable()> _
    Public Class StTratarEntradaUvaAuto
        Public DtEntradas As DataTable
        Public CosteKgUva As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtEntradas As DataTable, ByVal CosteKgUva As Double)
            Me.DtEntradas = DtEntradas
            Me.CosteKgUva = CosteKgUva
        End Sub
    End Class

    <Task()> Public Shared Sub TratarEntradaUvaAutomatica(ByVal data As StTratarEntradaUvaAuto, ByVal services As ServiceProvider)
        If Not IsNothing(data.DtEntradas) AndAlso data.DtEntradas.Rows.Count > 0 Then
            For Each dr As DataRow In data.DtEntradas.Rows
                If dr("CosteKgUvaCalculada") = False Then 'No bloqueado
                    Dim StEntrada As New StActualizarEntradasUva(dr("IDEntrada"), data.CosteKgUva, 0)
                    ProcessServer.ExecuteTask(Of StActualizarEntradasUva)(AddressOf ActualizarEntradasUva, StEntrada, services)
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub TratarEntradaUvaManualEntradasCalculadas(ByVal dtEntradas As DataTable, ByVal services As ServiceProvider)
        If Not IsNothing(dtEntradas) AndAlso dtEntradas.Rows.Count > 0 Then
            For Each dr As DataRow In dtEntradas.Rows
                If dr("CosteKgUvaCalculada") = True Then
                    Dim StActua As New StActuaEntradaUvaCalc(dr("IDEntrada"), Nz(dr("CantidadMarca1"), 0), 1)
                    ProcessServer.ExecuteTask(Of StActuaEntradaUvaCalc)(AddressOf ActualizarEntradasUvaCalculada, StActua, services)
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub TratarEntradaUvaDeshacerManualEntradasCalculadas(ByVal dtEntradas As DataTable, ByVal services As ServiceProvider)
        If Not IsNothing(dtEntradas) AndAlso dtEntradas.Rows.Count > 0 Then
            Dim f As New Filter
            For Each dr As DataRow In dtEntradas.Rows
                If dr("CosteKgUvaCalculada") = True Then
                    Dim StActua As New StActuaEntradaUvaCalc(dr("IDEntrada"), 0, 0)
                    ProcessServer.ExecuteTask(Of StActuaEntradaUvaCalc)(AddressOf ActualizarEntradasUvaCalculada, StActua, services)

                    'Averiguar que obraHija pertenece la entrada seleccionada, con la Entrada y la finca se obtiene la obraPadre,
                    'y con las fechas de la entrada se obtiene la obraHija 
                    f.Clear()
                    f.Add("IDFinca", FilterOperator.Equal, dr("IDFinca"))
                    f.Add("FechaInicio", FilterOperator.LessThanOrEqual, dr("Fecha"))
                    f.Add("FechaFin", FilterOperator.GreaterThanOrEqual, dr("Fecha"))
                    Dim dtObra As DataTable = New BE.DataEngine().Filter("vFrmBdgObraFinca", f)
                    If Not IsNothing(dtObra) AndAlso dtObra.Rows.Count > 0 Then
                        Dim StAccion As New StAccionCalcCosteKg(dtObra.Rows(0)("IDObra"), dtObra.Rows(0)("FechaInicio"), dtObra.Rows(0)("FechaFin"))
                        ProcessServer.ExecuteTask(Of StAccionCalcCosteKg)(AddressOf AccioncalcularCosteKg, StAccion, services)
                    End If
                End If
            Next
        End If
    End Sub

    <Serializable()> _
    Public Class StActuaEntradaUvaCalc
        Public IDEntrada As String
        Public CosteKgUva As Double
        Public Manual As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDEntrada As String, ByVal CosteKgUva As Double, ByVal Manual As Integer)
            Me.IDEntrada = IDEntrada
            Me.CosteKgUva = CosteKgUva
            Me.Manual = Manual
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarEntradasUvaCalculada(ByVal data As StActuaEntradaUvaCalc, ByVal services As ServiceProvider)
        Dim ClsBdgEnt As New BdgEntrada
        Dim f As New Filter
        f.Add("IDEntrada", FilterOperator.Equal, data.IDEntrada)
        Dim dtEntrada As DataTable = ClsBdgEnt.Filter(f)
        If Not IsNothing(dtEntrada) AndAlso dtEntrada.Rows.Count > 0 Then
            For Each dr As DataRow In dtEntrada.Rows
                dr("CosteKgUva") = data.CosteKgUva
                dr("CosteKgUvaManual") = data.Manual
            Next
            ClsBdgEnt.Update(dtEntrada)
        End If
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _E
    Public Const IDEntrada As String = "IDEntrada"
    Public Const Vendimia As String = "Vendimia"
    Public Const IDCartillista As String = "IDCartillista"
    Public Const NEntrada As String = "NEntrada"
    Public Const IDContador As String = "IDContador"
    Public Const Fecha As String = "Fecha"
    Public Const Hora As String = "Hora"
    Public Const IDMunicipio As String = "IDMunicipio"
    Public Const IDVariedad As String = "IDVariedad"
    Public Const Texto As String = "Texto"
    Public Const Bruto As String = "Bruto"
    Public Const BrutoB As String = "BrutoB"
    Public Const Tara As String = "Tara"
    Public Const Neto As String = "Neto"
    Public Const Declarado As String = "Declarado"
    Public Const TipoVariedad As String = "TipoVariedad"
    Public Const IDCuadrilla As String = "IDCuadrilla"
    Public Const NOperarios As String = "NOperarios"
    Public Const IDTractor As String = "IDTractor"
    Public Const IDRemolque As String = "IDRemolque"
    Public Const IDAnalisis As String = "IDAnalisis"
    Public Const IDFinca As String = "IDFinca"
    Public Const IDTransportista As String = "IDTransportista"
    Public Const Talon As String = "Talon"
    Public Const Excedente As String = "Excedente"
    Public Const IDProveedorFra As String = "IDProveedorFra"
    Public Const Envase As String = "Envase"
    Public Const Portes As String = "Portes"
    Public Const ImportePortes As String = "ImportePortes"
    Public Const IdBodega As String = "IdBodega"
    Public Const GradoTeorico As String = "GradoTeorico"
    Public Const IdDeposito As String = "IdDeposito"
    Public Const IDMovimiento As String = "IDMovimiento"
    Public Const IDTipoRecogida As String = "IDTipoRecogida"
    Public Const IncrementoTipoRecogida As String = "IncrementoTipoRecogida"
    Public Const IDExplotadorFincas As String = "IDExplotadorFincas"
End Class