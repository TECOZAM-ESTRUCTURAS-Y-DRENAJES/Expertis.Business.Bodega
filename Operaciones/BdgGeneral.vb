Public Enum OrigenOperacion
    OperacionPlanificada
    Depositos
    OrdenFabricacion
    Expediciones
End Enum


Public Class BdgGeneral

    Public Const NUM_DECIMALES_CANTIDADES As Integer = 8

#Region " Tareas Generales "

    <Task()> Public Shared Function AlmacenMaterial(ByVal IDArticulo As String, ByVal services As ServiceProvider) As String
        If Length(IDArticulo) > 0 Then
            Dim dataArt As New DataArtAlm(IDArticulo)
            Return ProcessServer.ExecuteTask(Of DataArtAlm, String)(AddressOf ArticuloAlmacen.AlmacenPredeterminadoArticulo, dataArt, services)
        End If
    End Function


    <Serializable()> _
  Public Class StObtenerPrecio
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Cantidad As Double
        Public Fecha As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, Optional ByVal Cantidad As Double = 1, Optional ByVal Fecha As Date = cnMinDate)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.Cantidad = Cantidad
            Me.Fecha = Fecha
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerPrecio(ByVal data As StObtenerPrecio, ByVal services As ServiceProvider) As Double
        If data.Fecha = cnMinDate Then data.Fecha = Today
        Dim StData As New DataCalculoTarifaComercial(data.IDArticulo, data.Cantidad, data.Fecha, data.IDAlmacen)
        ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf ProcesoComercial.TarifaCosteArticulo, StData, services)
        Return StData.DatosTarifa.PrecioCosteA
    End Function

    <Task()> Public Shared Function ObtenerCosteMaterial(ByVal data As BusinessRuleData, ByVal services As ServiceProvider) As Double
        If Nz(data.Current("Cantidad"), 0) + Nz(data.Current("Merma"), 0) <> 0 AndAlso data.Current.ContainsKey("Precio") Then
            Dim Fecha As Date = cnMinDate
            If Not data.Context Is Nothing AndAlso data.Context.ContainsKey("Fecha") Then
                Fecha = Nz(data.Context("Fecha"), cnMinDate)
            End If

            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfo As MonedaInfo = Monedas.GetMoneda(Monedas.MonedaA.ID, Fecha)
            Dim datPrecio As New BdgGeneral.StObtenerPrecio(data.Current("IDArticulo") & String.Empty, data.Current("IDAlmacen") & String.Empty, Nz(data.Current("Cantidad"), 0) + Nz(data.Current("Merma"), 0), Fecha)
            data.Current("Precio") = xRound(ProcessServer.ExecuteTask(Of BdgGeneral.StObtenerPrecio, Double)(AddressOf BdgGeneral.ObtenerPrecio, datPrecio, services), MonInfo.NDecimalesPrecio)
            Return Nz(data.Current("Precio"), 0)
        End If
    End Function

    <Serializable()> _
    Public Class StGetQDeposito
        Public IDArticulo As String
        Public IDUdMedida As String
        Public IDDeposito As String
        Public Cantidad As Double

        Public Sub New()

        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal IDUdMedida As String, ByVal IDDeposito As String, ByVal Cantidad As Double)
            Me.IDArticulo = IDArticulo
            Me.IDUdMedida = IDUdMedida
            Me.IDDeposito = IDDeposito
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Function GetQDeposito(ByVal data As StGetQDeposito, ByVal services As ServiceProvider) As Double
        Dim dtDep As DataTable = New Business.Bodega.BdgDeposito().SelOnPrimaryKey(data.IDDeposito)
        If Not dtDep Is Nothing AndAlso dtDep.Rows.Count > 0 Then
            If Length(dtDep.Rows(0)("IDUdMedida")) = 0 Then
                Return data.Cantidad
            Else
                Dim datFactor As New ArticuloUnidadAB.DatosFactorConversion(data.IDArticulo, data.IDUdMedida, dtDep.Rows(0)("IDUdMedida"))
                Dim dblFactor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datFactor, services)
                Return data.Cantidad * dblFactor
            End If
        End If
    End Function


    <Serializable()> _
    Public Class DataCalculoOcupacion
        Public IDDeposito As String
        Public IDArticulo As String
        Public Lote As String
        Public IDBarrica As String

        Public IDVino As Guid
        Public IDUDMedida As String
        Public Ocupacion As Double

        Public Sub New(ByVal IDDeposito As String, ByVal IDArticulo As String, ByVal Lote As String, ByVal IDBarrica As String, ByVal IDUDMedida As String, ByVal IDVino As Guid)
            Me.IDDeposito = IDDeposito
            Me.IDArticulo = IDArticulo
            Me.Lote = Lote
            Me.IDBarrica = IDBarrica
            Me.IDUDMedida = IDUDMedida
            Me.IDVino = IDVino
        End Sub
    End Class

    <Task()> Public Shared Function CalculoOcupacion(ByVal data As DataCalculoOcupacion, ByVal services As ServiceProvider) As Double
        'La Ocupación:
        '- Se obtendrá del vino introducido
        '- Estará en la unidad del depósito.
        data.Ocupacion = 0

        Dim datVinoEnDpto As New DataObtenerVinoEnDeposito(data.IDDeposito, data.IDArticulo, data.Lote, data.IDBarrica)
        datVinoEnDpto = ProcessServer.ExecuteTask(Of DataObtenerVinoEnDeposito, DataObtenerVinoEnDeposito)(AddressOf ObtenerVinoEnDeposito, datVinoEnDpto, services)
        data.IDVino = datVinoEnDpto.IDVino
        If Not (data.IDVino.Equals(Guid.Empty)) Then
            Dim FilDepVino As New Filter
            'FilDepVino.Add("IDDeposito", FilterOperator.Equal, data.IDDeposito)
            FilDepVino.Add("IDVino", FilterOperator.Equal, data.IDVino)
            Dim dt As DataTable = New BdgDepositoVino().Filter(FilDepVino)
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                Dim stGetQDeposito As New BdgGeneral.StGetQDeposito(data.IDArticulo, data.IDUDMedida, data.IDDeposito, Nz(dt.Rows(0)("Cantidad"), 0))
                data.Ocupacion = ProcessServer.ExecuteTask(Of BdgGeneral.StGetQDeposito, Double)(AddressOf BdgGeneral.GetQDeposito, stGetQDeposito, services)
            End If
        End If

        Return data.Ocupacion
    End Function

    <Serializable()> _
    Public Class DataObtenerVinoEnDeposito
        Public IDDeposito As String
        Public IDArticulo As String
        Public Lote As String
        Public IDBarrica As String

        Public dtVino As DataTable
        Public IDVino As Guid
        Public Sub New(ByVal IDDeposito As String, ByVal IDArticulo As String, ByVal Lote As String, ByVal IDBarrica As String)
            Me.IDDeposito = IDDeposito
            Me.IDArticulo = IDArticulo
            Me.Lote = Lote
            Me.IDBarrica = IDBarrica
        End Sub
    End Class
    <Task()> Public Shared Function ObtenerVinoEnDeposito(ByVal data As DataObtenerVinoEnDeposito, ByVal services As ServiceProvider) As DataObtenerVinoEnDeposito
        Dim gVino As Guid
        Dim ClsParam As New BdgParametro
        Dim stLoteAUtilizar As String = data.Lote
        If (Length(stLoteAUtilizar) = 0) Then stLoteAUtilizar = ClsParam.LotePorDefecto
        Dim blLoteExplicitoEnBotellero As Boolean = ClsParam.LoteExplicitoEnBotellero
        Dim dttDeposito As DataTable = New BdgDeposito().SelOnPrimaryKey(data.IDDeposito)
        If dttDeposito.Rows.Count > 0 Then
            If dttDeposito.Rows(0)("TipoDeposito") = TipoDeposito.Barricas AndAlso dttDeposito.Rows(0)("UsarBarricaComoLote") = True Then
                stLoteAUtilizar = data.IDBarrica
            ElseIf (dttDeposito.Rows(0)("TipoDeposito") = TipoDeposito.Botellero OrElse dttDeposito.Rows(0)("TipoDeposito") = TipoDeposito.Almacen) AndAlso blLoteExplicitoEnBotellero Then
                stLoteAUtilizar = data.Lote
            End If
        End If

        Dim FilDep As New Filter()
        FilDep.Add("IDDeposito", FilterOperator.Equal, data.IDDeposito)
        FilDep.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
        FilDep.Add("IDAlmacen", FilterOperator.Equal, ProcessServer.ExecuteTask(Of String, String)(AddressOf BdgDeposito.ObtenerAlmacenDeposito, data.IDDeposito, services))
        FilDep.Add("Lote", FilterOperator.Equal, stLoteAUtilizar)

        Dim dtVino As DataTable = New BE.DataEngine().Filter("NegBdgVinoEnDeposito", FilDep)
        If Not IsNothing(dtVino) AndAlso dtVino.Rows.Count > 0 Then
            data.IDVino = dtVino.Rows(0)("IDVino")
            data.dtVino = dtVino.Clone
            data.dtVino.Rows.Add(dtVino.Rows(0).ItemArray)
        End If
        Return data
    End Function

    '<Serializable()> _
    'Public Class StObtenerVinoDeposito
    '    Public IDDeposito As String
    '    Public IDArticulo As String
    '    Public Lote As String
    '    Public IDBarrica As String
    '    Public IDUDMedida As String
    '    Public Cantidad As Double

    '    Public dttVinos As DataTable

    '    Public Sub New(ByVal idDeposito As String, Optional ByVal idArticulo As String = "", Optional ByVal strLote As String = "", Optional ByVal idBarrica As String = "", _
    '                   Optional ByVal IDUDMedida As String = "", Optional ByVal Cantidad As Double = 0)
    '        Me.IDDeposito = idDeposito
    '        Me.IDArticulo = idArticulo
    '        Me.Lote = strLote
    '        Me.IDBarrica = idBarrica
    '        Me.IDUDMedida = IDUDMedida
    '        Me.Cantidad = Cantidad
    '    End Sub

    'End Class

    '<Task()> Public Shared Function ObtenerVinoEnDeposito(ByVal data As StObtenerVinoDeposito, ByVal service As ServiceProvider) As StObtenerVinoDeposito
    '    Dim gVino As Guid
    '    Dim oc As New BE.DataEngine
    '    Dim fDeposito As New Filter()
    '    fDeposito.Add("IDDeposito", data.IDDeposito)
    '    Dim stLoteAUtilizar As String = data.Lote 'oParam.LotePorDefecto
    '    If (Length(data.IDArticulo) > 0) Then
    '        Dim oParam As New BdgParametro
    '        If (Length(stLoteAUtilizar) = 0) Then
    '            stLoteAUtilizar = oParam.LotePorDefecto
    '        End If
    '        Dim blLoteExplicitoEnBotellero As Boolean = oParam.LoteExplicitoEnBotellero

    '        Dim bsnDeposito As New Business.Bodega.BdgDeposito
    '        Dim dttDeposito As DataTable = bsnDeposito.SelOnPrimaryKey(data.IDDeposito)
    '        If dttDeposito.Rows.Count > 0 Then
    '            If dttDeposito.Rows(0)("TipoDeposito") = TipoDeposito.Barricas AndAlso dttDeposito.Rows(0)("UsarBarricaComoLote") = True Then
    '                stLoteAUtilizar = data.IDBarrica
    '            ElseIf (dttDeposito.Rows(0)("TipoDeposito") = TipoDeposito.Botellero _
    '                        Or dttDeposito.Rows(0)("TipoDeposito") = TipoDeposito.Almacen) And blLoteExplicitoEnBotellero Then
    '                stLoteAUtilizar = data.Lote '??
    '            End If
    '        End If


    '        fDeposito.Add("IDDeposito", FilterOperator.Equal, data.IDDeposito)
    '        fDeposito.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)

    '        Dim dtDep As DataTable = New BdgDeposito().SelOnPrimaryKey(data.IDDeposito)

    '        Dim strAlmacen As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf Business.Bodega.BdgNave.Almacen, dtDep.Rows(0)("IDNave") & String.Empty, service)
    '        If Length(strAlmacen) > 0 Then
    '            fDeposito.Add("IDAlmacen", FilterOperator.Equal, strAlmacen)
    '        End If
    '        If Length(stLoteAUtilizar) > 0 Then
    '            fDeposito.Add("Lote", FilterOperator.Equal, stLoteAUtilizar)
    '        End If
    '    End If


    '    Dim dtVino As DataTable = oc.Filter("NegBdgVinoPlanEnDeposito", fDeposito)
    '    If Not IsNothing(dtVino) AndAlso dtVino.Rows.Count > 0 Then
    '        data.dttVinos = dtVino
    '    End If
    '    Return data
    'End Function


    <Serializable()> _
    Public Class DataActualizarPKCabecera
        Public dtActualizar As DataTable
        Public NOperacionProvisional As String
        Public NOperacionNew As String
        Public NOperacionField As String

        Public Sub New(ByVal dtActualizar As DataTable, ByVal NOperacionProvisional As String, ByVal NOperacionNew As String, ByVal NOperacionField As String)
            Me.dtActualizar = dtActualizar
            Me.NOperacionProvisional = NOperacionProvisional
            Me.NOperacionNew = NOperacionNew
            Me.NOperacionField = NOperacionField
        End Sub
    End Class
    <Task()> Public Shared Function ActualizarPKCabecera(ByVal data As DataActualizarPKCabecera, ByVal services As ServiceProvider) As DataTable
        If data Is Nothing OrElse data.dtActualizar Is Nothing Then Exit Function
        If data.dtActualizar.Columns.Contains(data.NOperacionField) Then
            Dim lst As List(Of DataRow) = (From c In data.dtActualizar Where c.IsNull(data.NOperacionField) OrElse c(data.NOperacionField) = data.NOperacionProvisional Select c).ToList
            If Not lst Is Nothing AndAlso lst.Count > 0 Then
                For Each dr As DataRow In lst
                    dr(data.NOperacionField) = data.NOperacionNew
                Next
            End If
        End If
        Return data.dtActualizar
    End Function


    <Serializable()> _
    Public Class DataBRule
        Public CurrentData As BusinessData
        Public ContextData As BusinessData
        Public Entity As String
        Public Columnname As String
        Public Value As Object

        'Public Sub New()
        'End Sub
        Public Sub New(ByVal CurrentData As BusinessData, ByVal ContextData As BusinessData, ByVal Entity As String, ByVal ColumnName As String, ByVal Value As Object)
            Me.CurrentData = CurrentData
            Me.ContextData = ContextData
            Me.Entity = Entity
            Me.Columnname = ColumnName
            Me.Value = Value
        End Sub
    End Class

    <Task()> Public Shared Function ApplyGeneralBRule(ByVal data As DataBRule, ByVal services As ServiceProvider) As BusinessData
        Dim ClsEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(data.Entity)
        Dim BDData As New BusinessData
        BDData = ClsEntity.ApplyBusinessRule(data.Columnname, data.Value, data.CurrentData, data.ContextData)
        Return BDData
    End Function

#End Region

#Region " Business Rules Cabecera de Operación (prevista y real) "

    'Public Class DataAsignarFechaPropuestaOperacion
    '    Public Row As DataRow
    '    Public EntidadOp As String
    '    Public FechaOrigen As Date

    '    Public Sub New(ByVal EntidadOp As String, ByVal Row As DataRow, Optional ByVal FechaOrigen As Date = cnMinDate)
    '        Me.EntidadOp = EntidadOp
    '        Me.Row = Row
    '        Me.FechaOrigen = FechaOrigen
    '    End Sub
    'End Class
    '<Task()> Public Shared Sub AsignarFechaPropuestaOperacion(ByVal data As DataAsignarFechaPropuestaOperacion, ByVal services As ServiceProvider)
    '    Dim datFecha As New DataGetFechaPropuestaOperacion(data.EntidadOp)
    '    data.Row("Fecha") = ProcessServer.ExecuteTask(Of DataGetFechaPropuestaOperacion, Date)(AddressOf GetFechaPropuestaOperacion, datFecha, services)
    'End Sub

    <Serializable()> _
    Public Class DataGetFechaPropuestaOperacion
        Public EntidadOp As String
        Public FechaOrigen As Date

        Public Sub New(ByVal EntidadOp As String)
            Me.EntidadOp = EntidadOp
            Me.FechaOrigen = cnMinDate
        End Sub
    End Class
    <Task()> Public Shared Function GetFechaPropuestaOperacion(ByVal data As DataGetFechaPropuestaOperacion, ByVal services As ServiceProvider) As Date
        Dim FechaOperacion As Date = Today
        If New BdgParametro().FechaNuevaOperacionIgualUltimaOperacion Then
            If data.FechaOrigen <> cnMinDate Then
                FechaOperacion = data.FechaOrigen
                Return New Date(FechaOperacion.Year, FechaOperacion.Month, FechaOperacion.Day, FechaOperacion.Hour, FechaOperacion.Minute, FechaOperacion.Second)
            Else
                If Length(data.EntidadOp) > 0 Then
                    Dim oBusinessEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(data.EntidadOp)
                    Dim dtOpAnterior As DataTable = oBusinessEntity.Filter("TOP 1 Fecha", , "Fecha DESC")
                    If Not dtOpAnterior Is Nothing AndAlso dtOpAnterior.Rows.Count > 0 Then
                        FechaOperacion = dtOpAnterior.Rows(0)("Fecha")
                    End If
                End If
            End If
        End If
        Return New Date(FechaOperacion.Year, FechaOperacion.Month, FechaOperacion.Day, Date.Now.Hour, Date.Now.Minute, Date.Now.Second)
    End Function

    <Task()> Public Shared Sub CambioIDTipoOperacion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        Dim OperacionReal As Boolean
        If Not data.Context Is Nothing AndAlso data.Context.ContainsKey("Origen") AndAlso Nz(data.Context("Origen"), 0) = enumBdgOrigenOperacion.Real Then
            OperacionReal = True
        End If

        Dim FieldImputacionMaterial As String = "ImputacionGlobalMat"
        Dim FieldImputacionMOD As String = "ImputacionGlobalMod"
        Dim FieldImputacionCentro As String = "ImputacionGlobalCentro"
        Dim FieldImputacionVarios As String = "ImputacionGlobalVarios"

        If OperacionReal Then
            FieldImputacionMaterial = "ImputacionRealMaterial"
            FieldImputacionMOD = "ImputacionRealMod"
            FieldImputacionCentro = "ImputacionRealCentro"
            FieldImputacionVarios = "ImputacionRealVarios"
        End If

        If Length(data.Current("IDTipoOperacion")) > 0 Then

            Dim TiposOperacion As EntityInfoCache(Of BdgTipoOperacionInfo) = services.GetService(Of EntityInfoCache(Of BdgTipoOperacionInfo))()
            Dim BdgTipoOpInfo As BdgTipoOperacionInfo = TiposOperacion.GetEntity(data.Current("IDTipoOperacion"))

            data.Current("IDAnalisis") = BdgTipoOpInfo.IDAnalisis
            data.Current("TipoMovimiento") = BdgTipoOpInfo.TipoMovimiento

            If OperacionReal Then
                If Nz(BdgTipoOpInfo.IDRule, -1) > 0 Then
                    If Not ProcessServer.ExecuteTask(Of Integer, Boolean)(AddressOf Solmicro.Expertis.Business.Rules.Rules.TienePermisos, BdgTipoOpInfo.IDRule, services) Then
                        ApplicationService.GenerateError("No tiene permisos para crear una Operación del Tipo |.", Quoted(BdgTipoOpInfo.DescTipoOperacion))
                    End If
                End If
            End If

            Dim datImputacion As StImputacion
            If OperacionReal Then
                datImputacion = New StImputacion(data.Current("IDTipoOperacion"), data.Current("NOperacion"), Business.Bodega.enumBdgOrigenOperacion.Real, True, data.Current("Fecha"))
            Else
                datImputacion = New StImputacion(data.Current("IDTipoOperacion"), data.Current("NOperacionPlan"), Business.Bodega.enumBdgOrigenOperacion.Planificada, True, data.Current("Fecha"))
            End If
            datImputacion = ProcessServer.ExecuteTask(Of StImputacion, StImputacion)(AddressOf Imputaciones, datImputacion, services)

            data.Current(FieldImputacionMaterial) = datImputacion.ImputarMaterialesGlobales 'drTipoOp("ImputacionRealMaterial")
            data.Current(FieldImputacionMOD) = datImputacion.ImputarMODGlobales 'drTipoOp("ImputacionRealMod")
            data.Current(FieldImputacionCentro) = datImputacion.ImputarCentrosGlobales ' drTipoOp("ImputacionRealCentro")
            data.Current(FieldImputacionVarios) = datImputacion.ImputarVariosGlobales  'drTipoOp("ImputacionRealVarios")

            data.Current("DTImputacionMaterialGlobal") = datImputacion.DtMaterialesGlobal
            data.Current("DTImputacionMODGlobal") = datImputacion.DtMODGlobal
            data.Current("DTImputacionCentroGlobal") = datImputacion.DtCentrosGlobal
            data.Current("DTImputacionVariosGlobal") = datImputacion.DtVariosGlobal

            'If data.Context Is Nothing OrElse Not data.Context.ContainsKey("OperacionNew") OrElse Not data.Context("OperacioNew") Then
            '    Dim ClsOper As New BdgOperacion
            '    'MATERIALES         
            '    If data.Current("ImputacionRealMaterial") AndAlso data.Context.ContainsKey("ADDMaterialGlobalReal") AndAlso data.Context("ADDMaterialGlobalReal") Then
            '        ClsOper.ApplyBusinessRule("ImputacionRealMaterial", data.Current("ImputacionRealMaterial"), data.Current, data.Context)
            '    End If
            '    'MATERIALES FIN

            '    'MOD Recoger los datos segun el tipo de operacion 
            '    If data.Current("ImputacionRealMod") AndAlso data.Context.ContainsKey("ADDModGlobalReal") AndAlso data.Context("ADDModGlobalReal") Then
            '        ClsOper.ApplyBusinessRule("ImputacionRealMod", data.Current("ImputacionRealMod"), data.Current, data.Context)
            '    End If
            '    'MOD FIN
            '    'CENTRO Recoger los datos segun el tipo de operacion           
            '    If data.Current("ImputacionRealCentro") AndAlso data.Context.ContainsKey("ADDCentroGlobalReal") AndAlso data.Context("ADDCentroGlobalReal") Then
            '        ClsOper.ApplyBusinessRule("ImputacionRealCentro", data.Current("ImputacionRealCentro"), data.Current, data.Context)
            '    End If
            '    'CENTRO FIN
            '    'VARIOS Recoger los datos segun el tipo de operacion 
            '    If data.Current("ImputacionRealVarios") AndAlso data.Context.ContainsKey("ADDVariosGlobalReal") AndAlso data.Context("ADDVariosGlobalReal") Then
            '        ClsOper.ApplyBusinessRule("ImputacionRealVarios", data.Current("ImputacionRealVarios"), data.Current, data.Context)
            '    End If
            '    'VARIOS FIN
            '    data.Current("ShowMensaje") = False
            'End If
        Else
            data.Current("IDAnalisis") = System.DBNull.Value

            data.Current(FieldImputacionMaterial) = False
            data.Current(FieldImputacionMOD) = False
            data.Current(FieldImputacionCentro) = False
            data.Current(FieldImputacionVarios) = False
        End If
    End Sub

    <Task()> Public Shared Function CambioFecha(ByVal Fecha As Date, ByVal services As ServiceProvider) As Date
        Dim FechaConHora As Date = Fecha
        If Nz(Fecha, cnMinDate) <> cnMinDate Then
            If Fecha.Hour = 0 AndAlso Fecha.Minute = 0 AndAlso Fecha.Second = 0 Then
                Dim Hora As Date = Now
                FechaConHora = New Date(Fecha.Year, Fecha.Month, Fecha.Day, Hora.Hour, Hora.Minute, Hora.Second)
            End If
        Else
            FechaConHora = Now
        End If

        Return FechaConHora
    End Function

#End Region

#Region " OperacionVino Origen y Destino de Planificadas y Reales "

    <Task()> Public Shared Sub CambioQDeposito(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        Dim datRecalQ As New DataRecalcularQ(data.ColumnName, data.Current)
        data.Current = ProcessServer.ExecuteTask(Of DataRecalcularQ, IPropertyAccessor)(AddressOf RecalcularQ, datRecalQ, services)

    End Sub

    <Task()> Public Shared Sub CambioCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        Dim datRecalQ As New DataRecalcularQ(data.ColumnName, data.Current)
        data.Current = ProcessServer.ExecuteTask(Of DataRecalcularQ, IPropertyAccessor)(AddressOf RecalcularQ, datRecalQ, services)
    End Sub

    <Task()> Public Shared Sub CambioLitros(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        Dim datRecalQ As New DataRecalcularQ(data.ColumnName, data.Current)
        data.Current = ProcessServer.ExecuteTask(Of DataRecalcularQ, IPropertyAccessor)(AddressOf RecalcularQ, datRecalQ, services)
    End Sub

    <Task()> Public Shared Sub CambioIDArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDArticulo")) = 0 Then
            data.Current("DescArticulo") = DBNull.Value
            data.Current("IDUdMedida") = DBNull.Value
            data.Current("Cantidad") = DBNull.Value
            data.Current("Lote") = DBNull.Value
        Else
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current("IDArticulo"))
            If Not ArtInfo Is Nothing AndAlso Length(ArtInfo.IDArticulo) > 0 Then
                data.Current("DescArticulo") = ArtInfo.DescArticulo
                data.Current("IDUdMedida") = ArtInfo.IDUDInterna
            End If
        End If
        Dim datRecalQ As New DataRecalcularQ("Litros", data.Current)
        data.Current = ProcessServer.ExecuteTask(Of DataRecalcularQ, IPropertyAccessor)(AddressOf RecalcularQ, datRecalQ, services)


        If Nz(data.Current("Destino"), False) AndAlso Length(data.Current("IDEstadoVino")) = 0 Then
            data.Current("IDEstadoVino") = ProcessServer.ExecuteTask(Of BusinessRuleData, String)(AddressOf GetOperacionEstadoVinoContexto, data, services)
        End If
    End Sub


    <Task()> Public Shared Function GetOperacionEstadoVinoContexto(ByVal data As BusinessRuleData, ByVal services As ServiceProvider) As String
        If data.Context.ContainsKey("IDTipoOperacion") AndAlso Length(data.Context("IDTipoOperacion")) > 0 Then

            Dim Origenes As DataTable
            If data.Context.ContainsKey("OperacionVinoOrigenes") Then
                Origenes = data.Context("OperacionVinoOrigenes")
            End If

            Dim datEstadoVino As New DataGetOperacionEstadoVino(data.Context("IDTipoOperacion"), Origenes)
            Return ProcessServer.ExecuteTask(Of DataGetOperacionEstadoVino, String)(AddressOf GetOperacionEstadoVino, datEstadoVino, services)
        End If

    End Function

    <Serializable()> _
    Public Class DataGetOperacionEstadoVino
        Public IDTipoOperacion As String
        Public OperacionVinoOrigenes As DataTable

        Public Sub New(ByVal IDTipoOperacion As String, ByVal OperacionVinoOrigenes As DataTable)
            Me.IDTipoOperacion = IDTipoOperacion
            Me.OperacionVinoOrigenes = OperacionVinoOrigenes
        End Sub
    End Class
    <Task()> Public Shared Function GetOperacionEstadoVino(ByVal data As DataGetOperacionEstadoVino, ByVal services As ServiceProvider) As String
        If Length(data.IDTipoOperacion) = 0 Then Exit Function

        Dim TiposOperacion As EntityInfoCache(Of BdgTipoOperacionInfo) = services.GetService(Of EntityInfoCache(Of BdgTipoOperacionInfo))()
        Dim TipoOpInfo As BdgTipoOperacionInfo = TiposOperacion.GetEntity(data.IDTipoOperacion)
        Dim IDEstadoVino As String = TipoOpInfo.IDEstadoVino
        If Not data.OperacionVinoOrigenes Is Nothing AndAlso data.OperacionVinoOrigenes.Rows.Count > 0 Then

            If Length(IDEstadoVino) = 0 AndAlso Not data.OperacionVinoOrigenes Is Nothing Then
                Dim EstadosVino As List(Of Object) = (From c In data.OperacionVinoOrigenes Where c.RowState <> DataRowState.Deleted Select c("IDEstadoVino") Distinct).ToList
                If Not EstadosVino Is Nothing AndAlso EstadosVino.Count = 1 Then
                    IDEstadoVino = EstadosVino(0) & String.Empty
                Else
                    IDEstadoVino = String.Empty
                End If
            End If

        End If
        Return IDEstadoVino
    End Function


    <Serializable()> _
    Public Class DataRecalcularQ
        Public Current As IPropertyAccessor
        Public CampoModificado As String
        Public Sub New(ByVal CampoModificado As String, ByVal Current As IPropertyAccessor)
            Me.Current = Current
            Me.CampoModificado = CampoModificado
        End Sub
    End Class

    <Task()> Public Shared Function RecalcularQ(ByVal data As DataRecalcularQ, ByVal services As ServiceProvider) As IPropertyAccessor
        Dim CamposRecalcular As String() = New String() {"Cantidad", "QDeposito", "Litros"}

        If Array.IndexOf(CamposRecalcular, data.CampoModificado) < 0 Then Return data.Current

        Dim strUDMedidaLitros As String = New BdgParametro().UnidadesCampoLitros()


        Dim dblFactor As Double = 1
        Dim IDUDMedida1 As String
        Dim IDUDMedida2 As String
        Select Case data.CampoModificado
            Case "Cantidad"
                IDUDMedida1 = Nz(data.Current("IDUdMedida"), strUDMedidaLitros)
                IDUDMedida2 = strUDMedidaLitros
            Case "QDeposito"
                IDUDMedida1 = Nz(data.Current("IDUdMedida"), strUDMedidaLitros)
                IDUDMedida2 = strUDMedidaLitros
            Case "Litros"
                IDUDMedida1 = strUDMedidaLitros
                IDUDMedida2 = Nz(data.Current("IDUdMedida"), strUDMedidaLitros)
        End Select


        If (Length(data.Current("IDArticulo")) > 0) Then
            Dim datFactor As New ArticuloUnidadAB.DatosFactorConversion(data.Current("IDArticulo"), IDUDMedida1, IDUDMedida2)
            dblFactor = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datFactor, services)
        End If
        If data.CampoModificado <> "Cantidad" Then
            'data.Current("Cantidad") = 0 'xRound(data.Current("Cantidad") * dblFactor, 3)
            If data.CampoModificado = "QDeposito" Then
                Dim QArticulo As Double
                If (Length(data.Current("IDDeposito")) > 0) Then
                    Dim Depositos As EntityInfoCache(Of BdgDepositoInfo) = services.GetService(Of EntityInfoCache(Of BdgDepositoInfo))()
                    ' Dim oDep As New Business.Bodega.BdgDeposito
                    Dim DptoInfo As BdgDepositoInfo = Depositos.GetEntity(data.Current("IDDeposito"))
                    'Dim dtDep As DataTable = oDep.SelOnPrimaryKey(data.Current("IDDeposito"))
                    'If dtDep.Rows.Count > 0 Then
                    If Length(DptoInfo.IDUDMedida) = 0 Then
                        QArticulo = data.Current("QDeposito")
                    Else
                        '//Obtenemos el Factor desde la unidad del depósito
                        Dim datFactorDeposito As New ArticuloUnidadAB.DatosFactorConversion(data.Current("IDArticulo"), DptoInfo.IDUDMedida, Nz(data.Current("IDUdMedida"), strUDMedidaLitros))
                        Dim dblFactorDeposito As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datFactorDeposito, services)

                        QArticulo = Nz(data.Current("QDeposito"), 0) * dblFactorDeposito
                    End If
                    ' End If
                End If
                data.Current("Cantidad") = xRound(QArticulo, 3)
            ElseIf data.CampoModificado = "Litros" Then
                data.Current("Cantidad") = xRound(Nz(data.Current("Litros"), 0) * dblFactor, 3)
            End If
        End If
        If data.CampoModificado <> "QDeposito" Then
            Dim params As New BdgGeneral.StGetQDeposito(data.Current("IDArticulo") & String.Empty, Nz(data.Current("IDUdMedida"), strUDMedidaLitros), data.Current("IDDeposito") & String.Empty, Nz(data.Current("Cantidad"), 0))
            Dim QDeposito As Double = ProcessServer.ExecuteTask(Of BdgGeneral.StGetQDeposito, Double)(AddressOf BdgGeneral.GetQDeposito, params, services)
            data.Current("QDeposito") = xRound(QDeposito, 3)
        End If
        If data.CampoModificado <> "Litros" Then
            data.Current("Litros") = xRound(Nz(data.Current("Cantidad"), 0) * dblFactor, 3)
        End If
        Return data.Current
    End Function

    <Serializable()> _
  Public Class DataOperacionDeposito
        Public IDDeposito As String
        Public IDArticulo As String
        Public DescArticulo As String
        Public IDUDMedida As String
        Public IDBarrica As String
        Public Lote As String
        Public IDVino As Guid
        Public Ocupacion As Double
        Public Cantidad As Double
        Public CantidadOriginal As Double
        Public QDeposito As Double
        Public TotalPendiente As Double
        Public Litros As Double
        Public IDTipoOperacion As String
        Public dtOperacionVinoPlanOrigen As DataTable
        Public Origen As DataOperacionDeposito
        Public FechaVino As Date

        Public Capacidad As Double
        Public IDUDDeposito As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDDeposito As String, ByVal IDArticulo As String, ByVal DescArticulo As String, ByVal IDUDMedida As String, _
                       ByVal IDBarrica As String, ByVal Lote As String, ByVal IDVino As Guid, ByVal Ocupacion As Double, ByVal Cantidad As Double, _
                       ByVal CantidadOriginal As Double, ByVal QDeposito As Double, ByVal TotalPendiente As Double, ByVal Litros As Double, _
                       ByVal IDTipoOperacion As String, ByVal dtOperacionVinoPlanOrigen As DataTable, ByVal Origen As DataOperacionDeposito)
            Me.IDDeposito = IDDeposito
            Me.IDArticulo = IDArticulo
            Me.DescArticulo = DescArticulo
            Me.IDUDMedida = IDUDMedida
            Me.IDBarrica = IDBarrica
            Me.Lote = Lote
            Me.IDVino = IDVino
            Me.Ocupacion = Ocupacion
            Me.Cantidad = Cantidad
            Me.CantidadOriginal = CantidadOriginal
            Me.QDeposito = QDeposito
            Me.TotalPendiente = TotalPendiente
            Me.Litros = Litros
            Me.IDTipoOperacion = IDTipoOperacion
            Me.dtOperacionVinoPlanOrigen = dtOperacionVinoPlanOrigen
            Me.Origen = Origen
        End Sub
    End Class

    <Task()> Public Shared Function CambioDatosDeposito(ByVal data As DataOperacionDeposito, ByVal services As ServiceProvider) As DataOperacionDeposito
        Dim IDUDMedidaLitros As String = New BdgParametro().UnidadesCampoLitros()
        Dim StDataArtDep As New DataObtenerVinoArtDep(data.IDDeposito, Nothing)
        data.dtOperacionVinoPlanOrigen = ProcessServer.ExecuteTask(Of DataObtenerVinoArtDep, DataTable)(AddressOf ObtenerVinoArticuloDeposito, StDataArtDep, services)

        Dim Depositos As EntityInfoCache(Of BdgDepositoInfo) = services.GetService(Of EntityInfoCache(Of BdgDepositoInfo))()
        Dim DeptoInfo As BdgDepositoInfo = Depositos.GetEntity(data.IDDeposito)
        data.Capacidad = DeptoInfo.Capacidad
        data.IDUDDeposito = DeptoInfo.IDUDMedida

        If data.dtOperacionVinoPlanOrigen.Rows.Count = 1 Then
            data.IDArticulo = data.dtOperacionVinoPlanOrigen.Rows(0)("IDArticulo")
            data.DescArticulo = data.dtOperacionVinoPlanOrigen.Rows(0)("DescArticulo")
            data.IDUDMedida = data.dtOperacionVinoPlanOrigen.Rows(0)("IDUdMedida")
            data.IDVino = data.dtOperacionVinoPlanOrigen.Rows(0)("IDVino")
            data.Lote = Nz(data.dtOperacionVinoPlanOrigen.Rows(0)("Lote"), String.Empty)
            If Length(data.dtOperacionVinoPlanOrigen.Rows(0)("Fecha")) > 0 Then data.FechaVino = data.dtOperacionVinoPlanOrigen.Rows(0)("Fecha")

            data.Cantidad = Nz(data.dtOperacionVinoPlanOrigen.Rows(0)("Cantidad"), 0)
            data.IDBarrica = data.dtOperacionVinoPlanOrigen.Rows(0)("IDBarrica") & String.Empty

            Dim current As New BusinessData
            current("IDArticulo") = data.IDArticulo
            current("IDUDMedida") = data.IDUDMedida
            current("Cantidad") = data.Cantidad
            current("IDDeposito") = data.IDDeposito
            current("QDeposito") = data.QDeposito
            current("Litros") = data.Litros

            Dim datRecalQ As New DataRecalcularQ("Cantidad", current)
            current = ProcessServer.ExecuteTask(Of DataRecalcularQ, IPropertyAccessor)(AddressOf RecalcularQ, datRecalQ, services)
            data.QDeposito = Nz(current("QDeposito"), 0)
            data.Litros = Nz(current("Litros"), 0)
            Dim datOcup As New BdgGeneral.DataCalculoOcupacion(data.IDDeposito, data.IDArticulo, data.Lote, data.IDBarrica, data.IDUDMedida, data.IDVino)
            data.Ocupacion = ProcessServer.ExecuteTask(Of BdgGeneral.DataCalculoOcupacion, Double)(AddressOf BdgGeneral.CalculoOcupacion, datOcup, services)
        ElseIf data.dtOperacionVinoPlanOrigen.Rows.Count = 0 Then
            data.Ocupacion = 0
        End If

        If Not data.Origen Is Nothing Then
            If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Business.Bodega.BdgTipoOperacion.ProponerArticulo, data.IDTipoOperacion, services) Then
                If DeptoInfo.MultiplesVinos Then
                    'Si no tiene ningun vino el deposito seleccionado se le asigna el de Origen
                    'If Length(data.IDArticulo) = 0 Then
                    data.IDArticulo = data.Origen.IDArticulo
                    data.DescArticulo = data.Origen.DescArticulo
                    data.IDUDMedida = data.Origen.IDUDMedida
                    data.IDVino = Guid.Empty
                    data.Lote = String.Empty
                    'End If
                Else
                    If Length(data.IDArticulo) = 0 Then
                        data.IDArticulo = data.Origen.IDArticulo
                        data.DescArticulo = data.Origen.DescArticulo
                        data.IDUDMedida = data.Origen.IDUDMedida
                        data.IDVino = Guid.Empty
                        data.Lote = String.Empty
                    End If
                End If
            End If
            data.IDBarrica = data.Origen.IDBarrica
            data.Litros = ProcessServer.ExecuteTask(Of DataOperacionDeposito, Double)(AddressOf QDestinoPropuesta, data, services)
        End If
        If Length(data.IDUDMedida) > 0 Then
            'Dim StGetFactor As New BdgWorkClass.StGetFactorConversion(data.IDArticulo, IDUDMedidaLitros, data.IDUDMedida)
            'data.Cantidad = xRound(Nz(data.Litros, 0) * ProcessServer.ExecuteTask(Of BdgWorkClass.StGetFactorConversion, Double)(AddressOf BdgWorkClass.GetFactorConversion, StGetFactor, services), 3)

            'Dim stGetQDeposito As New BdgOperacionGeneral.StGetQDeposito(data.IDArticulo, data.IDUDMedida, data.IDDeposito, Nz(data.Cantidad, 0))
            'data.QDeposito = ProcessServer.ExecuteTask(Of BdgOperacionGeneral.StGetQDeposito, Double)(AddressOf BdgOperacionGeneral.GetQDeposito, stGetQDeposito, services)
            Dim current As New BusinessData
            current("IDArticulo") = data.IDArticulo
            current("DescArticulo") = data.DescArticulo
            current("IDUDMedida") = data.IDUDMedida
            current("Cantidad") = data.Cantidad
            current("IDDeposito") = data.IDDeposito
            current("QDeposito") = data.QDeposito
            current("Litros") = data.Litros

            Dim datRecalQ As New DataRecalcularQ("Litros", current)
            current = ProcessServer.ExecuteTask(Of DataRecalcularQ, IPropertyAccessor)(AddressOf RecalcularQ, datRecalQ, services)
            data.Cantidad = Nz(current("Cantidad"), 0)
            data.QDeposito = Nz(current("QDeposito"), 0)

        End If

        Return data
    End Function

    <Serializable()> _
  Public Class DataObtenerVinoArtDep
        Public IDDeposito As String
        Public IDArticulo As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDDeposito As String, ByVal IDArticulo As String)
            Me.IDDeposito = IDDeposito
            Me.IDArticulo = IDArticulo
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerVinoArticuloDeposito(ByVal data As DataObtenerVinoArtDep, ByVal services As ServiceProvider) As DataTable
        Dim FilArtDep As New Filter
        FilArtDep.Add("IDDeposito", FilterOperator.Equal, data.IDDeposito)
        If Length(data.IDArticulo) > 0 Then FilArtDep.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
        Return New BE.DataEngine().Filter("frmBdgOperacionArticulo", FilArtDep, , "IDArticulo, Lote")
    End Function

    <Task()> Public Shared Function QDestinoPropuesta(ByVal data As DataOperacionDeposito, ByVal services As ServiceProvider) As Double
        Dim dtDeposito As DataTable = New Business.Bodega.BdgDeposito().SelOnPrimaryKey(data.IDDeposito)
        If Not dtDeposito Is Nothing AndAlso dtDeposito.Rows.Count > 0 Then
            Dim cantidadRestante As Double = Nz(data.TotalPendiente, 0)
            Dim capacidad As Double = dtDeposito.Rows(0)("Capacidad")
            Dim ocupacion As Double = ProcessServer.ExecuteTask(Of String, Double)(AddressOf Business.Bodega.BdgOperacion.DevolverOcupacion, data.IDDeposito, services)

            'El espacio restante que hay en el Deposito
            Dim cantidadDeposito As Double = capacidad - ocupacion

            'Si la cantidad libre en el Deposito es mayor que la cantidad a asignar, o es un Deposito sin limite.
            'Le asigna el al destion la cantiadad que Restante
            If cantidadRestante < cantidadDeposito Or dtDeposito.Rows(0)("SinLimite") Then
                Return cantidadRestante
            Else : Return cantidadDeposito
            End If
        End If
    End Function

    <Task()> Public Shared Sub CambioIDDeposito(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDDeposito")) = 0 Then
            Select Case CBool(data.Current("Destino"))
                Case True
                    data.Current("Litros") = DBNull.Value
                    data.Current("IDArticulo") = DBNull.Value
                    data.Current("IDEstadoVino") = System.DBNull.Value
                Case False
            End Select
        Else
            Dim dtDeposito As DataTable = New BdgDeposito().SelOnPrimaryKey(data.Current("IDDeposito"))
            If dtDeposito.Rows.Count > 0 Then
                data.Current("Capacidad") = dtDeposito.Rows(0)("Capacidad")
                data.Current("IDUDDeposito") = dtDeposito.Rows(0)("IDUDMedida")

                If Nz(data.Current("Destino"), False) AndAlso Length(data.Current("IDEstadoVino")) = 0 Then
                    data.Current("IDEstadoVino") = ProcessServer.ExecuteTask(Of BusinessRuleData, String)(AddressOf GetOperacionEstadoVinoContexto, data, services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDOrden(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDOrden")) > 0 Then
            If Length(data.Current("IDArticulo")) = 0 Then
                Dim ClsOF As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenFabricacion")
                Dim dtOrden As DataTable = ClsOF.SelOnPrimaryKey(data.Current("IDOrden"))
                If dtOrden.Rows.Count > 0 Then
                    data.Current("IDArticulo") = dtOrden.Rows(0)("IDArticulo")
                    data.Current("NOrden") = dtOrden.Rows(0)("NOrden")
                End If
            End If
            Dim dtArt As DataTable = New Articulo().SelOnPrimaryKey(data.Current("IDArticulo"))
            If Not dtArt Is Nothing AndAlso dtArt.Rows.Count > 0 Then
                data.Current("DescArticulo") = dtArt.Rows(0)("DescArticulo")
                data.Current("IDUdMedida") = dtArt.Rows(0)("IDUdInterna")
            End If
        End If
    End Sub


    <Task()> Public Shared Sub CambioIDBarrica(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDDeposito")) > 0 Then
            Dim dtDeposito As DataTable = New BdgDeposito().SelOnPrimaryKey(data.Current("IDDeposito"))
            If dtDeposito.Rows.Count > 0 AndAlso Nz(dtDeposito.Rows(0)("UsarBarricaComoLote"), False) Then
                data.Current("Lote") = data.Current("IDBarrica")
            End If
        End If
    End Sub

#End Region

#Region " Imputaciones "

#Region " B.Rules Imputaciones Materiales "

#Region " Globales "

    <Task()> Public Shared Sub CambioMaterialGlobal(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioMaterial, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData, Double)(AddressOf BdgGeneral.ObtenerCosteMaterial, data, services)
    End Sub

    <Task()> Public Shared Sub CambioCantidadGlobal(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioCantidadMerma, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData, Double)(AddressOf BdgGeneral.ObtenerCosteMaterial, data, services)
    End Sub

    <Task()> Public Shared Sub CambioMermaGlobal(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioCantidadMerma, data, services)

        Dim StData As New Business.Bodega.BdgOperacion.dataAplicarRulesMerma(data.Current("IDArticulo") & String.Empty, Nz(data.Current("Cantidad"), 0), Nz(data.Current("Merma"), 0))
        If Not ProcessServer.ExecuteTask(Of Business.Bodega.BdgOperacion.dataAplicarRulesMerma, Boolean)(AddressOf Business.Bodega.BdgOperacion.ValidarMermaMaterial, StData, services) Then
            ApplicationService.GenerateError("La Merma supera el % máximo permitido.")
        End If

        ProcessServer.ExecuteTask(Of BusinessRuleData, Double)(AddressOf BdgGeneral.ObtenerCosteMaterial, data, services)
    End Sub

    <Task()> Public Shared Sub CambioAlmacenGlobal(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData, Double)(AddressOf BdgGeneral.ObtenerCosteMaterial, data, services)
    End Sub

#End Region

#Region " Vino Material "

    <Task()> Public Shared Sub CambioMaterialVinoMaterial(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioMaterial, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData, Double)(AddressOf BdgGeneral.ObtenerCosteMaterial, data, services)
    End Sub

    <Task()> Public Shared Sub CambioCantidadVinoMaterial(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioCantidadMerma, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData, Double)(AddressOf BdgGeneral.ObtenerCosteMaterial, data, services)
    End Sub

    <Task()> Public Shared Sub CambioMermaVinoMaterial(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioCantidadMerma, data, services)

        Dim StData As New Business.Bodega.BdgOperacion.dataAplicarRulesMerma(data.Current("IDArticulo") & String.Empty, Nz(data.Current("Cantidad"), 0), Nz(data.Current("Merma"), 0))
        If Not ProcessServer.ExecuteTask(Of Business.Bodega.BdgOperacion.dataAplicarRulesMerma, Boolean)(AddressOf Business.Bodega.BdgOperacion.ValidarMermaMaterial, StData, services) Then
            ApplicationService.GenerateError("La Merma supera el % máximo permitido.")
        End If
        ProcessServer.ExecuteTask(Of BusinessRuleData, Double)(AddressOf BdgGeneral.ObtenerCosteMaterial, data, services)
    End Sub

    <Task()> Public Shared Sub CambioAlmacenVinoMaterial(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData, Double)(AddressOf BdgGeneral.ObtenerCosteMaterial, data, services)
    End Sub

#End Region

    <Task()> Public Shared Sub CambioMaterial(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDArticulo")) > 0 Then
            data.Current("IDAlmacen") = ProcessServer.ExecuteTask(Of String, String)(AddressOf BdgGeneral.AlmacenMaterial, data.Current("IDArticulo"), services)
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current("IDArticulo"))
            data.Current("DescArticulo") = ArtInfo.DescArticulo
            data.Current("GestionStockPorLotes") = ArtInfo.GestionStockPorLotes
            data.Current("Precinta") = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Business.Negocio.Articulo.EsPrecinta, data.Current("IDArticulo"), services)
        Else
            data.Current("IDAlmacen") = DBNull.Value
            data.Current("DescArticulo") = DBNull.Value
            data.Current("GestionStockPorLotes") = False
            data.Current("Precinta") = False
        End If
    End Sub

    <Task()> Public Shared Sub CambioCantidadMerma(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        data.Current = ProcessServer.ExecuteTask(Of IPropertyAccessor, IPropertyAccessor)(AddressOf ActualizarCantidadBruto, data.Current, services)
    End Sub

    <Task()> Public Shared Function ActualizarCantidadBruto(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider) As IPropertyAccessor
        data("CantidadBruto") = Nz(data("Cantidad"), 0) + Nz(data("Merma"), 0)
        Return data
    End Function

#End Region

#Region " B.Rules Imputaciones Centros "

    <Task()> Public Shared Sub CambioCentro(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf GetInfoCentro, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarPorCantidad, data, services)
    End Sub

    <Task()> Public Shared Sub CambioIncidencia(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf GetInfoCentro, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarPorCantidad, data, services)
    End Sub

    <Task()> Public Shared Sub GetInfoCentro(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDCentro")) > 0 Then
            Dim Centros As EntityInfoCache(Of CentroInfo) = services.GetService(Of EntityInfoCache(Of CentroInfo))()
            Dim CtroInfo As CentroInfo = Centros.GetEntity(data.Current("IDCentro"))
            If Not CtroInfo Is Nothing AndAlso Length(CtroInfo.IDCentro) > 0 Then
                If Length(data.Current("IDIncidencia")) = 0 Then
                    data.Current("Tasa") = CtroInfo.TasaEjecucionA
                Else : data.Current("Tasa") = CtroInfo.TasaPreparacionA
                End If
                data.Current("DescCentro") = CtroInfo.DescCentro

                data.Current("IdUdMedidaCentro") = CtroInfo.IDUdMedida
                If Not data.Context Is Nothing AndAlso data.Context.ContainsKey("IDUdMedida") Then
                    data.Current("IdUdMedidaArticulo") = data.Context("IDUdMedida") & String.Empty
                End If

                If Length(CtroInfo.UdTiempo) = 0 Then
                    data.Current("UDTiempo") = enumstdUdTiempo.Horas
                Else : data.Current("UDTiempo") = CtroInfo.UdTiempo
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarPorCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDIncidencia")) = 0 Then
            If Length(data.Current("IdUdMedidaCentro")) > 0 Then
                data.Current("PorCantidad") = True
            Else : data.Current("PorCantidad") = False
            End If
        Else : data.Current("PorCantidad") = False
        End If
    End Sub

#End Region

#Region " B.Rules Imputaciones MOD "

    <Task()> Public Shared Sub CambioOperario(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDOperario")) > 0 Then
            Dim Operarios As EntityInfoCache(Of OperarioInfo) = services.GetService(Of EntityInfoCache(Of OperarioInfo))()
            Dim OperInfo As OperarioInfo = Operarios.GetEntity(data.Current("IDOperario"))
            data.Current("DescOperario") = OperInfo.DescOperario
            If Length(OperInfo.IDCategoria) > 0 Then
                data.Current("IDCategoria") = OperInfo.IDCategoria
                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalculoTasa, data, services)
            Else : ApplicationService.GenerateError("El Operario {0} no tiene espeficicada una categoría", Quoted(OperInfo.IDOperario))
            End If
        Else
            data.Current("IDCategoria") = DBNull.Value
            data.Current("DescOperario") = DBNull.Value
            data.Current("IDHora") = DBNull.Value
            data.Current("Tasa") = 0
        End If
    End Sub

    <Task()> Public Shared Sub CalculoTasa(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDCategoria")) > 0 Then
            If Length(data.Current("IDHora")) = 0 Then
                Dim StrHoraPred As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf HoraCategoria.GetHoraPredeterminada, data.Current("IDCategoria"), services)
                If Length(StrHoraPred) > 0 Then data.Current("IDHora") = StrHoraPred
            End If
            Dim StHoraCat As New HoraCategoria.DatosPrecioHoraCatOper(data.Current("IDCategoria") & String.Empty, data.Current("IDHora") & String.Empty, data.Context("Fecha"), data.Current("IDOperario") & String.Empty)
            data.Current("Tasa") = ProcessServer.ExecuteTask(Of HoraCategoria.DatosPrecioHoraCatOper, Double)(AddressOf HoraCategoria.ObtenerPrecioHora, StHoraCat, services)
        Else : ApplicationService.GenerateError("No se ha encontrado Hora Predeterminada para el Operario: | con Categoría: |.", data.Current("IDOperario"), data.Current("IDCategoria"))
        End If
    End Sub

#End Region

#Region " B.Rules Imputaciones Varios "

    <Task()> Public Shared Sub CambioVarios(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDVarios")) > 0 Then

            Dim dtVarios As DataTable = New DataEngine().Filter("vfrmMntoVarios", New FilterItem("IDVarios", FilterOperator.Equal, data.Current("IDVarios")))
            If Not dtVarios Is Nothing AndAlso dtVarios.Rows.Count > 0 Then
                data.Current("DescVarios") = dtVarios.Rows(0)("DescVarios") & String.Empty
                data.Current("TipoCosteFV") = dtVarios.Rows(0)("TipoCosteFV")
                data.Current("TipoCosteDI") = dtVarios.Rows(0)("TipoCosteDI")
                data.Current("Fiscal") = dtVarios.Rows(0)("Fiscal")
            End If
        Else
            data.Current("DescVarios") = DBNull.Value
            data.Current("TipoCosteFV") = DBNull.Value
            data.Current("TipoCosteDI") = DBNull.Value
            data.Current("Fiscal") = DBNull.Value
        End If
    End Sub

#End Region

#Region " Materiales,Centros,Varios,MOD - Lineas/Globales "

    <Serializable()> _
    Public Class StImputacion
        Public IDTipoOperacion As String
        Public NOperacion As String

        Public Origen As enumBdgOrigenOperacion
        Public BlnGlobal As Boolean = True

        Public Fecha As Date?

        Public ImputarCentrosGlobales As Boolean?
        Public ImputarMaterialesGlobales As Boolean?
        Public ImputarVariosGlobales As Boolean?
        Public ImputarMODGlobales As Boolean?
        Public ForzarGlobales As Boolean

        Public DtMateriales As DataTable
        Public DtMaterialesLotes As DataTable
        Public DtMaterialesGlobal As DataTable
        Public DtMaterialesGlobalLotes As DataTable
        Public DtCentros As DataTable
        Public DtCentrosGlobal As DataTable
        Public DtMOD As DataTable
        Public DtMODGlobal As DataTable
        Public DtVarios As DataTable
        Public DtVariosGlobal As DataTable
        Public DtAnalisis As DataTable
        Public DtAnalisisVariable As DataTable


        Public LineaDestino As New List(Of StDestino)


        Public IDAnalisis As String
        Public IDOrden As Integer?
        Public Cantidad As Double
        Public TotalLitrosDestino As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDTipoOperacion As String, ByVal NOperacion As String, ByVal Origen As enumBdgOrigenOperacion, ByVal BlnGlobal As Boolean, Optional ByVal Fecha As Date = cnMinDate)
            Me.IDTipoOperacion = IDTipoOperacion
            Me.NOperacion = NOperacion
            Me.Origen = Origen
            Me.Fecha = Fecha
            Me.BlnGlobal = BlnGlobal
        End Sub
    End Class

    <Serializable()> _
    Public Class StDestino
        Public IDLineaOperacionDestino As Guid
        Public IDArticulo As String
        Public IDEstructura As String
        Public Litros As Double
        Public Cantidad As Double
        Public Sub New()
        End Sub

        Public Sub New(ByVal IDLineaOperacionDestino As Guid, ByVal IDArticulo As String, ByVal IDEstructura As String, ByVal Litros As Double, ByVal Cantidad As Double)
            Me.IDLineaOperacionDestino = IDLineaOperacionDestino
            Me.IDArticulo = IDArticulo
            Me.IDEstructura = IDEstructura
            Me.Litros = Litros
            Me.Cantidad = Cantidad
        End Sub
    End Class


    <Task()> Public Shared Function Imputaciones(ByVal data As StImputacion, ByVal services As ServiceProvider) As StImputacion
        '//Si no se ha indicado lo que debemos imputar, lo cogeremos de la Operacion o del Tipo de Operacion
        data = ProcessServer.ExecuteTask(Of StImputacion, StImputacion)(AddressOf GetMarcasImputacion, data, services)


        Select Case data.BlnGlobal
            Case True
                data.DtMaterialesGlobal = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf ImputacionGlobalMateriales, data, services)
                data.DtCentrosGlobal = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf ImputacionGlobalCentros, data, services)
                data.DtMODGlobal = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf ImputacionGlobalMOD, data, services)
                data.DtVariosGlobal = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf ImputacionGlobalVarios, data, services)
            Case False
                data.DtMateriales = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf ImputacionLineaMateriales, data, services)
                data.DtCentros = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf ImputacionLineaCentros, data, services)
                data.DtMOD = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf ImputacionLineaMOD, data, services)
                data.DtVarios = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf ImputacionLineaVarios, data, services)

                If data.Origen = enumBdgOrigenOperacion.Real Then
                    Dim DsAnalisis As List(Of DataTable) = ProcessServer.ExecuteTask(Of StImputacion, List(Of DataTable))(AddressOf ImputacionLineaAnalisis, data, services)
                    If Not DsAnalisis Is Nothing AndAlso DsAnalisis.Count > 0 Then
                        If Not DsAnalisis(0) Is Nothing AndAlso DsAnalisis(0).Rows.Count > 0 Then
                            data.DtAnalisis = DsAnalisis(0)
                        End If
                        If Not DsAnalisis(1) Is Nothing AndAlso DsAnalisis(1).Rows.Count > 0 Then
                            data.DtAnalisisVariable = DsAnalisis(1)
                        End If
                    End If
                End If
        End Select
        Return data
    End Function

#Region " Get Marcas Imputaciones "

    <Task()> Public Shared Function GetMarcasImputacion(ByVal data As StImputacion, ByVal services As ServiceProvider) As StImputacion
        data = ProcessServer.ExecuteTask(Of StImputacion, StImputacion)(AddressOf GetMarcasImputacionPorOperacion, data, services)
        data = ProcessServer.ExecuteTask(Of StImputacion, StImputacion)(AddressOf GetMarcasImputacionPorTipoOperacion, data, services)
        Return data
    End Function

    <Task()> Public Shared Function GetMarcasImputacionPorOperacion(ByVal data As StImputacion, ByVal services As ServiceProvider) As StImputacion
        If data.BlnGlobal AndAlso data.ImputarMaterialesGlobales Is Nothing AndAlso data.ImputarMODGlobales Is Nothing AndAlso data.ImputarCentrosGlobales Is Nothing AndAlso data.ImputarVariosGlobales Is Nothing Then
            If Length(data.NOperacion) > 0 Then
                Dim dtOperacion As DataTable
                Select Case data.Origen
                    Case enumBdgOrigenOperacion.Planificada
                        dtOperacion = New BdgOperacionPlan().SelOnPrimaryKey(data.NOperacion)
                    Case enumBdgOrigenOperacion.Real
                        dtOperacion = New BdgOperacion().SelOnPrimaryKey(data.NOperacion)
                End Select

                If dtOperacion.Rows.Count > 0 Then
                    Dim FieldMarcaMaterial As String
                    Dim FieldMarcaMOD As String
                    Dim FieldMarcaCentro As String
                    Dim FieldMarcaVarios As String

                    Select Case data.Origen
                        Case enumBdgOrigenOperacion.Planificada
                            FieldMarcaMaterial = "ImputacionGlobalMat"
                            FieldMarcaMOD = "ImputacionGlobalMod"
                            FieldMarcaCentro = "ImputacionGlobalCentro"
                            FieldMarcaVarios = "ImputacionGlobalVarios"
                        Case enumBdgOrigenOperacion.Real
                            FieldMarcaMaterial = "ImputacionRealMaterial"
                            FieldMarcaMOD = "ImputacionRealMod"
                            FieldMarcaCentro = "ImputacionRealCentro"
                            FieldMarcaVarios = "ImputacionRealVarios"
                    End Select

                    data.ImputarMaterialesGlobales = dtOperacion.Rows(0)(FieldMarcaMaterial)
                    data.ImputarMODGlobales = dtOperacion.Rows(0)(FieldMarcaMOD)
                    data.ImputarCentrosGlobales = dtOperacion.Rows(0)(FieldMarcaCentro)
                    data.ImputarVariosGlobales = dtOperacion.Rows(0)(FieldMarcaVarios)
                End If
            End If
        End If
        Return data
    End Function

    <Task()> Public Shared Function GetMarcasImputacionPorTipoOperacion(ByVal data As StImputacion, ByVal services As ServiceProvider) As StImputacion
        If data.BlnGlobal AndAlso data.ImputarMaterialesGlobales Is Nothing AndAlso data.ImputarMODGlobales Is Nothing AndAlso data.ImputarCentrosGlobales Is Nothing AndAlso data.ImputarVariosGlobales Is Nothing Then
            If Length(data.IDTipoOperacion) > 0 Then
                Dim TiposOperacion As EntityInfoCache(Of BdgTipoOperacionInfo) = services.GetService(Of EntityInfoCache(Of BdgTipoOperacionInfo))()
                Dim BdgTipoOpInfo As BdgTipoOperacionInfo = TiposOperacion.GetEntity(data.IDTipoOperacion)

                Dim FieldMarcaMaterial As String
                Dim FieldMarcaMOD As String
                Dim FieldMarcaCentro As String
                Dim FieldMarcaVarios As String

                Select Case data.Origen
                    Case enumBdgOrigenOperacion.Planificada
                        data.ImputarMaterialesGlobales = BdgTipoOpInfo.ImputacionPrevMaterial
                        data.ImputarMODGlobales = BdgTipoOpInfo.ImputacionPrevMod
                        data.ImputarCentrosGlobales = BdgTipoOpInfo.ImputacionPrevCentro
                        data.ImputarVariosGlobales = BdgTipoOpInfo.ImputacionPrevVarios
                    Case enumBdgOrigenOperacion.Real
                        data.ImputarMaterialesGlobales = BdgTipoOpInfo.ImputacionRealMaterial
                        data.ImputarMODGlobales = BdgTipoOpInfo.ImputacionRealMod
                        data.ImputarCentrosGlobales = BdgTipoOpInfo.ImputacionRealCentro
                        data.ImputarVariosGlobales = BdgTipoOpInfo.ImputacionRealVarios
                End Select
            End If
        Else
            '//Si alguno no es Nothing es que hemos especificado explicaitamente que queremos tratar uno de ellos, el resto los ponemos a false
            If data.ImputarMaterialesGlobales Is Nothing Then data.ImputarMaterialesGlobales = False
            If data.ImputarMODGlobales Is Nothing Then data.ImputarMODGlobales = False
            If data.ImputarCentrosGlobales Is Nothing Then data.ImputarCentrosGlobales = False
            If data.ImputarVariosGlobales Is Nothing Then data.ImputarVariosGlobales = False
        End If
        Return data
    End Function

#End Region

#Region "Tratamiento de Imputaciones Globales"

#Region " Materiales "

    <Task()> Public Shared Function ImputacionGlobalMateriales(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 AndAlso data.ImputarMaterialesGlobales Then
            data.DtMaterialesGlobal = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetEstructuraImputacionMaterialGlobal, data, services)

            Dim context As New BusinessData
            context("Fecha") = Nz(data.Fecha, Today)

            '//Materiales que NO son de Bodega
            Dim dtComponentes As DataTable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionMaterialPorTipoOperacion, data, services)
            For Each DrComponente As DataRow In dtComponentes.Select
                Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacionGlobal(data.Origen, data.NOperacion, data.IDTipoOperacion, context("Fecha"), data.DtMaterialesGlobal, DrComponente)
                ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacionGlobal)(AddressOf AddNuevaLineaImputacionMaterialGlobal, datNuevaLineaImputacion, services)
            Next
            Return data.DtMaterialesGlobal
        End If

    End Function

    <Task()> Public Shared Function GetEstructuraImputacionMaterialGlobal(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                Return New BE.DataEngine().Filter("frmBdgOperacionPlanMaterialGlobal", New NoRowsFilterItem)
            Case enumBdgOrigenOperacion.Real
                Return New BE.DataEngine().Filter("frmBdgOperacionMaterialRealGlobal", New NoRowsFilterItem)
        End Select
    End Function
    <Task()> Public Shared Function GetImputacionOperacionMaterialPorTipoOperacion(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 Then
            Dim FilView As New Filter
            FilView.Add("IDTipoOperacion", FilterOperator.Equal, data.IDTipoOperacion)

            If Not data.ForzarGlobales Then
                Select Case data.Origen
                    Case enumBdgOrigenOperacion.Planificada
                        FilView.Add("ImputacionPrevMaterial", True)
                    Case enumBdgOrigenOperacion.Real
                        FilView.Add("ImputacionRealMaterial", True)
                End Select
            End If
            Return New BE.DataEngine().Filter("NegBdgTipoOperacionMaterial", FilView)
        End If
    End Function

    Public Class DataAddNuevaLineaImputacionGlobal
        Public Origen As enumBdgOrigenOperacion '//Prevista/real

        Public NOperacion As String
        Public Fecha As Date
        Public IDTipoOperacion As String

        '//Row Origen de las imputaciones
        Public RowOrigen As DataRow

        '//
        Public dtImputacion As DataTable

        Public Sub New(ByVal Origen As enumBdgOrigenOperacion, ByVal NOperacion As String, ByVal IDTipoOperacion As String, ByVal Fecha As Date, ByVal dtImputacion As DataTable, ByVal RowOrigen As DataRow)
            Me.Origen = Origen
            Me.NOperacion = NOperacion
            Me.IDTipoOperacion = IDTipoOperacion
            Me.Fecha = Fecha

            Me.dtImputacion = dtImputacion
            Me.RowOrigen = RowOrigen
        End Sub
    End Class
    <Task()> Public Shared Sub AddNuevaLineaImputacionMaterialGlobal(ByVal data As DataAddNuevaLineaImputacionGlobal, ByVal services As ServiceProvider)

        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)

        Dim oEntidad As BusinessHelper
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                oEntidad = New BdgOperacionPlanMaterial
            Case enumBdgOrigenOperacion.Real
                oEntidad = New BdgOperacionMaterial
        End Select

        Dim DrNew As DataRow = data.dtImputacion.NewRow
        If DrNew.Table.Columns.Contains("Fecha") Then DrNew("Fecha") = context("Fecha")
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "NOperacionPlan", "NOperacion")) = data.NOperacion
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDOperacionPlanMaterial", "IDOperacionMaterial")) = Guid.NewGuid

        DrNew = oEntidad.ApplyBusinessRule("IDArticulo", data.RowOrigen("IDArticulo") & String.Empty, DrNew, context)
        DrNew = oEntidad.ApplyBusinessRule("Cantidad", xRound(Nz(data.RowOrigen("Cantidad"), 0), NUM_DECIMALES_CANTIDADES), DrNew, context)
        DrNew("RecalcularMaterial") = Nz(data.RowOrigen("RecalcularMaterial"), False)

        Dim Merma As Double = 0
        If data.RowOrigen.Table.Columns.Contains("Merma") Then
            Merma = xRound(Nz(data.RowOrigen("Merma"), 0), NUM_DECIMALES_CANTIDADES)
        End If
        DrNew = oEntidad.ApplyBusinessRule("Merma", Merma, DrNew, context)

        data.dtImputacion.Rows.Add(DrNew)
    End Sub

#End Region

#Region " Centros "

    <Task()> Public Shared Function ImputacionGlobalCentros(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 AndAlso data.ImputarCentrosGlobales Then
            data.DtCentrosGlobal = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetEstructuraImputacionCentrosGlobal, data, services)

            Dim context As New BusinessData
            context("Fecha") = Nz(data.Fecha, Today)

            Dim dtCentros As DataTable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionCentrosPorTipoOperacion, data, services)
            For Each drCentro As DataRow In dtCentros.Select
                Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacionGlobal(data.Origen, data.NOperacion, data.IDTipoOperacion, context("Fecha"), data.DtCentrosGlobal, drCentro)
                ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacionGlobal)(AddressOf AddNuevaLineaImputacionCentrosGlobal, datNuevaLineaImputacion, services)
            Next

            Return data.DtCentrosGlobal
        End If
    End Function

    <Task()> Public Shared Function GetEstructuraImputacionCentrosGlobal(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                Return New BE.DataEngine().Filter("frmBdgOperacionPlanCentroGlobal", New NoRowsFilterItem)
            Case enumBdgOrigenOperacion.Real
                Return New BE.DataEngine().Filter("frmBdgOperacionCentroRealGlobal", New NoRowsFilterItem)
        End Select
    End Function
    <Task()> Public Shared Function GetImputacionOperacionCentrosPorTipoOperacion(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 Then
            Dim FilView As New Filter
            FilView.Add("IDTipoOperacion", FilterOperator.Equal, data.IDTipoOperacion)

            If Not data.ForzarGlobales Then
                Select Case data.Origen
                    Case enumBdgOrigenOperacion.Planificada
                        FilView.Add("ImputacionPrevCentro", True)
                    Case enumBdgOrigenOperacion.Real
                        FilView.Add("ImputacionRealCentro", True)
                End Select
            End If

            Return New BE.DataEngine().Filter("frmBdgTipoOperacionCentro", FilView)
        End If
    End Function

    <Task()> Public Shared Sub AddNuevaLineaImputacionCentrosGlobal(ByVal data As DataAddNuevaLineaImputacionGlobal, ByVal services As ServiceProvider)

        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)

        Dim oEntidad As BusinessHelper
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                oEntidad = New BdgOperacionPlanCentro
            Case enumBdgOrigenOperacion.Real
                oEntidad = New BdgOperacionCentro
        End Select

        Dim DrNew As DataRow = data.dtImputacion.NewRow
        If DrNew.Table.Columns.Contains("Fecha") Then DrNew("Fecha") = context("Fecha")
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "NOperacionPlan", "NOperacion")) = data.NOperacion
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDOperacionPlanCentro", "IDOperacionCentro")) = Guid.NewGuid

        DrNew = oEntidad.ApplyBusinessRule("IDCentro", data.RowOrigen("IDCentro") & String.Empty, DrNew, context)
        ' DrNew = oEntidad.ApplyBusinessRule("IDIncidencia", data.RowOrigen("IDIncidencia") & String.Empty, DrNew, context)
        DrNew = oEntidad.ApplyBusinessRule("Tiempo", xRound(Nz(data.RowOrigen("Tiempo"), 0), NUM_DECIMALES_CANTIDADES), DrNew, context)
        DrNew = oEntidad.ApplyBusinessRule("PorCantidad", Nz(data.RowOrigen("PorCantidad"), False), DrNew, context)

        data.dtImputacion.Rows.Add(DrNew)
    End Sub

#End Region

#Region " MOD "

    <Task()> Public Shared Function ImputacionGlobalMOD(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 AndAlso data.ImputarMODGlobales Then
            data.DtMODGlobal = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetEstructuraImputacionMODGlobal, data, services)

            Dim context As New BusinessData
            context("Fecha") = Nz(data.Fecha, Today)

            Dim dtMOD As DataTable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionMODPorTipoOperacion, data, services)
            For Each drMOD As DataRow In dtMOD.Select
                Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacionGlobal(data.Origen, data.NOperacion, data.IDTipoOperacion, context("Fecha"), data.DtMODGlobal, drMOD)
                ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacionGlobal)(AddressOf AddNuevaLineaImputacionMODGlobal, datNuevaLineaImputacion, services)
            Next
            Return data.DtMODGlobal
        End If

    End Function

    <Task()> Public Shared Function GetEstructuraImputacionMODGlobal(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                Return New BE.DataEngine().Filter("frmBdgOperacionPlanMODGlobal", New NoRowsFilterItem)
            Case enumBdgOrigenOperacion.Real
                Return New BE.DataEngine().Filter("frmBdgOperacionMODRealGlobal", New NoRowsFilterItem)
        End Select
    End Function
    <Task()> Public Shared Function GetImputacionOperacionMODPorTipoOperacion(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 Then
            Dim FilView As New Filter
            FilView.Add("IDTipoOperacion", FilterOperator.Equal, data.IDTipoOperacion)

            If Not data.ForzarGlobales Then
                Select Case data.Origen
                    Case enumBdgOrigenOperacion.Planificada
                        FilView.Add("ImputacionPrevMOD", True)
                    Case enumBdgOrigenOperacion.Real
                        FilView.Add("ImputacionRealMOD", True)
                End Select
            End If

            Return New BE.DataEngine().Filter("frmBdgTipoOperacionMod", FilView)
        End If
    End Function

    <Task()> Public Shared Sub AddNuevaLineaImputacionMODGlobal(ByVal data As DataAddNuevaLineaImputacionGlobal, ByVal services As ServiceProvider)
        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)

        Dim oEntidad As BusinessHelper
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                oEntidad = New BdgOperacionPlanMOD
            Case enumBdgOrigenOperacion.Real
                oEntidad = New BdgOperacionMOD
        End Select

        Dim DrNew As DataRow = data.dtImputacion.NewRow
        If DrNew.Table.Columns.Contains("Fecha") Then DrNew("Fecha") = context("Fecha")
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "NOperacionPlan", "NOperacion")) = data.NOperacion
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDOperacionPlanMOD", "IDOperacionMOD")) = Guid.NewGuid

        DrNew = oEntidad.ApplyBusinessRule("IDOperario", data.RowOrigen("IDOperario") & String.Empty, DrNew, context)
        DrNew = oEntidad.ApplyBusinessRule("Tiempo", xRound(Nz(data.RowOrigen("Tiempo"), 0), NUM_DECIMALES_CANTIDADES), DrNew, context)

        If Length(data.RowOrigen("IDCategoria")) > 0 AndAlso DrNew("IDCategoria") & String.Empty <> data.RowOrigen("IDCategoria") Then
            DrNew = oEntidad.ApplyBusinessRule("IDCategoria", Nz(data.RowOrigen("IDCategoria"), 0), DrNew, context)
        End If

        data.dtImputacion.Rows.Add(DrNew)
    End Sub

#End Region

#Region " Varios "

    <Task()> Public Shared Function ImputacionGlobalVarios(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable

        If Length(data.IDTipoOperacion) > 0 AndAlso data.ImputarVariosGlobales Then
            data.DtVariosGlobal = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetEstructuraImputacionVariosGlobal, data, services)
            Dim context As New BusinessData
            context("Fecha") = Nz(data.Fecha, Today)

            Dim dtVarios As DataTable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVariosPorTipoOperacion, data, services)
            For Each drVarios As DataRow In dtVarios.Select
                Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacionGlobal(data.Origen, data.NOperacion, data.IDTipoOperacion, context("Fecha"), data.DtVariosGlobal, drVarios)
                ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacionGlobal)(AddressOf AddNuevaLineaImputacionVariosGlobal, datNuevaLineaImputacion, services)
            Next
            Return data.DtVariosGlobal
        End If

    End Function

    <Task()> Public Shared Function GetEstructuraImputacionVariosGlobal(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                Return New BE.DataEngine().Filter("frmBdgOperacionPlanVariosGlobal", New NoRowsFilterItem)
            Case enumBdgOrigenOperacion.Real
                Return New BE.DataEngine().Filter("frmBdgOperacionVariosRealGlobal", New NoRowsFilterItem)
        End Select
    End Function
    <Task()> Public Shared Function GetImputacionOperacionVariosPorTipoOperacion(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 Then
            Dim FilView As New Filter
            FilView.Add("IDTipoOperacion", FilterOperator.Equal, data.IDTipoOperacion)
            If Not data.ForzarGlobales Then

                Select Case data.Origen
                    Case enumBdgOrigenOperacion.Planificada
                        FilView.Add("ImputacionPrevVarios", True)
                    Case enumBdgOrigenOperacion.Real
                        FilView.Add("ImputacionRealVarios", True)
                End Select
            End If

            Return New BE.DataEngine().Filter("frmBdgTipoOperacionVarios", FilView)
        End If
    End Function

    <Task()> Public Shared Sub AddNuevaLineaImputacionVariosGlobal(ByVal data As DataAddNuevaLineaImputacionGlobal, ByVal services As ServiceProvider)
        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)

        Dim oEntidad As BusinessHelper
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                oEntidad = New BdgOperacionPlanVarios
            Case enumBdgOrigenOperacion.Real
                oEntidad = New BdgOperacionVarios
        End Select

        Dim DrNew As DataRow = data.dtImputacion.NewRow
        If DrNew.Table.Columns.Contains("Fecha") Then DrNew("Fecha") = context("Fecha")
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "NOperacionPlan", "NOperacion")) = data.NOperacion
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDOperacionPlanVarios", "IDOperacionVarios")) = Guid.NewGuid

        DrNew = oEntidad.ApplyBusinessRule("IDVarios", data.RowOrigen("IDVarios") & String.Empty, DrNew, context)
        DrNew = oEntidad.ApplyBusinessRule("Cantidad", xRound(Nz(data.RowOrigen("Cantidad"), 0), NUM_DECIMALES_CANTIDADES), DrNew, context)

        If DrNew.Table.Columns.Contains("Tasa") Then
            DrNew("Tasa") = data.RowOrigen("Tasa")
        End If

        data.dtImputacion.Rows.Add(DrNew)
    End Sub


#End Region

#End Region

#Region "Tratamiento de Imputaciones del Vino"

#Region " Tareas generales de Imputaciones del vino "

    Public Class DataAjustarCantidadImputaciones
        Public RegistroImputacionGlobal As DataRow
        Public FieldAjustar As String
        Public dtImputacionesAjustar As DataTable

        Public Sub New(ByVal FieldAjustar As String, ByVal dtImputacionesAjustar As DataTable, ByVal RegistroImputacionGlobal As DataRow)
            Me.FieldAjustar = FieldAjustar
            Me.dtImputacionesAjustar = dtImputacionesAjustar
            Me.RegistroImputacionGlobal = RegistroImputacionGlobal
        End Sub
    End Class
    <Task()> Public Shared Function AjustarCantidadImputaciones(ByVal data As DataAjustarCantidadImputaciones, ByVal services As ServiceProvider) As DataTable
        If Not data.dtImputacionesAjustar Is Nothing AndAlso data.dtImputacionesAjustar.Rows.Count > 0 Then
            Dim Total As Decimal = (Aggregate r In data.dtImputacionesAjustar Into Sum(CDec(Nz(r(data.FieldAjustar), 0))))
            If Total <> 0 AndAlso Total <> data.RegistroImputacionGlobal(data.FieldAjustar) Then
                Dim DblDif As Double = Math.Abs(Total - data.RegistroImputacionGlobal(data.FieldAjustar))
                data.dtImputacionesAjustar.Rows(data.dtImputacionesAjustar.Rows.Count - 1)(data.FieldAjustar) += DblDif
            End If
            Return data.dtImputacionesAjustar
        End If
    End Function

#End Region

#Region " Imputación de Materiales "

    <Task()> Public Shared Function GetEstructuraImputacionMaterial(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                Return New BE.DataEngine().Filter("frmBdgOperacionVinoPlanMaterial", New NoRowsFilterItem)
            Case enumBdgOrigenOperacion.Real
                Return New BE.DataEngine().Filter("frmBdgOperacionMaterial", New NoRowsFilterItem)
        End Select
    End Function
    <Task()> Public Shared Function ImputacionLineaMateriales(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 AndAlso Length(data.NOperacion) > 0 Then
            data.DtMateriales = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetEstructuraImputacionMaterial, data, services)

            '//Imputación de Materiales desde Imputaciones Globales o por Tipo Operación
            ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf ImputacionLineasMaterialesGenerales, data, services)

            If Nz(data.IDOrden, 0) > 0 Then
                '//Imputamos materiales de un OF
                ProcessServer.ExecuteTask(Of StImputacion)(AddressOf ImputacionLineasMaterialesOF, data, services)
            Else
                '//Imputamos Materiales de un articulo
                ProcessServer.ExecuteTask(Of StImputacion)(AddressOf ImputacionLineasMaterialesArticulo, data, services)
            End If


            Return data.DtMateriales
        End If
    End Function


    Public Class DataAddNuevaLineaImputacion
        Public Origen As enumBdgOrigenOperacion '//Prevista/real

        Public NOperacion As String
        Public Fecha As Date
        Public IDTipoOperacion As String

        Public RepartoGlobales As Boolean

        '//Datos Linea destino
        Public IDLineaOperacionVino As Guid
        Public Cantidad As Double
        Public Merma As Double
        Public Tiempo As Double

        '//Row Origen de las imputaciones
        Public RowOrigen As DataRow

        '//
        Public dtImputacion As DataTable

        Public Sub New(ByVal Origen As enumBdgOrigenOperacion, ByVal NOperacion As String, ByVal IDTipoOperacion As String, ByVal Fecha As Date, ByVal RepartoGlobales As Boolean, ByVal IDLineaOperacionVino As Guid, ByVal dtImputacion As DataTable, ByVal RowOrigen As DataRow)
            Me.Origen = Origen
            Me.NOperacion = NOperacion
            Me.IDTipoOperacion = IDTipoOperacion
            Me.Fecha = Fecha
            Me.RepartoGlobales = RepartoGlobales
            Me.IDLineaOperacionVino = IDLineaOperacionVino
            Me.dtImputacion = dtImputacion
            Me.RowOrigen = RowOrigen
        End Sub
    End Class
    <Task()> Public Shared Sub AddNuevaLineaImputacionMaterial(ByVal data As DataAddNuevaLineaImputacion, ByVal services As ServiceProvider)

        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)

        Dim FieldVinculoGlobal As String
        Dim oEntidad As BusinessHelper
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                oEntidad = New BdgOperacionVinoPlanMaterial
            Case enumBdgOrigenOperacion.Real
                oEntidad = New BdgVinoMaterial
                FieldVinculoGlobal = "IDOperacionMaterial"
        End Select


        Dim DrNew As DataRow = data.dtImputacion.NewRow
        If DrNew.Table.Columns.Contains("Fecha") Then DrNew("Fecha") = context("Fecha")

        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "NOperacionPlan", "NOperacion")) = data.NOperacion
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDOperacionVinoPlanMaterial", "IDVinoMaterial")) = Guid.NewGuid
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDLineaOperacionVinoPlan", "IDVino")) = data.IDLineaOperacionVino

        Dim IDArticulo As String = data.RowOrigen("IDArticulo") & String.Empty
        If data.RowOrigen.Table.Columns.Contains("IDComponente") AndAlso Length(data.RowOrigen("IDComponente")) > 0 Then
            IDArticulo = data.RowOrigen("IDComponente") & String.Empty
        End If

        DrNew = oEntidad.ApplyBusinessRule("IDArticulo", IDArticulo, DrNew, context)

        If data.RowOrigen.Table.Columns.Contains("IDAlmacen") Then
            DrNew = oEntidad.ApplyBusinessRule("IDAlmacen", data.RowOrigen("IDAlmacen") & String.Empty, DrNew, context)
        End If
        If data.RowOrigen.Table.Columns.Contains("Lote") Then
            DrNew = oEntidad.ApplyBusinessRule("Lote", data.RowOrigen("Lote") & String.Empty, DrNew, context)
        End If
        If data.RowOrigen.Table.Columns.Contains("Ubicacion") Then
            DrNew = oEntidad.ApplyBusinessRule("Ubicacion", data.RowOrigen("Ubicacion") & String.Empty, DrNew, context)
        End If

        DrNew = oEntidad.ApplyBusinessRule("Cantidad", xRound(data.Cantidad, NUM_DECIMALES_CANTIDADES), DrNew, context)
        DrNew = oEntidad.ApplyBusinessRule("Merma", xRound(data.Merma, NUM_DECIMALES_CANTIDADES), DrNew, context)

        If data.RowOrigen.Table.Columns.Contains("RecalcularMaterial") Then
            DrNew("RecalcularMaterial") = Nz(data.RowOrigen("RecalcularMaterial"), False)
        End If


        If DrNew.Table.Columns.Contains("IDTipoOperacionOrigen") Then DrNew("IDTipoOperacionOrigen") = data.IDTipoOperacion

        If data.RepartoGlobales Then
            If DrNew.Table.Columns.Contains(FieldVinculoGlobal) Then
                If data.RowOrigen.Table.Columns.Contains(FieldVinculoGlobal) AndAlso Length(data.RowOrigen(FieldVinculoGlobal)) > 0 Then
                    DrNew(FieldVinculoGlobal) = data.RowOrigen(FieldVinculoGlobal)
                End If
            End If
        End If


        If DrNew.Table.Columns.Contains("IDTipoEstructuraOrigen") AndAlso data.RowOrigen.Table.Columns.Contains("IDTipoEstructura") Then DrNew("IDTipoEstructuraOrigen") = data.RowOrigen("IDTipoEstructura")
        If DrNew.Table.Columns.Contains("IDEstructuraOrigen") AndAlso data.RowOrigen.Table.Columns.Contains("IDEstructura") Then DrNew("IDEstructuraOrigen") = data.RowOrigen("IDEstructura")

        data.dtImputacion.Rows.Add(DrNew)
    End Sub
    <Task()> Public Shared Function ImputacionLineasMaterialesLotes(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ClsArt As New Articulo
        For Each drMaterial As DataRow In data.DtMateriales.Select
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(drMaterial("IDArticulo"))
            If ArtInfo.GestionStockPorLotes AndAlso (drMaterial.RowState <> DataRowState.Deleted) AndAlso Length(drMaterial("IDOperacionMaterial")) > 0 Then

                Dim MaterialEnGlobal As List(Of DataRow) = (From c In data.DtMaterialesGlobal Where Not c.IsNull("IDOperacionMaterial") AndAlso c("IDOperacionMaterial").Equals(drMaterial("IDOperacionMaterial"))).ToList()
                If Not MaterialEnGlobal Is Nothing AndAlso MaterialEnGlobal.Count > 0 Then
                    Dim dblCantidadAcumulada As Double = 0
                    If Not data.DtMaterialesGlobalLotes Is Nothing AndAlso data.DtMaterialesGlobalLotes.Rows.Count > 0 Then
                        Dim LotesMaterialEnGlobal As List(Of DataRow) = (From c In data.DtMaterialesGlobalLotes Where Not c.IsNull("IDOperacionMaterial") AndAlso c("IDOperacionMaterial").Equals(drMaterial("IDOperacionMaterial"))).ToList()
                        If Not LotesMaterialEnGlobal Is Nothing AndAlso LotesMaterialEnGlobal.Count > 0 Then
                            For Each drLoteMaterial As DataRow In LotesMaterialEnGlobal

                                Dim DrNew As DataRow = data.DtMaterialesLotes.NewRow
                                DrNew("IDVinoMaterialLote") = Guid.NewGuid
                                DrNew("IDVinoMaterial") = drMaterial("IDVinoMaterial")
                                DrNew("Lote") = drLoteMaterial("Lote")
                                DrNew("Ubicacion") = drLoteMaterial("Ubicacion")

                                Dim dblTotalMaterial As Double = drMaterial("Cantidad") + Nz(drMaterial("Merma"), 0)
                                Dim dblTotalImpu As Double = Nz(MaterialEnGlobal(0)("Cantidad"), 0) + Nz(MaterialEnGlobal(0)("Merma"), 0)
                                If dblTotalImpu <> 0 Then
                                    DrNew("Cantidad") = drLoteMaterial("Cantidad") * dblTotalMaterial / dblTotalImpu
                                Else
                                    DrNew("Cantidad") = 0
                                End If
                                If (Length(drLoteMaterial("SeriePrecinta")) > 0) Then
                                    DrNew("SeriePrecinta") = drLoteMaterial("SeriePrecinta")
                                    DrNew("NDesde") = Nz(drLoteMaterial("NDesde"), 0) + dblCantidadAcumulada
                                    DrNew("NHasta") = Nz(DrNew("NDesde"), 0) + DrNew("Cantidad") - 1
                                End If
                                dblCantidadAcumulada += DrNew("Cantidad")
                                data.DtMaterialesLotes.Rows.Add(DrNew)
                            Next
                        End If
                    End If
                End If
            End If
        Next
        Return data.DtMaterialesLotes
    End Function

#Region " Imputación de Materiales desde Imputaciones Globales o por Tipo Operación "

    <Task()> Public Shared Function ImputacionLineasMaterialesGenerales(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        Dim blnRepartoGlobales As Boolean = False
        Dim dtMaterialesImputar As DataTable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVinoMaterialDesdeGlobales, data, services)
        If dtMaterialesImputar Is Nothing OrElse dtMaterialesImputar.Rows.Count = 0 Then
            dtMaterialesImputar = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVinoMaterialDesdeTipoOperacion, data, services)
        Else
            blnRepartoGlobales = (data.Origen = enumBdgOrigenOperacion.Real)
        End If

        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)

        '//Imputamos las lineas desde Globales o desde Tipo de Operacion
        If Not dtMaterialesImputar Is Nothing AndAlso dtMaterialesImputar.Rows.Count > 0 Then
            For Each DrMat As DataRow In dtMaterialesImputar.Select
                Dim dtMaterialesLineaGlobal As DataTable = data.DtMateriales.Clone

                For Each StDest As StDestino In data.LineaDestino
                    Dim Cantidad As Double = 0
                    Dim Merma As Double = 0
                    If data.TotalLitrosDestino <> 0 Then
                        Dim DblMerma As Double = 0
                        If DrMat.Table.Columns.Contains("Merma") Then DblMerma = Nz(DrMat("Merma"), 0)
                        Cantidad = xRound((Nz(StDest.Litros, 0) / data.TotalLitrosDestino) * Nz(DrMat("Cantidad"), 0), NUM_DECIMALES_CANTIDADES)
                        Merma = xRound((Nz(StDest.Litros, 0) / data.TotalLitrosDestino) * DblMerma, NUM_DECIMALES_CANTIDADES)
                    Else

                        Cantidad = xRound(Nz(DrMat("Cantidad"), 0), NUM_DECIMALES_CANTIDADES)
                        If DrMat.Table.Columns.Contains("Merma") Then Merma = xRound(Nz(DrMat("Merma"), 0), NUM_DECIMALES_CANTIDADES)
                    End If

                    Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacion(data.Origen, data.NOperacion, data.IDTipoOperacion, context("Fecha"), blnRepartoGlobales, StDest.IDLineaOperacionDestino, dtMaterialesLineaGlobal, DrMat)
                    datNuevaLineaImputacion.Cantidad = Cantidad
                    datNuevaLineaImputacion.Merma = Merma

                    ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacion)(AddressOf AddNuevaLineaImputacionMaterial, datNuevaLineaImputacion, services)
                Next

                If Not dtMaterialesLineaGlobal Is Nothing AndAlso dtMaterialesLineaGlobal.Rows.Count > 0 Then
                    Dim datAjustar As New DataAjustarCantidadImputaciones("Cantidad", dtMaterialesLineaGlobal, DrMat)
                    dtMaterialesLineaGlobal = ProcessServer.ExecuteTask(Of DataAjustarCantidadImputaciones, DataTable)(AddressOf AjustarCantidadImputaciones, datAjustar, services)

                    datAjustar = New DataAjustarCantidadImputaciones("Merma", dtMaterialesLineaGlobal, DrMat)
                    dtMaterialesLineaGlobal = ProcessServer.ExecuteTask(Of DataAjustarCantidadImputaciones, DataTable)(AddressOf AjustarCantidadImputaciones, datAjustar, services)

                    Dim lstImputaciones As List(Of DataRow) = (From c In dtMaterialesLineaGlobal Select c).ToList()
                    If Not lstImputaciones Is Nothing AndAlso lstImputaciones.Count > 0 Then
                        For Each dr As DataRow In lstImputaciones
                            If DrMat.Table.Columns.Contains("IDOperacionMaterial") Then dr("IDOperacionMaterial") = DrMat("IDOperacionMaterial")
                            data.DtMateriales.ImportRow(dr)
                        Next
                    End If
                End If
            Next
        End If
        Return data.DtMateriales
    End Function
    <Task()> Public Shared Function GetImputacionOperacionVinoMaterialDesdeGlobales(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If data.Origen = enumBdgOrigenOperacion.Planificada Then Exit Function
        If Not data.DtMaterialesGlobal Is Nothing Then
            Return data.DtMaterialesGlobal.Copy
        End If
    End Function
    <Task()> Public Shared Function GetImputacionOperacionVinoMaterialDesdeTipoOperacion(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 AndAlso Not Nz(data.ImputarMaterialesGlobales, False) Then
            Return New BE.DataEngine().Filter("NegBdgTipoOperacionMaterial", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))
        End If
    End Function

#End Region

#Region " Imputación de Materiales desde OFs "

    <Task()> Public Shared Sub ImputacionLineasMaterialesOF(ByVal data As StImputacion, ByVal services As ServiceProvider)
        data.DtMateriales = New BE.DataEngine().Filter("frmBdgOperacionMaterial", New NoRowsFilterItem)

        Dim dtComponentes As DataTable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVinoMaterialDesdeOF, data, services)
        If Not dtComponentes.Columns.Contains("RecalcularMaterial") Then
            dtComponentes.Columns.Add("RecalcularMaterial", GetType(Boolean))
        End If

        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)

        For Each drComponente As DataRow In dtComponentes.Select
            drComponente("RecalcularMaterial") = True
            For Each StDest As StDestino In data.LineaDestino
                Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacion(data.Origen, data.NOperacion, data.IDTipoOperacion, context("Fecha"), False, StDest.IDLineaOperacionDestino, data.DtMateriales, drComponente)
                datNuevaLineaImputacion.Cantidad = xRound(drComponente("Cantidad") * data.Cantidad * (1 + (drComponente("Merma") / 100)), NUM_DECIMALES_CANTIDADES)
                datNuevaLineaImputacion.Merma = xRound(drComponente("Merma"), NUM_DECIMALES_CANTIDADES)

                ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacion)(AddressOf AddNuevaLineaImputacionMaterial, datNuevaLineaImputacion, services)
            Next
        Next
    End Sub
    <Task()> Public Shared Function GetImputacionOperacionVinoMaterialDesdeOF(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        'Dim FilMatsOF As New Filter
        'FilMatsOF.Add("IDOrden", FilterOperator.Equal, data.IDOrden)
        'FilMatsOF.Add("EsVino", FilterOperator.Equal, False) 'TODO: igual hay que pasar esta variable data.blEsVino
        'Return New BE.DataEngine().Filter("NegBdgOrdenMateriales", FilMatsOF)

        'Dim dtMateriales As DataTable
        Dim dataOrigenOF As New DataMaterialesOF(data.IDOrden, data.Cantidad, False, Nothing)
        ProcessServer.ExecuteTask(Of DataMaterialesOF)(AddressOf MaterialesOF, dataOrigenOF, services)
        Return dataOrigenOF.dtMateriales
    End Function

#End Region

#Region " Imputación de Materiales desde Estructura del Articulo "

    <Task()> Public Shared Function GetImputacionOperacionVinoMaterialDesdeArticulo(ByVal data As StDestino, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDArticulo) > 0 Then
            Dim DtComponentes As DataTable
            Dim DrArticulo As DataRow = New Articulo().GetItemRow(data.IDArticulo)
            If DrArticulo("TipoEstructura") Then
                If Length(DrArticulo("IDTipoEstructura")) > 0 Then
                    DtComponentes = New BE.DataEngine().Filter("NegBdgArticuloMaterial", New FilterItem("IDTipoEstructura", DrArticulo("IDTipoEstructura")))
                End If
            Else
                If Len(data.IDEstructura) = 0 Then
                    Dim FilArtEst As New Filter
                    FilArtEst.Add("IDArticulo", data.IDArticulo)
                    FilArtEst.Add("Principal", True)
                    Dim DtEst As DataTable = New ArticuloEstructura().Filter(FilArtEst)
                    If Not DtEst Is Nothing AndAlso DtEst.Rows.Count > 0 Then
                        data.IDEstructura = DtEst.Rows(0)("IDEstructura")
                    End If
                End If
                If Len(data.IDEstructura) <> 0 Then
                    Dim fltr As New Filter
                    fltr.Add("IDArticulo", data.IDArticulo)
                    fltr.Add("IDEstructura", data.IDEstructura)
                    DtComponentes = New BE.DataEngine().Filter("NegBdgArticuloMaterial", fltr)
                End If
            End If
            Return DtComponentes
        End If
    End Function


    <Task()> Public Shared Sub ImputacionLineasMaterialesArticulo(ByVal data As StImputacion, ByVal services As ServiceProvider)
        For Each StDest As StDestino In data.LineaDestino
            Dim dtComponentes As DataTable = ProcessServer.ExecuteTask(Of StDestino, DataTable)(AddressOf GetImputacionOperacionVinoMaterialDesdeArticulo, StDest, services)

            If Not dtComponentes Is Nothing AndAlso dtComponentes.Rows.Count > 0 Then
                If Not dtComponentes.Columns.Contains("RecalcularMaterial") Then
                    dtComponentes.Columns.Add("RecalcularMaterial", GetType(Boolean))
                End If

                Dim context As New BusinessData
                context("Fecha") = Nz(data.Fecha, Today)

                For Each drComponente As DataRow In dtComponentes.Select
                    drComponente("RecalcularMaterial") = True

                    Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacion(data.Origen, data.NOperacion, data.IDTipoOperacion, context("Fecha"), False, StDest.IDLineaOperacionDestino, data.DtMateriales, drComponente)
                    datNuevaLineaImputacion.Cantidad = xRound(drComponente("Cantidad") * StDest.Cantidad * (1 + (drComponente("Merma") / 100)), NUM_DECIMALES_CANTIDADES)
                    datNuevaLineaImputacion.Merma = xRound(drComponente("Merma"), NUM_DECIMALES_CANTIDADES)

                    ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacion)(AddressOf AddNuevaLineaImputacionMaterial, datNuevaLineaImputacion, services)
                Next

            End If
        Next
    End Sub

#End Region

#End Region

#Region " Imputaciones Centros "

    <Task()> Public Shared Sub AddNuevaLineaImputacionCentro(ByVal data As DataAddNuevaLineaImputacion, ByVal services As ServiceProvider)

        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)

        Dim FieldVinculoGlobal As String
        Dim oEntidad As BusinessHelper
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                oEntidad = New BdgOperacionVinoPlanCentro
            Case enumBdgOrigenOperacion.Real
                oEntidad = New BdgVinoCentro
                FieldVinculoGlobal = "IDOperacionCentro"
        End Select


        Dim DrNew As DataRow = data.dtImputacion.NewRow
        If DrNew.Table.Columns.Contains("Fecha") Then DrNew("Fecha") = context("Fecha")

        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "NOperacionPlan", "NOperacion")) = data.NOperacion
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDOperacionVinoPlanCentro", "IDVinoCentro")) = Guid.NewGuid
        If Not data.IDLineaOperacionVino.Equals(Guid.Empty) Then
            DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDLineaOperacionVinoPlan", "IDVino")) = data.IDLineaOperacionVino
        End If


        DrNew = oEntidad.ApplyBusinessRule("IDCentro", data.RowOrigen("IDCentro"), DrNew, context)
        DrNew = oEntidad.ApplyBusinessRule("Tiempo", xRound(data.Tiempo, NUM_DECIMALES_CANTIDADES), DrNew, context)



        If DrNew.Table.Columns.Contains("Tasa") AndAlso data.RowOrigen.Table.Columns.Contains("TasaEjecucionA") AndAlso Not IsDBNull(data.RowOrigen("TasaEjecucionA")) Then
            DrNew("Tasa") = data.RowOrigen("TasaEjecucionA")
        End If
        If data.RowOrigen.Table.Columns.Contains("IdUdMedidaCentro") AndAlso Not IsDBNull(data.RowOrigen("IdUdMedidaCentro")) Then
            DrNew("IdUdMedidaCentro") = data.RowOrigen("IdUdMedidaCentro")
        End If
        If data.RowOrigen.Table.Columns.Contains("UdTiempoEjec") AndAlso Not IsDBNull(data.RowOrigen("UdTiempoEjec")) Then
            DrNew("UDTiempo") = data.RowOrigen("UdTiempoEjec")
        End If
        If data.RowOrigen.Table.Columns.Contains("UDTiempo") AndAlso Not IsDBNull(data.RowOrigen("UDTiempo")) Then
            DrNew("UDTiempo") = data.RowOrigen("UDTiempo")
        End If
        If data.RowOrigen.Table.Columns.Contains("PorCantidad") AndAlso Not IsDBNull(data.RowOrigen("PorCantidad")) Then
            DrNew("PorCantidad") = Nz(data.RowOrigen("PorCantidad"), False)
        Else
            If Length(data.RowOrigen("IdUdMedidaCentro")) > 0 Then
                DrNew("PorCantidad") = True
            Else : DrNew("PorCantidad") = False
            End If
        End If

        If Nz(DrNew("PorCantidad"), False) Then
            DrNew = oEntidad.ApplyBusinessRule("Cantidad", xRound(data.Cantidad, NUM_DECIMALES_CANTIDADES), DrNew, context)
        End If



        If DrNew.Table.Columns.Contains("IDTipoOperacionOrigen") Then DrNew("IDTipoOperacionOrigen") = data.IDTipoOperacion

        If data.RepartoGlobales Then
            If DrNew.Table.Columns.Contains(FieldVinculoGlobal) Then
                If data.RowOrigen.Table.Columns.Contains(FieldVinculoGlobal) AndAlso Length(data.RowOrigen(FieldVinculoGlobal)) > 0 Then
                    DrNew(FieldVinculoGlobal) = data.RowOrigen(FieldVinculoGlobal)
                End If
            End If
        End If

        data.dtImputacion.Rows.Add(DrNew)
    End Sub

    <Task()> Public Shared Function ImputacionLineaCentros(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 AndAlso Length(data.NOperacion) > 0 Then
            data.DtCentros = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetEstructuraImputacionCentros, data, services)

            '//Imputación de Centros desde Imputaciones Globales o por Tipo Operación
            ProcessServer.ExecuteTask(Of StImputacion)(AddressOf ImputacionLineasCentrosGenerales, data, services)

            '//Imputamos Centros de una OF
            If Nz(data.IDOrden, 0) > 0 Then ProcessServer.ExecuteTask(Of StImputacion)(AddressOf ImputacionLineasCentrosOF, data, services)

            Return data.DtCentros
        End If
    End Function


    <Task()> Public Shared Sub ImputacionLineasCentrosGenerales(ByVal data As StImputacion, ByVal services As ServiceProvider)
        Dim blnRepartoGlobales As Boolean = False
        Dim dtCentrosImputar As DataTable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVinoCentroDesdeGlobales, data, services)
        If dtCentrosImputar Is Nothing OrElse dtCentrosImputar.Rows.Count = 0 Then
            dtCentrosImputar = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVinoCentroDesdeTipoOperacion, data, services)
        Else
            blnRepartoGlobales = (data.Origen = enumBdgOrigenOperacion.Real)
        End If

        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)


        '//Imputamos las lineas desde Globales o desde Tipo de Operacion
        If Not dtCentrosImputar Is Nothing AndAlso dtCentrosImputar.Rows.Count > 0 Then
            For Each DrCentro As DataRow In dtCentrosImputar.Select
                Dim dtCentrosLineaGlobal As DataTable = data.DtCentros.Clone

                For Each StDest As StDestino In data.LineaDestino
                    Dim Tiempo As Double
                    If data.TotalLitrosDestino <> 0 Then
                        Tiempo = (Nz(StDest.Litros, 0) / data.TotalLitrosDestino) * Nz(DrCentro("Tiempo"), 0) ', "##,##0.00")
                    Else
                        Tiempo = Nz(DrCentro("Tiempo"), 0)
                    End If

                    Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacion(data.Origen, data.NOperacion, data.IDTipoOperacion, context("Fecha"), blnRepartoGlobales, StDest.IDLineaOperacionDestino, dtCentrosLineaGlobal, DrCentro)
                    datNuevaLineaImputacion.Tiempo = Tiempo
                    datNuevaLineaImputacion.Cantidad = StDest.Cantidad

                    ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacion)(AddressOf AddNuevaLineaImputacionCentro, datNuevaLineaImputacion, services)
                Next


                If Not dtCentrosLineaGlobal Is Nothing AndAlso dtCentrosLineaGlobal.Rows.Count > 0 Then
                    Dim datAjustar As New DataAjustarCantidadImputaciones("Tiempo", dtCentrosLineaGlobal, DrCentro)
                    dtCentrosLineaGlobal = ProcessServer.ExecuteTask(Of DataAjustarCantidadImputaciones, DataTable)(AddressOf AjustarCantidadImputaciones, datAjustar, services)

                    Dim lstImputaciones As List(Of DataRow) = (From c In dtCentrosLineaGlobal Select c).ToList()
                    If Not lstImputaciones Is Nothing AndAlso lstImputaciones.Count > 0 Then
                        For Each dr As DataRow In lstImputaciones
                            data.DtCentros.ImportRow(dr)
                        Next
                    End If
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Function GetEstructuraImputacionCentros(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                Return New BE.DataEngine().Filter("frmBdgOperacionVinoPlanCentro", New NoRowsFilterItem)
            Case enumBdgOrigenOperacion.Real
                Return New BE.DataEngine().Filter("frmBdgOperacionCentro", New NoRowsFilterItem)
        End Select
    End Function

    <Task()> Public Shared Function GetImputacionOperacionVinoCentroDesdeGlobales(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If data.Origen = enumBdgOrigenOperacion.Planificada Then Exit Function

        If Not data.DtCentrosGlobal Is Nothing AndAlso data.DtCentrosGlobal.Rows.Count > 0 Then
            Return data.DtCentrosGlobal.Copy
        End If
    End Function

    <Task()> Public Shared Function GetImputacionOperacionVinoCentroDesdeTipoOperacion(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 AndAlso Not Nz(data.ImputarCentrosGlobales, False) Then
            Return New BE.DataEngine().Filter("frmBdgTipoOperacionCentro", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))
        End If
    End Function

#Region " Imputacion de Centros desde OF "


    <Task()> Public Shared Sub ImputacionLineasCentrosOF(ByVal data As StImputacion, ByVal services As ServiceProvider)
        Dim dtCentrosOF As DataTable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVinoCentroDesdeOF, data, services)
        If dtCentrosOF Is Nothing OrElse dtCentrosOF.Rows.Count = 0 Then Exit Sub

        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)

        For Each StDest As StDestino In data.LineaDestino
            For Each drCentrosOrigen As DataRow In dtCentrosOF.Rows
                Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacion(data.Origen, data.NOperacion, data.IDTipoOperacion, context("Fecha"), False, StDest.IDLineaOperacionDestino, data.DtCentros, drCentrosOrigen)
                datNuevaLineaImputacion.Cantidad = drCentrosOrigen("TiempoEjecUnit")
                ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacion)(AddressOf AddNuevaLineaImputacionCentro, datNuevaLineaImputacion, services)
            Next
        Next

    End Sub
    <Task()> Public Shared Function GetImputacionOperacionVinoCentroDesdeOF(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Nz(data.IDOrden, 0) > 0 Then
            Return AdminData.GetData("NegBdgOrdenCentro", New NumberFilterItem("IDOrden", data.IDOrden))
        End If
    End Function

#End Region


#End Region

#Region " Imputaciones MOD "

    <Task()> Public Shared Function ImputacionLineaMOD(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 AndAlso Length(data.NOperacion) > 0 Then
            data.DtMOD = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetEstructuraImputacionMOD, data, services)

            '//Imputación de Centros desde Imputaciones Globales o por Tipo Operación
            ProcessServer.ExecuteTask(Of StImputacion)(AddressOf ImputacionLineasMODGenerales, data, services)

            Return data.DtMOD
        End If

    End Function

    <Task()> Public Shared Sub ImputacionLineasMODGenerales(ByVal data As StImputacion, ByVal services As ServiceProvider)
        Dim blnRepartoGlobales As Boolean = False
        Dim dtMODImputar As DataTable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVinoMODDesdeGlobales, data, services)
        If dtMODImputar Is Nothing OrElse dtMODImputar.Rows.Count = 0 Then
            dtMODImputar = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVinoMODDesdeTipoOperacion, data, services)
        Else
            blnRepartoGlobales = (data.Origen = enumBdgOrigenOperacion.Real)
        End If

        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)


        '//Imputamos las lineas desde Globales o desde Tipo de Operacion
        If Not dtMODImputar Is Nothing AndAlso dtMODImputar.Rows.Count > 0 Then
            For Each drMOD As DataRow In dtMODImputar.Select
                Dim dtMODLineaGlobal As DataTable = data.DtMOD.Clone

                For Each StDest As StDestino In data.LineaDestino
                    Dim Tiempo As Double
                    If data.TotalLitrosDestino <> 0 Then
                        Tiempo = (Nz(StDest.Litros, 0) / data.TotalLitrosDestino) * Nz(drMOD("Tiempo"), 0) ', "##,##0.00")
                    Else
                        Tiempo = Nz(drMOD("Tiempo"), 0)
                    End If

                    Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacion(data.Origen, data.NOperacion, data.IDTipoOperacion, context("Fecha"), blnRepartoGlobales, StDest.IDLineaOperacionDestino, dtMODLineaGlobal, drMOD)
                    datNuevaLineaImputacion.Cantidad = Tiempo
                    ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacion)(AddressOf AddNuevaLineaImputacionMOD, datNuevaLineaImputacion, services)
                Next


                If Not dtMODLineaGlobal Is Nothing AndAlso dtMODLineaGlobal.Rows.Count > 0 Then
                    Dim datAjustar As New DataAjustarCantidadImputaciones("Tiempo", dtMODLineaGlobal, drMOD)
                    dtMODLineaGlobal = ProcessServer.ExecuteTask(Of DataAjustarCantidadImputaciones, DataTable)(AddressOf AjustarCantidadImputaciones, datAjustar, services)

                    Dim lstImputaciones As List(Of DataRow) = (From c In dtMODLineaGlobal Select c).ToList()
                    If Not lstImputaciones Is Nothing AndAlso lstImputaciones.Count > 0 Then
                        For Each dr As DataRow In lstImputaciones
                            data.DtMOD.ImportRow(dr)
                        Next
                    End If
                End If
            Next
        End If
    End Sub


    <Task()> Public Shared Function GetEstructuraImputacionMOD(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                Return New BE.DataEngine().Filter("frmBdgOperacionVinoPlanMOD", New NoRowsFilterItem)
            Case enumBdgOrigenOperacion.Real
                Return New BE.DataEngine().Filter("frmBdgOperacionMod", New NoRowsFilterItem)
        End Select
    End Function

    <Task()> Public Shared Function GetImputacionOperacionVinoMODDesdeGlobales(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If data.Origen = enumBdgOrigenOperacion.Planificada Then Exit Function

        If Not data.DtMODGlobal Is Nothing AndAlso data.DtMODGlobal.Rows.Count > 0 Then
            Return data.DtMODGlobal.Copy
        End If
    End Function

    <Task()> Public Shared Function GetImputacionOperacionVinoMODDesdeTipoOperacion(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 AndAlso Not Nz(data.ImputarMODGlobales, False) Then
            Return New BE.DataEngine().Filter("frmBdgTipoOperacionMod", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))
        End If
    End Function

    <Task()> Public Shared Sub AddNuevaLineaImputacionMOD(ByVal data As DataAddNuevaLineaImputacion, ByVal services As ServiceProvider)
        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)

        Dim FieldVinculoGlobal As String
        Dim oEntidad As BusinessHelper
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                oEntidad = New BdgOperacionVinoPlanMOD
            Case enumBdgOrigenOperacion.Real
                oEntidad = New BdgVinoMod
                FieldVinculoGlobal = "IDOperacionMOD"
        End Select


        Dim DrNew As DataRow = data.dtImputacion.NewRow
        If DrNew.Table.Columns.Contains("Fecha") Then DrNew("Fecha") = context("Fecha")

        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "NOperacionPlan", "NOperacion")) = data.NOperacion
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDOperacionVinoPlanMOD", "IDVinoMOD")) = Guid.NewGuid
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDLineaOperacionVinoPlan", "IDVino")) = data.IDLineaOperacionVino

        DrNew = oEntidad.ApplyBusinessRule("IDOperario", data.RowOrigen("IDOperario"), DrNew, context)

        If Length(data.RowOrigen("IDCategoria")) > 0 AndAlso DrNew("IDCategoria") & String.Empty <> data.RowOrigen("IDCategoria") Then
            DrNew = oEntidad.ApplyBusinessRule("IDCategoria", data.RowOrigen("IDCategoria"), DrNew, context)
        End If

        DrNew = oEntidad.ApplyBusinessRule("Tiempo", xRound(data.Cantidad, NUM_DECIMALES_CANTIDADES), DrNew, context)

        If DrNew.Table.Columns.Contains("IDTipoOperacionOrigen") Then DrNew("IDTipoOperacionOrigen") = data.IDTipoOperacion

        If data.RepartoGlobales Then
            If DrNew.Table.Columns.Contains(FieldVinculoGlobal) Then
                If data.RowOrigen.Table.Columns.Contains(FieldVinculoGlobal) AndAlso Length(data.RowOrigen(FieldVinculoGlobal)) > 0 Then
                    DrNew(FieldVinculoGlobal) = data.RowOrigen(FieldVinculoGlobal)
                End If
            End If
        End If

        data.dtImputacion.Rows.Add(DrNew)
    End Sub

#End Region

#Region " Imputaciones Varios "


    <Task()> Public Shared Function ImputacionLineaVarios(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 AndAlso Length(data.NOperacion) > 0 Then
            data.DtVarios = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetEstructuraImputacionVarios, data, services)

            '//Imputación de Centros desde Imputaciones Globales o por Tipo Operación
            ProcessServer.ExecuteTask(Of StImputacion)(AddressOf ImputacionLineasVariosGenerales, data, services)

            Return data.DtVarios
        End If

    End Function

    <Task()> Public Shared Sub ImputacionLineasVariosGenerales(ByVal data As StImputacion, ByVal services As ServiceProvider)
        Dim blnRepartoGlobales As Boolean = False
        Dim dtVariosImputar As DataTable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVinoVariosDesdeGlobales, data, services)
        If dtVariosImputar Is Nothing OrElse dtVariosImputar.Rows.Count = 0 Then
            dtVariosImputar = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVinoVariosDesdeTipoOperacion, data, services)
        Else
            blnRepartoGlobales = (data.Origen = enumBdgOrigenOperacion.Real)
        End If

        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)


        '//Imputamos las lineas desde Globales o desde Tipo de Operacion
        If Not dtVariosImputar Is Nothing AndAlso dtVariosImputar.Rows.Count > 0 Then
            For Each drVarios As DataRow In dtVariosImputar.Select
                Dim dtVariosLineaGlobal As DataTable = data.DtVarios.Clone

                For Each StDest As StDestino In data.LineaDestino
                    Dim Cantidad As Double

                    If data.TotalLitrosDestino <> 0 Then
                        Cantidad = (Nz(StDest.Litros, 0) / data.TotalLitrosDestino) * Nz(drVarios("Cantidad"), 0)  ', "##,##0.00")
                    Else
                        Cantidad = Nz(drVarios("Cantidad"), 0)
                    End If

                    Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacion(data.Origen, data.NOperacion, data.IDTipoOperacion, context("Fecha"), blnRepartoGlobales, StDest.IDLineaOperacionDestino, dtVariosLineaGlobal, drVarios)
                    datNuevaLineaImputacion.Cantidad = Cantidad
                    ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacion)(AddressOf AddNuevaLineaImputacionVarios, datNuevaLineaImputacion, services)
                Next


                If Not dtVariosLineaGlobal Is Nothing AndAlso dtVariosLineaGlobal.Rows.Count > 0 Then
                    Dim datAjustar As New DataAjustarCantidadImputaciones("Cantidad", dtVariosLineaGlobal, drVarios)
                    dtVariosLineaGlobal = ProcessServer.ExecuteTask(Of DataAjustarCantidadImputaciones, DataTable)(AddressOf AjustarCantidadImputaciones, datAjustar, services)

                    Dim lstImputaciones As List(Of DataRow) = (From c In dtVariosLineaGlobal Select c).ToList()
                    If Not lstImputaciones Is Nothing AndAlso lstImputaciones.Count > 0 Then
                        For Each dr As DataRow In lstImputaciones
                            data.DtVarios.ImportRow(dr)
                        Next
                    End If
                End If
            Next
        End If
    End Sub


    <Task()> Public Shared Function GetEstructuraImputacionVarios(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                Return New BE.DataEngine().Filter("frmBdgOperacionVinoPlanVarios", New NoRowsFilterItem)
            Case enumBdgOrigenOperacion.Real
                Return New BE.DataEngine().Filter("frmBdgOperacionVarios", New NoRowsFilterItem)
        End Select
    End Function

    <Task()> Public Shared Function GetImputacionOperacionVinoVariosDesdeGlobales(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If data.Origen = enumBdgOrigenOperacion.Planificada Then Exit Function

        If Not data.DtVariosGlobal Is Nothing AndAlso data.DtVariosGlobal.Rows.Count > 0 Then
            Return data.DtVariosGlobal.Copy
        End If
    End Function

    <Task()> Public Shared Function GetImputacionOperacionVinoVariosDesdeTipoOperacion(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 AndAlso Not Nz(data.ImputarVariosGlobales, False) Then
            Return New BE.DataEngine().Filter("frmBdgTipoOperacionVarios", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))
        End If
    End Function

    <Task()> Public Shared Sub AddNuevaLineaImputacionVarios(ByVal data As DataAddNuevaLineaImputacion, ByVal services As ServiceProvider)
        Dim context As New BusinessData
        context("Fecha") = Nz(data.Fecha, Today)

        Dim FieldVinculoGlobal As String
        Dim oEntidad As BusinessHelper
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Planificada
                oEntidad = New BdgOperacionVinoPlanVarios
            Case enumBdgOrigenOperacion.Real
                oEntidad = New BdgVinoVarios
                FieldVinculoGlobal = "IDOperacionVarios"
        End Select


        Dim DrNew As DataRow = data.dtImputacion.NewRow
        If DrNew.Table.Columns.Contains("Fecha") Then DrNew("Fecha") = context("Fecha")

        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "NOperacionPlan", "NOperacion")) = data.NOperacion
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDOperacionVinoPlanVarios", "IDVinoVarios")) = Guid.NewGuid
        DrNew(IIf(data.Origen = enumBdgOrigenOperacion.Planificada, "IDLineaOperacionVinoPlan", "IDVino")) = data.IDLineaOperacionVino

        DrNew = oEntidad.ApplyBusinessRule("IDVarios", data.RowOrigen("IDVarios"), DrNew, context)
        DrNew = oEntidad.ApplyBusinessRule("Cantidad", xRound(data.Cantidad, NUM_DECIMALES_CANTIDADES), DrNew, context)

        If DrNew.Table.Columns.Contains("Tasa") Then
            DrNew("Tasa") = data.RowOrigen("Tasa")
        End If

        If DrNew.Table.Columns.Contains("IDTipoOperacionOrigen") Then DrNew("IDTipoOperacionOrigen") = data.IDTipoOperacion

        If data.RepartoGlobales Then
            If DrNew.Table.Columns.Contains(FieldVinculoGlobal) Then
                If data.RowOrigen.Table.Columns.Contains(FieldVinculoGlobal) AndAlso Length(data.RowOrigen(FieldVinculoGlobal)) > 0 Then
                    DrNew(FieldVinculoGlobal) = data.RowOrigen(FieldVinculoGlobal)
                End If
            End If
        End If

        data.dtImputacion.Rows.Add(DrNew)
    End Sub

#End Region

#Region " Imputaciones Análisis "

    <Task()> Public Shared Function ImputacionLineaAnalisis(ByVal data As StImputacion, ByVal services As ServiceProvider) As List(Of DataTable)
        If Length(data.IDAnalisis) > 0 AndAlso Length(data.NOperacion) > 0 Then
            If data.DtAnalisisVariable Is Nothing Then data.DtAnalisisVariable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetEstructuraImputacionAnalisis, data, services)

            '//Imputación de Analisis  
            ProcessServer.ExecuteTask(Of StImputacion)(AddressOf ImputacionLineasAnalisisVino, data, services)

            Dim LstData As New List(Of DataTable)
            LstData.Add(data.DtAnalisis)
            LstData.Add(data.DtAnalisisVariable)
            Return LstData
        End If
    End Function


    <Task()> Public Shared Function GetEstructuraImputacionAnalisis(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        Select Case data.Origen
            Case enumBdgOrigenOperacion.Real
                Return New BE.DataEngine().Filter("frmBdgOperacionAnalisis", New NoRowsFilterItem)
        End Select
    End Function

    <Task()> Public Shared Function GetImputacionOperacionVinoAnalisis(ByVal data As StImputacion, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDAnalisis) > 0 Then
            Return New BE.DataEngine().Filter("frmBdgAnalisisVariable", New FilterItem("IDAnalisis", FilterOperator.Equal, data.IDAnalisis))
        End If
    End Function

    <Task()> Public Shared Sub ImputacionLineasAnalisisVino(ByVal data As StImputacion, ByVal services As ServiceProvider)
        If Length(data.IDAnalisis) > 0 Then
            Dim context As New BusinessData
            context("Fecha") = Nz(data.Fecha, Today)

            data.DtAnalisis = New BdgVinoAnalisis().AddNew

            Dim dtVariables As DataTable = ProcessServer.ExecuteTask(Of StImputacion, DataTable)(AddressOf GetImputacionOperacionVinoAnalisis, data, services)
            For Each StDest As StDestino In data.LineaDestino
                Dim datNuevaLineaImputacion As New DataAddNuevaLineaImputacionAnalisis(data.NOperacion, data.IDAnalisis, context("Fecha"), StDest.IDLineaOperacionDestino, data.DtAnalisis, data.DtAnalisisVariable, dtVariables)
                ProcessServer.ExecuteTask(Of DataAddNuevaLineaImputacionAnalisis)(AddressOf AddNuevaLineaImputacionAnalisis, datNuevaLineaImputacion, services)
            Next

        End If
    End Sub

    Public Class DataAddNuevaLineaImputacionAnalisis
        Public NOperacion As String
        Public Fecha As Date
        Public IDAnalisis As String

        '//Datos Linea destino
        Public IDLineaOperacionVino As Guid
        Public Cantidad As Double
        Public Merma As Double

        '//Row Origen de las imputaciones
        Public dtOrigenImputacion As DataTable

        '//
        Public dtAnalisis As DataTable
        Public dtAnalisisVariable As DataTable

        Public Sub New(ByVal NOperacion As String, ByVal IDAnalisis As String, ByVal Fecha As Date, ByVal IDLineaOperacionVino As Guid, ByVal dtAnalisis As DataTable, ByVal dtAnalisisVariable As DataTable, ByVal dtOrigenImputacion As DataTable)

            Me.NOperacion = NOperacion
            Me.IDAnalisis = IDAnalisis
            Me.Fecha = Fecha

            Me.IDLineaOperacionVino = IDLineaOperacionVino

            Me.dtAnalisis = dtAnalisis
            Me.dtAnalisisVariable = dtAnalisisVariable
            Me.dtOrigenImputacion = dtOrigenImputacion
        End Sub
    End Class
    <Task()> Public Shared Sub AddNuevaLineaImputacionAnalisis(ByVal data As DataAddNuevaLineaImputacionAnalisis, ByVal services As ServiceProvider)
        Dim context As New BusinessData
        context("Fecha") = data.Fecha


        Dim IDVinoAnalisis As Guid = Guid.NewGuid
        Dim dtNewAnalisis As DataTable = New BdgVinoAnalisis().AddNewForm
        dtNewAnalisis.Rows(0)("IDAnalisis") = data.IDAnalisis
        dtNewAnalisis.Rows(0)("IDVino") = data.IDLineaOperacionVino
        dtNewAnalisis.Rows(0)("IDVinoAnalisis") = IDVinoAnalisis
        dtNewAnalisis.Rows(0)("NOperacion") = data.NOperacion
        If Nz(context("Fecha"), cnMinDate) <> cnMinDate Then dtNewAnalisis.Rows(0)("Fecha") = context("Fecha")
        data.dtAnalisis.Rows.Add(dtNewAnalisis.Rows(0).ItemArray)


        For Each DrVar As DataRow In data.dtOrigenImputacion.Select
            Dim DrNewVar As DataRow = data.dtAnalisisVariable.NewRow
            DrNewVar("IDVinoAnalisis") = IDVinoAnalisis

            If data.dtAnalisisVariable.Columns.Contains("IDVino") Then DrNewVar("IDVino") = data.IDLineaOperacionVino

            If data.dtAnalisisVariable.Columns.Contains("NOperacion") Then DrNewVar("NOperacion") = data.NOperacion
            If data.dtAnalisisVariable.Columns.Contains("Fecha") Then If Nz(context("Fecha"), cnMinDate) <> cnMinDate Then DrNewVar("Fecha") = context("Fecha")

            For Each dc As DataColumn In data.dtAnalisisVariable.Columns
                If data.dtAnalisisVariable.Columns.Contains(dc.ColumnName) AndAlso _
                   data.dtOrigenImputacion.Columns.Contains(dc.ColumnName) Then

                    DrNewVar(dc.ColumnName) = DrVar(dc.ColumnName)

                End If
                If data.dtAnalisisVariable.Columns.Contains("IDVino") Then DrNewVar("IDVino") = data.IDLineaOperacionVino
            Next
            data.dtAnalisisVariable.Rows.Add(DrNewVar)
        Next
    End Sub

#End Region

#End Region

#End Region

#End Region

#Region " Ordenes de Fabricación "

    <Serializable()> _
    Public Class DataDatosOfParaOperacion
        Public IDOrden As Integer
        Public dblCantidad As Double
        Public dtOrigen As DataTable
        Public dtDestino As DataTable
        Public dtMateriales As DataTable
        Public dtCentros As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDOrden As Integer, ByVal dblCantidad As Double)
            Me.IDOrden = IDOrden
            Me.dblCantidad = dblCantidad
        End Sub
    End Class

    <Task()> Public Shared Function DatosOfParaOperacion(ByVal data As DataDatosOfParaOperacion, ByVal services As ServiceProvider) As DataDatosOfParaOperacion

        'Origen
        Dim dataOrigenOF As New DataMaterialesOF(data.IDOrden, data.dblCantidad, True, data.dtOrigen)
        ProcessServer.ExecuteTask(Of DataMaterialesOF)(AddressOf MaterialesOF, dataOrigenOF, services)
        data.dtOrigen = dataOrigenOF.dtMateriales
        If Not data.dtOrigen Is Nothing Then
            For Each dr As DataRow In data.dtOrigen.Rows
                If Length(dr("IDAlmacen")) > 0 AndAlso Length(dr("IDArticulo")) > 0 AndAlso Length(dr("Lote")) > 0 AndAlso Length(dr("Ubicacion")) > 0 Then
                    Dim dataVQ As New ProcesoStocks.DataVinoQ(dr("IDArticulo"), dr("Ubicacion"), dr("Lote"), dr("IDAlmacen"))
                    dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, services)
                    If Not dataVQ Is Nothing Then
                        dr("IDVino") = dataVQ.IDVino
                    Else
                        ApplicationService.GenerateError("No se ha encontrado el vino identificado con{0} Depósito: {1}{2} Artículo: {3}{4} Lote: {5}{6} Almacén: {7}", vbNewLine, dr("Ubicacion"), vbNewLine, dr("IDArticulo"), vbNewLine, dr("Lote"), vbNewLine, dr("IDAlmacen"))
                    End If
                End If
            Next
        End If

        'Destino
        Dim dataDestinoOF As New DataCabeceraOF(data.IDOrden)
        ProcessServer.ExecuteTask(Of DataCabeceraOF)(AddressOf CabeceraOF, dataDestinoOF, services)
        data.dtDestino = dataDestinoOF.dtOrdenFabricacion
        If Not data.dtDestino Is Nothing Then
            data.dtDestino.Columns.Add("IDVino", GetType(Guid))
            For Each dr As DataRow In data.dtDestino.Rows
                If Length(dr("IDAlmacen")) > 0 AndAlso Length(dr("IDArticulo")) > 0 AndAlso Length(dr("Lote")) > 0 AndAlso Length(dr("IDUbicacion")) > 0 Then
                    Dim dataVQ As New ProcesoStocks.DataVinoQ(dr("IDArticulo"), dr("IDUbicacion"), dr("Lote"), dr("IDAlmacen"))
                    dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, services)
                    If Not dataVQ Is Nothing Then
                        dr("IDVino") = dataVQ.IDVino
                    End If
                End If
            Next
        End If



        'Materiales
        dataOrigenOF = New DataMaterialesOF(data.IDOrden, data.dblCantidad, False, data.dtMateriales)
        ProcessServer.ExecuteTask(Of DataMaterialesOF)(AddressOf MaterialesOF, dataOrigenOF, services)
        data.dtMateriales = dataOrigenOF.dtMateriales

        Return data
    End Function


    <Serializable()> _
   Public Class DataMaterialesOF
        Public IDOrden As Integer
        Public dblCantidad As Double
        Public dtMateriales As DataTable
        Public blEsVino As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDOrden As Integer, ByVal dblCantidad As Double, ByVal blEsVino As Boolean, ByVal dtMateriales As DataTable)
            Me.IDOrden = IDOrden
            Me.dblCantidad = dblCantidad
            Me.blEsVino = blEsVino
            Me.dtMateriales = dtMateriales
        End Sub
    End Class

    <Task()> Public Shared Sub MaterialesOF(ByVal data As DataMaterialesOF, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDOrden", data.IDOrden))
        f.Add(New BooleanFilterItem("EsVino", FilterOperator.Equal, data.blEsVino))
        If data.blEsVino Then
            f.Add(New IsNullFilterItem("IDAlmacen", False))
            f.Add(New IsNullFilterItem("IDArticulo", False))
            f.Add(New IsNullFilterItem("Lote", False))
            f.Add(New IsNullFilterItem("Ubicacion", False))
        End If
        Dim dtComponentes As DataTable = AdminData.GetData("NegBdgOrdenMateriales", f)
        If data.dtMateriales Is Nothing Then
            data.dtMateriales = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetDTMateriales, Nothing, services)
        End If
        For Each drComponente As DataRow In dtComponentes.Rows
            Dim drMaterial As DataRow = data.dtMateriales.NewRow
            data.dtMateriales.Rows.Add(drMaterial)
            drMaterial("IDAlmacen") = drComponente("IDAlmacen")
            drMaterial("IDArticulo") = drComponente("IDArticulo")
            drMaterial("Lote") = drComponente("Lote")
            drMaterial("Ubicacion") = drComponente("Ubicacion")
            drMaterial("IDDeposito") = drComponente("Ubicacion")
            drMaterial("Cantidad") = xRound(drComponente("Cantidad") * data.dblCantidad * (1 + (drComponente("Merma") / 100)), NUM_DECIMALES_CANTIDADES)
            drMaterial("Merma") = xRound(drComponente("Merma"), NUM_DECIMALES_CANTIDADES)
            drMaterial("DescArticulo") = drComponente("DescArticulo")
            drMaterial("GestionStockPorLotes") = drComponente("GestionStockPorLotes")
            drMaterial("RecalcularMaterial") = True
            Dim StPrecio As New BdgGeneral.StObtenerPrecio(drMaterial("IDArticulo"), drMaterial("IDAlmacen"))
            drMaterial("Precio") = ProcessServer.ExecuteTask(Of BdgGeneral.StObtenerPrecio, Double)(AddressOf BdgGeneral.ObtenerPrecio, StPrecio, services)
            drMaterial("GestionStockPorLotes") = drComponente("GestionStockPorLotes")
        Next
    End Sub


    <Task()> Public Shared Function GetDTMateriales(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim VM As New Bodega.BdgVinoMaterial
        Dim materiales As DataTable = VM.AddNew()
        materiales.Columns.Add("Lote", GetType(String))
        materiales.Columns.Add("Ubicacion", GetType(String))
        materiales.Columns.Add("IDDeposito", GetType(String))
        materiales.Columns.Add("DescArticulo", GetType(String))
        materiales.Columns.Add("GestionStockPorLotes", GetType(Boolean))
        materiales.Columns.Add("Importe", GetType(Double))
        Return materiales
    End Function


    <Serializable()> _
    Public Class DataCabeceraOF
        Public IDOrden As Integer
        Public dtOrdenFabricacion As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDOrden As Integer)
            Me.IDOrden = IDOrden
        End Sub
    End Class

    <Task()> Public Shared Sub CabeceraOF(ByVal data As DataCabeceraOF, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDOrden", data.IDOrden))
        Dim dtOrden As DataTable = New BE.DataEngine().Filter("NegBdgOrdenCabecera", f)
        If Not dtOrden Is Nothing AndAlso dtOrden.Rows.Count > 0 Then
            data.dtOrdenFabricacion = dtOrden.Copy
        End If
    End Sub

#End Region

End Class

Public Enum enumBdgOrigenOperacion
    Real = 0
    Planificada = 1
End Enum