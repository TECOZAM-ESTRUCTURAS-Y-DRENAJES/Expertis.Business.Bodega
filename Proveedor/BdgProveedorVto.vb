Public Class BdgProveedorVto

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgProveedorVto"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarBorrado)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarEntradaVtoRelacionadas)
    End Sub

    <Task()> Public Shared Sub ActualizarBorrado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_PV.IDProveedorGrupo)) > 0 Then
            Dim ClsProv As New BdgProveedorVto
            Dim DrPV As DataRow = ClsProv.GetItemRow(data(_PV.IDVendimiaVto), data(_PV.IDProveedorGrupo), data(_PV.TipoVariedad))
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarDesglosado, DrPV, services)
            AdminData.SetData(DrPV.Table)
        End If
    End Sub

    <Task()> Public Shared Sub BorrarEntradaVtoRelacionadas(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsBdgEntVto As New BdgEntradaVto
        Dim FilEntVto As New Filter
        FilEntVto.Add(New GuidFilterItem("IDVendimiaVto", data("IDVendimiaVto")))
        FilEntVto.Add(New StringFilterItem("IDProveedor", data("IDProveedor")))
        FilEntVto.Add(New NumberFilterItem("TipoVariedad", data("TipoVariedad")))
        Dim DtEntVto As DataTable = ClsBdgEntVto.Filter(FilEntVto)
        If Not DtEntVto Is Nothing AndAlso DtEntVto.Rows.Count > 0 Then
            ClsBdgEntVto.Delete(DtEntVto)
        End If
    End Sub

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("Kilos", AddressOf ActualizarImportes)
        Obrl.Add("KilosSO", AddressOf ActualizarImportes)
        Obrl.Add("KilosO", AddressOf ActualizarImportes)
        Obrl.Add("KilosExc", AddressOf ActualizarImportes)
        Obrl.Add("Precio", AddressOf ActualizarImportes)
        Obrl.Add("PrecioSO", AddressOf ActualizarImportes)
        Obrl.Add("PrecioO", AddressOf ActualizarImportes)
        Obrl.Add("PrecioExc", AddressOf ActualizarImportes)
        Obrl.Add("KilosFra", AddressOf ActualizarImportes)
        Obrl.Add("PrecioFra", AddressOf ActualizarImportes)
        Return Obrl
    End Function

    <Task()> Public Shared Sub ActualizarImportes(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim Kilos, KilosSO, KilosO, KilosExc, KilosFra As Double
        Dim Importe, ImporteSO, ImporteO, ImporteExc, ImporteFra As Double
        Dim Precio, PrecioSO, PrecioO, PrecioExc, PrecioFra As Double
        data.Current(data.ColumnName) = data.Value

        Kilos = Nz(data.Current(_PV.Kilos), 0)
        KilosSO = Nz(data.Current(_PV.KilosSO), 0)
        KilosO = Nz(data.Current(_PV.KilosO), 0)
        KilosExc = Nz(data.Current(_PV.KilosExc), 0)

        Precio = Nz(data.Current(_PV.Precio), 0)
        PrecioSO = Nz(data.Current(_PV.PrecioSO), 0)
        PrecioO = Nz(data.Current(_PV.PrecioO), 0)
        PrecioExc = Nz(data.Current(_PV.PrecioExc), 0)

        Importe = Kilos * Precio
        ImporteSO = KilosSO * PrecioSO
        ImporteO = KilosO * PrecioO
        ImporteExc = KilosExc * PrecioExc

        data.Current(_PV.Importe) = Importe
        data.Current(_PV.ImporteSO) = ImporteSO
        data.Current(_PV.ImporteO) = ImporteO
        data.Current(_PV.ImporteExc) = ImporteExc

        If CBool(data.Current(_PV.Desglosado)) Or Length(data.Current(_PV.IDProveedorGrupo)) <> 0 Then
            data.Current(_PV.ImporteFra) = Nz(data.Current(_PV.KilosFra), 0) * Nz(data.Current(_PV.PrecioFra), 0)
        Else
            data.Current(_PV.KilosFra) = Kilos + KilosExc
            data.Current(_PV.ImporteFra) = Importe + ImporteSO + ImporteO + ImporteExc

            If Kilos + KilosExc > 0 Then
                data.Current(_PV.PrecioFra) = (Importe + ImporteSO + ImporteO + ImporteExc) / (Kilos + KilosExc)
            Else
                data.Current(_PV.PrecioFra) = 0
            End If
        End If
    End Sub

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRegistroExistente)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDProveedor")) = 0 Then ApplicationService.GenerateError("El Proveedor es un dato obligatorio.")
        If Nz(data("TipoVariedad"), -1) = -1 Then ApplicationService.GenerateError("La Variedad es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarRegistroExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added OrElse _
          (data.RowState = DataRowState.Modified AndAlso (data("IDProveedor") <> data("IDProveedor", DataRowVersion.Original) OrElse data("TipoVariedad") <> data("TipoVariedad", DataRowVersion.Original))) Then
            Dim f As New Filter
            f.Add(New FilterItem("IDVendimiaVto", data("IDVendimiaVto")))
            f.Add(New FilterItem("IDProveedor", data("IDProveedor")))
            f.Add(New FilterItem("TipoVariedad", data("TipoVariedad")))
            Dim dt As DataTable = New BdgProveedorVto().Filter(f)
            If dt.Rows.Count > 0 Then
                Dim Variedad As String = IIf(data("TipoVariedad") = BdgTipoVariedad.Blanca, AdminData.GetMessageText("Blanca"), AdminData.GetMessageText("Tinta"))
                ApplicationService.GenerateError("Ya existe un registro para el proveedor {0} y la variedad {1}.", Quoted(data("IDProveedor")), Quoted(Variedad))
            End If
        End If
    End Sub


#End Region

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarDatos)
    End Sub

    <Task()> Public Shared Sub ActualizarDatos(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsProvVto As New BdgProveedorVto
        Dim DrPV As DataRow = ClsProvVto.GetItemRow(data(_PV.IDVendimiaVto), data(_PV.IDProveedor), data(_PV.TipoVariedad))
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarDesglosado, DrPV, services)
        AdminData.SetData(DrPV.Table)
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub DeleteVto(ByVal IDVendimiaVto As Guid, ByVal services As ServiceProvider)
        'TODO - Revisar Borrados
        Dim fDeleteEV As Filter = New Filter
        fDeleteEV.Add(New GuidFilterItem(_EVto.IDVendimiaVto, IDVendimiaVto))
        fDeleteEV.Add(New IsNullFilterItem(_EVto.IDLineaFactura))
        fDeleteEV.Add(New IsNullFilterItem(_EVto.IDLineaFacturaExc))
        fDeleteEV.Add(New IsNullFilterItem(_EVto.IDLineaFacturaO))
        fDeleteEV.Add(New IsNullFilterItem(_EVto.IDLineaFacturaSO))
        AdminData.Execute("delete from tbBdgEntradaVto where " & AdminData.ComposeFilter(fDeleteEV))

        Dim oDlFltr As Filter = New Filter
        oDlFltr.Add(New GuidFilterItem(_PV.IDVendimiaVto, IDVendimiaVto))
        oDlFltr.Add(New IsNullFilterItem(_PV.IDFactura))
        AdminData.Execute("delete from tbBdgProveedorVto where " & AdminData.ComposeFilter(oDlFltr))
    End Sub

    <Serializable()> _
    Public Class StCalcularVencimientos
        Public Vendimia As Integer
        Public TipoCalculo As BdgTipoCalculo
        Public IDVendimiaVto As Guid
        Public PrecioKiloT As Double
        Public PrecioKiloB As Double
        Public IDTarifaT As String
        Public IDTarifaB As String
        Public Filter As Filter
        Public TipoOrigenCalculo As enumBdgTipoCalculoVtos

        Public Sub New()
        End Sub

        Public Sub New(ByVal Vendimia As Integer, ByVal TipoCalculo As BdgTipoCalculo, ByVal IDVendimiaVto As Guid, ByVal PrecioKiloT As Double, _
                       ByVal PrecioKiloB As Double, ByVal IDTarifaT As String, ByVal IDTarifaB As String, Optional ByVal Filter As Filter = Nothing, _
                       Optional ByVal tipoOrigenCalculo As enumBdgTipoCalculoVtos = enumBdgTipoCalculoVtos.PorCabecera)
            Me.Vendimia = Vendimia
            Me.TipoCalculo = TipoCalculo
            Me.IDVendimiaVto = IDVendimiaVto
            Me.PrecioKiloT = PrecioKiloT
            Me.PrecioKiloB = PrecioKiloB
            Me.IDTarifaT = IDTarifaT
            Me.IDTarifaB = IDTarifaB
            Me.Filter = Filter
            Me.TipoOrigenCalculo = tipoOrigenCalculo
        End Sub
    End Class

    <Task()> Public Shared Sub CalcularVencimientos(ByVal data As StCalcularVencimientos, ByVal services As ServiceProvider)
        If Not data Is Nothing Then
            Select Case data.TipoOrigenCalculo
                Case enumBdgTipoCalculoVtos.PorCabecera
                    ProcessServer.ExecuteTask(Of StCalcularVencimientos)(AddressOf CalcularVencimientosEstandar, data, services)
                Case enumBdgTipoCalculoVtos.PorCartillista
                    ProcessServer.ExecuteTask(Of StCalcularVencimientos)(AddressOf CalcularVencimientosEstandar, data, services)
                Case enumBdgTipoCalculoVtos.PorFacturacion
                    ProcessServer.ExecuteTask(Of StCalcularVencimientos)(AddressOf CalcularVencimientosFacturacion, data, services)
            End Select
        End If
    End Sub

    <Task()> Public Shared Sub CalcularVencimientosEstandar(ByVal data As StCalcularVencimientos, ByVal services As ServiceProvider)
        Dim ClsBdgProvVto As New BdgProveedorVto
        Dim ClsBdgEntrVto As New BdgEntradaVto
        Dim dtProvVto As DataTable = ClsBdgProvVto.AddNew
        Dim dtEntradaVto As DataTable = ClsBdgEntrVto.AddNew

        Dim fVendimia As New Filter
        fVendimia.Add(New NumberFilterItem(_E.Vendimia, FilterOperator.Equal, data.Vendimia))
        If Not data.Filter Is Nothing Then fVendimia.Add(data.Filter)

        dtProvVto.DefaultView.Sort = _PV.IDProveedor & ", " & _PV.TipoVariedad

        Dim dtEntradaCartillista As DataTable
        Dim dtEntradaVariable As DataTable

        Dim oPrm As New BdgParametro
        Dim blnFacturarEnDetalle As Boolean = oPrm.CalculoFraEnDetalle()
        Dim blnAplicarDesglose As Boolean = oPrm.DesgloseEnFacturacion

        ProcessServer.ExecuteTask(Of Guid)(AddressOf DeleteVto, data.IDVendimiaVto, services)

        '//Registros facturados que no se han borrado
        Dim dtFrado As DataTable = New BdgProveedorVto().Filter(New GuidFilterItem(_EVto.IDVendimiaVto, data.IDVendimiaVto), _PV.IDProveedor & ", " & _PV.TipoVariedad)
        Dim dvFrado As DataView = dtFrado.DefaultView
        dvFrado.Sort = _PV.IDProveedor & ", " & _PV.TipoVariedad


        If blnFacturarEnDetalle Then
            dtEntradaCartillista = New BE.DataEngine().Filter("NegBdgFacturacionCartillista", fVendimia, , _EC.IDEntrada)
            dtEntradaCartillista.DefaultView.Sort = _EC.IDEntrada
        End If

        dtEntradaVariable = New BE.DataEngine().Filter("NegBdgFacturacionVariable", fVendimia, , _EA.IDEntrada & ", " & _EA.IDVariable)
        dtEntradaVariable.DefaultView.Sort = _EA.IDEntrada & ", " & _EA.IDVariable

        Dim strVarGrado As String = oPrm.VariableGrado

        Dim dtEntrada As DataTable = AdminData.GetData("NegBdgFacturacion", fVendimia, , _C.IDProveedor & ", " & _E.TipoVariedad)
        For Each drEntrada As DataRow In dtEntrada.Rows
            Dim strIDProveedor As String = String.Empty
            Dim strIDProveedorFra As String
            Dim intTipoUva As BdgTipoVariedad
            Dim rwPV As DataRow
            Dim oTrf As Tarifa
            Dim strIDTarifa As String
            Dim oP As BdgProveedor
            Dim dblPrecioDeclarado, dblPrecioExcedente, dblPrecioOrigen As Double
            Dim strIDVariableTarifa As String = strVarGrado
            Dim dblVariableValor As Double
            Dim dblVariableValorTarifa As Double
            Dim IDEntrada As Long = drEntrada(_E.IDEntrada)
            Dim dblNeto As Double = drEntrada(_E.Neto)
            Dim dblExcedente As Double = drEntrada(_E.Excedente)
            Dim dblDclrdo As Double
            Dim dblKilosDeclarado, dblKilosOrigen, dblKilosSinOrigen As Double

            If strIDProveedor <> Nz(drEntrada(_C.IDProveedor), String.Empty) OrElse intTipoUva <> drEntrada(_E.TipoVariedad) Then
                strIDProveedor = drEntrada(_C.IDProveedor)
                If drEntrada.IsNull(_E.IDProveedorFra) Then
                    strIDProveedorFra = strIDProveedor
                Else
                    strIDProveedorFra = drEntrada(_E.IDProveedorFra)
                End If
                intTipoUva = drEntrada(_E.TipoVariedad)

                If oP Is Nothing Then oP = New BdgProveedor

                Dim DataGetTarifa As New BdgProveedor.DataGetTarifa(strIDProveedorFra, drEntrada(_E.IDVariedad), data.Vendimia)
                oTrf = ProcessServer.ExecuteTask(Of BdgProveedor.DataGetTarifa, Business.Bodega.Tarifa)(AddressOf BdgProveedor.GetTarifa, DataGetTarifa, services)
                If intTipoUva = BdgTipoVariedad.Tinta Then
                    dblPrecioOrigen = oTrf.PrecioOrigenT
                    dblPrecioExcedente = oTrf.PrecioExcedenteT
                Else
                    dblPrecioOrigen = oTrf.PrecioOrigenB
                    dblPrecioExcedente = oTrf.PrecioExcedenteB
                End If

                Dim StDefView As New StGetRowFromView
                StDefView.vw = dtProvVto.DefaultView
                ReDim StDefView.Keys(1)
                StDefView.Keys(0) = strIDProveedorFra
                StDefView.Keys(1) = intTipoUva
                rwPV = ProcessServer.ExecuteTask(Of StGetRowFromView, DataRow)(AddressOf GetRowFromView, StDefView, services)
                If rwPV Is Nothing Then
                    '//Si el registro no está previamente facturado se crea una linea
                    Dim StDefViewPV As New StGetRowFromView
                    StDefViewPV.vw = dvFrado
                    ReDim StDefViewPV.Keys(1)
                    StDefViewPV.Keys(0) = strIDProveedorFra
                    StDefViewPV.Keys(1) = intTipoUva
                    If ProcessServer.ExecuteTask(Of StGetRowFromView, DataRow)(AddressOf GetRowFromView, StDefViewPV, services) Is Nothing Then
                        Dim StNewrowPV As New StNewRow(dtProvVto, data.IDVendimiaVto, strIDProveedorFra, intTipoUva, dblPrecioExcedente, dblPrecioOrigen)
                        rwPV = ProcessServer.ExecuteTask(Of StNewRow, DataRow)(AddressOf NewRow, StNewrowPV, services)
                    Else
                        rwPV = Nothing
                    End If
                End If
            End If

            If data.TipoCalculo = BdgTipoCalculo.Predeterminado Then
                If intTipoUva = BdgTipoVariedad.Tinta Then
                    strIDTarifa = oTrf.IDTarifaT
                Else
                    strIDTarifa = oTrf.IDTarifaB
                End If
            ElseIf data.TipoCalculo = BdgTipoCalculo.TarifaEspecifica Then
                If intTipoUva = BdgTipoVariedad.Tinta Then
                    strIDTarifa = data.IDTarifaT
                Else
                    strIDTarifa = data.IDTarifaB
                End If
            Else
                If intTipoUva = BdgTipoVariedad.Tinta Then
                    dblPrecioDeclarado = data.PrecioKiloT
                Else
                    dblPrecioDeclarado = data.PrecioKiloB
                End If
            End If

            If data.TipoCalculo <> BdgTipoCalculo.PorKilo Then
                'Obtener los datos de la tarifa.
                Dim bdEntrada As New BusinessData(drEntrada)
                Dim DataDatosTarifa As New DataDatosTarifa(IDEntrada, strIDTarifa, strVarGrado, drEntrada(_C.IDProveedor), dtEntradaVariable, data.TipoCalculo, fVendimia)
                ProcessServer.ExecuteTask(Of DataDatosTarifa)(AddressOf DatosTarifa, DataDatosTarifa, services)
                strIDVariableTarifa = DataDatosTarifa.strIDVariableTarifa
                dblVariableValor = DataDatosTarifa.dblVariableValor
                dblVariableValorTarifa = DataDatosTarifa.dblVariableValorTarifa
                dblPrecioDeclarado = DataDatosTarifa.dblPrecioDeclarado
            End If

            If blnFacturarEnDetalle Then
                dblDclrdo = 0
                For Each oRwV As DataRowView In dtEntradaCartillista.DefaultView.FindRows(IDEntrada)
                    If oRwV(_C.IDProveedor) = strIDProveedor Then
                        dblDclrdo += oRwV(_EC.Declarado)
                        'eliminamos las filas tratadas de la lista para tratar las que queden al final
                        dtEntradaCartillista.Rows.Remove(oRwV.Row)
                    End If
                Next
            Else
                dblDclrdo = drEntrada(_E.Declarado)
            End If

            If dblNeto > dblDclrdo Then
                dblKilosDeclarado = dblDclrdo - dblExcedente
                dblKilosOrigen = 0
                dblKilosSinOrigen = dblNeto - dblDclrdo
            Else
                dblKilosDeclarado = dblNeto - dblExcedente
                dblKilosOrigen = dblDclrdo - dblNeto
                dblKilosSinOrigen = 0
            End If

            'Guardar el detalle por Entrada.
            Dim DataEntradaVto As New DataEntradaVto(dtEntradaVto, IDEntrada, data.IDVendimiaVto, _
                                                     strIDProveedorFra, intTipoUva, dblVariableValor, _
                                                     dblVariableValorTarifa, strIDVariableTarifa, strIDTarifa, _
                                                     dblPrecioDeclarado, dblKilosDeclarado, dblPrecioExcedente, _
                                                     dblExcedente, dblPrecioOrigen, dblKilosOrigen, _
                                                     dblPrecioDeclarado - dblPrecioOrigen, dblKilosSinOrigen)
            ProcessServer.ExecuteTask(Of DataEntradaVto)(AddressOf BdgEntradaVto.NuevaEntradaVto, DataEntradaVto, services)

            If Not rwPV Is Nothing Then
                rwPV(_PV.Kilos) += dblKilosDeclarado
                rwPV(_PV.Importe) += dblPrecioDeclarado * dblKilosDeclarado
                rwPV(_PV.KilosExc) += dblExcedente
                rwPV(_PV.KilosO) += dblKilosOrigen
                rwPV(_PV.ImporteO) += dblPrecioOrigen * dblKilosOrigen
                rwPV(_PV.KilosSO) += dblKilosSinOrigen
                rwPV(_PV.ImporteSO) += (dblPrecioDeclarado - dblPrecioOrigen) * dblKilosSinOrigen
            End If
        Next

        '//se procesan los registros de cartillistas que no se hayan procesado
        '//estos son registros de declaración de cantidades para proveedores que no son el de cabecera de entrada
        If blnFacturarEnDetalle Then
            dtEntradaCartillista.DefaultView.Sort = _C.IDProveedor
            Dim oP As BdgProveedor = New BdgProveedor
            For Each oRwv As DataRowView In dtEntradaCartillista.DefaultView
                Dim strIDPrv As String
                Dim oTrf As Tarifa
                Dim dblPrecioO As Double
                Dim rwPV As DataRow

                If strIDPrv <> oRwv(_C.IDProveedor) Then
                    strIDPrv = oRwv(_C.IDProveedor)
                    Dim StGet As New BdgProveedor.DataGetTarifa(strIDPrv, oRwv(_EC.IDVariedad), data.Vendimia)
                    oTrf = ProcessServer.ExecuteTask(Of BdgProveedor.DataGetTarifa, Bodega.Tarifa)(AddressOf BdgProveedor.GetTarifa, StGet, services)
                End If

                If CType(oRwv(_EC.TipoVariedad), BdgTipoVariedad) = BdgTipoVariedad.Tinta Then
                    dblPrecioO = oTrf.PrecioOrigenT
                Else : dblPrecioO = oTrf.PrecioOrigenB
                End If

                Dim StViewPV As New StGetRowFromView
                StViewPV.vw = dtProvVto.DefaultView
                ReDim StViewPV.Keys(1)
                StViewPV.Keys(0) = oRwv(_C.IDProveedor)
                StViewPV.Keys(1) = oRwv(_EC.TipoVariedad)
                rwPV = ProcessServer.ExecuteTask(Of StGetRowFromView, DataRow)(AddressOf GetRowFromView, StViewPV, services)
                If rwPV Is Nothing Then
                    '//Si el registro no está previamente facturado se crea una linea
                    Dim StViewGrado As New StGetRowFromView
                    StViewGrado.vw = dvFrado
                    ReDim StViewGrado.Keys(1)
                    StViewGrado.Keys(0) = oRwv(_C.IDProveedor)
                    StViewGrado.Keys(1) = oRwv(_EC.TipoVariedad)
                    If ProcessServer.ExecuteTask(Of StGetRowFromView, DataRow)(AddressOf GetRowFromView, StViewGrado, services) Is Nothing Then
                        Dim StNew As New StNewRow(dtProvVto, data.IDVendimiaVto, oRwv(_C.IDProveedor), oRwv(_EC.TipoVariedad), 0, dblPrecioO)
                        rwPV = ProcessServer.ExecuteTask(Of StNewRow, DataRow)(AddressOf NewRow, StNew, services)
                    Else : rwPV = Nothing
                    End If
                End If

                rwPV(_PV.KilosO) += oRwv(_EC.Declarado)
                rwPV(_PV.ImporteO) += dblPrecioO * oRwv(_EC.Declarado)
            Next
        End If

        '//Ajustar segun máximos en cartilla
        Dim dtMx As DataTable = AdminData.GetData("NegBdgFacturacionMax", fVendimia, , _C.IDProveedor)
        Dim vwMx As DataView = New DataView(dtMx)
        vwMx.Sort = _C.IDProveedor
        For Each rwPv As DataRow In dtProvVto.Rows
            '//Solo si no ha hay una asignación de Excedente manual
            If rwPv(_PV.KilosExc) = 0 Then
                Dim StView As New StGetRowFromView
                StView.vw = vwMx
                ReDim StView.Keys(0)
                StView.Keys(0) = rwPv(_PV.IDProveedor)
                Dim rwMx As DataRow = ProcessServer.ExecuteTask(Of StGetRowFromView, DataRow)(AddressOf GetRowFromView, StView, services)
                Dim dblMax As Double = 0
                If Not rwMx Is Nothing Then
                    If CType(rwPv(_PV.TipoVariedad), Bodega.BdgTipoVariedad) = BdgTipoVariedad.Tinta Then
                        dblMax = rwMx(_CV.MaxT)
                        If dblMax = 0 Then ApplicationService.GenerateError("El Cupo Máximo de Tinta para el Proveedor | es 0.", rwPv(_PV.IDProveedor))
                    Else
                        dblMax = rwMx(_CV.MaxB)
                        If dblMax = 0 Then ApplicationService.GenerateError("El Cupo Máximo de Blanca para el Proveedor | es 0.", rwPv(_PV.IDProveedor))
                    End If
                End If
                If rwPv(_PV.Kilos) > dblMax Then
                    rwPv(_PV.KilosExc) = rwPv(_PV.Kilos) - dblMax
                    rwPv(_PV.Kilos) = dblMax
                    rwPv(_PV.Importe) = rwPv(_PV.Importe) * (rwPv(_PV.Kilos) / (dblMax + rwPv(_PV.KilosExc)))
                End If

                If rwPv(_PV.Kilos) <> 0 Then rwPv(_PV.Precio) = rwPv(_PV.Importe) / rwPv(_PV.Kilos)
                'If rwPV(_PV.KilosExc) <> 0 Then rwPV(_PV.PrecioExc) = rwPV(_PV.Precio) - rwPV(_PV.PrecioO)
                'If rwPV(_PV.KilosO) <> 0 Then rwPV(_PV.PrecioO) = rwPV(_PV.ImporteO) / rwPV(_PV.KilosO)
                If rwPv(_PV.KilosSO) <> 0 Then rwPv(_PV.PrecioSO) = rwPv(_PV.ImporteSO) / rwPv(_PV.KilosSO)

                ClsBdgProvVto.ApplyBusinessRule(_PV.Kilos, rwPv(_PV.Kilos), rwPv)
            End If
        Next

        '//Aplicar desglose de facturas
        If blnAplicarDesglose Then ProcessServer.ExecuteTask(Of DataTable)(AddressOf AplicarDesglose, dtProvVto, services)

        'se actualiza el vto d la vendimia
        Dim oVV As BdgVendimiaVto = New BdgVendimiaVto
        Dim rwVV As DataRow = oVV.GetItemRow(data.IDVendimiaVto)
        rwVV(_Vvto.TipoCalculo) = data.TipoCalculo
        rwVV(_Vvto.PrecioKiloT) = data.PrecioKiloT
        rwVV(_Vvto.PrecioKiloB) = data.PrecioKiloB
        If Len(data.IDTarifaT) Then
            rwVV(_Vvto.IDTarifaT) = data.IDTarifaT
        Else
            rwVV(_Vvto.IDTarifaT) = DBNull.Value
        End If
        If Len(data.IDTarifaB) Then
            rwVV(_Vvto.IDTarifaB) = data.IDTarifaB
        Else
            rwVV(_Vvto.IDTarifaB) = DBNull.Value
        End If
        rwVV(_Vvto.Facturado) = True
        BusinessHelper.UpdateTable(rwVV.Table)
        BusinessHelper.UpdateTable(dtEntradaVto)
        BusinessHelper.UpdateTable(dtProvVto)
    End Sub

    <Task()> Public Shared Sub CalcularVencimientosFacturacion(ByVal data As StCalcularVencimientos, ByVal services As ServiceProvider)

        Dim ClsBdgProvVto As New BdgProveedorVto
        Dim ClsBdgEntrVto As New BdgEntradaVto
        Dim dtProvVto As DataTable = ClsBdgProvVto.AddNew
        Dim dtEntradaVto As DataTable = ClsBdgEntrVto.AddNew

        Dim fVendimia As New Filter
        fVendimia.Add(New NumberFilterItem(_E.Vendimia, FilterOperator.Equal, data.Vendimia))
        If Not data.Filter Is Nothing Then fVendimia.Add(data.Filter)

        ProcessServer.ExecuteTask(Of Guid)(AddressOf DeleteVto, data.IDVendimiaVto, services)

        Dim dtEntradaVariable As DataTable = New BE.DataEngine().Filter("NegBdgFacturacionVariable", fVendimia, , _EA.IDEntrada & ", " & _EA.IDVariable)
        dtEntradaVariable.DefaultView.Sort = _EA.IDEntrada & ", " & _EA.IDVariable

        Dim strVarGrado As String = New BdgParametro().VariableGrado

        Dim fProvAgrup As New Filter
        fProvAgrup.Add(fVendimia)
        fProvAgrup.Add(New GuidFilterItem(_EVto.IDVendimiaVto, data.IDVendimiaVto))
        Dim dtProvAgrup As DataTable = New BE.DataEngine().Filter("NegBdgFacturacionProveedorAgrup", fProvAgrup)
        For Each drvProvAgrup As DataRowView In dtProvAgrup.DefaultView

            If Not drvProvAgrup Is Nothing Then

                Dim dblImporteDeclarado As Double = 0
                Dim dblImporteExcedente As Double = 0
                Dim dblImporteOrigen As Double = 0
                Dim dblImporteSinOrigen As Double = 0

                Dim f As New Filter
                f.Add(fVendimia)
                f.Add(New StringFilterItem("IDProveedor", FilterOperator.Equal, drvProvAgrup("IDProveedor")))
                f.Add(New NumberFilterItem("TipoVariedad", FilterOperator.Equal, drvProvAgrup("TipoVariedad")))
                Dim dtProvFact As DataTable = New BE.DataEngine().Filter("NegBdgFacturacionProveedor", f)
                For Each drvProvFact As DataRowView In dtProvFact.DefaultView

                    If Not drvProvFact Is Nothing Then

                        Dim strIDTarifa As String = String.Empty
                        Dim dblPrecioDeclarado As Double = 0
                        Dim dblPrecioExcedente As Double = 0
                        Dim dblPrecioOrigen As Double = 0
                        Dim dblPrecioSinOrigen As Double = 0
                        Dim strIDVariableTarifa As String = strVarGrado
                        Dim dblVariableValor As Double = 0
                        Dim dblVariableValorTarifa As Double = 0

                        Dim DataGetTarifa As New BdgProveedor.DataGetTarifa(drvProvFact("IDProveedor"), drvProvFact("IDVariedad"), drvProvFact("Vendimia"))
                        Dim clsTarifa As Tarifa = ProcessServer.ExecuteTask(Of BdgProveedor.DataGetTarifa, Business.Bodega.Tarifa)(AddressOf BdgProveedor.GetTarifa, DataGetTarifa, services)
                        If drvProvFact("TipoVariedad") = BdgTipoVariedad.Tinta Then
                            strIDTarifa = clsTarifa.IDTarifaT
                            dblPrecioOrigen = clsTarifa.PrecioOrigenT
                            dblPrecioExcedente = clsTarifa.PrecioExcedenteT
                        Else
                            strIDTarifa = clsTarifa.IDTarifaB
                            dblPrecioOrigen = clsTarifa.PrecioOrigenB
                            dblPrecioExcedente = clsTarifa.PrecioExcedenteB
                        End If

                        If data.TipoCalculo = BdgTipoCalculo.Predeterminado And drvProvFact("TienePrecioFijo") = True Then
                            strIDTarifa = String.Empty
                            dblPrecioDeclarado = drvProvFact("Precio")
                        ElseIf data.TipoCalculo = BdgTipoCalculo.Predeterminado Then
                            If drvProvFact("TipoVariedad") = BdgTipoVariedad.Tinta Then
                                strIDTarifa = clsTarifa.IDTarifaT
                            Else
                                strIDTarifa = clsTarifa.IDTarifaB
                            End If
                        ElseIf data.TipoCalculo = BdgTipoCalculo.TarifaEspecifica Then
                            If drvProvFact("TipoVariedad") = BdgTipoVariedad.Tinta Then
                                strIDTarifa = data.IDTarifaT
                            Else
                                strIDTarifa = data.IDTarifaB
                            End If
                        Else
                            If drvProvFact("TipoVariedad") = BdgTipoVariedad.Tinta Then
                                dblPrecioDeclarado = data.PrecioKiloT
                            Else
                                dblPrecioDeclarado = data.PrecioKiloB
                            End If
                        End If

                        If data.TipoCalculo <> BdgTipoCalculo.PorKilo And drvProvFact("TienePrecioFijo") = False Then
                            'Obtener los datos de la tarifa.
                            Dim DataDatosTarifa As New DataDatosTarifa(drvProvFact("IDEntrada"), strIDTarifa, strVarGrado, drvProvFact("IDProveedor"), dtEntradaVariable, data.TipoCalculo, fVendimia)
                            ProcessServer.ExecuteTask(Of DataDatosTarifa)(AddressOf DatosTarifa, DataDatosTarifa, services)
                            strIDVariableTarifa = DataDatosTarifa.strIDVariableTarifa
                            dblVariableValor = DataDatosTarifa.dblVariableValor
                            dblVariableValorTarifa = DataDatosTarifa.dblVariableValorTarifa
                            dblPrecioDeclarado = DataDatosTarifa.dblPrecioDeclarado
                        End If

                        dblPrecioSinOrigen = dblPrecioDeclarado - dblPrecioOrigen

                        dblImporteDeclarado += dblPrecioDeclarado * drvProvFact("CantidadDeclarado")
                        dblImporteExcedente += dblPrecioExcedente * drvProvFact("CantidadExcedente")
                        dblImporteOrigen += dblPrecioOrigen * drvProvFact("CantidadOrigen")
                        dblImporteSinOrigen += dblPrecioSinOrigen * drvProvFact("CantidadSinOrigen")

                        'Guardar el detalle por Entrada.
                        Dim DataEntradaVto As New DataEntradaVto(dtEntradaVto, drvProvFact("IDEntrada"), data.IDVendimiaVto, _
                                                                 drvProvFact("IDProveedor"), drvProvFact("TipoVariedad"), dblVariableValor, _
                                                                 dblVariableValorTarifa, strIDVariableTarifa, strIDTarifa, _
                                                                 dblPrecioDeclarado, drvProvFact("CantidadDeclarado"), dblPrecioExcedente, _
                                                                 drvProvFact("CantidadExcedente"), dblPrecioOrigen, drvProvFact("CantidadOrigen"), _
                                                                 dblPrecioSinOrigen, drvProvFact("CantidadSinOrigen"))
                        ProcessServer.ExecuteTask(Of DataEntradaVto)(AddressOf BdgEntradaVto.NuevaEntradaVto, DataEntradaVto, services)
                    End If
                Next

                Dim dtrNewRowProvVto As DataRow = dtProvVto.NewRow

                dtrNewRowProvVto(_PV.IDVendimiaVto) = data.IDVendimiaVto
                dtrNewRowProvVto(_PV.IDProveedor) = drvProvAgrup("IDProveedor")
                dtrNewRowProvVto(_PV.TipoVariedad) = drvProvAgrup("TipoVariedad")
                dtrNewRowProvVto(_PV.Fecha) = Date.Today
                dtrNewRowProvVto(_PV.Desglosado) = 0 'TODO - Habría que meter registros en la tabla nueva tbBdgProveedorVtoLinea*

                dtrNewRowProvVto(_PV.Kilos) = drvProvAgrup("CantidadDeclarado")
                If drvProvAgrup("CantidadDeclarado") <> 0 Then dtrNewRowProvVto(_PV.Precio) = dblImporteDeclarado / drvProvAgrup("CantidadDeclarado")
                dtrNewRowProvVto(_PV.Importe) = dblImporteDeclarado

                dtrNewRowProvVto(_PV.KilosExc) = drvProvAgrup("CantidadExcedente")
                If drvProvAgrup("CantidadExcedente") <> 0 Then dtrNewRowProvVto(_PV.PrecioExc) = dblImporteExcedente / drvProvAgrup("CantidadExcedente")
                dtrNewRowProvVto(_PV.ImporteExc) = dblImporteExcedente

                dtrNewRowProvVto(_PV.KilosO) = drvProvAgrup("CantidadOrigen")
                If drvProvAgrup("CantidadOrigen") <> 0 Then dtrNewRowProvVto(_PV.PrecioO) = dblImporteOrigen / drvProvAgrup("CantidadOrigen")
                dtrNewRowProvVto(_PV.ImporteO) = dblImporteOrigen

                dtrNewRowProvVto(_PV.KilosSO) = drvProvAgrup("CantidadSinOrigen")
                If drvProvAgrup("CantidadSinOrigen") <> 0 Then dtrNewRowProvVto(_PV.PrecioSO) = dblImporteSinOrigen / drvProvAgrup("CantidadSinOrigen")
                dtrNewRowProvVto(_PV.ImporteSO) = dblImporteSinOrigen

                dtrNewRowProvVto(_PV.KilosFra) = drvProvAgrup("CantidadFactura")
                dtrNewRowProvVto(_PV.ImporteFra) = dblImporteDeclarado + dblImporteExcedente + dblImporteOrigen + dblImporteSinOrigen
                If dtrNewRowProvVto(_PV.KilosFra) <> 0 Then dtrNewRowProvVto(_PV.PrecioFra) = dtrNewRowProvVto(_PV.ImporteFra) / dtrNewRowProvVto(_PV.KilosFra)

                dtProvVto.Rows.Add(dtrNewRowProvVto)
            End If
        Next

        ClsBdgProvVto.Update(dtProvVto)
        ClsBdgEntrVto.Update(dtEntradaVto)

    End Sub

    <Serializable()> _
Public Class DataDatosTarifa
        Public IDEntrada As Long
        Public strIDTarifa As String
        Public strVarGrado As String
        Public strIDProveedor As String
        Public dtEntradaVariable As DataTable
        Public TipoCalculo As BdgTipoCalculo
        Public fVendimia As Filter

        Public strIDVariableTarifa As String
        Public dblVariableValor As Double
        Public dblVariableValorTarifa As Double
        Public dblPrecioDeclarado As Double


        Public Sub New(ByVal IDEntrada As Long, ByVal strIDTarifa As String, ByVal strVarGrado As String, ByVal strIDProveedor As String, _
                       ByVal dtEntradaVariable As DataTable, ByVal TipoCalculo As BdgTipoCalculo, ByVal fVendimia As Filter)
            Me.IDEntrada = IDEntrada
            Me.strIDTarifa = strIDTarifa
            Me.strVarGrado = strVarGrado
            Me.strIDProveedor = strIDProveedor
            Me.dtEntradaVariable = dtEntradaVariable
            Me.TipoCalculo = TipoCalculo
            Me.fVendimia = fVendimia
        End Sub
    End Class

    <Task()> Public Shared Sub DatosTarifa(ByVal data As DataDatosTarifa, ByVal services As ServiceProvider)
        If Length(data.strIDTarifa) > 0 Then
            Dim oScrpt As ScriptEngine = New ScriptEngine

            Dim fEnt As New Filter
            fEnt.Add(data.fVendimia)
            fEnt.Add(New NumberFilterItem("IDEntrada", FilterOperator.Equal, data.IDEntrada))
            Dim dtE As DataTable = AdminData.GetData("NegBdgFacturacion", fEnt)
            Dim dtEntrada As DataTable = dtE.Clone
            dtEntrada.ImportRow(dtE.Rows(0))

            Dim ClsBdgTar As New BdgTarifa
            Dim drTarifa As DataRow = ClsBdgTar.GetItemRow(data.strIDTarifa)
            Select Case drTarifa("TipoTarifa")
                Case BdgTarifa.enumBdgTipoTarifa.Formula, BdgTarifa.enumBdgTipoTarifa.Grado
                    data.strIDVariableTarifa = data.strVarGrado
                Case BdgTarifa.enumBdgTipoTarifa.PorKilogrado, BdgTarifa.enumBdgTipoTarifa.PorVariables
                    data.strIDVariableTarifa = drTarifa("IDVariable")
            End Select

            Dim StGetRow As New StGetRowFromView
            StGetRow.vw = data.dtEntradaVariable.DefaultView
            ReDim StGetRow.Keys(1)
            StGetRow.Keys(0) = data.IDEntrada
            StGetRow.Keys(1) = data.strIDVariableTarifa
            Dim rwGrd As DataRow = ProcessServer.ExecuteTask(Of StGetRowFromView, DataRow)(AddressOf GetRowFromView, StGetRow, services)
            If rwGrd Is Nothing Then
                ApplicationService.GenerateError("No se encontró el valor de la variable (|) en la Entrada |.", data.strIDVariableTarifa, dtE.Rows(0)(_E.NEntrada))
            ElseIf rwGrd.IsNull(_EA.ValorNumerico) Then
                ApplicationService.GenerateError("La variable (|) no tiene valor en la Entrada |.", data.strIDVariableTarifa, dtE.Rows(0)(_E.NEntrada))
            Else
                data.dblVariableValor = rwGrd(_EA.ValorNumerico)
            End If

            Dim udtPrec As BdgTarifa.udtPrecioTarifa
            Dim StPrecioEntrada As New BdgTarifa.StPrecioEntrada(data.strIDTarifa, data.dblVariableValor, data.IDEntrada, dtEntrada, data.dtEntradaVariable.DefaultView, oScrpt)
            udtPrec = ProcessServer.ExecuteTask(Of BdgTarifa.StPrecioEntrada, BdgTarifa.udtPrecioTarifa)(AddressOf BdgTarifa.PrecioEntrada, StPrecioEntrada, services)
            data.dblPrecioDeclarado = udtPrec.Precio
            data.dblVariableValorTarifa = udtPrec.GradoTarifa
        Else
            If data.TipoCalculo = BdgTipoCalculo.Predeterminado Then
                ApplicationService.GenerateError("No existe Tarifa para el Proveedor |.", data.strIDProveedor)
            ElseIf data.TipoCalculo = BdgTipoCalculo.TarifaEspecifica Then
                ApplicationService.GenerateError("No existe Tarifa Específica.")
            End If
        End If


        'Dim oBdgTrf As BdgTarifa = New BdgTarifa
        'Dim oScrpt As ScriptEngine = New ScriptEngine

        'Dim drEntrada As DataRow = data.dtEntrada.NewRow
        'For Each dc As DataColumn In data.dtEntrada.Columns
        '    If data.bdEntrada.ContainsKey(dc.ColumnName) Then
        '        drEntrada(dc.ColumnName) = data.bdEntrada(dc.ColumnName)
        '    End If
        'Next

        'Dim ClsBdgTar As New BdgTarifa
        'Dim drTarifa As DataRow = ClsBdgTar.GetItemRow(data.strIDTarifa)
        'Select Case drTarifa("TipoTarifa")
        '    Case BdgTarifa.enumBdgTipoTarifa.Formula, BdgTarifa.enumBdgTipoTarifa.Grado
        '        data.strIDVariableTarifa = data.strVarGrado
        '    Case BdgTarifa.enumBdgTipoTarifa.PorKilogrado, BdgTarifa.enumBdgTipoTarifa.PorVariables
        '        data.strIDVariableTarifa = drTarifa("IDVariable")
        'End Select

        'Dim StGetRow As New StGetRowFromView
        'StGetRow.vw = data.dtEntradaVariable.DefaultView
        'ReDim StGetRow.Keys(1)
        'StGetRow.Keys(0) = data.IDEntrada
        'StGetRow.Keys(1) = data.strIDVariableTarifa
        'Dim rwGrd As DataRow = ProcessServer.ExecuteTask(Of StGetRowFromView, DataRow)(AddressOf GetRowFromView, StGetRow, services)
        'If rwGrd Is Nothing Then
        '    ApplicationService.GenerateError("No se encontró el valor de la variable Grado (|) en la entrada |.", data.strVarGrado, drEntrada(_E.NEntrada))
        '    'dblGrado = 0
        'ElseIf rwGrd.IsNull(_EA.ValorNumerico) Then
        '    ApplicationService.GenerateError("La variable Grado (|) no tiene valor en la entrada |.", data.strVarGrado, drEntrada(_E.NEntrada))
        'Else
        '    data.dblVariableValor = rwGrd(_EA.ValorNumerico)
        'End If

        'If Length(data.strIDTarifa) > 0 Then
        '    Dim udtPrec As BdgTarifa.udtPrecioTarifa

        '    Dim dtEntradaClone As DataTable = data.dtEntrada.Clone
        '    dtEntradaClone.ImportRow(drEntrada)
        '    Dim StPrecioEntrada As New BdgTarifa.StPrecioEntrada(data.strIDTarifa, data.dblVariableValor, data.IDEntrada, dtEntradaClone, data.dtEntradaVariable.DefaultView, oScrpt)
        '    udtPrec = ProcessServer.ExecuteTask(Of BdgTarifa.StPrecioEntrada, BdgTarifa.udtPrecioTarifa)(AddressOf BdgTarifa.PrecioEntrada, StPrecioEntrada, services)
        '    data.dblPrecioDeclarado = udtPrec.Precio
        '    data.dblVariableValorTarifa = udtPrec.GradoTarifa
        'Else
        '    If data.TipoCalculo = BdgTipoCalculo.Predeterminado Then
        '        ApplicationService.GenerateError("No existe Tarifa para el Proveedor |.", drEntrada(_C.IDProveedor))
        '    ElseIf data.TipoCalculo = BdgTipoCalculo.TarifaEspecifica Then
        '        ApplicationService.GenerateError("No existe Tarifa Específica.")
        '    End If
        'End If

    End Sub

    <Serializable()> _
    Public Class StGetRowFromView
        Public vw As DataView
        Public Keys() As Object
    End Class

    <Task()> Public Shared Function GetRowFromView(ByVal data As StGetRowFromView, ByVal services As ServiceProvider) As DataRow
        Dim i As Integer = data.vw.Find(data.Keys)
        If i >= 0 Then
            Return data.vw(i).Row
        End If
    End Function

    <Serializable()> _
    Public Class StNewRow
        Public Dt As DataTable
        Public IDVendimiaVto As Guid
        Public IDProveedor As String
        Public TipoVariedad As BdgTipoVariedad
        Public PrecioExcdt As Double
        Public PrecioO As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal Dt As DataTable, ByVal IDVendimiaVto As Guid, ByVal IDProveedor As String, ByVal TipoVariedad As BdgTipoVariedad, _
                       ByVal PrecioExcdt As Double, ByVal PrecioO As Double)
            Me.Dt = Dt
            Me.IDVendimiaVto = IDVendimiaVto
            Me.IDProveedor = IDProveedor
            Me.TipoVariedad = TipoVariedad
            Me.PrecioExcdt = PrecioExcdt
            Me.PrecioO = PrecioO
        End Sub
    End Class

    <Task()> Public Shared Function NewRow(ByVal data As StNewRow, ByVal services As ServiceProvider) As DataRow
        Dim rwPV As DataRow = data.Dt.NewRow
        data.Dt.Rows.Add(rwPV)
        rwPV(_PV.IDVendimiaVto) = data.IDVendimiaVto
        rwPV(_PV.IDProveedor) = data.IDProveedor
        rwPV(_PV.TipoVariedad) = data.TipoVariedad
        rwPV(_PV.Fecha) = Date.Today
        rwPV(_PV.Kilos) = 0
        rwPV(_PV.Precio) = 0
        rwPV(_PV.Importe) = 0
        rwPV(_PV.KilosExc) = 0
        rwPV(_PV.PrecioExc) = data.PrecioExcdt
        rwPV(_PV.ImporteExc) = 0
        rwPV(_PV.KilosO) = 0
        rwPV(_PV.PrecioO) = data.PrecioO
        rwPV(_PV.ImporteO) = 0
        rwPV(_PV.KilosSO) = 0
        rwPV(_PV.PrecioSO) = 0
        rwPV(_PV.ImporteSO) = 0
        rwPV(_PV.KilosFra) = 0
        rwPV(_PV.PrecioFra) = 0
        rwPV(_PV.ImporteFra) = 0
        rwPV(_PV.Desglosado) = False
        Return rwPV
    End Function

    <Task()> Public Shared Sub AplicarDesglose(ByVal dtPV As DataTable, ByVal services As ServiceProvider)
        Dim oFD As New BdgFraDesglose
        Dim ClsBdgProvVto As New BdgProveedorVto
        Dim dtFD As DataTable = oFD.Filter
        Dim dvFD As DataView = dtFD.DefaultView
        dvFD.Sort = _FD.IDProveedor & ", " & _FD.TipoVariedad

        For Each rwPV As DataRow In dtPV.Select
            Dim dblResto As Double = 100
            Dim blnDesglosado As Boolean = False
            Dim dblKgsFra As Double = rwPV(_PV.KilosFra)

            For Each rwFD As DataRowView In dvFD.FindRows(New Object() {rwPV(_PV.IDProveedor), rwPV(_PV.TipoVariedad)})
                Dim dblPrctj As Double = rwFD(_FD.Porcentaje)
                dblResto -= dblPrctj
                blnDesglosado = True

                Dim StNew As New StNewRow(dtPV, rwPV(_PV.IDVendimiaVto), rwFD(_FD.IDProveedorFra), rwPV(_PV.TipoVariedad), rwPV(_PV.PrecioExc), rwPV(_PV.PrecioO))

                Dim nwRwPV As DataRow = ProcessServer.ExecuteTask(Of StNewRow, DataRow)(AddressOf NewRow, StNew, services)
                nwRwPV(_PV.IDProveedorGrupo) = rwPV(_PV.IDProveedor)
                nwRwPV(_PV.KilosFra) = dblKgsFra * (dblPrctj / 100)
                ClsBdgProvVto.ApplyBusinessRule(_PV.PrecioFra, rwPV(_PV.PrecioFra), nwRwPV)
            Next

            If blnDesglosado Then
                rwPV(_PV.Desglosado) = blnDesglosado
                ClsBdgProvVto.ApplyBusinessRule(_PV.KilosFra, dblKgsFra * (dblResto / 100), rwPV)
            End If
        Next
    End Sub

    <Task()> Public Shared Sub ActualizarDesglosado(ByVal DrPV As DataRow, ByVal services As ServiceProvider)
        Dim ClsProvVto As New BdgProveedorVto
        Dim oFltr As New Filter
        oFltr.Add(New GuidFilterItem(_PV.IDVendimiaVto, DrPV(_PV.IDVendimiaVto)))
        oFltr.Add(New NumberFilterItem(_PV.TipoVariedad, DrPV(_PV.TipoVariedad)))
        oFltr.Add(New StringFilterItem(_PV.IDProveedorGrupo, DrPV(_PV.IDProveedor)))
        Dim dt As DataTable = ClsProvVto.Filter(oFltr)
        Dim dblQDsgls As Double
        For Each oRw As DataRow In dt.Select
            dblQDsgls += oRw(_PV.KilosFra)
        Next

        '//Nos aseguramos un valor correcto para KilosFra
        ClsProvVto.ApplyBusinessRule(_PV.Desglosado, False, DrPV)

        If dblQDsgls > 0 Then
            DrPV(_PV.KilosFra) -= dblQDsgls
            ClsProvVto.ApplyBusinessRule(_PV.Desglosado, True, DrPV)
        End If
    End Sub

    <Serializable()> _
    Public Class StCrearFacturas
        Public dtProvVto As DataTable
        Public IDContador As String
        Public TipoOrigenCalculo As enumBdgTipoCalculoVtos
        Public blActualizarCostekgEntradaUva As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal dtProvVto As DataTable, ByVal TipoOrigenCalculo As enumBdgTipoCalculoVtos, ByVal IDContador As String, ByVal blActualizarCostekgEntradaUva As Boolean)
            Me.dtProvVto = dtProvVto
            Me.TipoOrigenCalculo = TipoOrigenCalculo
            Me.IDContador = IDContador
            Me.blActualizarCostekgEntradaUva = blActualizarCostekgEntradaUva
        End Sub
    End Class

    <Task()> Public Shared Sub CrearFacturasEstandar(ByVal data As StCrearFacturas, ByVal services As ServiceProvider)
        If data.dtProvVto Is Nothing Then Exit Sub
        data.dtProvVto.DefaultView.Sort = _PV.IDProveedor & ", " & _PV.TipoVariedad

        Dim oFCC As New FacturaCompraCabecera
        Dim dtFCC As DataTable
        Dim dtFCL As DataTable = New FacturaCompraLinea().AddNew
        Dim dtPV As DataTable
        Dim strIDProv As String = String.Empty
        Dim VtoVendimia As Guid
        Dim TipoVto As BdgTipoVto
        Dim dtAnticipos As DataTable

        For Each rvwPV As DataRowView In data.dtProvVto.DefaultView
            Dim rwFCC As DataRow
            Dim rwPV As DataRow
            Dim strIDArtT As String
            Dim strIDArtB As String
            Dim strIDArt As String

            If strIDProv <> rvwPV(_PV.IDProveedor) Then
                strIDProv = rvwPV(_PV.IDProveedor)

                '//Crear la cabecera de factura
                Dim dtFCCAux As DataTable = oFCC.AddNewForm
                If dtFCC Is Nothing Then
                    dtFCC = dtFCCAux
                Else
                    dtFCC.ImportRow(dtFCCAux.Rows(0))
                End If
                rwFCC = dtFCC.Rows(dtFCC.Rows.Count - 1)

                If Len(data.IDContador) <> 0 Then
                    rwFCC("IDContador") = data.IDContador
                    rwFCC("NFactura") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data.IDContador, services)
                    rwFCC("SuFactura") = rwFCC("NFactura")
                End If
                oFCC.ApplyBusinessRule("IDProveedor", strIDProv, rwFCC)

            End If

            '//Obtener los artículos para las lineas de fra
            If Not VtoVendimia.Equals(rvwPV(_PV.IDVendimiaVto)) Then
                VtoVendimia = rvwPV(_PV.IDVendimiaVto)
                Dim oVV As BdgVendimiaVto = New BdgVendimiaVto
                Dim rwVV As DataRow = oVV.GetItemRow(VtoVendimia)
                Dim oVDM As BdgVendimia = New BdgVendimia
                Dim rwVDM As DataRow = oVDM.GetItemRow(rwVV(_Vvto.Vendimia))

                If Not rwVDM.IsNull(_VDM.IDArticuloT) Then strIDArtT = rwVDM(_VDM.IDArticuloT)
                If Not rwVDM.IsNull(_VDM.IDArticuloB) Then strIDArtB = rwVDM(_VDM.IDArticuloB)

                TipoVto = CType(rwVV(_Vvto.TipoVto), BdgTipoVto)
                '//Si se trata de una liquidación se buscan los anticipos de la vendimia
                If TipoVto = BdgTipoVto.Liquidacion Then
                    Dim f As Filter = New Filter
                    f.Add(New NumberFilterItem(_Vvto.Vendimia, rwVV(_Vvto.Vendimia)))
                    f.Add(New NumberFilterItem(_Vvto.TipoVto, BdgTipoVto.Anticipo))
                    dtAnticipos = oVV.Filter(f)
                End If
            End If

            '//Crear la linea de fra
            Dim TipoVariedad As BdgTipoVariedad = CType(rvwPV(_PV.TipoVariedad), BdgTipoVariedad)
            If TipoVariedad = BdgTipoVariedad.Tinta Then
                strIDArt = strIDArtT
            Else
                strIDArt = strIDArtB
            End If
            Select Case TipoVto
                Case BdgTipoVto.Anticipo
                    Dim DtLin As DataTable = rwFCC.Table.Clone
                    DtLin.ImportRow(rwFCC)
                    Dim StCrearLinea1 As New StCrearLineaFra(strIDArt, DtLin, rvwPV(_PV.Kilos), rvwPV(_PV.Precio))
                    dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLinea1, services))
                    Dim StCrearLinea2 As New StCrearLineaFra(strIDArt, DtLin, rvwPV(_PV.KilosExc), rvwPV(_PV.PrecioExc))
                    dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLinea2, services))
                Case BdgTipoVto.Liquidacion
                    Dim DtLin As DataTable = rwFCC.Table.Clone
                    DtLin.ImportRow(rwFCC)
                    Dim StCrearLineaLiq As New StCrearLineaFra(strIDArt, DtLin, rvwPV(_PV.KilosFra), rvwPV(_PV.PrecioFra))
                    dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLineaLiq, services))

                    '//Se agregan las lineas de anticipo con importe negativo
                    For Each rwAnt As DataRow In dtAnticipos.Rows
                        Dim dtProvAnt As DataTable = New BdgProveedorVto().SelOnPrimaryKey(rwAnt(_Vvto.IDVendimiaVto), strIDProv, TipoVariedad)
                        If dtProvAnt.Rows.Count > 0 Then
                            Dim rwProvAnt As DataRow = dtProvAnt.Rows(0)
                            If Not rwProvAnt.IsNull(_PV.IDFactura) Then
                                Dim DtLin1 As DataTable = rwFCC.Table.Clone
                                DtLin1.ImportRow(rwFCC)
                                Dim StCrearLinea1 As New StCrearLineaFra(strIDArt, DtLin1, -rwProvAnt(_PV.Kilos), rwProvAnt(_PV.Precio))
                                dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLinea1, services))
                                Dim StCrearLinea2 As New StCrearLineaFra(strIDArt, DtLin1, -rwProvAnt(_PV.KilosExc), rwProvAnt(_PV.PrecioExc))
                                dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLinea2, services))
                            End If
                        End If
                    Next
                Case BdgTipoVto.Bonificacion
                    Dim DtLin As DataTable = rwFCC.Table.Clone
                    DtLin.ImportRow(rwFCC)
                    Dim StCrearLineaBon As New StCrearLineaFra(strIDArt, DtLin, rvwPV(_PV.KilosFra), rvwPV(_PV.PrecioFra))
                    dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLineaBon, services))
            End Select

            '//Actualizar la relación Vencimiento/Factura
            Dim rwAux As DataRow = New BdgProveedorVto().GetItemRow(rvwPV(_PV.IDVendimiaVto), rvwPV(_PV.IDProveedor), rvwPV(_PV.TipoVariedad))
            If dtPV Is Nothing Then
                rwPV = rwAux
                dtPV = rwPV.Table
            Else
                dtPV.ImportRow(rwAux)
                rwPV = dtPV.Rows(dtPV.Rows.Count - 1)
            End If

            rwPV(_PV.IDFactura) = rwFCC("IDFactura")
        Next

        If Not dtFCC Is Nothing Then
            AdminData.BeginTx()
            Dim DtCab As DataTable = dtFCC.Clone
            Dim DtLin As DataTable = dtFCL.Clone
            For Each Dr As DataRow In dtFCC.Select
                DtCab.ImportRow(Dr)
                For Each DrLin As DataRow In dtFCL.Select("IDFactura = " & Dr("IDFactura"))
                    DtLin.ImportRow(DrLin)
                Next
                Dim up As New UpdatePackage
                up.Add(DtCab)
                up.Add(DtLin)
                oFCC.Update(up)
                DtCab.Rows.Clear()
                DtLin.Rows.Clear()
            Next
            Dim ClsBdgProvVto As New BdgProveedorVto
            ClsBdgProvVto.Update(dtPV)
            AdminData.CommitTx()
        End If
    End Sub

    <Task()> Public Shared Sub CrearFacturas(ByVal data As StCrearFacturas, ByVal services As ServiceProvider)
        If data.dtProvVto Is Nothing Then Exit Sub
        data.dtProvVto.DefaultView.Sort = _PV.IDProveedor & ", " & _PV.TipoVariedad

        Dim oFCC As New FacturaCompraCabecera
        Dim dtFCC As DataTable
        Dim dtFCL As DataTable = New FacturaCompraLinea().AddNew
        Dim dtPV As DataTable
        Dim dtEV As DataTable

        Dim strIDProv As String = String.Empty
        Dim VtoVendimia As Guid
        Dim TipoVto As BdgTipoVto
        Dim dtAnticipos As DataTable

        Dim BEDataEngine As New BE.DataEngine

        For Each drvProvVto As DataRowView In data.dtProvVto.DefaultView
            Dim rwFCC As DataRow
            Dim rowProvVto As DataRow
            Dim strIDArtT As String
            Dim strIDArtB As String
            Dim strIDArt As String

            If strIDProv <> drvProvVto(_PV.IDProveedor) Then
                strIDProv = drvProvVto(_PV.IDProveedor)

                '//Crear la cabecera de factura
                Dim dtFCCAux As DataTable = oFCC.AddNewForm
                If dtFCC Is Nothing Then
                    dtFCC = dtFCCAux
                Else
                    dtFCC.ImportRow(dtFCCAux.Rows(0))
                End If
                rwFCC = dtFCC.Rows(dtFCC.Rows.Count - 1)

                If Len(data.IDContador) <> 0 Then
                    rwFCC("IDContador") = data.IDContador
                    Dim cont As Contador.CounterTx = ProcessServer.ExecuteTask(Of String, Contador.CounterTx)(AddressOf Contador.CounterValueTx, rwFCC("IDContador"), services)
                    If Not cont Is Nothing Then
                        rwFCC("NFactura") = cont.strCounterValue
                        rwFCC("SuFactura") = rwFCC("NFactura")
                    End If
                End If
                oFCC.ApplyBusinessRule("IDProveedor", strIDProv, rwFCC)

            End If

            '//Obtener los artículos para las lineas de fra
            If Not VtoVendimia.Equals(drvProvVto(_PV.IDVendimiaVto)) Then
                VtoVendimia = drvProvVto(_PV.IDVendimiaVto)
                Dim oVV As BdgVendimiaVto = New BdgVendimiaVto
                Dim rwVV As DataRow = oVV.GetItemRow(VtoVendimia)
                Dim oVDM As BdgVendimia = New BdgVendimia
                Dim rwVDM As DataRow = oVDM.GetItemRow(rwVV(_Vvto.Vendimia))

                If Not rwVDM.IsNull(_VDM.IDArticuloT) Then strIDArtT = rwVDM(_VDM.IDArticuloT)
                If Not rwVDM.IsNull(_VDM.IDArticuloB) Then strIDArtB = rwVDM(_VDM.IDArticuloB)

                TipoVto = CType(rwVV(_Vvto.TipoVto), BdgTipoVto)
                '//Si se trata de una liquidación se buscan los anticipos de la vendimia
                If TipoVto = BdgTipoVto.Liquidacion Then
                    Dim f As Filter = New Filter
                    f.Add(New NumberFilterItem(_Vvto.Vendimia, rwVV(_Vvto.Vendimia)))
                    f.Add(New NumberFilterItem(_Vvto.TipoVto, BdgTipoVto.Anticipo))
                    dtAnticipos = oVV.Filter(f)
                End If
            End If

            '//Crear la linea de fra
            Dim TipoVariedad As BdgTipoVariedad = CType(drvProvVto(_PV.TipoVariedad), BdgTipoVariedad)
            If TipoVariedad = BdgTipoVariedad.Tinta Then
                strIDArt = strIDArtT
            Else
                strIDArt = strIDArtB
            End If
            Select Case TipoVto
                Case BdgTipoVto.Anticipo
                    Dim DtLin As DataTable = rwFCC.Table.Clone
                    DtLin.ImportRow(rwFCC)
                    Dim StCrearLinea1 As New StCrearLineaFra(strIDArt, DtLin, drvProvVto(_PV.Kilos), drvProvVto(_PV.Precio))
                    dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLinea1, services))
                    Dim StCrearLinea2 As New StCrearLineaFra(strIDArt, DtLin, drvProvVto(_PV.KilosExc), drvProvVto(_PV.PrecioExc))
                    dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLinea2, services))
                Case BdgTipoVto.Liquidacion
                    If data.TipoOrigenCalculo = enumBdgTipoCalculoVtos.PorFacturacion Then
                        Dim f As New Filter
                        f.Add(New GuidFilterItem("IDVendimiaVto", FilterOperator.Equal, drvProvVto("IDVendimiaVto")))
                        f.Add(New StringFilterItem("IDProveedor", FilterOperator.Equal, drvProvVto("IDProveedor")))
                        f.Add(New NumberFilterItem("TipoVariedad", FilterOperator.Equal, drvProvVto("TipoVariedad")))
                        Dim dtEntradaVto As DataTable = BEDataEngine.Filter("tbBdgEntradaVto", f)
                        For Each drvEntradaVto As DataRowView In dtEntradaVto.DefaultView
                            If Not drvEntradaVto Is Nothing Then
                                Dim rowEntradaVto As DataRow

                                Dim DtLin As DataTable = rwFCC.Table.Clone
                                DtLin.ImportRow(rwFCC)
                                Dim StCrearLineaLiq1 As New StCrearLineaFra(strIDArt, DtLin, drvEntradaVto(_EVto.Kilos), drvEntradaVto(_EVto.Precio))
                                Dim drLin1 As DataRow = ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLineaLiq1, services)
                                dtFCL.ImportRow(drLin1)
                                Dim StCrearLineaLiq2 As New StCrearLineaFra(strIDArt, DtLin, drvEntradaVto(_EVto.KilosExc), drvEntradaVto(_EVto.PrecioExc))
                                Dim drLin2 As DataRow = ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLineaLiq2, services)
                                dtFCL.ImportRow(drLin2)
                                Dim StCrearLineaLiq3 As New StCrearLineaFra(strIDArt, DtLin, drvEntradaVto(_EVto.KilosO), drvEntradaVto(_EVto.PrecioO))
                                Dim drLin3 As DataRow = ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLineaLiq3, services)
                                dtFCL.ImportRow(drLin3)
                                Dim StCrearLineaLiq4 As New StCrearLineaFra(strIDArt, DtLin, drvEntradaVto(_EVto.KilosSO), drvEntradaVto(_EVto.PrecioSO))
                                Dim drLin4 As DataRow = ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLineaLiq4, services)
                                dtFCL.ImportRow(drLin4)

                                '//Actualizar la relación EntradaVto/Línea Factura
                                Dim rwAuxLin As DataRow = New BdgEntradaVto().GetItemRow(drvEntradaVto(_EVto.IDEntradaVto))
                                If dtEV Is Nothing Then
                                    rowEntradaVto = rwAuxLin
                                    dtEV = rowEntradaVto.Table
                                Else
                                    dtEV.ImportRow(rwAuxLin)
                                    rowEntradaVto = dtEV.Rows(dtEV.Rows.Count - 1)
                                End If

                                If Not drLin1 Is Nothing Then rowEntradaVto(_EVto.IDLineaFactura) = drLin1("IDLineaFactura")
                                If Not drLin2 Is Nothing Then rowEntradaVto(_EVto.IDLineaFacturaExc) = drLin2("IDLineaFactura")
                                If Not drLin3 Is Nothing Then rowEntradaVto(_EVto.IDLineaFacturaO) = drLin3("IDLineaFactura")
                                If Not drLin4 Is Nothing Then rowEntradaVto(_EVto.IDLineaFacturaSO) = drLin4("IDLineaFactura")

                                '//Actualizar el Coste Unitario en la Entrada de Uva
                                If data.blActualizarCostekgEntradaUva Then
                                    Dim datCalcCoste As New DataCalcularCostekgEntradaUva(drvProvVto("IDVendimiaVto"), drvProvVto("TipoVariedad"), drvEntradaVto(_EVto.IDEntrada))
                                    ProcessServer.ExecuteTask(Of DataCalcularCostekgEntradaUva)(AddressOf CalcularCostekgEntradaUva, datCalcCoste, services)
                                End If
                            End If
                        Next
                    Else
                        Dim DtLin As DataTable = rwFCC.Table.Clone
                        DtLin.ImportRow(rwFCC)
                        Dim StCrearLineaLiq As New StCrearLineaFra(strIDArt, DtLin, drvProvVto(_PV.KilosFra), drvProvVto(_PV.PrecioFra))
                        dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLineaLiq, services))
                    End If

                    '//Se agregan las lineas de anticipo con importe negativo
                    For Each rwAnt As DataRow In dtAnticipos.Rows
                        Dim dtProvAnt As DataTable = New BdgProveedorVto().SelOnPrimaryKey(rwAnt(_Vvto.IDVendimiaVto), strIDProv, TipoVariedad)
                        If dtProvAnt.Rows.Count > 0 Then
                            Dim rwProvAnt As DataRow = dtProvAnt.Rows(0)
                            If Not rwProvAnt.IsNull(_PV.IDFactura) Then
                                Dim DtLin1 As DataTable = rwFCC.Table.Clone
                                DtLin1.ImportRow(rwFCC)
                                Dim StCrearLinea1 As New StCrearLineaFra(strIDArt, DtLin1, -rwProvAnt(_PV.Kilos), rwProvAnt(_PV.Precio))
                                dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLinea1, services))
                                Dim StCrearLinea2 As New StCrearLineaFra(strIDArt, DtLin1, -rwProvAnt(_PV.KilosExc), rwProvAnt(_PV.PrecioExc))
                                dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLinea2, services))
                            End If
                        End If
                    Next
                Case BdgTipoVto.Bonificacion
                    Dim DtLin As DataTable = rwFCC.Table.Clone
                    DtLin.ImportRow(rwFCC)
                    Dim StCrearLineaBon As New StCrearLineaFra(strIDArt, DtLin, drvProvVto(_PV.KilosFra), drvProvVto(_PV.PrecioFra))
                    dtFCL.ImportRow(ProcessServer.ExecuteTask(Of StCrearLineaFra, DataRow)(AddressOf CrearLineaFra, StCrearLineaBon, services))
            End Select

            '//Actualizar la relación Vencimiento/Factura
            Dim rwAux As DataRow = New BdgProveedorVto().GetItemRow(drvProvVto(_PV.IDVendimiaVto), drvProvVto(_PV.IDProveedor), drvProvVto(_PV.TipoVariedad))
            If dtPV Is Nothing Then
                rowProvVto = rwAux
                dtPV = rowProvVto.Table
            Else
                dtPV.ImportRow(rwAux)
                rowProvVto = dtPV.Rows(dtPV.Rows.Count - 1)
            End If

            rowProvVto(_PV.IDFactura) = rwFCC("IDFactura")
        Next

        Dim dtEntradaUva As DataTable
        If data.blActualizarCostekgEntradaUva Then
            Dim Costes As CostesEntradas = services.GetService(Of CostesEntradas)()
            If Costes.Lista.Keys.Count > 0 Then
                Dim EV As New BdgEntrada
                For Each IDEntrada As Integer In Costes.Lista.Keys
                    Dim dtEntradaUvaAux As DataTable = EV.SelOnPrimaryKey(IDEntrada)
                    If dtEntradaUvaAux.Rows.Count > 0 Then
                        Dim CosteEntrada As Double = 0
                        CosteEntrada = Costes.Lista(IDEntrada)
                        dtEntradaUvaAux.Rows(0)("CosteKgUva") = CosteEntrada
                        If dtEntradaUva Is Nothing Then
                            dtEntradaUva = dtEntradaUvaAux.Copy
                        Else
                            dtEntradaUva.ImportRow(dtEntradaUvaAux.Rows(0))
                        End If
                    End If

                Next
            End If
        End If

        If Not dtFCC Is Nothing Then
            AdminData.BeginTx()
            Dim DtCab As DataTable = dtFCC.Clone
            Dim DtLin As DataTable = dtFCL.Clone
            For Each Dr As DataRow In dtFCC.Select
                DtCab.ImportRow(Dr)
                For Each DrLin As DataRow In dtFCL.Select("IDFactura = " & Dr("IDFactura"))
                    DtLin.ImportRow(DrLin)
                Next
                Dim up As New UpdatePackage
                up.Add(DtCab)
                up.Add(DtLin)
                oFCC.Update(up)
                DtCab.Rows.Clear()
                DtLin.Rows.Clear()
            Next
            Dim ClsBdgProvVto As New BdgProveedorVto
            ClsBdgProvVto.Update(dtPV)
            Dim ClsBdgEntradaVto As New BdgEntradaVto
            ClsBdgEntradaVto.Update(dtEV)
            Dim ClsBdgEntradaUva As New BdgEntrada
            ClsBdgEntradaUva.Update(dtEntradaUva)
            AdminData.CommitTx()
        End If
    End Sub

    <Serializable()> _
    Public Class DataCalcularCostekgEntradaUva
        Public IDVendimiaVto As Guid
        Public TipoVariedad As Integer
        Public IDEntrada As Integer

        Public Sub New(ByVal IDVendimiaVto As Guid, ByVal TipoVariedad As Integer, ByVal IDEntrada As Integer)
            Me.IDVendimiaVto = IDVendimiaVto
            Me.TipoVariedad = TipoVariedad
            Me.IDEntrada = IDEntrada
        End Sub
    End Class

    <Serializable()> _
    Public Class CostesEntradas
        Public mLista As New Dictionary(Of Integer, Double)
        Public Property Lista() As Dictionary(Of Integer, Double)
            Get
                Return mLista
            End Get
            Set(ByVal value As Dictionary(Of Integer, Double))
                mLista = value
            End Set
        End Property
    End Class

    <Task()> Public Shared Sub CalcularCostekgEntradaUva(ByVal data As DataCalcularCostekgEntradaUva, ByVal services As ServiceProvider)
        Dim Costes As CostesEntradas = services.GetService(Of CostesEntradas)()

        If Not Costes.Lista.ContainsKey(data.IDEntrada) Then
            Dim f As New Filter
            f.Add(New GuidFilterItem("IDVendimiaVto", data.IDVendimiaVto))
            f.Add(New NumberFilterItem("TipoVariedad", data.TipoVariedad))
            f.Add(New NumberFilterItem("IDEntrada", data.IDEntrada))

            Dim CostesEntradas As New Dictionary(Of Integer, Double)
            Dim dtEntradaVto As DataTable = New BE.DataEngine().Filter("tbBdgEntradaVto", f)
            If dtEntradaVto.Rows.Count > 0 Then

                Dim ProveedoresVariedad As List(Of DataRow) = (From c In dtEntradaVto Select c).ToList

                Dim PrecioxKilo As Double = 0
                For Each drProvVar As DataRow In ProveedoresVariedad
                    PrecioxKilo += (Nz(drProvVar("Precio"), 0) * Nz(drProvVar("Kilos"), 0))
                Next
                Dim TotalKilos As Double = (Aggregate c In dtEntradaVto Into Sum(CDbl(c("Kilos"))))
                Dim Coste As Double = 0
                If TotalKilos <> 0 Then
                    Coste = PrecioxKilo / TotalKilos
                End If

                Costes.Lista.Add(data.IDEntrada, Coste)
            End If
        End If

    End Sub

    <Serializable()> _
    Public Class StCrearLineaFra
        Public IDArt As String
        Public DtFCC As DataTable
        Public Q As Double
        Public P As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDArt As String, ByVal DtFCC As DataTable, ByVal Q As Double, ByVal P As Double)
            Me.IDArt = IDArt
            Me.DtFCC = DtFCC
            Me.Q = Q
            Me.P = P
        End Sub
    End Class

    <Task()> Public Shared Function CrearLineaFra(ByVal data As StCrearLineaFra, ByVal services As ServiceProvider) As DataRow
        If data.Q = 0 Then Exit Function
        Dim oFCL As New FacturaCompraLinea

        Dim rwFCL As DataRow = oFCL.AddNewForm.Rows(0)
        rwFCL("IDLineaFactura") = AdminData.GetAutoNumeric
        rwFCL("IDFactura") = data.DtFCC.Rows(0)("IDFactura")
        rwFCL("IDCentroGestion") = data.DtFCC.Rows(0)("IDCentroGestion")

        Dim ctx As New BusinessData(data.DtFCC.Rows(0))
        rwFCL = oFCL.ApplyBusinessRule("IDArticulo", data.IDArt, rwFCL, ctx)
        rwFCL = oFCL.ApplyBusinessRule("Cantidad", data.Q, rwFCL, ctx)
        rwFCL = oFCL.ApplyBusinessRule("Precio", data.P, rwFCL, ctx)
        Return rwFCL
    End Function

#End Region

End Class

Public Enum BdgTipoCalculo
    Predeterminado
    PorKilo
    TarifaEspecifica
End Enum

Public Enum enumBdgTipoCalculoVtos
    PorCabecera
    PorCartillista
    PorFacturacion
End Enum

<Serializable()> _
Public Class _PV
    Public Const IDVendimiaVto As String = "IDVendimiaVto"
    Public Const IDProveedor As String = "IDProveedor"
    Public Const TipoVariedad As String = "TipoVariedad"
    Public Const Fecha As String = "Fecha"
    Public Const Kilos As String = "Kilos"
    Public Const Precio As String = "Precio"
    Public Const Importe As String = "Importe"
    Public Const KilosSO As String = "KilosSO"
    Public Const PrecioSO As String = "PrecioSO"
    Public Const ImporteSO As String = "ImporteSO"
    Public Const KilosO As String = "KilosO"
    Public Const PrecioO As String = "PrecioO"
    Public Const ImporteO As String = "ImporteO"
    Public Const KilosExc As String = "KilosExc"
    Public Const PrecioExc As String = "PrecioExc"
    Public Const ImporteExc As String = "ImporteExc"
    Public Const KilosFra As String = "KilosFra"
    Public Const PrecioFra As String = "PrecioFra"
    Public Const ImporteFra As String = "ImporteFra"
    Public Const IDFactura As String = "IDFactura"
    Public Const IDProveedorGrupo As String = "IDProveedorGrupo"
    Public Const Desglosado As String = "Desglosado"
End Class