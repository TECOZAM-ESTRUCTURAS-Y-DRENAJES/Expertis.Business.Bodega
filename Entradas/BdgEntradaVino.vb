Public Class BdgEntradaVino
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaVino"


    Public Overloads Function GetItemRow(ByVal NEntrada As Integer) As DataRow
        Dim dt As DataTable = New BdgEntradaVino().SelOnPrimaryKey(NEntrada)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe la entrada de vino |", NEntrada)
        Else : Return dt.Rows(0)
        End If
    End Function

#Region " RegisterAddnewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DataCont As New Contador.DatosDefaultCounterValue(data, "BdgEntradaVino", _EVn.NEntrada)
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, DataCont, services)
        data(_EVn.Fecha) = Date.Today
        data(_EVn.TipoPrecio) = TipoPrecioContrato.PorLitro
        data(_EVn.Cantidad) = 0
        data(_EVn.Grado) = 0
        data(_EVn.Precio) = 0
        data(_EVn.PrecioPorte) = 0
        data(_EVn.Importe) = 0
        data(_EVn.ImportePorte) = 0
    End Sub

#End Region

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarQRecibida)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarMovimientosDelete)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        'deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarContador)
    End Sub

    <Task()> Public Shared Sub ActualizarMovimientosDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim oEVD As New BdgEntradaVinoDeposito
        Dim dtEVD As DataTable = oEVD.Filter(New NumberFilterItem(_EVD.NEntrada, data(_EVn.NEntrada)))
        If dtEVD.Rows.Count > 0 Then
            oEVD.Delete(dtEVD)
            If Length(data(_EVn.IDMovimiento)) = 0 Then
                Dim StEj As New BdgWorkClass.StEjecutarMovimientos(data(_EVn.NEntrada), data(_EVn.Fecha))
                ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientos, Integer)(AddressOf BdgWorkClass.EjecutarMovimientos, StEj, services)
            Else
                Dim StEj As New BdgWorkClass.StEjecutarMovimientosNumero(data(_EVn.IDMovimiento), data(_EVn.NEntrada), data(_EVn.Fecha))
                ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientosNumero)(AddressOf BdgWorkClass.EjecutarMovimientosNumero, StEj, services)
            End If
        End If
    End Sub

    '<Task()> Public Shared Sub ActualizarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    Dim NEntrada As String = data(_EVn.NEntrada) & String.Empty
    '    Dim IDContador As String = data(_EVn.IDContador) & String.Empty
    '    If Length(IDContador) > 0 Then
    '        Dim dataContador As New Contador.DatosDecrementCounter(IDContador, NEntrada)
    '        ProcessServer.ExecuteTask(Of Contador.DatosDecrementCounter)(AddressOf Contador.DecrementCounter, dataContador, services)
    '    End If
    'End Sub

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDContratoLinea", AddressOf CambioContratoLinea)
        Obrl.Add("IDDeposito", AddressOf CambioIDDeposito)
        Obrl.Add("IDArticulo", AddressOf BdgEntradaVinoDeposito.CambioIDArticulo)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioContratoLinea(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If TypeOf data.Current("IDContratoLinea") Is Guid Then
            Dim lineaContrato As Guid = CType(data.Current("IDContratoLinea"), Guid)
            Dim dt As DataTable = AdminData.GetData("advBdgContrato", New GuidFilterItem(_EVn.IDContratoLinea, lineaContrato))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("TipoPrecio") = dt.Rows(0)("TipoPrecio")
                data.Current("Precio") = dt.Rows(0)("Precio")
                data.Current("IDArticulo") = dt.Rows(0)("IDArticulo")
                data.Current("DescArticulo") = dt.Rows(0)("DescArticulo")
                data.Current("IDProveedor") = dt.Rows(0)("IDProveedor")
                data.Current("PrecioPorte") = dt.Rows(0)("PrecioPorte")
                data.Current("NContrato") = dt.Rows(0)("NContrato")
            End If
        Else
            data.Current("NContrato") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDDeposito(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDDeposito")) > 0 Then

        Else
            data.Current("TipoDeposito") = System.DBNull.Value
        End If
    End Sub

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarArticulo)
    End Sub

    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim StrIDAnalisis As String = New BdgParametro().AnalisisEntradaVino()
            If Length(StrIDAnalisis) > 0 Then
                ProcessServer.ExecuteTask(Of String)(AddressOf BdgAnalisis.ValidatePrimaryKey, StrIDAnalisis, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_EVn.IDArticulo)) = 0 Then ApplicationService.GenerateError("No se ha establecido el artículo de la entrada de vino")
    End Sub

    <Task()> Public Shared Sub ValidarProveedor(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_EVn.IDProveedor)) = 0 Then ApplicationService.GenerateError("No se ha establecido el proveedor de la entrada de vino")
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of UpdatePackage, DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.CrearDocumento)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.ComprobarTotales)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.AsignarContador)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.CalcularImporte)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.ActualizarMovimientos)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.ActualizarDescArticulo)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.ActualizarContratoLinea)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.AsignarValorNumericoAnalisis)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.AsignarVinosDepositos)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.CambioFechaEntradaVino)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.ActualizarEstadoVino)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf Comunes.UpdateDocument)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DocumentoEntradaVino)(AddressOf ProcesoEntradaVino.AsignarMovimientosDepositos)
    End Sub

#End Region

#Region " Funciones Públicas "

    <Task()> Public Shared Sub ActualizarQRecibida(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim entrada As DataRow = New BdgEntradaVino().GetItemRow(data("NEntrada"))
        If Length(data(_EVn.IDContratoLinea)) > 0 Then
            Dim bdgContrato As New BdgContratoLinea
            Dim rwContrato As DataRow = bdgContrato.GetItemRow(data(_EVn.IDContratoLinea))
            rwContrato("QRecibida") -= data(_EVn.Cantidad)
            bdgContrato.Update(rwContrato.Table)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarMovimientosEntrada(ByVal DrEntradaVino As DataRow, ByVal services As ServiceProvider)
        If Not DrEntradaVino.IsNull("IDMovimiento") Then
            Dim oMI As MonedaInfo = ProcessServer.ExecuteTask(Of Date, MonedaInfo)(AddressOf Moneda.MonedaB, DrEntradaVino(_EVn.Fecha), services)
            Dim Precio As Double = (DrEntradaVino(_EVn.Importe) + DrEntradaVino(_EVn.ImportePorte)) / DrEntradaVino(_EVn.Cantidad)
            Dim PrecioB As Double = xRound(Precio * oMI.CambioB, oMI.NDecimalesPrecio)

            Dim blnCambioImportes As Boolean = AreDifferents(DrEntradaVino(_EVn.Importe, DataRowVersion.Original), DrEntradaVino(_EVn.Importe)) _
                                       OrElse AreDifferents(DrEntradaVino(_EVn.ImportePorte, DataRowVersion.Original), DrEntradaVino(_EVn.ImportePorte))
            Dim blnCambioFecha As Boolean = Nz(DrEntradaVino("Fecha")) <> Nz(DrEntradaVino("Fecha", DataRowVersion.Original))

            Dim fMovtosEntradaVino As New Filter
            fMovtosEntradaVino.Add(New NumberFilterItem("IDMovimiento", DrEntradaVino("IDMovimiento")))
            fMovtosEntradaVino.Add(New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmEntFabrica))
            Dim dtMovs As DataTable = New BE.DataEngine().Filter("tbHistoricoMovimiento", fMovtosEntradaVino)
            For Each oRw As DataRow In dtMovs.Select
                Dim dataCorreccion As ProcesoStocks.DataActualizarMovimiento
                If blnCambioFecha AndAlso blnCambioImportes Then
                    dataCorreccion = New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, oRw("IDLineaMovimiento"), oRw(_EVn.Cantidad), DrEntradaVino(_EVn.Fecha), Precio, PrecioB)
                ElseIf blnCambioImportes Then
                    dataCorreccion = New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, oRw("IDLineaMovimiento"), Precio, PrecioB)
                ElseIf blnCambioFecha Then
                    dataCorreccion = New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, oRw("IDLineaMovimiento"), CDate(DrEntradaVino(_EVn.Fecha)))
                End If
                If Not dataCorreccion Is Nothing Then ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
            Next
        End If
    End Sub

    '<Task()> Public Shared Sub ActualizarFechaMovimientosEntrada(ByVal DrEntradaVino As DataRow, ByVal services As ServiceProvider)
    '    If Not DrEntradaVino.IsNull("IDMovimiento") Then
    '        Dim stocks As New ProcesoStocks
    '        Dim dtMovs As DataTable = New BE.DataEngine().Filter("tbHistoricoMovimiento", New NumberFilterItem("IDMovimiento", DrEntradaVino("IDMovimiento")))
    '        For Each oRw As DataRow In dtMovs.Select
    '            If oRw("IDTipoMovimiento") = enumTipoMovimiento.tmEntFabrica Then
    '                Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, oRw("IDLineaMovimiento"), CDate(oRw("Fecha")))
    '                ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
    '            End If
    '        Next
    '    End If
    'End Sub

    <Task()> Public Shared Function GetValorGradoDesdeUpdateContext(ByVal VariableGrado As String, ByVal services As ServiceProvider) As Double
        Return Double.NaN
    End Function

    <Serializable()> _
    Public Class StGetValorGrado
        Public NEntrada As Integer
        Public VariableGrado As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal NEntrada As Integer, ByVal VariableGrado As String)
            Me.NEntrada = NEntrada
            Me.VariableGrado = VariableGrado
        End Sub
    End Class

    <Task()> Public Shared Function GetValorGrado(ByVal data As StGetValorGrado, ByVal services As ServiceProvider) As Double
        Dim Valor As Double = ProcessServer.ExecuteTask(Of String, Double)(AddressOf GetValorGradoDesdeUpdateContext, data.VariableGrado, services)
        If Double.IsNaN(Valor) Then
            Dim dtEA As DataTable = New BdgEntradaVinoAnalisis().SelOnPrimaryKey(data.NEntrada, data.VariableGrado)
            If dtEA.Rows.Count > 0 Then
                If Length(dtEA.Rows(0)(_EA.ValorNumerico)) > 0 Then Return dtEA.Rows(0)(_EA.ValorNumerico)
            End If
        End If
        Return Valor
    End Function

#End Region

End Class

<Serializable()> _
Public Class _EVn
    Public Const NEntrada As String = "NEntrada"
    Public Const IDContador As String = "IDContador"
    Public Const IDProveedor As String = "IDProveedor"
    Public Const Fecha As String = "Fecha"
    Public Const Cantidad As String = "Cantidad"
    Public Const Texto As String = "Texto"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const DescArticulo As String = "DescArticulo"
    Public Const IDContratoLinea As String = "IDContratoLinea"
    Public Const Grado As String = "Grado"
    Public Const Facturado As String = "Facturado"
    Public Const Precio As String = "Precio"
    Public Const TipoPrecio As String = "TipoPrecio"
    Public Const Importe As String = "Importe"
    Public Const PrecioPorte As String = "PrecioPorte"
    Public Const ImportePorte As String = "ImportePorte"
    Public Const Matricula As String = "Matricula"
    Public Const NDAA As String = "NDAA"
    Public Const DiasBarrica As String = "DiasBarrica"
    Public Const DiasBotella As String = "DiasBotella"
    Public Const Lote As String = "Lote"
    Public Const IDMovimiento As String = "IDMovimiento"
End Class