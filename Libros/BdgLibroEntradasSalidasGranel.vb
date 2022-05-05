Public Class BdgLibroEntradasSalidasGranel

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgLibroEntradasSalidasGranel"

#End Region

#Region "Clases Publicas"

    <Serializable()> _
    Public Class DatosEntradasSalidas
        Public FechaDesde As Date
        Public FechaHasta As Date
        Public Vendimia As Long

        Public Sub New()

        End Sub

        Public Sub New(ByVal FechaDesde As Date, ByVal FechaHasta As Date)
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
        End Sub

        Public Sub New(ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal Vendimia As Long)
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
            Me.Vendimia = Vendimia
        End Sub
    End Class

#End Region

#Region "Funciones Publicas"
    <Task()> Public Shared Function CrearEntradasUva(ByVal data As DatosEntradasSalidas, ByVal services As ServiceProvider) As String

        Dim clsLibro As New BdgLibroEntradasSalidasGranel
        Dim dtLibro As DataTable = clsLibro.AddNew

        AdminData.BeginTx()

        'Borrar lineas existentes
        Dim f As New Filter
        f.Add("FechaMovimiento", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
        f.Add("FechaMovimiento", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
        f.Add("Declarado", FilterOperator.Equal, False, FilterType.Boolean)
        f.Add("TipoMovimiento", FilterOperator.Equal, enumBdgTipoMovimientoLibro.EntradaUva, FilterType.Numeric)
        Dim dtLibroBorrar As DataTable = New BdgLibroEntradasSalidasGranel().Filter(f)
        If dtLibroBorrar.Rows.Count > 0 Then
            clsLibro.Delete(dtLibroBorrar)
        End If

        'Crear nuevas lineas
        Dim f2 As New Filter
        f2.Add("Fecha", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
        f2.Add("Fecha", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
        Dim dtOrigen As DataTable = New BE.DataEngine().Filter("NegBdgLibroESGranelEntradaUva", f2)
        If dtOrigen.Rows.Count > 0 Then
            For Each oRw As DataRow In dtOrigen.Select
                Dim nwRw As DataRow = dtLibro.NewRow
                dtLibro.Rows.Add(nwRw)
                nwRw("FechaMovimiento") = oRw("Fecha")
                nwRw("Entrada") = enumBdgLibroEntradaSalida.Entrada
                nwRw("TipoMovimiento") = enumBdgTipoMovimientoLibro.EntradaUva
                nwRw("Procedencia") = "Entrada de Uva"
                If oRw("QBlanco") <> 0 Then
                    nwRw("QMostoBlanco") = oRw("QBlanco")
                    nwRw("IDUdMostoBlanco") = oRw("IDUdMedida")
                End If
                If oRw("QTinto") <> 0 Then
                    nwRw("QMostoTinto") = oRw("QTinto")
                    nwRw("IDUdMostoTinto") = oRw("IDUdMedida")
                End If
                nwRw("Texto") = oRw("DescVariedad")
                nwRw("Declarado") = False
            Next
            clsLibro.Update(dtLibro)
        Else
            ApplicationService.GenerateError("No hay Entradas de Uva en el periodo seleccionado.")
        End If

    End Function

    <Task()> Public Shared Function TratarUvaParaMosto(ByVal data As DatosEntradasSalidas, ByVal services As ServiceProvider) As String

        Dim clsLibro As New BdgLibroEntradasSalidasGranel
        Dim dtLibro As DataTable = clsLibro.AddNew

        'Acumular Q
        Dim dblQMostoBlanco As Double = 0
        Dim dblQMostoTinto As Double = 0
        Dim IDUDMedida As String = String.Empty
        Dim f As New Filter
        f.Add("FechaMovimiento", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
        f.Add("FechaMovimiento", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
        f.Add("Declarado", FilterOperator.Equal, True, FilterType.Boolean)
        f.Add("TipoMovimiento", FilterOperator.Equal, enumBdgTipoMovimientoLibro.EntradaUva, FilterType.Numeric)
        Dim dtOrigen As DataTable = New BdgLibroEntradasSalidasGranel().Filter(f)
        If dtOrigen.Rows.Count > 0 Then
            For Each oRw As DataRow In dtOrigen.Select
                dblQMostoBlanco += Nz(oRw("QMostoBlanco"), 0)
                dblQMostoTinto += Nz(oRw("QMostoTinto"), 0)
                IDUDMedida = Nz(oRw("IDUdMostoBlanco"), oRw("IDUdMostoTinto")) & String.Empty
            Next
        End If
        If dblQMostoBlanco + dblQMostoTinto = 0 Then
            ApplicationService.GenerateError("No hay Entradas de Uva Declaradas en el periodo seleccionado.")
        End If

        AdminData.BeginTx()

        'Crear salida uva
        Dim nwRw As DataRow = dtLibro.NewRow
        dtLibro.Rows.Add(nwRw)
        nwRw("FechaMovimiento") = data.FechaHasta
        nwRw("Entrada") = enumBdgLibroEntradaSalida.Salida
        nwRw("TipoMovimiento") = enumBdgTipoMovimientoLibro.SalidaUva
        nwRw("Procedencia") = "Uva para elaboración de Mosto"
        If dblQMostoBlanco <> 0 Then
            nwRw("QMostoBlanco") = dblQMostoBlanco
            nwRw("IDUdMostoBlanco") = IDUDMedida
        End If
        If dblQMostoTinto <> 0 Then
            nwRw("QMostoTinto") = dblQMostoTinto
            nwRw("IDUdMostoTinto") = IDUDMedida
        End If
        nwRw("Declarado") = False

        'Crear entrada mosto
        Dim dblRdtoBlanco As Double = 0
        Dim dblRdtoTinto As Double = 0
        Dim clsBdgVendimia As New BdgVendimia
        Dim drVendimia As DataRow = clsBdgVendimia.GetItemRow(data.Vendimia)
        dblRdtoBlanco = Nz(drVendimia("RdtoB"), 100) / 100
        dblRdtoTinto = Nz(drVendimia("RdtoT"), 100) / 100
        nwRw = dtLibro.NewRow
        dtLibro.Rows.Add(nwRw)
        nwRw("FechaMovimiento") = data.FechaHasta
        nwRw("Entrada") = enumBdgLibroEntradaSalida.Entrada
        nwRw("TipoMovimiento") = enumBdgTipoMovimientoLibro.EntradaMosto
        nwRw("Procedencia") = "Mosto procedente de Uva"
        If dblQMostoBlanco <> 0 Then
            nwRw("QMostoBlanco") = dblQMostoBlanco * dblRdtoBlanco
            nwRw("IDUdMostoBlanco") = IDUDMedida
        End If
        If dblQMostoTinto <> 0 Then
            nwRw("QMostoTinto") = dblQMostoTinto * dblRdtoTinto
            nwRw("IDUdMostoTinto") = IDUDMedida
        End If
        nwRw("Texto") = "Rendimiento Blanco: " & Nz(drVendimia("RdtoB"), 100) & "%; Rendimiento Tinto: " & Nz(drVendimia("RdtoT"), 100) & "%."
        nwRw("Declarado") = False

        clsLibro.Update(dtLibro)

    End Function

    <Task()> Public Shared Function CrearEntradasMosto(ByVal data As DatosEntradasSalidas, ByVal services As ServiceProvider) As String

        Dim clsLibro As New BdgLibroEntradasSalidasGranel
        Dim dtLibro As DataTable = clsLibro.AddNew

        AdminData.BeginTx()

        'Borrar lineas existentes
        Dim f As New Filter
        f.Add("FechaMovimiento", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
        f.Add("FechaMovimiento", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
        f.Add("Declarado", FilterOperator.Equal, False, FilterType.Boolean)
        f.Add("TipoMovimiento", FilterOperator.Equal, enumBdgTipoMovimientoLibro.CompraMosto, FilterType.Numeric)
        Dim dtLibroBorrar As DataTable = New BdgLibroEntradasSalidasGranel().Filter(f)
        If dtLibroBorrar.Rows.Count > 0 Then
            clsLibro.Delete(dtLibroBorrar)
        End If

        'Crear nuevas lineas
        Dim f2 As New Filter
        f2.Add("Fecha", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
        f2.Add("Fecha", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
        Dim dtOrigen As DataTable = New BE.DataEngine().Filter("NegBdgLibroESGranelCompraMosto", f2)
        If dtOrigen.Rows.Count > 0 Then
            For Each oRw As DataRow In dtOrigen.Select
                Dim nwRw As DataRow = dtLibro.NewRow
                dtLibro.Rows.Add(nwRw)
                nwRw("FechaMovimiento") = oRw("Fecha")
                nwRw("Entrada") = enumBdgLibroEntradaSalida.Entrada
                nwRw("TipoMovimiento") = enumBdgTipoMovimientoLibro.CompraMosto
                If Length(oRw("NDAA")) > 0 Then
                    nwRw("ClaseDocumentoCirculacion") = "DAA"
                    nwRw("SerieDocumentoCirculacion") = oRw("NDAA")
                End If
                nwRw("Procedencia") = "Compra Mosto(" & oRw("NEntrada") & "): " & oRw("DescArticulo")
                If oRw("QBlanco") <> 0 Then
                    nwRw("QMostoBlanco") = oRw("QBlanco")
                    nwRw("IDUdMostoBlanco") = oRw("IDUdMedida")
                End If
                If oRw("QTinto") <> 0 Then
                    nwRw("QMostoTinto") = oRw("QTinto")
                    nwRw("IDUdMostoTinto") = oRw("IDUdMedida")
                End If
                nwRw("Texto") = "Proveedor: " & oRw("DescProveedor")
                nwRw("Declarado") = False
            Next
            clsLibro.Update(dtLibro)
        Else
            ApplicationService.GenerateError("No hay Entradas de Mosto en el periodo seleccionado.")
        End If

    End Function

    <Task()> Public Shared Function TratarMostoParaVino(ByVal data As DatosEntradasSalidas, ByVal services As ServiceProvider) As String

        Dim clsLibro As New BdgLibroEntradasSalidasGranel
        Dim dtLibro As DataTable = clsLibro.AddNew

        'Acumular Q
        Dim dblQMostoBlanco As Double = 0
        Dim dblQMostoTinto As Double = 0
        Dim IDUDMedida As String = String.Empty
        Dim f As New Filter
        f.Add("FechaMovimiento", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
        f.Add("FechaMovimiento", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
        f.Add("Declarado", FilterOperator.Equal, True, FilterType.Boolean)
        Dim f2 As New Filter(FilterUnionOperator.Or)
        f2.Add("TipoMovimiento", FilterOperator.Equal, enumBdgTipoMovimientoLibro.EntradaMosto, FilterType.Numeric)
        f2.Add("TipoMovimiento", FilterOperator.Equal, enumBdgTipoMovimientoLibro.CompraMosto, FilterType.Numeric)
        f.Add(f2)

        Dim dtOrigen As DataTable = New BdgLibroEntradasSalidasGranel().Filter(f)
        If dtOrigen.Rows.Count > 0 Then
            For Each oRw As DataRow In dtOrigen.Select
                dblQMostoBlanco += Nz(oRw("QMostoBlanco"), 0)
                dblQMostoTinto += Nz(oRw("QMostoTinto"), 0)
                IDUDMedida = Nz(oRw("IDUdMostoBlanco"), oRw("IDUdMostoTinto")) & String.Empty
            Next
        End If
        If dblQMostoBlanco + dblQMostoTinto = 0 Then
            ApplicationService.GenerateError("No hay Entradas de Mosto Declaradas en el periodo seleccionado.")
        End If

        AdminData.BeginTx()

        'Crear salida mosto
        Dim nwRw As DataRow = dtLibro.NewRow
        dtLibro.Rows.Add(nwRw)
        nwRw("FechaMovimiento") = data.FechaHasta
        nwRw("Entrada") = enumBdgLibroEntradaSalida.Salida
        nwRw("TipoMovimiento") = enumBdgTipoMovimientoLibro.SalidaMosto
        nwRw("Procedencia") = "Mosto para elaboración de Vino"
        If dblQMostoBlanco <> 0 Then
            nwRw("QMostoBlanco") = dblQMostoBlanco
            nwRw("IDUdMostoBlanco") = IDUDMedida
        End If
        If dblQMostoTinto <> 0 Then
            nwRw("QMostoTinto") = dblQMostoTinto
            nwRw("IDUdMostoTinto") = IDUDMedida
        End If
        nwRw("Declarado") = False

        'Crear entrada elaboracion vino
        ''''''''Dim dblRdtoBlanco As Double = 0
        ''''''''Dim dblRdtoTinto As Double = 0
        ''''''''Dim clsBdgVendimia As New BdgVendimia
        ''''''''Dim drVendimia As DataRow = clsBdgVendimia.GetItemRow(data.Vendimia)
        ''''''''dblRdtoBlanco = Nz(drVendimia("RdtoB"), 100) / 100
        ''''''''dblRdtoTinto = Nz(drVendimia("RdtoT"), 100) / 100
        ''''''''nwRw = dtLibro.NewRow
        ''''''''dtLibro.Rows.Add(nwRw)
        ''''''''nwRw("FechaMovimiento") = data.FechaHasta
        ''''''''nwRw("Entrada") = enumBdgLibroEntradaSalida.Entrada
        ''''''''nwRw("TipoMovimiento") = enumBdgTipoMovimientoLibro.EntradaMosto
        ''''''''nwRw("Procedencia") = "Mosto procedente de Uva"
        ''''''''If dblQMostoBlanco <> 0 Then
        ''''''''    nwRw("QMostoBlanco") = dblQMostoBlanco * dblRdtoBlanco
        ''''''''    nwRw("IDUdMostoBlanco") = IDUDMedida
        ''''''''End If
        ''''''''If dblQMostoTinto <> 0 Then
        ''''''''    nwRw("QMostoTinto") = dblQMostoTinto * dblRdtoTinto
        ''''''''    nwRw("IDUdMostoTinto") = IDUDMedida
        ''''''''End If
        ''''''''nwRw("Texto") = "Rendimiento Blanco: " & Nz(drVendimia("RdtoB"), 100) & "%; Rendimiento Tinto: " & Nz(drVendimia("RdtoT"), 100) & "%."
        ''''''''nwRw("Declarado") = False

        clsLibro.Update(dtLibro)

    End Function

    <Task()> Public Shared Function CrearEntradasVino(ByVal data As DatosEntradasSalidas, ByVal services As ServiceProvider) As String

        Dim clsLibro As New BdgLibroEntradasSalidasGranel
        Dim dtLibro As DataTable = clsLibro.AddNew

        AdminData.BeginTx()

        'Borrar lineas existentes
        Dim f As New Filter
        f.Add("FechaMovimiento", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
        f.Add("FechaMovimiento", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
        f.Add("Declarado", FilterOperator.Equal, False, FilterType.Boolean)
        f.Add("TipoMovimiento", FilterOperator.Equal, enumBdgTipoMovimientoLibro.CompraVino, FilterType.Numeric)
        Dim dtLibroBorrar As DataTable = New BdgLibroEntradasSalidasGranel().Filter(f)
        If dtLibroBorrar.Rows.Count > 0 Then
            clsLibro.Delete(dtLibroBorrar)
        End If

        'Crear nuevas lineas
        Dim f2 As New Filter
        f2.Add("Fecha", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
        f2.Add("Fecha", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
        Dim dtOrigen As DataTable = New BE.DataEngine().Filter("NegBdgLibroESGranelCompraVino", f2)
        If dtOrigen.Rows.Count > 0 Then
            For Each oRw As DataRow In dtOrigen.Select
                Dim nwRw As DataRow = dtLibro.NewRow
                dtLibro.Rows.Add(nwRw)
                nwRw("FechaMovimiento") = oRw("Fecha")
                nwRw("Entrada") = enumBdgLibroEntradaSalida.Entrada
                nwRw("TipoMovimiento") = enumBdgTipoMovimientoLibro.CompraVino
                If Length(oRw("NDAA")) > 0 Then
                    nwRw("ClaseDocumentoCirculacion") = "DAA"
                    nwRw("SerieDocumentoCirculacion") = oRw("NDAA")
                End If
                nwRw("Procedencia") = "Compra Vino(" & oRw("NEntrada") & "): " & oRw("DescArticulo")
                If oRw("QBlanco") <> 0 Then
                    nwRw("QVCPRDBlanco") = oRw("QBlanco")
                    nwRw("IDUdVCPRDBlanco") = oRw("IDUdMedida")
                End If
                If oRw("QTinto") <> 0 Then
                    nwRw("QVCPRDTinto") = oRw("QTinto")
                    nwRw("IDUdVCPRDTinto") = oRw("IDUdMedida")
                End If
                nwRw("Texto") = "Proveedor: " & oRw("DescProveedor")
                nwRw("Declarado") = False
            Next
            clsLibro.Update(dtLibro)
        Else
            ApplicationService.GenerateError("No hay Entradas de Vino en el periodo seleccionado.")
        End If

    End Function

    <Task()> Public Shared Function Declarar(ByVal data As DatosEntradasSalidas, ByVal services As ServiceProvider) As String

        Dim clsLibro As New BdgLibroEntradasSalidasGranel

        AdminData.BeginTx()

        'Actualizar
        Dim f As New Filter
        f.Add("FechaMovimiento", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
        f.Add("FechaMovimiento", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
        f.Add("Declarado", FilterOperator.Equal, False, FilterType.Boolean)
        Dim dtLibro As DataTable = New BdgLibroEntradasSalidasGranel().Filter(f)
        If dtLibro.Rows.Count > 0 Then
            For Each oRw As DataRow In dtLibro.Select
                oRw("Declarado") = True
            Next
            clsLibro.Update(dtLibro)
        Else
            ApplicationService.GenerateError("No hay datos para declarar en el periodo seleccionado.")
        End If

    End Function
#End Region

#Region "Eventos Entidad"
    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarEstado)
    End Sub

    <Task()> Friend Shared Sub ComprobarEstado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Declarado") = True Then
            ApplicationService.GenerateError("Sólo se permite eliminar movimientos sin declarar.")
        End If
    End Sub

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Friend Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("FechaMovimiento") = Date.Today
        data("Entrada") = 1
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarZonaVitivinicola)
    End Sub

    <Task()> Public Shared Sub AsignarZonaVitivinicola(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("ZonaVitivinicola")) = 0 Then
                Dim clsBDGParametro As New BdgParametro
                data("ZonaVitivinicola") = clsBDGParametro.BdgZonaVitivinicolaLibros
            End If
        End If
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaMovimiento")) = 0 Then ApplicationService.GenerateError("La Fecha de Movimiento es obligatoria.")
        If Length(data("TipoMovimiento")) = 0 Then ApplicationService.GenerateError("El Tipo de Movimiento es obligatorio.")
        If Length(data("Entrada")) = 0 Then ApplicationService.GenerateError("La columna Entrada es obligatoria.")
        If Length(data("Procedencia")) = 0 Then ApplicationService.GenerateError("La Procedencia es obligatoria.")
    End Sub
#End Region

End Class

Public Enum enumBdgTipoMovimientoLibro
    EntradaUva = 1
    SalidaUva = 2
    EntradaMosto = 3
    CompraMosto = 4
    SalidaMosto = 5
    CompraVino = 6
End Enum

Public Enum enumBdgLibroEntradaSalida
    Salida = 0
    Entrada = 1
End Enum
