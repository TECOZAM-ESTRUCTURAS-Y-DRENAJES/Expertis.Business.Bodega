Imports Solmicro.Expertis.Business.Bodega.BdgDAACabecera
Imports Solmicro.Expertis.Business.Bodega.BdgDAALineaPaquete
Imports Solmicro.Expertis.Business.Bodega.BdgDAALineaOperacion

Public Class BdgDAALinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbDAALinea"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

#End Region

#Region "Clases Públicas"

    <Serializable()> _
    Public Class StCrearDAALineas
        Public IDDAA As Guid
        Public DttIDAlbaran As DataTable
        Public TipoEnvioDAA As BdgDAACabecera.enumTipoEnvioDAA
        Public NDaaOrigen As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDDAA As Guid, ByVal DttIDAlbaran As DataTable, ByVal TipoEnvioDAA As BdgDAACabecera.enumTipoEnvioDAA, Optional ByVal NDaaOrigen As String = "")
            Me.IDDAA = IDDAA
            Me.DttIDAlbaran = DttIDAlbaran
            Me.TipoEnvioDAA = TipoEnvioDAA
            Me.NDaaOrigen = NDaaOrigen
        End Sub
    End Class

    <Serializable()> _
    Public Class StValidarDAALineas
        Public RegistrosEmpresas As RegistroEmpresaInfo
        Public TipoEnvioDAA As BdgDAACabecera.enumTipoEnvioDAA
        Public Origen As enumOrigenDAA

        Public CampoID As String = "IDAlbaran"
        Public VistaOrigenLineas As String = CN_VistaDAAAlbaran

        Public Sub New()
        End Sub

        Public Sub New(ByVal Origen As enumOrigenDAA, ByVal registrosEmpresas As RegistroEmpresaInfo, ByVal TipoEnvioDAA As BdgDAACabecera.enumTipoEnvioDAA, ByVal campoID As String, ByVal vistaOrigenLineas As String)
            Me.Origen = Origen
            Me.RegistrosEmpresas = registrosEmpresas
            Me.TipoEnvioDAA = TipoEnvioDAA
            Me.CampoID = campoID
            Me.VistaOrigenLineas = vistaOrigenLineas
        End Sub
    End Class

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added AndAlso Length(data("IDDaaLinea")) = 0 Then data("IDDaaLinea") = Guid.NewGuid
    End Sub

    '<Task()> Public Shared Sub CrearDAALineas(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider)
    '    'Comprobaciones
    '    If (data.Cabecera Is Nothing OrElse data.Cabecera.Rows.Count = 0) Then
    '        ApplicationService.GenerateError("No se ha podido crear la cabecera del DAA.")
    '    End If

    '    Dim bsnDAALin As New BdgDAALinea
    '    data.Lineas = bsnDAALin.AddNew()

    '    'Borrar lineas existentes
    '    Dim f As New Filter
    '    f.Add(New StringFilterItem("IDDAA", data.Cabecera.Rows(0)("IDDAA")))
    '    Dim dttBorrables As DataTable = bsnDAALin.Filter(f)
    '    bsnDAALin.Delete(dttBorrables)

    '    'TODO - TEMPORAL - Hay que modificarlo por la info por defecto
    '    Dim bsnParam As New BdgParametro
    '    Dim strCodigoZonaVitivinicola As String = bsnParam.EmcsCodigoZonaVitivinicola()
    '    Dim strCodigoManipulacionVitivinicola As String = bsnParam.EmcsCodigoManipulacionVitivinicola()

    '    'Agrupación
    '    'TODO - de momento lo hacemos sin agrupar - ya veremos
    '    'GUARDAR DATOS
    '    Dim i As Integer = 1
    '    For Each dtrRow As DataRow In data.LineasOrigen.Rows
    '        Dim newRow As DataRow = data.Lineas.NewRow
    '        data.Lineas.Rows.Add(newRow)
    '        newRow("IDDaaLinea") = Guid.NewGuid
    '        newRow("IDDaa") = data.Cabecera.Rows(0)("IDDaa")
    '        newRow("NDaa") = data.Cabecera.Rows(0)("NDaa")
    '        newRow("CodNC1") = dtrRow("CodigoNC")                                               'CN code
    '        newRow("Descripcion") = String.Format("{0} {1} {2}", _
    '                                              dtrRow("DescSubfamilia"), _
    '                                              dtrRow("QServida"), _
    '                                              dtrRow("DescUdMedida"))                             'Descripcion
    '        newRow("NumeroRegistro") = Format(i, "000")                                         'BodyRecordNumber

    '        Dim dblGrado As Double = ProcessServer.ExecuteTask(Of String, Double)(AddressOf CalcularGrado, dtrRow("IDArticulo"), services) * dtrRow("QServida") '* newRow("Litros")
    '        Dim dblVolumen As Double = dtrRow("QServida") '* oRw("Litros")

    '        newRow("Grado") = xRound((dblGrado / dblVolumen), 2)                                'AlcoholicStrength
    '        newRow("PesoBruto_Kg") = dtrRow("PesoBruto") * dtrRow("QServida")                   'GrossWeigth
    '        newRow("PesoNeto_Kg") = dtrRow("PesoNeto") * dtrRow("QServida")                     'Net Weigth
    '        newRow("Volumen_Litros") = newRow("PesoNeto_Kg")

    '        newRow("IDPartidaEstadistica") = dtrRow("IDPartidaEstadistica")
    '        newRow("IDUdMedida") = dtrRow("IDUdMedida")
    '        newRow("CodigoProductoImpuestos") = dtrRow("CodigoProductoImpuestos")               'ExciseProductCode
    '        newRow("IDTipoEmbalaje") = dtrRow("IDTipoEmbalaje")                                 'KindOfPackages
    '        newRow("NumeroBultos") = dtrRow("QServida")                                         'NumberOfPackages
    '        newRow("CategoriaProductoVitivinicola") = dtrRow("CategoriaProductoVitivinicola")   'WineProductCategory
    '        newRow("CodigoZonaVitivinicola") = strCodigoZonaVitivinicola                        'WineGrowingZoneCode
    '        newRow("CodigoManipulacionVitivinicola") = strCodigoManipulacionVitivinicola        'WineOperationCode
    '        newRow("EpigrafeProducto") = dtrRow("EpigrafeProducto")                             'EpigrafeProducto

    '        newRow("CodigoProductoImpuestos") = dtrRow("CodigoProductoImpuestos")

    '        'TODO - rellenar?
    '        newRow("GradoPlat") = ""
    '        newRow("MarcaFiscal") = ""
    '        newRow("MarcaFiscalFlag") = ""
    '        newRow("TamanoProductor") = ""
    '        newRow("Densidad") = ""
    '        newRow("NombreMarca") = ""
    '        newRow("IdentificadorPrecinto") = ""
    '        newRow("InformacionPrecinto") = ""
    '        newRow("InfoExtraVitivinicola") = ""
    '        newRow("CodigoBiocarburante") = ""
    '        newRow("PorcentajeBiocarburante") = ""

    '        i += 1
    '    Next

    '    bsnDAALin.Update(data.Lineas)
    'End Sub

    '<Task()> Public Shared Function CrearDAALineasAgrupadas(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
    '    Dim ClsDAALin As New BdgDAALinea
    '    Dim dtDAALineas As DataTable = ClsDAALin.AddNew()
    '    Dim dtAlbaranLineas As DataTable
    '    Dim clsParametro As New BdgParametro

    '    AdminData.BeginTx()

    '    ''Borrar lineas existentes
    '    'dtDAALineas = ClsDAALin.Filter(New GuidFilterItem("IDDAA", FilterOperator.Equal, data.IDDAA))
    '    'If dtDAALineas.Rows.Count > 0 Then
    '    '    For Each rwDAALineas As DataRow In dtDAALineas.Rows
    '    '        rwDAALineas.Delete()
    '    '    Next
    '    'End If

    '    ''Filtro para todos los albaranes
    '    'Dim FiltroAV As New Filter(FilterUnionOperator.Or)
    '    'For Each rwAlbaran As DataRow In data.DttIDAlbaran.Rows
    '    '    FiltroAV.Add("IDAlbaran", rwAlbaran("IDAlbaran"))
    '    'Next

    '    ''Actualizar los AlbaranesVentaCabecera
    '    'Dim oAlbVentaCab As New AlbaranVentaCabecera
    '    'Dim dtAVC As DataTable
    '    'dtAVC = oAlbVentaCab.Filter(FiltroAV)
    '    'For Each rwAVC As DataRow In dtAVC.Rows
    '    '    rwAVC("IDDAA") = data.IDDAA
    '    'Next
    '    'BusinessHelper.UpdateTable(dtAVC)

    '    'Se crea una línea de DAA por cada CodigoNC.
    '    'En la descripción se pone una fila por cada combinación de Partida Estadística, Unidad de Medida y Sufijo.
    '    'Si el parámetro BDGDAASUBF = 1 en la descripción se pone una fila por cada combinación de Partida Estadística,
    '    'Unidad de Medida, Sufijo y Subfamilia.

    '    Dim stCamposAgrup As String = String.Empty
    '    Dim Filtro As New Filter

    '    Dim blEmcsAgruparLineas As Boolean = clsParametro.EmcsAgruparLineas
    '    If blEmcsAgruparLineas Then
    '        If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSIntracomunitario Then
    '            stCamposAgrup = "CodigoProductoImpuestos, CodigoNC, IDTipoEmbalaje, CategoriaProductoVitivinicola, EpigrafeProducto"
    '        ElseIf data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSInterno Then
    '            stCamposAgrup = "CodigoNC, IDTipoEmbalaje, CategoriaProductoVitivinicola, EpigrafeProducto" 'TODO => GRADO??
    '        Else
    '            stCamposAgrup = "CodigoNC"
    '        End If

    '        Filtro.Add(New IsNullFilterItem("CodigoNC", False))
    '        If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSIntracomunitario Or data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSInterno Then
    '            Filtro.Add(New IsNullFilterItem("IDTipoEmbalaje", False))
    '            Filtro.Add(New IsNullFilterItem("CategoriaProductoVitivinicola", False))
    '            Filtro.Add("PesoBruto", FilterOperator.GreaterThan, 0)
    '        End If
    '        If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSIntracomunitario Then
    '            Filtro.Add(New IsNullFilterItem("CodigoProductoImpuestos", False))
    '            Filtro.Add("PesoNeto", FilterOperator.GreaterThan, 0)
    '        End If
    '    Else
    '        stCamposAgrup = "IDLineaAlbaran"
    '    End If

    '    Dim dtGruposCodigoNC As DataTable
    '    Dim stSql As String = String.Empty
    '    stSql = "select " & stCamposAgrup
    '    stSql = stSql & " from NegBdgAlbaranLineas"
    '    stSql = stSql & " where " & Filtro.Compose(New AdoFilterComposer)
    '    stSql = stSql & " group by " & stCamposAgrup
    '    dtGruposCodigoNC = AdminData.GetData(stSql, False)

    '    If dtGruposCodigoNC.Rows.Count > 0 Then

    '        'Preparar parámetros
    '        'Dim oPrm As BdgParametro = New BdgParametro
    '        Dim strCodigoZonaVitivinicola As String = String.Empty
    '        Dim strCodigoManipulacionVitivinicola As String = String.Empty
    '        If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSIntracomunitario Or data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSInterno Then
    '            strCodigoZonaVitivinicola = clsParametro.EmcsCodigoZonaVitivinicola()
    '            strCodigoManipulacionVitivinicola = clsParametro.EmcsCodigoManipulacionVitivinicola()
    '        End If
    '        'Traemos todas las líneas de AV que sean de artículos de bodega
    '        dtAlbaranLineas = AdminData.GetData("NegBdgAlbaranLineas", Filtro)
    '        dtAlbaranLineas.Columns.Add("Litros", GetType(Double))

    '        'Calcular litros
    '        Dim stError As String
    '        Dim strUDDaa As String = String.Empty
    '        strUDDaa = clsParametro.DaaUnidad
    '        Dim oArt As New Articulo
    '        For Each oRw As DataRow In dtAlbaranLineas.Rows
    '            Dim StUns As New ArticuloUnidadAB.DatosFactorConversion(oRw("IDArticulo"), oRw("IDUDMedida"), strUDDaa)
    '            oRw("Litros") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf Articulo.FactorDeConversionUnidadesInternas, StUns, services)
    '        Next

    '        Dim stNDaa As String = String.Empty
    '        Dim blAgrupSubFamilia As Boolean = clsParametro.AgrupSubFamilia
    '        Dim nwRw As DataRow
    '        Dim i As Integer = 1
    '        For Each rwGrupo As DataRow In dtGruposCodigoNC.Rows

    '            'PREPARAR DATOS
    '            'Variables
    '            Dim strIDPartidaEstAnterior As String = String.Empty
    '            Dim strIDUDMedidaAnterior As String = String.Empty
    '            Dim strSufijoAnterior As String = String.Empty
    '            Dim strDescSubFamAnterior As String = String.Empty

    '            Dim strDescripcion As String = String.Empty
    '            Dim strDescUDMedidaAnterior As String = String.Empty
    '            Dim strDescPartidaEstAnterior As String = String.Empty

    '            Dim intGrado As Double = 0
    '            Dim intBruto As Double = 0
    '            Dim intNeto As Double = 0
    '            Dim intVolumen As Double = 0
    '            Dim intCajas As Double = 0
    '            Dim intCajasIDUDMedida As Double = 0
    '            Dim intItemVol As Double = 0

    '            Dim strIDPartidaEst As String = String.Empty
    '            Dim strIDUDMedida As String = String.Empty
    '            Dim strCodigoProductoImpuestos As String = String.Empty
    '            Dim strIDTipoEmbalaje As String = String.Empty
    '            Dim strCategoriaProductoVitivinicola As String = String.Empty
    '            Dim strEpigrafeProducto As String = String.Empty
    '            Dim strCodigoNC As String = String.Empty

    '            Dim blnCambio As Boolean = False

    '            'Filtrar por CodigoNC dentro de las líneas de AV
    '            Dim stOrden As String = String.Empty
    '            Dim f As New Filter

    '            If blEmcsAgruparLineas Then
    '                If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSIntracomunitario Or data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSInterno Then
    '                    f.Add(New StringFilterItem("IDTipoEmbalaje", rwGrupo("IDTipoEmbalaje")))
    '                    f.Add(New StringFilterItem("CategoriaProductoVitivinicola", rwGrupo("CategoriaProductoVitivinicola")))
    '                    f.Add(New StringFilterItem("EpigrafeProducto", rwGrupo("EpigrafeProducto")))
    '                    stOrden = "IDUDMedida"
    '                End If
    '                If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSIntracomunitario Then
    '                    f.Add(New StringFilterItem("CodigoProductoImpuestos", rwGrupo("CodigoProductoImpuestos")))
    '                End If
    '                If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.Modelo500 Then
    '                    stOrden = "IDPartidaEstadistica, IDUDMedida, Sufijo, DescSubFamilia"
    '                End If
    '                f.Add(New StringFilterItem("CodigoNC", rwGrupo("CodigoNC")))
    '            Else
    '                f.Add(New NumberFilterItem("IDLineaAlbaran", rwGrupo("IDLineaAlbaran")))
    '            End If

    '            For Each oRw As DataRow In dtAlbaranLineas.Select(f.Compose(New AdoFilterComposer), stOrden)
    '                If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.Modelo500 Then
    '                    blnCambio = strIDPartidaEstAnterior <> oRw("IDPartidaEstadistica") _
    '                                    OrElse strIDUDMedidaAnterior <> oRw("IDUdMedida") _
    '                                    OrElse strSufijoAnterior <> oRw("Sufijo") & String.Empty
    '                    If blAgrupSubFamilia Then
    '                        blnCambio = blnCambio OrElse strDescSubFamAnterior <> oRw("DescSubFamilia") & String.Empty
    '                    End If
    '                    If blnCambio Then
    '                        If Length(strIDPartidaEstAnterior) > 0 Then
    '                            'No es primer registro
    '                            If Len(strDescripcion) > 0 Then
    '                                strDescripcion &= vbCrLf
    '                            End If
    '                            strDescripcion &= intCajas & " " & strDescUDMedidaAnterior & " " & strDescPartidaEstAnterior
    '                            If blAgrupSubFamilia Then
    '                                strDescripcion &= " " & strDescSubFamAnterior
    '                            End If
    '                            intCajas = 0
    '                        End If

    '                        strIDPartidaEstAnterior = oRw("IDPartidaEstadistica") & String.Empty
    '                        strIDUDMedidaAnterior = oRw("IDUdMedida") & String.Empty
    '                        strSufijoAnterior = oRw("Sufijo") & String.Empty
    '                        If blAgrupSubFamilia Then strDescSubFamAnterior = oRw("DescSubFamilia") & String.Empty

    '                        strDescUDMedidaAnterior = oRw("DescUdMedida") & String.Empty
    '                        strDescPartidaEstAnterior = oRw("DescPartidaEstadistica") & String.Empty
    '                    End If
    '                End If

    '                If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSIntracomunitario Or data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSInterno Then
    '                    blnCambio = strIDUDMedidaAnterior <> oRw("IDUdMedida")
    '                    If blnCambio Then
    '                        If Length(strIDUDMedidaAnterior) > 0 Then
    '                            'No es primer registro
    '                            If Len(strDescripcion) > 0 Then
    '                                strDescripcion &= ";" & vbCrLf
    '                            End If
    '                            strDescripcion &= intCajasIDUDMedida & " " & strDescUDMedidaAnterior
    '                            intCajasIDUDMedida = 0
    '                        End If

    '                        strIDUDMedidaAnterior = oRw("IDUdMedida") & String.Empty
    '                        strDescUDMedidaAnterior = oRw("DescUdMedida") & String.Empty
    '                    End If
    '                End If

    '                intGrado += (ProcessServer.ExecuteTask(Of String, Double)(AddressOf CalcularGrado, oRw("IDArticulo"), services) * oRw("QServida") * oRw("Litros"))
    '                intItemVol = oRw("QServida") * oRw("Litros")
    '                intVolumen += intItemVol
    '                intBruto += (oRw("PesoBruto") * oRw("QServida"))
    '                intNeto += (oRw("PesoNeto") * oRw("QServida"))
    '                intCajas += oRw("QServida")
    '                intCajasIDUDMedida += oRw("QServida")

    '                If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSIntracomunitario Or data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSInterno Then
    '                    'strDescripcion = intCajas & " " & oRw("DescUdMedida") & " " & oRw("DescPartidaEstadistica")

    '                    strIDPartidaEst = oRw("IDPartidaEstadistica") & String.Empty
    '                    strIDUDMedida = oRw("IDUdMedida") & String.Empty

    '                    strCodigoProductoImpuestos = oRw("CodigoProductoImpuestos") & String.Empty
    '                    strIDTipoEmbalaje = oRw("IDTipoEmbalaje") & String.Empty
    '                    strCategoriaProductoVitivinicola = oRw("CategoriaProductoVitivinicola") & String.Empty
    '                    strEpigrafeProducto = oRw("EpigrafeProducto") & String.Empty
    '                End If

    '                strCodigoNC = oRw("CodigoNC") & String.Empty
    '            Next

    '            If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.Modelo500 Then
    '                If Len(strDescripcion) > 0 Then
    '                    strDescripcion &= vbCrLf
    '                End If
    '                strDescripcion &= intCajas & " " & strDescUDMedidaAnterior & " " & strDescPartidaEstAnterior
    '                If blAgrupSubFamilia Then
    '                    strDescripcion &= " " & strDescSubFamAnterior
    '                End If
    '            End If

    '            If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSIntracomunitario Or data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSInterno Then
    '                If Len(strDescripcion) > 0 Then
    '                    strDescripcion &= ";" & vbCrLf
    '                End If
    '                strDescripcion &= intCajasIDUDMedida & " " & strDescUDMedidaAnterior
    '            End If

    '            Dim dbGrado As Double = xRound((intGrado / intVolumen), 2)

    '            'GUARDAR DATOS
    '            nwRw = dtDAALineas.NewRow
    '            dtDAALineas.Rows.Add(nwRw)
    '            nwRw("IDDaa") = data.IDDAA
    '            nwRw("CodNC1") = strCodigoNC
    '            nwRw("Descripcion") = strDescripcion & String.Empty
    '            nwRw("CodNC2") = "V0"
    '            nwRw("CodNC3") = "S"
    '            nwRw("Grado") = dbGrado
    '            nwRw("Volumen_Litros") = intNeto 'intVolumen
    '            nwRw("PesoBruto_Kg") = intBruto
    '            nwRw("PesoNeto_Kg") = intNeto

    '            stNDaa = data.NDaaOrigen

    '            'NDaa cambia si hay más de tres líneas (sólo modelo 500)
    '            If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.Modelo500 Then
    '                Dim nLins As Integer = dtDAALineas.Rows.Count
    '                If nLins > 3 Then
    '                    If (nLins - 1) Mod 3 = 0 Then
    '                        Dim c As Contador.DefaultCounter = ProcessServer.ExecuteTask(Of String, Contador.DefaultCounter)(AddressOf Contador.GetDefaultCounterValue, GetType(BdgDAACabecera).Name, services)
    '                        stNDaa = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, c.CounterID, services)
    '                    End If
    '                End If
    '            End If
    '            nwRw("NDaa") = stNDaa

    '            If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSIntracomunitario Or data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSInterno Then
    '                'BodyRecordUniqueReference
    '                nwRw("NumeroRegistro") = Format(i, "000")

    '                nwRw("IDPartidaEstadistica") = strIDPartidaEst
    '                nwRw("IDUdMedida") = strIDUDMedida

    '                'ExciseProductCode
    '                nwRw("CodigoProductoImpuestos") = strCodigoProductoImpuestos

    '                'KindOfPackages
    '                nwRw("IDTipoEmbalaje") = strIDTipoEmbalaje

    '                'NumberOfPackages
    '                nwRw("NumeroBultos") = intCajas

    '                'WineProductCategory
    '                nwRw("CategoriaProductoVitivinicola") = strCategoriaProductoVitivinicola

    '                'WineGrowingZoneCode
    '                nwRw("CodigoZonaVitivinicola") = strCodigoZonaVitivinicola

    '                'WineOperationCode
    '                nwRw("CodigoManipulacionVitivinicola") = strCodigoManipulacionVitivinicola

    '                'EpigrafeProducto
    '                nwRw("EpigrafeProducto") = strEpigrafeProducto
    '            End If
    '            i = i + 1
    '        Next
    '    End If
    '    ClsDAALin.Update(dtDAALineas)
    '    Return data
    'End Function

    <Task()> Public Shared Function IncluirKitEnDAA(ByVal data As DataRow, ByVal services As ServiceProvider) As Boolean
        Return Nz(data("Grado"), 0) <> 0 AndAlso Length(data("IDPartidaEstadistica")) > 0
    End Function

    '<Task()> Public Shared Sub CrearDAALineas(ByVal data As StCrearDAALineas, ByVal services As ServiceProvider)
    <Task()> Public Shared Function CrearDAALineas(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        Dim ClsDAALin As New BdgDAALinea
        Dim dtDAALineas As DataTable = ClsDAALin.AddNew()
        Dim dtOrigenLineas As DataTable
        Dim clsParametro As New BdgParametro

        AdminData.BeginTx()

        'Borrar lineas existentes
        'dtDAALineas = ClsDAALin.Filter(New GuidFilterItem("IDDAA", FilterOperator.Equal, data.Cabecera.Rows(0)("IDDaa")))
        'If dtDAALineas.Rows.Count > 0 Then
        '    For Each rwDAALineas As DataRow In dtDAALineas.Rows
        '        rwDAALineas.Delete()
        '    Next
        'End If

        Dim stCamposAgrup As String = String.Empty
        Dim Filtro As New Filter
        Filtro.Add(New IsNullFilterItem("CodigoNC", False))
        Filtro.Add(New IsNullFilterItem("IDTipoEmbalaje", False))
        Filtro.Add(New IsNullFilterItem("CategoriaProductoVitivinicola", False))
        Filtro.Add(New IsNullFilterItem("CodigoProductoImpuestos", False))
        Filtro.Add("PesoBruto", FilterOperator.GreaterThan, 0)
        Filtro.Add("PesoNeto", FilterOperator.GreaterThan, 0)

        Dim strCodigoZonaVitivinicola As String = String.Empty
        Dim strCodigoManipulacionVitivinicola As String = String.Empty

        strCodigoZonaVitivinicola = clsParametro.EmcsCodigoZonaVitivinicola()
        strCodigoManipulacionVitivinicola = clsParametro.EmcsCodigoManipulacionVitivinicola()

        '''''''Traemos todas las líneas de AV que sean de artículos de bodega
        '''''''TODO - Aquí habría que realizar esto para cada grupo albaran - empresa y agregar los registros
        dtOrigenLineas = AdminData.GetData(data.VistaOrigenLineas, New NoRowsFilterItem)
        If Not dtOrigenLineas.Columns.Contains("IDBaseDatos") Then dtOrigenLineas.Columns.Add("IDBaseDatos", GetType(String))

        Dim dtbbdd As DataTable = AdminData.GetUserDataBases
        Dim idbasedatosoriginal As Guid = AdminData.GetConnectionInfo.IDDataBase
        For Each dtrRegistroEmpresa As DataRow In data.RegistrosEmpresas.Listado.Rows
            Try
                AdminData.CommitTx(True)
                AdminData.SetCurrentConnection(dtrRegistroEmpresa("IDBaseDatos"))
                AdminData.BeginTx()
                'todo - ojo idalbaran!
                Dim lstKitsIncluidos As New List(Of Integer)

                Dim fLineas As New Filter
                fLineas.Add(Filtro)
                fLineas.Add(data.CampoID, dtrRegistroEmpresa("IDRegistro"))
                Dim dtttemp As DataTable = AdminData.GetData(data.VistaOrigenLineas, fLineas)
                If Not dtttemp.Columns.Contains("IDBaseDatos") Then dtttemp.Columns.Add("IDBaseDatos", GetType(Guid))
                If Not dtttemp Is Nothing AndAlso dtttemp.Rows.Count > 0 Then

                    Dim datAddComponentes As New DataAddLineasPedidoComponentes(data.Origen, dtttemp, Filtro)
                    Dim dtttempComp As DataTable = ProcessServer.ExecuteTask(Of DataAddLineasPedidoComponentes, DataTable)(AddressOf BdgDAALinea.AddLineasPedidoComponentes, datAddComponentes, services)
                    If Not dtttempComp Is Nothing AndAlso dtttempComp.Rows.Count > 0 Then
                        dtttemp = dtttempComp
                    End If

                    Dim datLineasDAA As New DataGetLineasToDAA(data.Origen, dtrRegistroEmpresa("IDBaseDatos"), dtttemp, dtOrigenLineas)
                    dtOrigenLineas = ProcessServer.ExecuteTask(Of DataGetLineasToDAA, DataTable)(AddressOf GetLineasToDAA, datLineasDAA, services)
                End If
            Catch ex As Exception
                AdminData.RollBackTx()

            Finally
                AdminData.CommitTx(True)
                AdminData.SetCurrentConnection(idbasedatosoriginal)

                AdminData.BeginTx()
            End Try
        Next

        Dim blAgrupSubFamilia As Boolean = clsParametro.AgrupSubFamilia
        Dim nwRw As DataRow
        Dim i As Integer = 1
        For Each drOrigenLineas As DataRow In dtOrigenLineas.Select()

            'GUARDAR DATOS
            nwRw = dtDAALineas.NewRow
            dtDAALineas.Rows.Add(nwRw)
            nwRw("IDDaaLinea") = Guid.NewGuid
            nwRw("IDDaa") = data.Cabecera.Rows(0)("IDDaa")

            '17
            'BodyRecordUniqueReference
            nwRw("NumeroRegistro") = Format(i, "000")
            'ExciseProductCode
            nwRw("CodigoProductoImpuestos") = drOrigenLineas("CodigoProductoImpuestos") & String.Empty
            'CnCode
            nwRw("CodNC1") = drOrigenLineas("CodigoNC")
            'GrossWeight
            nwRw("PesoBruto_Kg") = (drOrigenLineas("PesoBruto") * drOrigenLineas("QServida"))
            'NetWeight
            nwRw("PesoNeto_Kg") = drOrigenLineas("PesoNeto") * drOrigenLineas("QServida")
            'Quantity
            nwRw("Volumen_Litros") = nwRw("PesoNeto_Kg")
            'AlcoholicStrength
            nwRw("Grado") = drOrigenLineas("Grado")
            'EpigrafeProducto
            nwRw("EpigrafeProducto") = drOrigenLineas("EpigrafeProducto") 'strEpigrafeProducto

            '17.1
            'Tabla líneas paquete
            Dim dataPaquete As New stCrearDAALineaPaqueteInfo(nwRw("IDDaaLinea").ToString(), drOrigenLineas("IDTipoEmbalaje"), drOrigenLineas("QServida"), String.Empty, String.Empty)
            Dim dttLineasPaquete As DataTable = ProcessServer.ExecuteTask(Of stCrearDAALineaPaqueteInfo, DataTable)(AddressOf BdgDAALineaPaquete.CrearDaaLineasPaquete, dataPaquete, services)
            If Not dttLineasPaquete Is Nothing Then
                If data.LineasPaquete Is Nothing Then data.LineasPaquete = dttLineasPaquete.Clone
                For Each dr As DataRow In dttLineasPaquete.Rows
                    data.LineasPaquete.ImportRow(dr)
                Next
            End If

            '17.2
            'WineProductCategory
            nwRw("CategoriaProductoVitivinicola") = drOrigenLineas("CategoriaProductoVitivinicola") 'strCategoriaProductoVitivinicola
            'WineGrowingZoneCode
            If drOrigenLineas.Table.Columns.Contains("ZonaProductoVitivinicola") AndAlso Length(drOrigenLineas("ZonaProductoVitivinicola")) > 0 Then
                nwRw("CodigoZonaVitivinicola") = drOrigenLineas("ZonaProductoVitivinicola")
            Else
                nwRw("CodigoZonaVitivinicola") = strCodigoZonaVitivinicola
            End If

            '17.2.1
            'WineOperationCode
            Dim dataOperacion As New stCrearDAALineaOperacionInfo(nwRw("IDDaaLinea").ToString(), data.DefaultDAAOperacion)
            Dim dttLineasOperacion As DataTable = ProcessServer.ExecuteTask(Of stCrearDAALineaOperacionInfo, DataTable)(AddressOf BdgDAALineaOperacion.CrearDaaLineasOperacion, dataOperacion, services)
            If Not dttLineasOperacion Is Nothing Then
                If data.LineasOperacion Is Nothing Then data.LineasOperacion = dttLineasOperacion.Clone
                For Each dr As DataRow In dttLineasOperacion.Rows
                    data.LineasOperacion.ImportRow(dr)
                Next
            End If

            'Expertis
            nwRw("Descripcion") = drOrigenLineas("Descripcion") & String.Empty
            'nwRw("IDPartidaEstadistica") = rwGrupo("IDPartidaEstadistica") & String.Empty
            nwRw("IDUdMedida") = drOrigenLineas("IDUdMedida") & String.Empty

            'Seguimiento Origen
            nwRw("TipoOrigen") = data.Origen
            nwRw("IDLinea") = drOrigenLineas(data.CampoAgrupacionOrigenLineasPorDefecto)
            nwRw("IDCabecera") = drOrigenLineas(data.CampoID)
            nwRw("IDBaseDatos") = drOrigenLineas("IDBaseDatos")

            'TODO - rellenar?
            nwRw("GradoPlat") = DBNull.Value
            nwRw("MarcaFiscal") = ""
            nwRw("MarcaFiscalFlag") = ""
            nwRw("TamanoProductor") = ""
            nwRw("Densidad") = 0
            nwRw("NombreMarca") = ""
            'nwRw("IdentificadorPrecinto") = ""
            'nwRw("InformacionPrecinto") = ""
            nwRw("InfoExtraVitivinicola") = ""
            nwRw("CodigoBiocarburante") = ""
            nwRw("PorcentajeBiocarburante") = 0
            nwRw("TienePerdidasExcesos") = 0

            i = i + 1
        Next
        data.Lineas = dtDAALineas.Copy
        Return data
    End Function

    <Serializable()> _
    Public Class DataGetLineasToDAA
        Public Origen As enumOrigenDAA
        Public DatosOrigen As DataTable
        Public IDBaseDatos As Guid

        Public LineasToDAA As DataTable

        Public Sub New(ByVal Origen As enumOrigenDAA, ByVal IDBaseDatos As Guid, ByVal DatosOrigen As DataTable, ByVal LineasToDAA As DataTable)
            Me.Origen = Origen
            Me.IDBaseDatos = IDBaseDatos
            Me.DatosOrigen = DatosOrigen
            Me.LineasToDAA = LineasToDAA
        End Sub
    End Class
    <Task()> Public Shared Function GetLineasToDAA(ByVal data As DataGetLineasToDAA, ByVal services As ServiceProvider) As DataTable

        Dim lstKitsIncluidos As New List(Of Integer)

        Dim LineasNormales As List(Of DataRow) = (From c In data.DatosOrigen _
                                                            Where Not c.IsNull("TipoLinea") AndAlso _
                                                                  c("TipoLinea") = enumavlTipoLineaAlbaran.avlNormal).ToList()
        For Each drLinea As DataRow In LineasNormales
            If Not CType(data.IDBaseDatos, Guid).Equals(Guid.Empty) Then
                drLinea("IDBaseDatos") = data.IDBaseDatos
            End If
            data.LineasToDAA.ImportRow(drLinea)
        Next


        Dim LineasKits As List(Of DataRow) = (From c In data.DatosOrigen _
                                              Where Not c.IsNull("TipoLinea") AndAlso _
                                                    c("TipoLinea") = enumavlTipoLineaAlbaran.avlKit).ToList()
        For Each drLinea As DataRow In LineasKits
            If ProcessServer.ExecuteTask(Of DataRow, Boolean)(AddressOf IncluirKitEnDAA, drLinea, services) Then
                If Not CType(data.IDBaseDatos, Guid).Equals(Guid.Empty) Then
                    drLinea("IDBaseDatos") = data.IDBaseDatos
                End If
                data.LineasToDAA.ImportRow(drLinea)

                Dim FieldPadre As String = "IDLineaAlbaran"
                If data.Origen = enumOrigenDAA.Pedido Then
                    FieldPadre = "IDLineaPedido"
                End If

                lstKitsIncluidos.Add(drLinea(FieldPadre))
            End If
        Next

        Dim LineasComponentes As List(Of DataRow) = (From c In data.DatosOrigen _
                                                    Where Not c.IsNull("TipoLinea") AndAlso _
                                                          c("TipoLinea") = enumavlTipoLineaAlbaran.avlComponente).ToList()

        For Each drLinea As DataRow In LineasComponentes
            If Nz(drLinea("IDLineaPadre"), 0) <> 0 AndAlso Not lstKitsIncluidos.Contains(drLinea("IDLineaPadre")) Then
                If Not CType(data.IDBaseDatos, Guid).Equals(Guid.Empty) Then
                    drLinea("IDBaseDatos") = data.IDBaseDatos
                End If

                data.LineasToDAA.ImportRow(drLinea)
            End If
        Next
        Return data.LineasToDAA
    End Function


    <Task()> Public Shared Function CalcularGrado(ByVal IDIDArticulo As String, ByVal services As ServiceProvider) As Double
        Dim dtt As DataTable = New ArticuloCaracteristica().SelOnPrimaryKey(IDIDArticulo, New BdgParametro().GradoArticulo)
        If Not dtt Is Nothing And dtt.Rows.Count > 0 Then
            Return dtt.Rows(0)("Valor")
        Else : ApplicationService.GenerateError("El artículo | no tiene la característica Grado.", IDIDArticulo)
        End If
    End Function

    <Task()> Public Shared Function ValidarDAALineas(ByVal data As StValidarDAALineas, ByVal services As ServiceProvider) As ClassErrors()
        Dim guidCurrentDBID As Guid = AdminData.GetConnectionInfo.IDDataBase

        'Dim stError As String
        Dim Errores(-1) As ClassErrors

        For Each dtrRegistroEmpresa As DataRow In data.RegistrosEmpresas.Listado.Rows
            Try
                'CONEXIÓN
                AdminData.CommitTx(True)
                AdminData.SetCurrentConnection(dtrRegistroEmpresa("IDBaseDatos"))
                AdminData.BeginTx()

                Dim filtro As New Filter
                filtro.Add(data.CampoID, dtrRegistroEmpresa("IDRegistro"))

                '*********************************
                'TODO - validar empresa origen
                '*********************************
                Dim dtOrigenLineas As DataTable
                dtOrigenLineas = New BE.DataEngine().Filter(data.VistaOrigenLineas, filtro) '"NegBdgAlbaranLineas", filtro)

                'Validar datos
                Dim stErrorArticulo As String
                Dim oArt As New Articulo
                If dtOrigenLineas.Rows.Count > 0 Then

                    If Length(dtOrigenLineas.Rows(0)("IDModoTransporte")) = 0 Then
                        Dim Err As New ClassErrors("", AdminData.GetMessageText("No se ha indicado el Modo de Transporte."))
                        ReDim Preserve Errores(Errores.Length)
                        Errores(Errores.Length - 1) = Err
                        ''stError = stError & "El Albarán no tiene el Modo de Transporte." & vbCrLf & vbCrLf
                        'stError = stError & AdminData.GetMessageText("No se ha indicado el Modo de Transporte.") & vbCrLf & vbCrLf
                    End If


                    Dim datAddComponentes As New DataAddLineasPedidoComponentes(data.Origen, dtOrigenLineas)
                    Dim dtOrigenLineasComp As DataTable = ProcessServer.ExecuteTask(Of DataAddLineasPedidoComponentes, DataTable)(AddressOf BdgDAALinea.AddLineasPedidoComponentes, datAddComponentes, services)
                    If Not dtOrigenLineasComp Is Nothing AndAlso dtOrigenLineasComp.Rows.Count > 0 Then
                        dtOrigenLineas = dtOrigenLineasComp
                    End If
                End If

                Dim dtLineasDAA As DataTable = dtOrigenLineas.Clone
                'Decidir qué líneas van al DAA en función de los kits
                Dim datLineasDAA As New DataGetLineasToDAA(data.Origen, Guid.Empty, dtOrigenLineas, dtLineasDAA)
                dtLineasDAA = ProcessServer.ExecuteTask(Of DataGetLineasToDAA, DataTable)(AddressOf GetLineasToDAA, datLineasDAA, services)

                For Each oRw As DataRow In dtLineasDAA.Rows
                    'VALIDAR CAMPOS
                    stErrorArticulo = String.Empty

                    If Nz(oRw("Grado")) <= 0 Then
                        'stErrorArticulo = stErrorArticulo & AdminData.GetMessageText("    - No tiene la característica de Grado.") & vbCrLf
                        Dim Err As New ClassErrors("Artículo " & oRw("IDArticulo") & String.Empty, AdminData.GetMessageText("No tiene la característica de Grado. El Artículo no se incluirá en el DAA."))
                        ReDim Preserve Errores(Errores.Length)
                        Errores(Errores.Length - 1) = Err

                    End If
                    If Length(oRw("CodigoNC")) = 0 Then
                        'stErrorArticulo = stErrorArticulo & Engine.ParseFormatString(AdminData.GetMessageText("    - La Partida Estadística {0} no tiene el CódigoNC."), oRw("IDPartidaEstadistica")) & vbCrLf
                        Dim Err As New ClassErrors("Artículo " & oRw("IDArticulo") & String.Empty, Engine.ParseFormatString(AdminData.GetMessageText("La Partida Estadística {0} no tiene el CódigoNC. El Artículo no se incluirá en el DAA."), oRw("IDPartidaEstadistica")))
                        ReDim Preserve Errores(Errores.Length)
                        Errores(Errores.Length - 1) = Err
                    End If
                    If Length(oRw("IDTipoEmbalaje")) = 0 Then
                        'stErrorArticulo = stErrorArticulo & Engine.ParseFormatString(AdminData.GetMessageText("    - La Unidad {0} no tiene el Tipo de Embalaje."), oRw("IDUDMedida")) & vbCrLf
                        Dim Err As New ClassErrors("Artículo " & oRw("IDArticulo") & String.Empty, Engine.ParseFormatString(AdminData.GetMessageText("La Unidad {0} no tiene el Tipo de Embalaje. El Artículo no se incluirá en el DAA."), oRw("IDUDMedida")))
                        ReDim Preserve Errores(Errores.Length)
                        Errores(Errores.Length - 1) = Err
                    End If
                    If Length(oRw("CategoriaProductoVitivinicola")) = 0 Then
                        'stErrorArticulo = stErrorArticulo & Engine.ParseFormatString(AdminData.GetMessageText("    - La Partida Estadística {0} no tiene la Categoría del Producto Vitivinícola."), oRw("IDPartidaEstadistica")) & vbCrLf
                        Dim Err As New ClassErrors("Artículo " & oRw("IDArticulo") & String.Empty, Engine.ParseFormatString(AdminData.GetMessageText("La Partida Estadística {0} no tiene la Categoría del Producto Vitivinícola. El Artículo no se incluirá en el DAA."), oRw("IDPartidaEstadistica")))
                        ReDim Preserve Errores(Errores.Length)
                        Errores(Errores.Length - 1) = Err
                    End If
                    If Length(oRw("EpigrafeProducto")) = 0 Then
                        'stErrorArticulo = stErrorArticulo & Engine.ParseFormatString(AdminData.GetMessageText("    - La Partida Estadística {0} no tiene el Epígrafe Nacional del Producto."), oRw("IDPartidaEstadistica")) & vbCrLf
                        Dim Err As New ClassErrors("Artículo " & oRw("IDArticulo") & String.Empty, Engine.ParseFormatString(AdminData.GetMessageText("La Partida Estadística {0} no tiene el Epígrafe Nacional del Producto. El Artículo no se incluirá en el DAA."), oRw("IDPartidaEstadistica")))
                        ReDim Preserve Errores(Errores.Length)
                        Errores(Errores.Length - 1) = Err
                    End If
                    If data.TipoEnvioDAA = BdgDAACabecera.enumTipoEnvioDAA.EMCSIntracomunitario Then
                        If Length(oRw("CodigoProductoImpuestos") & String.Empty) = 0 Then
                            'stErrorArticulo = stErrorArticulo & Engine.ParseFormatString(AdminData.GetMessageText("    - La Partida Estadística {0} no tiene el Código de Producto de Impuestos Especiales."), oRw("IDPartidaEstadistica")) & vbCrLf
                            Dim Err As New ClassErrors("Artículo " & oRw("IDArticulo") & String.Empty, Engine.ParseFormatString(AdminData.GetMessageText("La Partida Estadística {0} no tiene el Código de Producto de Impuestos Especiales. El Artículo no se incluirá en el DAA."), oRw("IDPartidaEstadistica")))
                            ReDim Preserve Errores(Errores.Length)
                            Errores(Errores.Length - 1) = Err
                        End If
                    End If
                    If Nz(oRw("PesoBruto")) <= 0 Then
                        'stErrorArticulo = stErrorArticulo & AdminData.GetMessageText("    - No tiene Peso Bruto.") & vbCrLf
                        Dim Err As New ClassErrors("Artículo " & oRw("IDArticulo") & String.Empty, AdminData.GetMessageText("No tiene Peso Bruto. El Artículo no se incluirá en el DAA."))
                        ReDim Preserve Errores(Errores.Length)
                        Errores(Errores.Length - 1) = Err
                    End If
                    If Nz(oRw("PesoNeto")) <= 0 Then
                        'stErrorArticulo = stErrorArticulo & AdminData.GetMessageText("    - No tiene Peso Neto.") & vbCrLf
                        Dim Err As New ClassErrors("Artículo " & oRw("IDArticulo") & String.Empty, AdminData.GetMessageText("No tiene Peso Neto. El Artículo no se incluirá en el DAA."))
                        ReDim Preserve Errores(Errores.Length)
                        Errores(Errores.Length - 1) = Err
                    End If
                    'If Len(stErrorArticulo) > 0 Then
                    '    stError = stError & AdminData.GetMessageText("ARTICULO ") & oRw("IDArticulo") & ":" & vbCrLf
                    '    stError = stError & stErrorArticulo & vbCrLf
                    'End If
                Next
                ' End If
            Catch ex As Exception
                Dim Err As New ClassErrors("Error de Proceso", ex.Message)
                ReDim Preserve Errores(Errores.Length)
                Errores(Errores.Length - 1) = Err
                AdminData.RollBackTx()
                ApplicationService.GenerateError(ex.Message)
            Finally
                AdminData.CommitTx(True)
                AdminData.SetCurrentConnection(guidCurrentDBID)
                AdminData.BeginTx()
            End Try
        Next

        Return Errores 'stError
    End Function


    <Serializable()> _
    Public Class DataAddLineasPedidoComponentes
        Public Origen As enumOrigenDAA
        Public dtOrigenLineas As DataTable
        Public Filtro As Filter

        Public Sub New(ByVal Origen As enumOrigenDAA, ByVal dtOrigenLineas As DataTable, Optional ByVal Filtro As Filter = Nothing)
            Me.Origen = Origen
            Me.dtOrigenLineas = dtOrigenLineas
            Me.Filtro = Filtro
        End Sub
    End Class
    <Task()> Public Shared Function AddLineasPedidoComponentes(ByVal data As DataAddLineasPedidoComponentes, ByVal services As ServiceProvider) As DataTable
        If data.dtOrigenLineas Is Nothing Then Exit Function
        If data.Origen = enumOrigenDAA.Pedido Then
            '//En el caso de los pedidos, no tenemos las lineas de los componentes, por lo que las meteremos sólo si no tenemos que meter el kit.
            Dim LineasKits As List(Of DataRow) = (From c In data.dtOrigenLineas _
                                                  Where Not c.IsNull("TipoLinea") AndAlso _
                                                        c("TipoLinea") = enumavlTipoLineaAlbaran.avlKit).ToList()
            For Each drLinea As DataRow In LineasKits
                If Not ProcessServer.ExecuteTask(Of DataRow, Boolean)(AddressOf IncluirKitEnDAA, drLinea, services) Then
                    Dim f As New Filter
                    f.Add(New StringFilterItem("IDArticulo", drLinea("IDArticulo")))
                    If Not data.Filtro Is Nothing AndAlso data.Filtro.Count > 0 Then f.Add(data.Filtro)
                    Dim dtComponentes As DataTable = New BE.DataEngine().Filter("vNegBdgDAAArticuloComponentesPrimerNivel", f)
                    If dtComponentes.Rows.Count > 0 Then
                        For Each drComponente As DataRow In dtComponentes.Rows
                            Dim drNew As DataRow = data.dtOrigenLineas.NewRow
                            drNew.ItemArray = drLinea.ItemArray

                            drNew("IDLineaPedido") = -1 * drLinea("IDLineaPedido")
                            drNew("IDArticulo") = drComponente("IDComponente")
                            drNew("Descripcion") = drComponente("Descripcion")
                            drNew("DescSubfamilia") = drComponente("DescSubfamilia")

                            drNew("IDPartidaEstadistica") = drComponente("IDPartidaEstadistica")
                            drNew("DescPartidaEstadistica") = drComponente("DescPartidaEstadistica")
                            drNew("CodigoNC") = drComponente("CodigoNC")
                            drNew("Sufijo") = drComponente("Sufijo")
                            drNew("CodigoProductoImpuestos") = drComponente("CodigoProductoImpuestos")

                            drNew("IDUdMedida") = drComponente("IDUdMedida")
                            drNew("DescUdMedida") = drComponente("DescUdMedida")
                            drNew("IDTipoEmbalaje") = drComponente("IDTipoEmbalaje")

                            drNew("QServida") = drComponente("QServida")
                            drNew("PesoNeto") = drComponente("PesoNeto")
                            drNew("PesoBruto") = drComponente("PesoBruto")
                            drNew("Grado") = drComponente("Grado")

                            drNew("TipoLinea") = enumavlTipoLineaAlbaran.avlComponente
                            drNew("IDLineaPadre") = drLinea("IDLineaPedido")

                            data.dtOrigenLineas.Rows.Add(drNew)
                        Next
                    End If
                End If
            Next
            Return data.dtOrigenLineas
        End If
    End Function

#End Region

End Class