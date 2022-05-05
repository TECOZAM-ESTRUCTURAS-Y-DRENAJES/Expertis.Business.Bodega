Public Class BdgOperacionVino

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgOperacionVino"

#End Region

#Region "Eventos Entidad"

#Region "Tareas RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        ' deleteProcess.AddTask(Of DataRow)(AddressOf BorradoEntidadesRelacionadas)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarOperacionVinoMaterial)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarOperacionVinoCentro)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarAnalisisVino)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarOFRelacionada)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarPorTipoOperacion)
        deleteProcess.AddTask(Of DataRow)(AddressOf CambioOcupacion)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarMovimientos)
    End Sub

    <Task()> Public Shared Sub BorradoEntidadesRelacionadas(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim strIDTipoOp As String = New BdgOperacion().GetItemRow(data("NOperacion"))("IDTipoOperacion")
        Dim TipoOp As Business.Bodega.TipoMovimiento = New BdgTipoOperacion().GetItemRow(strIDTipoOp)("TipoMovimiento")

        If CBool(data("Destino")) Then
            Dim FilVino As New Filter
            FilVino.Add(New GuidFilterItem("IDVino", data("IDVino")))
            FilVino.Add("NOperacion", data("NOperacion"))

            Dim oVMat As New BdgVinoMaterial
            Dim dtVMat As DataTable = oVMat.Filter(FilVino)
            oVMat.Delete(dtVMat)

            Dim oVMod As New BdgVinoMod
            Dim dtVMod As DataTable = oVMod.Filter(FilVino)
            oVMod.Delete(dtVMod)

            Dim oVMaq As New BdgVinoCentro
            Dim dtVMaq As DataTable = oVMaq.Filter(FilVino)
            oVMaq.Delete(dtVMaq)

            Dim oVVrs As New BdgVinoVarios
            Dim dtVVrs As DataTable = oVVrs.Filter(FilVino)
            oVVrs.Delete(dtVVrs)

            Dim oVA As New BdgVinoAnalisis
            Dim dtVA As DataTable = oVA.Filter(FilVino)
            oVA.Delete(dtVA)

            Dim oVEla As New BdgVinoMaterial
            Dim dtVEla As DataTable = oVEla.Filter(FilVino)
            oVEla.Delete(dtVEla)

            Dim oCosteN As New BdgVinoCentro
            Dim dtCosteN As DataTable = oCosteN.Filter(FilVino)
            oCosteN.Delete(dtCosteN)
        End If
    End Sub

    <Task()> Public Shared Sub BorrarOperacionVinoMaterial(ByVal data As DataRow, ByVal services As ServiceProvider)
        If CBool(Nz(data("Destino"), False)) Then
            Dim f As New Filter
            f.Add(New FilterItem("IDVino", data("IDVino")))
            f.Add(New FilterItem("NOperacion", data("NOperacion")))

            Dim oVA As New BdgVinoMaterial
            Dim dtVA As DataTable = oVA.Filter(f)
            oVA.Delete(dtVA)
        End If
    End Sub

    <Task()> Public Shared Sub BorrarOperacionVinoCentro(ByVal data As DataRow, ByVal services As ServiceProvider)
        If CBool(Nz(data("Destino"), False)) Then
            Dim f As New Filter
            f.Add(New FilterItem("IDVino", data("IDVino")))
            f.Add(New FilterItem("NOperacion", data("NOperacion")))

            Dim oVA As New BdgVinoCentro
            Dim dtVA As DataTable = oVA.Filter(f)
            oVA.Delete(dtVA)
        End If
    End Sub

    <Task()> Public Shared Sub BorrarAnalisisVino(ByVal data As DataRow, ByVal services As ServiceProvider)
        If CBool(Nz(data("Destino"), False)) Then
            Dim f As New Filter
            f.Add(New FilterItem("IDVino", data("IDVino")))
            f.Add(New FilterItem("NOperacion", data("NOperacion")))

            Dim oVA As New BdgVinoAnalisis
            Dim dtVA As DataTable = oVA.Filter(f)
            oVA.Delete(dtVA)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarOFRelacionada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If CBool(data("Destino")) Then
            If Length(data("IDOrden")) > 0 Then
                Dim ClsObj As BE.BusinessHelper = CreateBusinessObject("OrdenFabricacion")
                Dim DtOF As DataTable = ClsObj.SelOnPrimaryKey(data("IDOrden"))
                If Not DtOF Is Nothing AndAlso DtOF.Rows.Count > 0 Then
                    DtOF.Rows(0)("QIniciada") -= data("Cantidad")
                    DtOF.Rows(0)("QFabricada") -= data("Cantidad")
                    ClsObj.Update(DtOF)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarPorTipoOperacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If CBool(data("Destino")) Then
            Dim strIDTipoOp As String = New BdgOperacion().GetItemRow(data("NOperacion"))("IDTipoOperacion")
            Dim TipoOp As Business.Bodega.TipoMovimiento = New BdgTipoOperacion().GetItemRow(strIDTipoOp)("TipoMovimiento")
            If TipoOp = Business.Bodega.TipoMovimiento.SinMovimiento Then
                ProcessServer.ExecuteTask(Of Guid)(AddressOf BdgVino.ActualizarUltimoEstadoVino, data("IDVino"), services)
            ElseIf TipoOp = Business.Bodega.TipoMovimiento.CrearOrigen Then
                Dim StDeshacer As New BdgWorkClass.StDeshaceEstructura(data("IDVino"), data("NOperacion"))
                ProcessServer.ExecuteTask(Of BdgWorkClass.StDeshaceEstructura)(AddressOf BdgWorkClass.DeshacerEstructura, StDeshacer, services)
            Else
                Dim StDeshacer As New BdgWorkClass.StDeshacerVino(data("IDVino"), data("NOperacion"), data("NOperacion"), data("Cantidad"))
                ProcessServer.ExecuteTask(Of BdgWorkClass.StDeshacerVino)(AddressOf BdgWorkClass.DeshacerVino, StDeshacer, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioOcupacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not CBool(data("Destino")) Then
            Dim oVinoWC As New BdgWorkClass
            Dim dblQ As Double = data("Cantidad")
            Dim StCambio As New BdgWorkClass.StCambiarOcupacion(data("IDVino"), dblQ)
            ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambio, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarMovimientos(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim Fecha As Date = Today
        Dim dtOperacion As DataTable = New BdgOperacion().SelOnPrimaryKey(data("NOperacion"))
        If dtOperacion.Rows.Count > 0 Then
            Fecha = New Date(CDate(dtOperacion.Rows(0)("Fecha")).Year, CDate(dtOperacion.Rows(0)("Fecha")).Month, CDate(dtOperacion.Rows(0)("Fecha")).Day)

            Dim StEj As New BdgWorkClass.StEjecutarMovimientosNumero(Nz(dtOperacion.Rows(0)("IDMovimiento"), 0), dtOperacion.Rows(0)("NOperacion"), Fecha)
            ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientosNumero)(AddressOf BdgWorkClass.EjecutarMovimientosNumero, StEj, services)
        End If
    End Sub

#End Region

#Region "Tareas GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDDeposito", AddressOf BdgGeneral.CambioIDDeposito)
        oBrl.Add("IDOrden", AddressOf BdgGeneral.CambioIDOrden)
        oBrl.Add("QDeposito", AddressOf BdgGeneral.CambioQDeposito)
        oBrl.Add("Cantidad", AddressOf BdgGeneral.CambioCantidad)
        oBrl.Add("Litros", AddressOf BdgGeneral.CambioLitros)
        oBrl.Add("IDArticulo", AddressOf BdgGeneral.CambioIDArticulo)
        oBrl.Add("IDBarrica", AddressOf BdgGeneral.CambioIDBarrica)
        Return oBrl
    End Function

#End Region

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function SelOnVino(ByVal IDVino As Guid, ByVal services As ServiceProvider) As DataTable
        Return New BdgOperacionVino().Filter(New GuidFilterItem("IDVino", FilterOperator.Equal, IDVino))
    End Function

    <Task()> Public Shared Function SelOnNOperacion(ByVal NOperacion As String, ByVal services As ServiceProvider) As DataTable
        Return New BdgOperacionVino().Filter(New StringFilterItem("NOperacion", NOperacion))
    End Function

    <Serializable()> _
    Public Class StSelOnNOperacionVinoOrigen
        Public NOperacion As String
        Public IDVino As Guid

        Public Sub New()
        End Sub

        Public Sub New(ByVal NOperacion As String, ByVal IDVino As Guid)
            Me.NOperacion = NOperacion
            Me.IDVino = IDVino
        End Sub
    End Class

    <Task()> Public Shared Function SelOnNOperacionVinoOrigen(ByVal data As StSelOnNOperacionVinoOrigen, ByVal services As ServiceProvider) As DataTable
        Dim oFltr As New Filter
        oFltr.Add(New GuidFilterItem("IDVino", data.IDVino))
        oFltr.Add(New StringFilterItem("NOperacion", data.NOperacion))
        oFltr.Add(New BooleanFilterItem("Destino", False))
        Return New BdgOperacionVino().Filter(oFltr)
    End Function

#Region " Modificar Artículo / Depósito "

    <Serializable()> _
    Public Class DataModificarOperacionVinoDestino
        Public IDVino As Guid
        Public NOperacion As String

        Public IDArticuloNew As String
        Public LoteNew As String
        Public IDBarricaNew As String
        Public IDDepositoNew As String

        Public Sub New(ByVal IDVino As Guid, ByVal NOperacion As String, ByVal IDDepositoNew As String)
            Me.IDVino = IDVino
            Me.NOperacion = NOperacion
            Me.IDDepositoNew = IDDepositoNew
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal NOperacion As String, ByVal IDArticuloNew As String, ByVal LoteNew As String, Optional ByVal IDBarricaNew As String = "")
            Me.IDVino = IDVino
            Me.NOperacion = NOperacion
            Me.IDArticuloNew = IDArticuloNew
            Me.LoteNew = LoteNew
            Me.IDBarricaNew = IDBarricaNew
        End Sub
    End Class

    '<Task()> Public Shared Sub ModificarArticulo(ByVal data As DataModificarDestino, ByVal services As ServiceProvider)
    '    ProcessServer.ExecuteTask(Of DataModificarDestino)(AddressOf ModificarOperacionVinoDestino, data, services)
    'End Sub

    '<Task()> Public Shared Sub ModificarArticuloLote(ByVal data As DataModificarDestino, ByVal services As ServiceProvider)
    '    ProcessServer.ExecuteTask(Of DataModificarDestino)(AddressOf ModificarOperacionVinoDestino, data, services)
    'End Sub

    '<Task()> Public Shared Sub ModificarDeposito(ByVal data As DataModificarDestino, ByVal services As ServiceProvider)
    '    ProcessServer.ExecuteTask(Of DataModificarDestino)(AddressOf ModificarOperacionVinoDestino, data, services)
    'End Sub

    <Task()> Public Shared Sub ModificarOperacionVinoDestino(ByVal data As DataModificarOperacionVinoDestino, ByVal services As ServiceProvider)
        If Length(data.NOperacion) = 0 Then Exit Sub

        Dim drVino As DataRow = New BdgVino().GetItemRow(data.IDVino)
        If ((Length(data.IDArticuloNew) > 0 AndAlso drVino("IDArticulo") <> data.IDArticuloNew) OrElse _
           (Length(data.LoteNew) > 0 AndAlso drVino("Lote") <> data.LoteNew) OrElse _
           (Length(data.IDBarricaNew) > 0 AndAlso drVino("IDBarrica") <> data.IDBarricaNew)) OrElse _
           (Length(data.IDDepositoNew) > 0 AndAlso drVino("IDDeposito") <> data.IDDepositoNew) Then

            Dim Doc As New DocumentoBdgOperacionReal(data.NOperacion)
            Dim cnCHAR_SEPARATOR As String = "/"
            If Not Doc.GetOperacionVinoDestino Is Nothing AndAlso Doc.GetOperacionVinoDestino.Count > 0 Then

                Dim RegEliminar As List(Of DataRow) = (From c In Doc.GetOperacionVinoDestino.ToList Where Not c.IsNull("IDVino") AndAlso c("IDVino") = data.IDVino Select c).ToList
                If Not RegEliminar Is Nothing AndAlso RegEliminar.Count > 0 Then
                    Dim datDelete As New DataPrepararDatosEliminarIDVino(data.IDVino, Doc)
                    datDelete.Separator = cnCHAR_SEPARATOR
                    Dim lstEliminar As Dictionary(Of String, DataTable) = ProcessServer.ExecuteTask(Of Object, Dictionary(Of String, DataTable))(AddressOf PrepararDatosEliminarIDVino, datDelete, services)

                    Dim dtOperacionVino As DataTable = Doc.dtOperacionVino.Clone
                    Dim dtOperacionVinoMaterial As DataTable = Doc.dtOperacionVinoMaterial.Clone
                    Dim dtOperacionVinoMaterialLote As DataTable = Doc.dtOperacionVinoMaterialLote.Clone
                    Dim dtOperacionVinoMOD As DataTable = Doc.dtOperacionVinoMOD.Clone
                    Dim dtOperacionVinoCentro As DataTable = Doc.dtOperacionVinoCentro.Clone
                    Dim dtOperacionVinoVarios As DataTable = Doc.dtOperacionVinoVarios.Clone
                    Dim dtOperacionVinoAnalisis As DataTable = Doc.dtOperacionVinoAnalisis.Clone
                    Dim dtOperacionVinoAnalisisVariable As DataTable = Doc.dtOperacionVinoAnalisisVariable.Clone

                    '//Preparamos el registro nuevo
                    Dim context As New BusinessData(Doc.HeaderRow)
                    Dim OPV As New BdgOperacionVino
                    Dim drNewOpVino As DataRow = dtOperacionVino.NewRow
                    For Each col As DataColumn In dtOperacionVino.Columns
                        drNewOpVino(col.ColumnName) = RegEliminar(0)(col.ColumnName)
                    Next
                    If Length(data.IDDepositoNew) > 0 AndAlso drNewOpVino("IDDeposito") <> data.IDDepositoNew Then
                        drNewOpVino = OPV.ApplyBusinessRule("IDDeposito", data.IDDepositoNew, drNewOpVino, context)
                        Dim CostesIDVino As ProcesoBdgOperacion.DataCostesIDVino = services.GetService(Of ProcesoBdgOperacion.DataCostesIDVino)()
                        If Not CostesIDVino.CostesVendimia Is Nothing AndAlso CostesIDVino.CostesVendimia.Rows.Count > 0 Then
                            For Each dr As DataRow In CostesIDVino.CostesVendimia.Rows
                                dr("IDDeposito") = data.IDDepositoNew
                            Next
                        End If
                    End If
                    If Length(data.IDArticuloNew) > 0 AndAlso drNewOpVino("IDArticulo") <> data.IDArticuloNew Then
                        drNewOpVino = OPV.ApplyBusinessRule("IDArticulo", data.IDArticuloNew, drNewOpVino, context)
                    End If
                    If Length(data.LoteNew) > 0 AndAlso drNewOpVino("Lote") <> data.LoteNew Then
                        drNewOpVino = OPV.ApplyBusinessRule("Lote", data.LoteNew, drNewOpVino, context)
                    End If
                    If Length(data.IDBarricaNew) > 0 AndAlso drNewOpVino("IDBarrica") <> data.IDBarricaNew Then
                        drNewOpVino = OPV.ApplyBusinessRule("IDBarrica", data.IDBarricaNew, drNewOpVino, context)
                    End If
                    dtOperacionVino.Rows.Add(drNewOpVino)


                    AdminData.BeginTx()
                    For Each key As String In lstEliminar.Keys.Reverse
                        Dim EntidadDel() As String = key.Split(cnCHAR_SEPARATOR)
                        Dim NOperacion As String = EntidadDel(0) & String.Empty
                        Dim Entidad As String = EntidadDel(1)
                        If Length(NOperacion) > 0 Then
                            If lstEliminar.ContainsKey(key) Then
                                Dim dtImputacionIDVino As DataTable = lstEliminar(key)
                                Dim blnTratar As Boolean = True
                                Dim blnQuitarMovimiento As Boolean = False
                                If Not dtImputacionIDVino Is Nothing AndAlso dtImputacionIDVino.Rows.Count > 0 Then
                                    Dim dtActualizar As DataTable
                                    Select Case Entidad
                                        Case Doc.EntidadOperacionVinoMaterial
                                            dtActualizar = dtOperacionVinoMaterial
                                            blnQuitarMovimiento = True
                                        Case Doc.EntidadOperacionVinoMaterialLote
                                            dtActualizar = dtOperacionVinoMaterialLote
                                            blnQuitarMovimiento = True
                                        Case Doc.EntidadOperacionVinoMOD
                                            dtActualizar = dtOperacionVinoMOD
                                        Case Doc.EntidadOperacionVinoCentro
                                            dtActualizar = dtOperacionVinoCentro
                                        Case Doc.EntidadOperacionVinoVarios
                                            dtActualizar = dtOperacionVinoVarios
                                        Case Doc.EntidadOperacionVinoAnalisis
                                            dtActualizar = dtOperacionVinoAnalisis
                                        Case Doc.EntidadOperacionVinoAnalisisVariable
                                            dtActualizar = dtOperacionVinoAnalisisVariable
                                        Case Else
                                            blnTratar = False
                                    End Select

                                    If blnTratar Then
                                        For Each dr As DataRow In dtImputacionIDVino.Rows
                                            If blnQuitarMovimiento Then dr("IDLineaMovimiento") = System.DBNull.Value
                                            dtActualizar.Rows.Add(dr.ItemArray)
                                        Next
                                    End If
                                End If
                            End If
                        End If
                    Next


                    ProcessServer.ExecuteTask(Of Object)(AddressOf ProcesoBdgOperacion.DeleteCostesVino, Nothing, services) '//Sin operacion
                    ProcessServer.ExecuteTask(Of Object)(AddressOf ProcesoBdgOperacion.DeleteAnalisisVinoNoOperacion, Nothing, services)
                   
                    '//Preparamos registro para quitar el stock del vino
                    Dim dt As DataTable = Doc.dtOperacionVino.Clone
                    dt.ImportRow(RegEliminar(0))
                    dt.Columns.Add("IDAlmacen", GetType(String))
                    Dim IDAlmacen As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf BdgWorkClass.GetAlmacenDeposito, RegEliminar(0)("IDDeposito") & String.Empty, services)
                    dt.Rows(0)("IDAlmacen") = IDAlmacen

                    OPV.Delete(dt, services)
                    '//

                    Dim pck As New UpdatePackage(Doc.HeaderRow.Table)
                    pck.Add(dtOperacionVino)
                    If Not dtOperacionVinoMaterial Is Nothing AndAlso dtOperacionVinoMaterial.Rows.Count > 0 Then pck.Add(dtOperacionVinoMaterial)
                    If Not dtOperacionVinoMaterialLote Is Nothing AndAlso dtOperacionVinoMaterialLote.Rows.Count > 0 Then pck.Add(dtOperacionVinoMaterialLote)
                    If Not dtOperacionVinoMOD Is Nothing AndAlso dtOperacionVinoMOD.Rows.Count > 0 Then pck.Add(dtOperacionVinoMOD)
                    If Not dtOperacionVinoCentro Is Nothing AndAlso dtOperacionVinoCentro.Rows.Count > 0 Then pck.Add(dtOperacionVinoCentro)
                    If Not dtOperacionVinoVarios Is Nothing AndAlso dtOperacionVinoVarios.Rows.Count > 0 Then pck.Add(dtOperacionVinoVarios)
                    If Not dtOperacionVinoAnalisis Is Nothing AndAlso dtOperacionVinoAnalisis.Rows.Count > 0 Then pck.Add(dtOperacionVinoAnalisis)
                    If Not dtOperacionVinoAnalisisVariable Is Nothing AndAlso dtOperacionVinoAnalisisVariable.Rows.Count > 0 Then pck.Add(dtOperacionVinoAnalisisVariable)

                    Dim Op As New BdgOperacion
                    Op.Update(pck, services)

                    ProcessServer.ExecuteTask(Of Object)(AddressOf ProcesoBdgOperacion.ActualizarAnalisisVinoNoOperacion, Nothing, services)
                End If
            End If
        End If
    End Sub

    Public Class DataPrepararDatosEliminarIDVino
        Public Doc As DocumentoBdgOperacionReal
        Public IDVino As Guid
        Public Separator As String

        Public Sub New(ByVal IDVino As Guid, Optional ByVal Doc As DocumentoBdgOperacionReal = Nothing)
            Me.IDVino = IDVino
            Me.Doc = Doc
        End Sub
    End Class
    <Task()> Public Shared Function PrepararDatosEliminarIDVino(ByVal data As DataPrepararDatosEliminarIDVino, ByVal services As ServiceProvider) As Dictionary(Of String, DataTable)
        Dim lstEliminar As Dictionary(Of String, DataTable) = ProcessServer.ExecuteTask(Of DataPrepararDatosEliminarIDVino, Dictionary(Of String, DataTable))(AddressOf PrepararDatosEliminarIDVinoOperacion, data, services)
        Dim lstEliminarCostes As Dictionary(Of String, DataTable) = ProcessServer.ExecuteTask(Of DataPrepararDatosEliminarIDVino, Dictionary(Of String, DataTable))(AddressOf PrepararDatosEliminarIDVinoCostes, data, services)
        If lstEliminar Is Nothing Then lstEliminar = New Dictionary(Of String, DataTable)
        If Not lstEliminarCostes Is Nothing AndAlso lstEliminarCostes.Count > 0 Then
            For Each entidad As String In lstEliminarCostes.Keys
                lstEliminar.Add(entidad, lstEliminarCostes(entidad))
            Next
        End If

        Dim lstEliminarAnalisisNoOp As Dictionary(Of String, DataTable) = ProcessServer.ExecuteTask(Of DataPrepararDatosEliminarIDVino, Dictionary(Of String, DataTable))(AddressOf PrepararDatosEliminarAnalisisNoOp, data, services)
        If Not lstEliminarAnalisisNoOp Is Nothing AndAlso lstEliminarAnalisisNoOp.Count > 0 Then
            For Each entidad As String In lstEliminarAnalisisNoOp.Keys
                lstEliminar.Add(entidad, lstEliminarAnalisisNoOp(entidad))
            Next
        End If

        Return lstEliminar
    End Function
    <Task()> Public Shared Function PrepararDatosEliminarIDVinoOperacion(ByVal data As DataPrepararDatosEliminarIDVino, ByVal services As ServiceProvider) As Dictionary(Of String, DataTable)
        If data.Doc Is Nothing Then Exit Function

        Dim lstEliminar As New Dictionary(Of String, DataTable)
        Dim dtDelete As DataTable
        Dim NOperacion As String = data.Doc.HeaderRow("NOperacion")
        Dim cnCHAR_SEPARATOR As String = Nz(data.Separator, "/")

        Dim RegEliminar As List(Of DataRow) = (From c In data.Doc.dtOperacionVino Where Not c.IsNull("IDVino") AndAlso c("IDVino") = data.IDVino AndAlso c("Destino") = True Select c).ToList
        If Not RegEliminar Is Nothing AndAlso RegEliminar.Count > 0 Then
            dtDelete = RegEliminar(0).Table.Clone
            dtDelete.ImportRow(RegEliminar(0))
            lstEliminar.Add(NOperacion & cnCHAR_SEPARATOR & data.Doc.EntidadOperacionVino, dtDelete.Copy)

            '//Materiales y sus Lotes
            If dtDelete.Rows.Count > 0 Then
                Dim dtDeleteLotes As DataTable = data.Doc.dtOperacionVinoMaterialLote.Clone
                Dim MaterialesVino As List(Of DataRow) = (From c In data.Doc.dtOperacionVinoMaterial Where c("IDVino") = data.IDVino AndAlso c.IsNull("IDOperacionMaterial") Select c).ToList
                dtDelete = data.Doc.dtOperacionVinoMaterial.Clone
                For Each dr As DataRow In MaterialesVino
                    dtDelete.ImportRow(dr)

                    If Not data.Doc.dtOperacionVinoMaterialLote Is Nothing AndAlso data.Doc.dtOperacionVinoMaterialLote.Rows.Count > 0 Then
                        Dim LotesMaterialesVino As List(Of DataRow) = (From c In data.Doc.dtOperacionVinoMaterialLote Where c("IDVinoMaterial") = dr("IDVinoMaterial") Select c).ToList
                        If Not LotesMaterialesVino Is Nothing AndAlso LotesMaterialesVino.Count > 0 Then
                            'dtDeleteLotes = data.Doc.dtOperacionVinoMaterialLote.Clone
                            For Each drLote As DataRow In LotesMaterialesVino
                                dtDeleteLotes.ImportRow(drLote)
                            Next
                        End If
                    End If
                Next

                lstEliminar.Add(NOperacion & cnCHAR_SEPARATOR & data.Doc.EntidadOperacionVinoMaterial, dtDelete.Copy)
                If Not dtDeleteLotes Is Nothing AndAlso dtDeleteLotes.Rows.Count > 0 Then
                    lstEliminar.Add(NOperacion & cnCHAR_SEPARATOR & data.Doc.EntidadOperacionVinoMaterialLote, dtDeleteLotes.Copy)
                End If
            End If

            '//MOD
            Dim MODVino As List(Of DataRow) = (From c In data.Doc.dtOperacionVinoMOD Where c("IDVino") = data.IDVino AndAlso c.IsNull("IDOperacionMOD") Select c).ToList
            dtDelete = data.Doc.dtOperacionVinoMOD.Clone
            For Each dr As DataRow In MODVino
                dtDelete.ImportRow(dr)
            Next
            lstEliminar.Add(NOperacion & cnCHAR_SEPARATOR & data.Doc.EntidadOperacionVinoMOD, dtDelete.Copy)

            '//Centros  y sus Tasas (las Tasas en cascada)
            Dim CentrosVino As List(Of DataRow) = (From c In data.Doc.dtOperacionVinoCentro Where c("IDVino") = data.IDVino AndAlso c.IsNull("IDOperacionCentro") Select c).ToList
            dtDelete = data.Doc.dtOperacionVinoCentro.Clone
            For Each dr As DataRow In CentrosVino
                dtDelete.ImportRow(dr)
            Next
            lstEliminar.Add(NOperacion & cnCHAR_SEPARATOR & data.Doc.EntidadOperacionVinoCentro, dtDelete.Copy)

            '//Varios
            Dim VariosVino As List(Of DataRow) = (From c In data.Doc.dtOperacionVinoVarios Where c("IDVino") = data.IDVino AndAlso c.IsNull("IDOperacionVarios") Select c).ToList
            dtDelete = data.Doc.dtOperacionVinoVarios.Clone
            For Each dr As DataRow In VariosVino
                dtDelete.ImportRow(dr)
            Next
            lstEliminar.Add(NOperacion & cnCHAR_SEPARATOR & data.Doc.EntidadOperacionVinoVarios, dtDelete.Copy)

            '//Analisis y sus variables
            Dim dtDeleteAnVariable As DataTable = data.Doc.dtOperacionVinoAnalisisVariable.Clone
            Dim AnalisisVino As List(Of DataRow) = (From c In data.Doc.dtOperacionVinoAnalisis Where c("IDVino") = data.IDVino Select c).ToList
            dtDelete = data.Doc.dtOperacionVinoAnalisis.Clone
            For Each dr As DataRow In AnalisisVino
                dtDelete.ImportRow(dr)

                If Not data.Doc.dtOperacionVinoAnalisisVariable Is Nothing AndAlso data.Doc.dtOperacionVinoAnalisisVariable.Rows.Count > 0 Then
                    Dim VariablesAnalisis As List(Of DataRow) = (From c In data.Doc.dtOperacionVinoAnalisisVariable Where c("IDVinoAnalisis") = dr("IDVinoAnalisis") Select c).ToList
                    If Not VariablesAnalisis Is Nothing AndAlso VariablesAnalisis.Count > 0 Then
                        For Each drVariable As DataRow In VariablesAnalisis
                            dtDeleteAnVariable.ImportRow(drVariable)
                        Next
                    End If
                End If
            Next

            lstEliminar.Add(NOperacion & cnCHAR_SEPARATOR & data.Doc.EntidadOperacionVinoAnalisis, dtDelete.Copy)
            If Not dtDeleteAnVariable Is Nothing AndAlso dtDeleteAnVariable.Rows.Count > 0 Then
                lstEliminar.Add(NOperacion & cnCHAR_SEPARATOR & data.Doc.EntidadOperacionVinoAnalisisVariable, dtDeleteAnVariable.Copy)
            End If
        End If

        Return lstEliminar
    End Function
   
    <Task()> Public Shared Function PrepararDatosEliminarIDVinoCostes(ByVal data As DataPrepararDatosEliminarIDVino, ByVal services As ServiceProvider) As Dictionary(Of String, DataTable)
        Dim lstEliminar As New Dictionary(Of String, DataTable)
        Dim cnCHAR_SEPARATOR As String = Nz(data.Separator, "/")

        '//Costes Elaboración   
        Dim fCostes As New Filter
        fCostes.Add(New GuidFilterItem(_OV.IDVino, data.IDVino))
        fCostes.Add(New IsNullFilterItem("NOperacion"))
        Dim dtCostesVendimia As DataTable
        Dim dtElaboracion As DataTable = New BdgVinoMaterial().Filter(fCostes)
        If dtElaboracion.Rows.Count > 0 Then
            Dim CVH As New BdgCosteVendimiaHist
            Dim IDVM As Guid
            Dim MaterialesElaboracion As List(Of DataRow) = (From c In dtElaboracion Order By c("IDVinoMaterial") Select c).ToList
            For Each drE As DataRow In MaterialesElaboracion
                If Not Nz(IDVM, Guid.Empty).Equals(drE("IDVinoMaterial")) Then
                    Dim f As New Filter
                    f.Add(New GuidFilterItem("IDVinoMaterial", FilterOperator.Equal, drE("IDVinoMaterial")))
                    Dim dtHis As DataTable = CVH.Filter(f)
                    If dtHis.Rows.Count > 0 Then
                        If dtCostesVendimia Is Nothing Then dtCostesVendimia = dtHis.Clone
                        dtCostesVendimia.ImportRow(dtHis.Rows(0))
                    End If
                End If
            Next

            If Not dtElaboracion Is Nothing AndAlso dtElaboracion.Rows.Count > 0 Then
                lstEliminar.Add(cnCHAR_SEPARATOR & GetType(BdgVinoMaterial).Name, dtElaboracion.Copy)
            End If
            If Not dtCostesVendimia Is Nothing AndAlso dtCostesVendimia.Rows.Count > 0 Then
                lstEliminar.Add(cnCHAR_SEPARATOR & GetType(BdgCosteVendimiaHist).Name, dtCostesVendimia.Copy)
            End If
        End If

        '//Coste de estancia en nave 
        Dim dtCostesNaveTasas As DataTable
        Dim dtCostesNave As DataTable = New BdgVinoCentro().Filter(fCostes)
        If Not dtCostesNave Is Nothing AndAlso dtCostesNave.Rows.Count > 0 Then
            Dim VCT As New BdgVinoCentroTasa
            lstEliminar.Add(cnCHAR_SEPARATOR & GetType(BdgVinoCentro).Name, dtCostesNave.Copy)

            Dim CostesNave As List(Of DataRow) = (From c In dtCostesNave Order By c("IdVinoCentro") Select c).ToList
            For Each dr As DataRow In CostesNave
                Dim f As New Filter
                f.Add(New GuidFilterItem("IdVinoCentro", FilterOperator.Equal, dr("IdVinoCentro")))
                Dim dtTasas As DataTable = VCT.Filter(f)
                If dtTasas.Rows.Count > 0 Then
                    If dtCostesNaveTasas Is Nothing Then dtCostesNaveTasas = dtTasas.Clone
                    dtCostesNaveTasas.ImportRow(dtTasas.Rows(0))
                End If
            Next
        End If

        Dim CostesIDVino As ProcesoBdgOperacion.DataCostesIDVino = services.GetService(Of ProcesoBdgOperacion.DataCostesIDVino)()
        If Not dtElaboracion Is Nothing AndAlso dtElaboracion.Rows.Count > 0 Then CostesIDVino.CostesElaboracion = dtElaboracion.Copy
        If Not dtCostesVendimia Is Nothing AndAlso dtCostesVendimia.Rows.Count > 0 Then CostesIDVino.CostesVendimia = dtCostesVendimia.Copy
        If Not dtCostesNave Is Nothing AndAlso dtCostesNave.Rows.Count > 0 Then CostesIDVino.CostesEstanciaNave = dtCostesNave.Copy
        If Not dtCostesNaveTasas Is Nothing AndAlso dtCostesNaveTasas.Rows.Count > 0 Then CostesIDVino.CostesEstanciaNaveTasas = dtCostesNaveTasas.Copy

        If Not dtElaboracion Is Nothing AndAlso dtElaboracion.Rows.Count > 0 Then CostesIDVino.CostesElaboracionDel = dtElaboracion.Copy
        If Not dtCostesNave Is Nothing AndAlso dtCostesNave.Rows.Count > 0 Then CostesIDVino.CostesEstanciaNaveDel = dtCostesNave.Copy

        Return lstEliminar
    End Function

    <Task()> Public Shared Function PrepararDatosEliminarAnalisisNoOp(ByVal data As DataPrepararDatosEliminarIDVino, ByVal services As ServiceProvider) As Dictionary(Of String, DataTable)
        Dim lstEliminar As New Dictionary(Of String, DataTable)
        Dim cnCHAR_SEPARATOR As String = Nz(data.Separator, "/")

        '//Analisis del Vino sin Operación
        Dim fAnalisisVino As New Filter
        fAnalisisVino.Add(New GuidFilterItem(_OV.IDVino, data.IDVino))
        fAnalisisVino.Add(New IsNullFilterItem("NOperacion"))

        Dim dtAnalisisVariable As DataTable
        Dim dtAnalisis As DataTable = New BdgVinoAnalisis().Filter(fAnalisisVino)
        If dtAnalisis.Rows.Count > 0 Then
            Dim BVV As New BdgVinoVariable

            Dim Analisis As List(Of Object) = (From c In dtAnalisis Order By c("IDVinoAnalisis") Select c("IDVinoAnalisis")).ToList
            For Each IDVinoAnalisis As Guid In Analisis
                Dim f As New Filter
                f.Add(New GuidFilterItem("IDVinoAnalisis", FilterOperator.Equal, IDVinoAnalisis))
                Dim dtAnalisisVariableAux As DataTable = BVV.Filter(f)
                If dtAnalisisVariableAux.Rows.Count > 0 Then
                    If dtAnalisisVariable Is Nothing Then dtAnalisisVariable = dtAnalisisVariableAux.Clone
                    Dim AnalisisVariable As List(Of DataRow) = (From c In dtAnalisisVariableAux Select c).ToList
                    For Each dr As DataRow In AnalisisVariable
                        dtAnalisisVariable.ImportRow(dr)
                    Next
                End If
            Next

            If Not dtAnalisis Is Nothing AndAlso dtAnalisis.Rows.Count > 0 Then
                lstEliminar.Add(cnCHAR_SEPARATOR & GetType(BdgVinoAnalisis).Name, dtAnalisis.Copy)
            End If
            If Not dtAnalisisVariable Is Nothing AndAlso dtAnalisisVariable.Rows.Count > 0 Then
                lstEliminar.Add(cnCHAR_SEPARATOR & GetType(BdgVinoVariable).Name, dtAnalisisVariable.Copy)
            End If
        End If


        Dim CostesIDVino As ProcesoBdgOperacion.DataAnalisisIDVino = services.GetService(Of ProcesoBdgOperacion.DataAnalisisIDVino)()
        If Not dtAnalisis Is Nothing AndAlso dtAnalisis.Rows.Count > 0 Then CostesIDVino.Analisis = dtAnalisis.Copy
        If Not dtAnalisisVariable Is Nothing AndAlso dtAnalisisVariable.Rows.Count > 0 Then CostesIDVino.AnalisisVariable = dtAnalisisVariable.Copy

        If Not dtAnalisis Is Nothing AndAlso dtAnalisis.Rows.Count > 0 Then CostesIDVino.AnalisisDel = dtAnalisis.Copy
        'If Not dtAnalisisVariable Is Nothing AndAlso dtAnalisisVariable.Rows.Count > 0 Then CostesIDVino.AnalisisVariableDel = dtAnalisisVariable.Copy

        Return lstEliminar
    End Function

#End Region

#Region " Modificar Articulo / Depósito  OLD"


    '<Serializable()> _
    'Public Class StModificarArticulo
    '    Public IDVino As Guid
    '    Public ArticuloNew As String
    '    Public Origen As DataTable
    '    Public Destino As DataTable
    '    Public LoteNew As String
    '    Public BarricaNew As String

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDVino As Guid, ByVal ArticuloNew As String, ByVal Origen As DataTable, ByVal Destino As DataTable, _
    '                   Optional ByVal LoteNew As String = "", Optional ByVal BarricaNew As String = "")
    '        Me.IDVino = IDVino
    '        Me.ArticuloNew = ArticuloNew
    '        Me.Origen = Origen
    '        Me.Destino = Destino
    '        Me.LoteNew = LoteNew
    '        Me.BarricaNew = BarricaNew
    '    End Sub
    'End Class

    '<Task()> Public Shared Sub ModificarArticulo(ByVal data As StModificarArticulo, ByVal services As ServiceProvider)
    '    ProcessServer.ExecuteTask(Of StModificarArticulo)(AddressOf ModificarArticuloLote, data, services)
    'End Sub

    '<Task()> Public Shared Sub ModificarArticuloLote(ByVal data As StModificarArticulo, ByVal services As ServiceProvider)
    '    ProcessServer.ExecuteTask(Of StModificarArticulo)(AddressOf ModificarArticuloBarricaLote, data, services)
    'End Sub

    '<Task()> Public Shared Sub ModificarArticuloBarricaLote(ByVal data As StModificarArticulo, ByVal services As ServiceProvider)
    '    AdminData.BeginTx()
    '    Dim drVino As DataRow = New BdgVino().GetItemRow(data.IDVino)
    '    If drVino("IDArticulo") <> data.ArticuloNew Or drVino("Lote") <> data.LoteNew Then
    '        Dim StInfo As New StRestaurarInfo(drVino.Table, data.Destino)
    '        Dim info As ModificacionesInfo = ProcessServer.ExecuteTask(Of StRestaurarInfo, ModificacionesInfo)(AddressOf RestaurarInfo, StInfo, services)
    '        info.IDVino = data.IDVino
    '        info.ArticuloNew = data.ArticuloNew
    '        info.LoteNew = data.LoteNew
    '        info.BarricaNew = data.BarricaNew

    '        Dim StCrear As New BdgOperacion.StCrearOperacionInfo(data.Origen, data.Destino, info)
    '        ProcessServer.ExecuteTask(Of BdgOperacion.StCrearOperacionInfo)(AddressOf BdgOperacion.CrearOperacionInfo, StCrear, services)
    '    End If
    'End Sub

    '<Serializable()> _
    'Public Class StRestaurarInfo
    '    Public DtVino As DataTable
    '    Public Destino As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal DtVino As DataTable, ByVal Destino As DataTable)
    '        Me.DtVino = DtVino
    '        Me.Destino = Destino
    '    End Sub
    'End Class

    '<Task()> Public Shared Function RestaurarInfo(ByVal data As StRestaurarInfo, ByVal services As ServiceProvider) As ModificacionesInfo
    '    'Añadir aquí todo el cuerpo de ModificarDeposito que sea común con ModificarArticulo
    '    Dim info As New ModificacionesInfo
    '    Dim strNOperacion As String = data.Destino.Rows(0)("NOperacion")

    '    info.drOperacion = New BdgOperacion().GetItemRow(strNOperacion)

    '    Dim f As Filter = New Filter
    '    f.Add(New GuidFilterItem(_OV.IDVino, data.DtVino.Rows(0)("IDVino")))
    '    f.Add(New StringFilterItem(_OV.NOperacion, strNOperacion))

    '    info.dtMateriales = New BdgVinoMaterial().Filter(f)
    '    If Not info.dtMateriales Is Nothing AndAlso info.dtMateriales.Rows.Count > 0 Then
    '        Dim IDsVinoMaterial As List(Of Object) = (From c In info.dtMateriales Where c.RowState <> DataRowState.Deleted Select c("IDVinoMaterial") Distinct).ToList
    '        If Not IDsVinoMaterial Is Nothing AndAlso IDsVinoMaterial.Count > 0 Then
    '            Dim fMateriales As New Filter
    '            fMateriales.Add(New InListFilterItem("IDVinoMaterial", IDsVinoMaterial.ToArray, FilterType.Guid))
    '            info.dtMaterialesLotes = New BdgVinoMaterialLote().Filter(fMateriales)
    '        End If

    '    End If
    '    info.dtMod = New BdgVinoMod().Filter(f)
    '    info.dtMaquinas = New BdgVinoCentro().Filter(f)
    '    info.dtVarios = New BdgVinoVarios().Filter(f)
    '    info.dtAnalisis = New BE.DataEngine().Filter("frmBdgOperacionAnalisis", f)
    '    info.TipoOperacion = New BdgTipoOperacion().GetItemRow(info.drOperacion("IDTipoOperacion"))("TipoMovimiento")

    '    Dim ff As New Filter
    '    ff.Add(New GuidFilterItem(_OV.IDVino, data.DtVino.Rows(0)("IDVino")))
    '    ff.Add(New IsNullFilterItem("NOperacion"))
    '    info.dtElaboracion = New BdgVinoMaterial().Filter(ff)

    '    ff.Clear()
    '    Dim clsCVH As New BdgCosteVendimiaHist
    '    info.dtElaboracionHist = clsCVH.AddNew
    '    Dim IDVM As String = String.Empty
    '    For Each drE As DataRow In info.dtElaboracion.Select(Nothing, "IDVinoMaterial")
    '        If IDVM <> drE("IDVinoMaterial").ToString Then
    '            ff.Clear()
    '            ff.Add(New GuidFilterItem("IDVinoMaterial", FilterOperator.Equal, drE("IDVinoMaterial")))
    '            Dim dtHis As DataTable = clsCVH.Filter(ff)
    '            If Not IsNothing(dtHis) AndAlso dtHis.Rows.Count > 0 Then
    '                info.dtElaboracionHist.ImportRow(dtHis.Rows(0))
    '            End If
    '        End If
    '    Next

    '    ff.Clear()
    '    ff.Add(New GuidFilterItem(_OV.IDVino, data.DtVino.Rows(0)("IDVino")))
    '    ff.Add(New IsNullFilterItem("NOperacion"))
    '    info.dtCosteNave = New BdgVinoCentro().Filter(ff)

    '    Dim oVinoWC As New BdgWorkClass

    '    Dim DrFil() As DataRow = data.Destino.Select(f.Compose(New AdoFilterComposer))
    '    If DrFil.Length > 0 Then
    '        Dim DtOPVino As DataTable = data.Destino.Clone
    '        DtOPVino.TableName = "BdgOperacionVino"
    '        DtOPVino.ImportRow(DrFil(0))
    '        Dim ClsOPVino As New BdgOperacionVino
    '        ClsOPVino.Delete(DtOPVino)
    '    End If
    '    If info.TipoOperacion <> Business.Bodega.TipoMovimiento.SinMovimiento Then
    '        Dim StEj As New BdgWorkClass.StEjecutarMovimientosNumero(info.drOperacion("IDMovimiento"), strNOperacion, info.drOperacion("Fecha"))
    '        ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientosNumero)(AddressOf BdgWorkClass.EjecutarMovimientosNumero, StEj, services)
    '    End If
    '    Return info
    'End Function

    '<Serializable()> _
    'Public Class StModificarDeposito
    '    Public IDVino As Guid
    '    Public DepositoNew As String
    '    Public Origen As DataTable
    '    Public Destino As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDVino As Guid, ByVal DepositoNew As String, ByVal Origen As DataTable, ByVal Destino As DataTable)
    '        Me.IDVino = IDVino
    '        Me.DepositoNew = DepositoNew
    '        Me.Origen = Origen
    '        Me.Destino = Destino
    '    End Sub
    'End Class

    '<Task()> Public Shared Sub ModificarDeposito(ByVal data As StModificarDeposito, ByVal services As ServiceProvider)
    '    AdminData.BeginTx()
    '    Dim drVino As DataRow = New BdgVino().GetItemRow(data.IDVino)
    '    If drVino("IDDeposito") <> data.DepositoNew Then
    '        Dim StInfo As New StRestaurarInfo(drVino.Table, data.Destino)
    '        Dim info As ModificacionesInfo = ProcessServer.ExecuteTask(Of StRestaurarInfo, ModificacionesInfo)(AddressOf RestaurarInfo, StInfo, services)
    '        info.IDVino = data.IDVino
    '        info.DepositoNew = data.DepositoNew
    '        Dim StCrear As New BdgOperacion.StCrearOperacionInfo(data.Origen, data.Destino, info)
    '        ProcessServer.ExecuteTask(Of BdgOperacion.StCrearOperacionInfo)(AddressOf BdgOperacion.CrearOperacionInfo, StCrear, services)
    '    End If
    'End Sub

#End Region

#End Region

End Class

<Serializable()> _
Public Class ModificacionesInfo
    Public IDVino As Guid
    Public drOperacion As DataRow
    Public dtMateriales As DataTable
    Public dtMaterialesLotes As DataTable
    Public dtMod As DataTable
    Public dtMaquinas As DataTable
    Public dtVarios As DataTable
    Public dtAnalisis As DataTable
    Public dtElaboracion As DataTable
    Public dtElaboracionHist As DataTable
    Public dtCosteNave As DataTable
    Public dtTasaCentro As DataTable
    Public TipoOperacion As Business.Bodega.TipoMovimiento
    Public DepositoNew As String
    Public ArticuloNew As String
    Public LoteNew As String
    Public BarricaNew As String
End Class

<Serializable()> _
Public Class _OV
    Public Const NOperacion As String = "NOperacion"
    Public Const IDVino As String = "IDVino"
    Public Const Cantidad As String = "Cantidad"
    Public Const Destino As String = "Destino"
    Public Const Merma As String = "Merma"
    Public Const Ocupacion As String = "Ocupacion"
    Public Const IDEstructura As String = "IDEstructura"
    Public Const IDOrden As String = "IDOrden"
    Public Const QDeposito As String = "QDeposito"
    Public Const Litros As String = "Litros"
    Public Const IDBarrica As String = "IDBarrica"
    Public Const IDEstadoVino As String = "IDEstadoVino"
End Class