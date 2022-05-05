Public Class ObraParteAgrupado
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbObraParteAgrupado"

    <Task()> Public Shared Sub BorrarParteAgrupado(ByVal IDParteAgrupado As String, ByVal services As ServiceProvider)
        Dim o As New ObraParteAgrupado
        Dim dt As DataTable = o.SelOnPrimaryKey(IDParteAgrupado)
        If dt.Rows.Count > 0 Then o.Delete(dt)
    End Sub

#Region " RegisterAddnewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDParteAgrupado") = Guid.NewGuid 'AdminData.GetAutoNumeric
        data("RepartoManual") = 1
        Dim IDContador As String = ProcessServer.ExecuteTask(Of ContadorEntidad, String)(AddressOf CentroGestion.GetContadorPredeterminadoCGestionUsuario, ContadorEntidad.AlbaranVenta, services)
        data("NParteAgrupado") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, IDContador, services)
        data("Fecha") = Date.Today
    End Sub

#End Region

#Region " GuardarParteAgrupado "

    <Serializable()> _
    Public Class dataParteAgrupado
        Public CurrentData As DataTable
        Public IDFinca As Guid
        Public IDFincaHija As Guid
        Public IDZonaFinca As String
        Public LineasParteAgrupado As DataTable
        Public LineasParteAgrupadoObservatorio As DataTable
        Public LineasParteAgrupadoMAT As DataTable
        Public LineasParteAgrupadoMATLote As DataTable
        Public LineasParteAgrupadoMOD As DataTable
        Public LineasParteAgrupadoCEN As DataTable
        Public LineasParteAgrupadoVAR As DataTable
        Public LineasAnalisisValor As DataTable

        Public Sub New(ByVal CurrentData As DataTable)
            Me.CurrentData = CurrentData
        End Sub
    End Class
    <Task()> Public Shared Sub GuardarParteAgrupado(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider)
        If Not data Is Nothing AndAlso Not data.CurrentData Is Nothing Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ValidarDatosObligatoriosObraParteAgrupado, data.CurrentData.Rows(0), services)

            Dim Added As Boolean = (data.CurrentData.Rows(0).RowState = DataRowState.Added)

            data.CurrentData.TableName = "ObraParteAgrupado"
            ObraParteAgrupado.UpdateTable(data.CurrentData)
            If Added Then
                ProcessServer.ExecuteTask(Of dataParteAgrupado)(AddressOf GenerarLineasParteAgrupado, data, services)
            End If
            'Else
            ProcessServer.ExecuteTask(Of dataParteAgrupado)(AddressOf TratarPartesAgrupadosLineas, data, services)
            ProcessServer.ExecuteTask(Of dataParteAgrupado)(AddressOf TratarPartesAgrupadosObservatorios, data, services)
            ProcessServer.ExecuteTask(Of dataParteAgrupado)(AddressOf TratarPartesAgrupadosMateriales, data, services)
            ProcessServer.ExecuteTask(Of dataParteAgrupado)(AddressOf TratarPartesAgrupadosMaterialesLotes, data, services)
            ProcessServer.ExecuteTask(Of dataParteAgrupado)(AddressOf TratarPartesAgrupadosMod, data, services)
            ProcessServer.ExecuteTask(Of dataParteAgrupado)(AddressOf TratarPartesAgrupadosCentros, data, services)
            ProcessServer.ExecuteTask(Of dataParteAgrupado)(AddressOf TratarPartesAgrupadosVarios, data, services)
            ProcessServer.ExecuteTask(Of dataParteAgrupado)(AddressOf TratarPartesAgrupadosAnalisisLineaValor, data, services)
            'End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatoriosObraParteAgrupado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("NParteAgrupado")) = 0 Then ApplicationService.GenerateError("El Nº Parte es un dato obligatorio.")
        If Length(data("Fecha")) = 0 Then ApplicationService.GenerateError("La Fecha es un dato obligatorio.")
        If Length(data("IDSubtipoTrabajo")) = 0 Then ApplicationService.GenerateError("El Tratamiento es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub GenerarLineasParteAgrupado(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider)
        If ((Length(data.IDFinca) = 0 OrElse data.IDFinca = Guid.Empty) AndAlso (Length(data.IDFincaHija) = 0 OrElse data.IDFinca = Guid.Empty) AndAlso Length(data.IDZonaFinca) = 0) Then Return
        Dim dtLineas As DataTable = ProcessServer.ExecuteTask(Of dataParteAgrupado, DataTable)(AddressOf SimulacionLineasParteAgrupado, data, services)
        If dtLineas.Rows.Count > 0 Then ObraParteAgrupadoLinea.UpdateTable(dtLineas)
    End Sub

    <Task()> Public Shared Function SimulacionLineasParteAgrupado(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider) As DataTable
        Dim f As New Filter
        If Not data.IDFinca.Equals(Guid.Empty) Then
            Dim f_OR_3 As New Filter(FilterUnionOperator.Or)
            Dim dtHijos As DataTable = New BdgFinca().Filter(New GuidFilterItem("IDFincaPadre", data.IDFinca), , "IDFinca")
            If dtHijos.Rows.Count > 0 Then
                For Each drHijo As DataRow In dtHijos.Rows
                    f_OR_3.Add(New GuidFilterItem("IDFinca", drHijo("IDFinca")))
                Next
                f.Add(f_OR_3)
            End If
            If Not data.IDFincaHija.Equals(Guid.Empty) Then f.Add(New GuidFilterItem("IDFinca", data.IDFincaHija))
        End If

        If Not Length(data.IDZonaFinca) = 0 Then f.Add(New StringFilterItem("IDZonaFinca", data.IDZonaFinca))
        f.Add("FincaTrabajo", True)

        Dim dtLineas As DataTable = New ObraParteAgrupadoLinea().AddNew
        Dim dtFinca As DataTable = New BdgFinca().Filter(f)
        If dtFinca.Rows.Count > 0 Then
            For Each drFinca As DataRow In dtFinca.Rows
                Dim drLinea As DataRow = dtLineas.NewRow
                drLinea("IDParteAgrupadoLinea") = Guid.NewGuid
                drLinea("IDParteAgrupado") = data.CurrentData.Rows(0)("IDParteAgrupado")
                drLinea("IDFinca") = drFinca("IDFinca")

                f.Clear()
                f.Add(New GuidFilterItem("IDFinca", drLinea("IDFinca")))
                f.Add(New StringFilterItem("IDTipoObra", data.CurrentData.Rows(0)("IDTipoObra")))
                f.Add(New StringFilterItem("IDTipoTrabajo", data.CurrentData.Rows(0)("IDTipoTrabajo")))
                f.Add(New StringFilterItem("IDSubTipoTrabajo", data.CurrentData.Rows(0)("IDSubTipoTrabajo")))
                Dim dt As DataTable = New BE.DataEngine().Filter("vBdgNegDatosFincaObraCampaña", f)
                If dt.Rows.Count > 0 Then
                    drLinea("IDTrabajo") = dt.Rows(0)("IDTrabajo")
                    Dim dtt As DataTable = New ObraTrabajo().Filter(New NumberFilterItem("IDTrabajoPadre", drLinea("IDTrabajo")))
                    If dtt.Rows.Count > 0 Then
                        drLinea("IDTrabajoGenerado") = dtt.Rows(0)("IDTrabajo")
                        drLinea("IDTrabajo") = dtt.Rows(0)("IDTrabajo")
                    End If
                End If
                dtLineas.Rows.Add(drLinea.ItemArray)
            Next
        End If

        Return dtLineas
    End Function

    <Task()> Public Shared Function PropuestaLineasParteAgrupado(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider) As DataTable

        Dim dtLineas As DataTable = New BE.DataEngine().Filter("vBdgMntoObraParteAgrupadoLineas", New NoRowsFilterItem)
        'dtLineas.Columns.Add("ID", GetType(Integer))
        'dtLineas.Columns.Add("Marca", GetType(Boolean))

        If ((Length(data.IDFinca) = 0 OrElse data.IDFinca = Guid.Empty) AndAlso (Length(data.IDFincaHija) = 0 OrElse data.IDFinca = Guid.Empty) AndAlso Length(data.IDZonaFinca) = 0) Then Return dtLineas

        Dim dt As DataTable = ProcessServer.ExecuteTask(Of dataParteAgrupado, DataTable)(AddressOf SimulacionLineasParteAgrupado, data, services)
        If dt.Rows.Count > 0 Then
            Dim i As Integer = 1
            For Each dr As DataRow In dt.Rows
                Dim drLinea As DataRow = dtLineas.NewRow
                drLinea("Marca") = False
                'drLinea("ID") = i
                drLinea("IDParteAgrupadoLinea") = dr("IDParteAgrupadoLinea")
                drLinea("IDParteAgrupado") = dr("IDParteAgrupado")
                drLinea("IDFinca") = dr("IDFinca")

                drLinea = New ObraParteAgrupadoLinea().ApplyBusinessRule("IDFinca", drLinea("IDFinca"), drLinea)
                'i += 1
                dtLineas.Rows.Add(drLinea.ItemArray)
            Next
        End If

        Return dtLineas
    End Function

    <Task()> Public Shared Sub TratarPartesAgrupadosLineas(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider)
        If Not data.LineasParteAgrupado Is Nothing Then
            Dim dtDelete As DataTable = data.LineasParteAgrupado.GetChanges(DataRowState.Deleted)
            If Not dtDelete Is Nothing Then ProcessServer.ExecuteTask(Of DataTable)(AddressOf BorrarLineasPartesAgrupados, dtDelete, services)

            Dim dtUpdate As DataTable = data.LineasParteAgrupado.GetChanges(DataRowState.Modified Or DataRowState.Added)
            If Not dtUpdate Is Nothing Then
                Dim o As New ObraParteAgrupadoLinea
                dtUpdate.TableName = "ObraParteAgrupadoLinea"
                o.Update(dtUpdate)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub TratarPartesAgrupadosObservatorios(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider)
        If Not data.LineasParteAgrupadoObservatorio Is Nothing Then
            Dim dtDelete As DataTable = data.LineasParteAgrupadoObservatorio.GetChanges(DataRowState.Deleted)
            If Not dtDelete Is Nothing Then ProcessServer.ExecuteTask(Of DataTable)(AddressOf BorrarLineasPartesAgrupadosObservatorios, dtDelete, services)

            Dim dtUpdate As DataTable = data.LineasParteAgrupadoObservatorio.GetChanges(DataRowState.Modified Or DataRowState.Added)
            If Not dtUpdate Is Nothing Then
                For Each DrUpd As DataRow In dtUpdate.Select(String.Empty, String.Empty, DataViewRowState.Added)
                    DrUpd("IDParteAgrupadoObservatorio") = Guid.NewGuid
                Next
                Dim o As New ObraParteAgrupadoObservatorio
                dtUpdate.TableName = "ObraParteAgrupadoObservatorio"
                o.Update(dtUpdate)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub BorrarLineasPartesAgrupados(ByVal data As DataTable, ByVal services As ServiceProvider)
        Dim dtDelete As DataTable = data.GetChanges(DataRowState.Deleted)
        If Not dtDelete Is Nothing AndAlso dtDelete.Rows.Count > 0 Then
            dtDelete.RejectChanges()
            Dim o As New ObraParteAgrupadoLinea
            dtDelete.TableName = GetType(ObraParteAgrupadoLinea).Name
            o.Delete(dtDelete)
        End If
    End Sub

    <Task()> Public Shared Sub BorrarLineasPartesAgrupadosObservatorios(ByVal data As DataTable, ByVal services As ServiceProvider)
        Dim dtDelete As DataTable = data.GetChanges(DataRowState.Deleted)
        If Not dtDelete Is Nothing AndAlso dtDelete.Rows.Count > 0 Then
            dtDelete.RejectChanges()
            Dim o As New ObraParteAgrupadoObservatorio
            dtDelete.TableName = GetType(ObraParteAgrupadoObservatorio).Name
            o.Delete(dtDelete)
        End If
    End Sub

    <Task()> Public Shared Sub TratarPartesAgrupadosMateriales(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider)
        If Not data.LineasParteAgrupadoMAT Is Nothing Then
            Dim dtDelete As DataTable = data.LineasParteAgrupadoMAT.GetChanges(DataRowState.Deleted)
            If Not dtDelete Is Nothing Then ProcessServer.ExecuteTask(Of DataTable)(AddressOf BorrarMaterialesPartesAgrupados, dtDelete, services)

            Dim dtUpdate As DataTable = data.LineasParteAgrupadoMAT.GetChanges(DataRowState.Modified Or DataRowState.Added)
            If Not dtUpdate Is Nothing Then
                Dim o As New ObraParteAgrupadoMaterial
                dtUpdate.TableName = "ObraParteAgrupadoMaterial"
                o.Update(dtUpdate)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub TratarPartesAgrupadosMaterialesLotes(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider)
        If Not data.LineasParteAgrupadoMATLote Is Nothing Then
            Dim dtLotes As DataTable = data.LineasParteAgrupadoMATLote.GetChanges(DataRowState.Modified Or DataRowState.Added)
            If Not dtLotes Is Nothing Then
                Dim o As New ObraParteAgrupadoMaterialLote
                dtLotes.TableName = "ObraParteAgrupadoMaterialLote"
                o.Update(dtLotes)
            End If
        End If
    End Sub


    <Task()> Public Shared Sub BorrarMaterialesPartesAgrupados(ByVal data As DataTable, ByVal services As ServiceProvider)
        Dim dtDelete As DataTable = data.GetChanges(DataRowState.Deleted)
        If Not dtDelete Is Nothing AndAlso dtDelete.Rows.Count > 0 Then
            dtDelete.RejectChanges()
            Dim o As New ObraParteAgrupadoMaterial
            dtDelete.TableName = GetType(ObraParteAgrupadoMaterial).Name
            o.Delete(dtDelete)
        End If
    End Sub

    <Task()> Public Shared Sub TratarPartesAgrupadosMod(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider)
        If Not data.LineasParteAgrupadoMOD Is Nothing Then
            Dim dtDelete As DataTable = data.LineasParteAgrupadoMOD.GetChanges(DataRowState.Deleted)
            If Not dtDelete Is Nothing Then ProcessServer.ExecuteTask(Of DataTable)(AddressOf BorrarModPartesAgrupados, dtDelete, services)

            Dim dtUpdate As DataTable = data.LineasParteAgrupadoMOD.GetChanges(DataRowState.Modified Or DataRowState.Added)
            If Not dtUpdate Is Nothing Then
                Dim o As New ObraParteAgrupadoMod
                dtUpdate.TableName = "ObraParteAgrupadoMod"
                o.Validate(dtUpdate)
                o.Update(dtUpdate)
            End If
        End If
    End Sub
    <Task()> Public Shared Sub BorrarModPartesAgrupados(ByVal data As DataTable, ByVal services As ServiceProvider)
        Dim dtDelete As DataTable = data.GetChanges(DataRowState.Deleted)
        If Not dtDelete Is Nothing AndAlso dtDelete.Rows.Count > 0 Then
            dtDelete.RejectChanges()
            Dim o As New ObraParteAgrupadoMod
            dtDelete.TableName = GetType(ObraParteAgrupadoMod).Name
            o.Delete(dtDelete)
        End If
    End Sub

    <Task()> Public Shared Sub TratarPartesAgrupadosCentros(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider)
        If Not data.LineasParteAgrupadoCEN Is Nothing Then
            Dim dtDelete As DataTable = data.LineasParteAgrupadoCEN.GetChanges(DataRowState.Deleted)
            If Not dtDelete Is Nothing Then ProcessServer.ExecuteTask(Of DataTable)(AddressOf BorrarCentrosPartesAgrupados, dtDelete, services)

            Dim dtUpdate As DataTable = data.LineasParteAgrupadoCEN.GetChanges(DataRowState.Modified Or DataRowState.Added)
            If Not dtUpdate Is Nothing Then
                Dim o As New ObraParteAgrupadoCentro
                dtUpdate.TableName = "ObraParteAgrupadoCentro"
                o.Update(dtUpdate)
            End If
        End If
    End Sub
    <Task()> Public Shared Sub BorrarCentrosPartesAgrupados(ByVal data As DataTable, ByVal services As ServiceProvider)
        Dim dtDelete As DataTable = data.GetChanges(DataRowState.Deleted)
        If Not dtDelete Is Nothing AndAlso dtDelete.Rows.Count > 0 Then
            dtDelete.RejectChanges()
            Dim o As New ObraParteAgrupadoCentro
            dtDelete.TableName = GetType(ObraParteAgrupadoCentro).Name
            o.Delete(dtDelete)
        End If
    End Sub

    <Task()> Public Shared Sub TratarPartesAgrupadosVarios(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider)
        If Not data.LineasParteAgrupadoVAR Is Nothing Then
            Dim dtDelete As DataTable = data.LineasParteAgrupadoVAR.GetChanges(DataRowState.Deleted)
            If Not dtDelete Is Nothing Then ProcessServer.ExecuteTask(Of DataTable)(AddressOf BorrarVariosPartesAgrupados, dtDelete, services)

            Dim dtUpdate As DataTable = data.LineasParteAgrupadoVAR.GetChanges(DataRowState.Modified Or DataRowState.Added)
            If Not dtUpdate Is Nothing Then
                Dim o As New ObraParteAgrupadoVarios
                dtUpdate.TableName = "ObraParteAgrupadoVarios"
                o.Update(dtUpdate)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub TratarPartesAgrupadosAnalisisLineaValor(ByVal data As dataParteAgrupado, ByVal services As ServiceProvider)
        If Not data.LineasAnalisisValor Is Nothing Then
            Dim dtUpdate As DataTable = data.LineasAnalisisValor.GetChanges(DataRowState.Modified)
            If Not dtUpdate Is Nothing Then
                Dim o As New BdgAnalisisLineaValor
                dtUpdate.TableName = "BdgAnalisisLineaValor"
                o.Update(dtUpdate)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub BorrarVariosPartesAgrupados(ByVal data As DataTable, ByVal services As ServiceProvider)
        Dim dtDelete As DataTable = data.GetChanges(DataRowState.Deleted)
        If Not dtDelete Is Nothing AndAlso dtDelete.Rows.Count > 0 Then
            dtDelete.RejectChanges()
            Dim o As New ObraParteAgrupadoVarios
            dtDelete.TableName = GetType(ObraParteAgrupadoVarios).Name
            o.Delete(dtDelete)
        End If
    End Sub

#End Region

#Region " Generar Trabajos "

    <Serializable()> _
    Public Class dataDatosParteAgrupado
        Public IDParteAgrupado As Guid
        Public dtLineas As DataTable
        Public dtMateriales As DataTable
        Public dtMod As DataTable
        Public dtCentros As DataTable
        Public dtVarios As DataTable
        Public FechaParte As Date
        Public IDAnalisis As String
        Friend GenerarPartesEnFincaPadre As Boolean

        Public Sub New(ByVal IDParteAgrupado As Guid, ByVal dtLineasAgrup As DataTable, ByVal IDAnalisis As String)
            Me.IDParteAgrupado = IDParteAgrupado
            Me.dtLineas = dtLineasAgrup
            Me.FechaParte = Date.Today
            If Length(IDAnalisis) > 0 Then Me.IDAnalisis = IDAnalisis
        End Sub

        Public Sub New(ByVal IDParteAgrupado As Guid, ByVal dtLineasAgrup As DataTable, ByVal FechaParte As Date, ByVal IDAnalisis As String)
            Me.IDParteAgrupado = IDParteAgrupado
            Me.dtLineas = dtLineasAgrup
            Me.FechaParte = FechaParte
            If Length(IDAnalisis) > 0 Then Me.IDAnalisis = IDAnalisis
        End Sub
    End Class
    <Task()> Public Shared Sub GenerarPartesTrabajos(ByVal data As dataDatosParteAgrupado, ByVal services As ServiceProvider)
        If Not data.dtLineas Is Nothing AndAlso data.dtLineas.Rows.Count > 0 Then
            AdminData.BeginTx()

            Dim bsnParteAgrupado As New ObraParteAgrupado()
            Dim bsnParam As New BdgParametro
            Dim blnAutoactualizacionStocks As Boolean = bsnParam.AutoactualizacionStocksPartesAgrupados()
            Dim bsnOT As New ObraTrabajo
            data.GenerarPartesEnFincaPadre = bsnParam.GestionPorFincaPadre()
            Dim Where As String = New IsNullFilterItem("IDFincaPadre", data.GenerarPartesEnFincaPadre).Compose(New AdoFilterComposer)
            data.dtLineas.Columns.Add("FechaInicio", GetType(Date))
            Dim dtrParte As DataRow = bsnParteAgrupado.GetItemRow(data.IDParteAgrupado)

            Dim datainversion As New ObraSubtipoTrabajo.stObtenerTrabajoInversion
            datainversion.IDTipoObra = dtrParte("IDTipoObra")
            datainversion.IDTipoTrabajo = dtrParte("IdTipoTrabajo")
            datainversion.IDSubtipoTrabajo = dtrParte("IDSubtipoTrabajo")
            If Length(dtrParte("Fecha")) > 0 Then
                datainversion.FechaInicioTrabajo = dtrParte("Fecha")
            Else
                datainversion.FechaInicioTrabajo = Today
            End If
            datainversion.FechaFinTrabajo = Today

            Dim blnInversion As Boolean = ProcessServer.ExecuteTask(Of ObraSubtipoTrabajo.stObtenerTrabajoInversion, Boolean) _
                                            (AddressOf ObraSubtipoTrabajo.ObtenerTrabajoEsInversion, datainversion, services)
            'For Each drLinea As DataRow In data.dtLineas.Select(Where, "IDObra")
            Dim dtrLineas() As DataRow = data.dtLineas.Select()
            If dtrLineas Is Nothing OrElse dtrLineas.Length = 0 Then Return


            'todas las fincas incluidas deben tener obra y obra campaña
            Dim fValidacion As New Filter
            fValidacion.Add(New IsNullFilterItem("IDObra", False))
            fValidacion.Add(New IsNullFilterItem("IDObraCampaña", False))
            Dim dtrLineasCount() As DataRow = data.dtLineas.Select(fValidacion.Compose(New AdoFilterComposer))
            If dtrLineasCount.Length <> dtrLineas.Length Then
                ApplicationService.GenerateError("Todas las fincas seleccionadas deben tener obra y obra de campaña asignada.")
            End If

            For Each drLinea As DataRow In dtrLineas
                Dim dtObra As DataTable = New BE.DataEngine().Filter("vBdgNegDatosFincaObraCampaña", New GuidFilterItem("IDFinca", drLinea("IDFinca")))
                If dtObra.Rows.Count > 0 AndAlso Length(dtObra.Rows(0)("IDObra")) > 0 Then
                    'LA OBRA A ELEGIR ES LA OBRA PADRE FINCA SI:
                    ' - EL SUBTIPO DE TRABAJO ELEGIDO ES INVERSIÓN
                    ' - LA FINCA ES DE INVERSIÓN PARA LA FECHA ACTUAL

                    If blnInversion Then
                        drLinea("IDObra") = dtObra.Rows(0)("IDObraPadre")
                    Else
                        drLinea("IDObra") = dtObra.Rows(0)("IDObra")
                    End If


                    If Length(drLinea("IDParteAgrupado")) = 0 Then drLinea("IDParteAgrupado") = data.IDParteAgrupado
                    drLinea("FechaInicio") = data.FechaParte

                    If Length(drLinea("IDTrabajoPadre")) > 0 Then
                        drLinea("IDTrabajoGenerado") = drLinea("IDTrabajo")
                        drLinea("IDTrabajo") = drLinea("IDTrabajoPadre")
                    End If

                    'trabajo
                    '1.	Buscar el trabajo de primer nivel que coincida con IDObra (la que corresponda) y el IDTipoObra / IDTipoTrabajo / IDSubtipoTrabajo seleccionado en el parte agrupado.
                    '    a.	Si no existe existe, se crea
                    '2.	Buscar el trabajo de siguiente nivel para la fecha del parte agrupado
                    '    a.	Sino existe, se crea, estableciéndole como padre el obtenido en el punto 1
                    '3.	Generar imputaciones a ese trabajo obtenido en el punto 2
                    If Length(drLinea("IDTrabajo")) = 0 Then
                        Dim f As New Filter
                        f.Add(New IsNullFilterItem("IDTrabajoPadre"))
                        f.Add("IDObra", drLinea("IDObra"))
                        f.Add("IDTipoObra", dtrParte("IDTipoObra"))
                        f.Add("IDTipoTrabajo", dtrParte("IDTipoTrabajo"))
                        If Length(dtrParte("IDSubtipoTrabajo")) > 0 Then f.Add("IDSubtipoTrabajo", dtrParte("IDSubtipoTrabajo"))
                        Dim dttTrabajo As DataTable = bsnOT.Filter(f)
                        If (dttTrabajo Is Nothing OrElse dttTrabajo.Rows.Count = 0) Then
                            drLinea("IDTrabajo") = ProcessServer.ExecuteTask(Of DataRow, Integer)(AddressOf GenerarTrabajo, drLinea, services)
                        Else
                            drLinea("IDTrabajo") = dttTrabajo.Rows(0)("IDTrabajo")
                            drLinea("CodTrabajo") = dttTrabajo.Rows(0)("CodTrabajo")
                        End If
                    End If

                    If Length(drLinea("IDTrabajoGenerado")) = 0 Then 'OrElse (Nz(drLinea("Estado"), enumotEstado.otTerminado) = enumotEstado.otTerminado) Then
                        Dim f As New Filter
                        f.Add("Estado", FilterOperator.NotEqual, enumotEstado.otTerminado) 'también puede ser que el trabajo esté finalizado?
                        f.Add("IDTrabajoPadre", drLinea("IDTrabajo"))
                        f.Add("IDObra", drLinea("IDObra"))
                        f.Add("IDTipoObra", dtrParte("IDTipoObra"))
                        f.Add("IDTipoTrabajo", dtrParte("IDTipoTrabajo"))
                        f.Add("IDSubtipoTrabajo", dtrParte("IDSubtipoTrabajo"))
                        Dim dttTrabajo As DataTable = bsnOT.Filter(f)
                        If (dttTrabajo Is Nothing OrElse dttTrabajo.Rows.Count = 0) Then
                            If Length(drLinea("CodTrabajo")) = 0 Then
                                Dim dtrTrabajoAux As DataRow = bsnOT.GetItemRow(drLinea("IDTrabajo"))
                                If Not dtrTrabajoAux Is Nothing Then
                                    drLinea("CodTrabajo") = dtrTrabajoAux("CodTrabajo") & String.Empty
                                End If
                            End If
                            drLinea("IDTrabajoGenerado") = ProcessServer.ExecuteTask(Of DataRow, Integer)(AddressOf GenerarTrabajo, drLinea, services)
                        Else
                            drLinea("IDTrabajoGenerado") = dttTrabajo.Rows(0)("IDTrabajo")
                        End If

                    End If

                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf TratarParteAgrupadoLinea, drLinea, services)

                    Dim dataControl As New dataGenerarControl(drLinea("IDObra"), drLinea("IDTrabajoGenerado"), drLinea("IDTrabajo"), _
                                                              drLinea("Porcentaje"), data.IDParteAgrupado, data.FechaParte)
                    dataControl.dtConcepto = data.dtMateriales
                    ProcessServer.ExecuteTask(Of dataGenerarControl)(AddressOf GenerarObraMaterialControl, dataControl, services)
                    dataControl.dtConcepto = data.dtMod
                    ProcessServer.ExecuteTask(Of dataGenerarControl)(AddressOf GenerarObraModControl, dataControl, services)
                    dataControl.dtConcepto = data.dtCentros
                    ProcessServer.ExecuteTask(Of dataGenerarControl)(AddressOf GenerarObraCentrosControl, dataControl, services)
                    dataControl.dtConcepto = data.dtVarios
                    ProcessServer.ExecuteTask(Of dataGenerarControl)(AddressOf GenerarObraVariosControl, dataControl, services)

                    ProcessServer.ExecuteTask(Of Integer)(AddressOf ObraCabecera.RecalcularObra, drLinea("IDObra"), services)

                    ''actualizar el stock si lo marca el parámetro
                    If blnAutoactualizacionStocks Then
                        Dim f As New Filter
                        'f.Add("Actualizado", 0)
                        f.Add("IDTrabajo", drLinea("IDTrabajoGenerado"))
                        f.Add("IDParteAgrupado", drLinea("IDParteAgrupado"))
                        'f.Add(New IsNullFilterItem("IDParteAgrupadoMat", False)) 'TODO => SÓLO ESCOGER LOS QUE SE GENEREN EN ESTE BUCLE

                        'Dim dttMat As DataTable = New ObraMaterialControl().Filter(f) 'cogemos los mat generados para este trabajo
                        Dim dttMat As DataTable = New DataEngine().Filter("NegBdgActualizacionStocksPartesAgrupados", f)
                        ProcessServer.ExecuteTask(Of DataTable)(AddressOf ActualizacionStocksParteAgrupado, dttMat, services)
                    End If

                    'cerrar el trabajo si es necesario
                    If (Nz(drLinea("CerrarLinea"), False)) Then
                        Dim dttTrabajo As DataTable = bsnOT.SelOnPrimaryKey(drLinea("IDTrabajoGenerado"))
                        If Not dttTrabajo Is Nothing AndAlso dttTrabajo.Rows.Count > 0 Then
                            dttTrabajo.Rows(0)("Estado") = enumotEstado.otTerminado
                            bsnOT.Update(dttTrabajo)
                        End If
                    End If
                End If
            Next
            'marcar el parte como generado
            Dim dttParte As DataTable = bsnParteAgrupado.SelOnPrimaryKey(data.IDParteAgrupado)
            If Not dttParte Is Nothing AndAlso dttParte.Rows.Count > 0 Then
                dttParte.Rows(0)("Estado") = enumBdgEstadoParteAgrupado.TrabajoGenerado
                bsnParteAgrupado.Update(dttParte)
            End If

        End If
    End Sub

    <Serializable()> _
    Public Class DataGenerarAnaliticaFinca
        Public IDParteAgrupadoLinea As Guid
        Public IDAnalisis As String
        Public Fecha As Date
        Public IDFinca As Guid

        Public Sub New(ByVal IDParteAgrupadoLinea As Guid, ByVal IDAnalisis As String, ByVal Fecha As Date, ByVal IDFinca As Guid)
            Me.IDParteAgrupadoLinea = IDParteAgrupadoLinea
            Me.IDAnalisis = IDAnalisis
            Me.Fecha = Fecha
            Me.IDFinca = IDFinca
        End Sub
    End Class

    <Task()> Public Shared Sub GenerarAnaliticaFinca(ByVal data As DataGenerarAnaliticaFinca, ByVal services As ServiceProvider)
        If Length(data.IDAnalisis) > 0 Then
            Dim dtA As DataTable = New BdgAnalisis().SelOnPrimaryKey(data.IDAnalisis)
            If Not dtA Is Nothing AndAlso dtA.Rows.Count > 0 Then

                Dim ClsBdgAC As New BdgAnalisisCabecera
                Dim dtBdgAC As DataTable = ClsBdgAC.AddNewForm
                dtBdgAC.Rows(0)("IDAnalisis") = data.IDAnalisis
                dtBdgAC.Rows(0)("Fecha") = data.Fecha
                dtBdgAC.Rows(0)("Estado") = enumBdgEstadoAnalisisGen.Solicitado
                dtBdgAC.Rows(0)("TipoAnalisis") = enumBdgTipoAnalisisCabGen.Finca
                dtBdgAC.Rows(0)("Valor") = data.IDFinca.ToString

                Dim ClsBdgALV As BdgAnalisisLineaValor
                Dim dtBdgALV As DataTable
                Dim dttVars As DataTable = New BE.DataEngine().Filter("tbBdgAnalisisVariable", New StringFilterItem("IDAnalisis", FilterOperator.Equal, data.IDAnalisis), , "IDVariable")
                If Not dttVars Is Nothing AndAlso dttVars.Rows.Count > 0 Then
                    ClsBdgALV = New BdgAnalisisLineaValor
                    dtBdgALV = ClsBdgALV.AddNew
                    For Each drBdgALV As DataRow In dttVars.Rows
                        Dim drNew As DataRow = dtBdgALV.NewRow
                        drNew("IDAnalisisCabecera") = dtBdgAC.Rows(0)("IDAnalisisCabecera")
                        drNew("IDVariable") = drBdgALV("IDVariable")
                        drNew("Valor") = String.Empty
                        drNew("ValorNumerico") = 0
                        drNew("Orden") = drBdgALV("Orden")
                        dtBdgALV.Rows.Add(drNew)
                    Next
                End If
                ClsBdgAC.Update(dtBdgAC)
                If Not dtBdgALV Is Nothing AndAlso dtBdgALV.Rows.Count > 0 Then ClsBdgALV.Update(dtBdgALV)

                Dim dt As DataTable = New ObraParteAgrupadoLinea().SelOnPrimaryKey(data.IDParteAgrupadoLinea)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    dt.Rows(0)("IDAnalisisCabecera") = dtBdgAC.Rows(0)("IDAnalisisCabecera")
                    dt.TableName = "ObraParteAgrupadoLinea"
                    BusinessHelper.UpdateTable(dt)
                End If
            End If
        End If

    End Sub

    <Serializable()> _
    Public Class DataGenerarAnaliticaObservatorio
        Public IDParteAgrupadoObservatorio As Guid
        Public IDAnalisis As String
        Public Fecha As Date
        Public IDObservatorio As String

        Public Sub New(ByVal IDParteAgrupadoObservatorio As Guid, ByVal IDAnalisis As String, ByVal Fecha As Date, ByVal IDObservatorio As String)
            Me.IDParteAgrupadoObservatorio = IDParteAgrupadoObservatorio
            Me.IDAnalisis = IDAnalisis
            Me.Fecha = Fecha
            Me.IDObservatorio = IDObservatorio
        End Sub
    End Class

    <Task()> Public Shared Sub GenerarAnaliticaObservatorio(ByVal data As DataGenerarAnaliticaObservatorio, ByVal services As ServiceProvider)
        If Length(data.IDAnalisis) > 0 Then
            Dim dtA As DataTable = New BdgAnalisis().SelOnPrimaryKey(data.IDAnalisis)
            If Not dtA Is Nothing AndAlso dtA.Rows.Count > 0 Then

                Dim ClsBdgAC As New BdgAnalisisCabecera
                Dim dtBdgAC As DataTable = ClsBdgAC.AddNewForm
                dtBdgAC.Rows(0)("IDAnalisis") = data.IDAnalisis
                dtBdgAC.Rows(0)("Fecha") = data.Fecha
                dtBdgAC.Rows(0)("TipoAnalisis") = enumBdgTipoAnalisisCabGen.Observatorio
                dtBdgAC.Rows(0)("Valor") = data.IDObservatorio

                Dim ClsBdgALV As BdgAnalisisLineaValor
                Dim dtBdgALV As DataTable
                Dim dttVars As DataTable = New BE.DataEngine().Filter("tbBdgAnalisisVariable", New StringFilterItem("IDAnalisis", FilterOperator.Equal, data.IDAnalisis), , "IDVariable")
                If Not dttVars Is Nothing AndAlso dttVars.Rows.Count > 0 Then
                    ClsBdgALV = New BdgAnalisisLineaValor
                    dtBdgALV = ClsBdgALV.AddNew
                    For Each drBdgALV As DataRow In dttVars.Rows
                        Dim drNew As DataRow = dtBdgALV.NewRow
                        drNew("IDAnalisisCabecera") = dtBdgAC.Rows(0)("IDAnalisisCabecera")
                        drNew("IDVariable") = drBdgALV("IDVariable")
                        drNew("Valor") = String.Empty
                        drNew("ValorNumerico") = 0
                        drNew("Orden") = drBdgALV("Orden")
                        dtBdgALV.Rows.Add(drNew)
                    Next
                End If
                ClsBdgAC.Update(dtBdgAC)
                If Not dtBdgALV Is Nothing AndAlso dtBdgALV.Rows.Count > 0 Then ClsBdgALV.Update(dtBdgALV)

                Dim ClsObraParte As New ObraParteAgrupadoObservatorio
                Dim dt As DataTable = ClsObraParte.SelOnPrimaryKey(data.IDParteAgrupadoObservatorio)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    dt.Rows(0)("IDAnalisisCabecera") = dtBdgAC.Rows(0)("IDAnalisisCabecera")
                    ClsObraParte.Update(dt)
                End If
            End If
        End If
    End Sub


    <Task()> Public Shared Sub ActualizacionStocksParteAgrupado(ByVal data As DataTable, ByVal services As ServiceProvider)
        If data Is Nothing OrElse data.Rows.Count = 0 Then Return

        Dim dt As New DataTable
        dt.Columns.Add("IDLineaMaterialControl", GetType(Integer))
        dt.Columns.Add("QStock", GetType(Double))
        dt.Columns.Add("QPendiente", GetType(Double))
        dt.Columns.Add("Lotes", GetType(String))
        dt.Columns.Add("Series", GetType(String))
        dt.Columns.Add("NObra", GetType(String))
        dt.Columns.Add("Fecha", GetType(Date))
        dt.RemotingFormat = SerializationFormat.Binary


        Dim de As New DataEngine
        For Each dtrMat As DataRow In data.Rows
            Dim dr As DataRow = dt.NewRow
            dr("IDLineaMaterialControl") = dtrMat("IDLineaMaterialControl")
            dr("QStock") = dtrMat("QReal") - dtrMat("QActualizado")
            dr("QPendiente") = 0
            Dim f As New Filter
            f.Add("IDLineaMaterialControl", dtrMat("IDLineaMaterialControl"))
            f.Add("Lote", dtrMat("Lote"))
            Dim dtLotes As DataTable = de.Filter("NegBdgObraParteAgrupadoLote", f) 'bsnOPAML.Filter(f) 'mLotes.LoteCollection.GetDataTable(drMarcados("IDLineaMaterialControl"))
            If Not dtLotes Is Nothing AndAlso dtLotes.Rows.Count > 0 Then
                dr("Lotes") = DataTableToXml(dtLotes)
            End If

            'Dim DtSeries As DataTable = mSeries.SerieCollection.GetDataTable(drMarcados("IDLineaMaterialControl"))
            'If Not DtSeries Is Nothing AndAlso DtSeries.Rows.Count > 0 Then
            '    dr("Series") = DataTableToXml(DtSeries)
            'End If

            dr("NObra") = dtrMat("NObra")

            dr("Fecha") = dtrMat("Fecha")
            dt.Rows.Add(dr)
        Next
        Dim log() As StockUpdateData = ProcessServer.ExecuteTask(Of DataTable, StockUpdateData())(AddressOf ActualizacionStocks.ActualizarStocks, dt, services)
        'actualizar el campo stocks a posteriori
        Dim bsnOMCL As New ObraMaterialControlLote
        For Each dtrMat As DataRow In data.Rows
            If (Length("Lote") > 0) Then
                Dim f As New Filter
                f.Add("IDLineaMaterialControl", dtrMat("IDLineaMaterialControl"))
                f.Add("Lote", dtrMat("Lote"))
                Dim dttOMCL As DataTable = bsnOMCL.Filter(f)
                If Not dttOMCL Is Nothing AndAlso dttOMCL.Rows.Count > 0 Then
                    dttOMCL.Rows(0)("IDParteAgrupadoMatLote") = dtrMat("IDParteAgrupadoMatLote")
                End If
                bsnOMCL.Update(dttOMCL)
            End If
        Next


    End Sub

    <Serializable()> _
    Public Class dataDatosImportacionPT
        Public CFinca As String
        Public IDOperario As String
        Public IDCategoria As String
        Public Fecha As Date
        Public CodTrabajo As String
        Public Horas As Integer
        Public Tasa As Integer
        Public Observaciones As String
        Public IDHora As Date
        Public dtExcel As DataTable
        'Friend GenerarPartesEnFincaPadre As Boolean

        Public Sub New(ByVal dtExcel As DataTable)
            Me.dtExcel = dtExcel
        End Sub
        
    End Class

    <Task()> Public Shared Function GenerarPartesTrabajosExcel(ByVal data As dataDatosImportacionPT, ByVal services As ServiceProvider) As Boolean

        If Not data.dtExcel Is Nothing AndAlso data.dtExcel.Rows.Count > 0 Then
            AdminData.BeginTx()
            Dim bsnObraTrabajo As New ObraTrabajo
            For Each drExcel As DataRow In data.dtExcel.Select
                '01. OBTENER IDOBRACAMPAÑA
                Dim f As New Filter
                f.Add("CFinca", FilterOperator.Equal, drExcel("CFinca").ToString)
                Dim bsnFinca As New BdgFinca
                Dim dtfinca As DataTable = bsnFinca.Filter(f)
                Dim dtObra As DataTable = New BE.DataEngine().Filter("vBdgNegDatosFincaObraCampaña", New GuidFilterItem("IDFinca", dtfinca.Rows(0)("IDFinca")))
                If dtObra Is Nothing OrElse dtObra.Rows.Count = 0 Then Return False

                '02. COMPROBAR QUE EL TRABAJO TIENE UN TRABAJO HIJO ASIGNADO
                '02.1 OBTENER EL IDTRABAJO QUE NOS HAN ENVIADO 
                f.Clear()
                f.Add("CodTrabajo", drExcel("CodTrabajo").ToString)
                f.Add("IDObra", dtObra.Rows(0)("IDObra"))
                Dim dtTrabajo As DataTable = bsnObraTrabajo.Filter(f)
                If dtTrabajo Is Nothing OrElse dtTrabajo.Rows.Count = 0 Then Return False

                '02.2 MIRAR SI ALGÚN TRABAJO TIENE ESE COMO PADRE
                'TODO => OJO CON TRABAJOS CERRADOS
                f.Clear()
                f.Add("IDTrabajoPadre", dtTrabajo.Rows(0)("IDTrabajo"))
                Dim dtTrabajoImputacion As DataTable = bsnObraTrabajo.Filter(f)
                Dim IDTrabajoGenerado As Integer
                If dtTrabajoImputacion Is Nothing OrElse dtTrabajoImputacion.Rows.Count = 0 Then
                    '03. SINO, CREAR UNO ' TODO

                    Dim generatrabajo As New stGenerarTrabajo(dtObra.Rows(0)("IDObra"), dtTrabajo.Rows(0)("IDTipoObra") & String.Empty, dtTrabajo.Rows(0)("IDTipoTrabajo") & String.Empty, _
                                             dtTrabajo.Rows(0)("IDSubTipoTrabajo") & String.Empty, dtTrabajo.Rows(0)("IDTrabajo"), drExcel("CodTrabajo").ToString, drExcel("Fecha"))
                    IDTrabajoGenerado = ProcessServer.ExecuteTask(Of stGenerarTrabajo, Integer)(AddressOf GenerarTrabajoExcel, generatrabajo, services)
                Else
                    IDTrabajoGenerado = dtTrabajo.Rows(0)("IDTrabajo")
                End If

                '04. PREPARAMOS DATOS PARA METER EN TBOBRAMODCONTROL
                Dim dataControl As New dataGenerarControl(dtObra.Rows(0)("IDObra"), IDTrabajoGenerado, dtTrabajo.Rows(0)("IDTrabajo"), _
                                                                       100, Guid.Empty, drExcel("Fecha"))
                Dim ObraParteAgrup As New ObraParteAgrupado()
                Dim StData As New BusinessData()
                StData = ObraParteAgrup.ObtenerTasa(drExcel, services)

                dataControl.dtConcepto = ObraParteAgrup.GenerarDatosConcepto(drExcel, StData.Item("TasaRealModA"), services)
                ProcessServer.ExecuteTask(Of dataGenerarControl)(AddressOf GenerarObraModControl, dataControl, services)

                '05. RECALCULAR OBRA
                ProcessServer.ExecuteTask(Of Integer)(AddressOf ObraCabecera.RecalcularObra, dtObra.Rows(0)("IDObra"), services)

                '06. ACTUALIZAR STOCK SI ES NECESARIO
                'If blnAutoactualizacionStocks Then
                '    Dim f As New Filter
                '    'f.Add("Actualizado", 0)
                '    f.Add("IDTrabajo", drLinea("IDTrabajoGenerado"))
                '    f.Add("IDParteAgrupado", drLinea("IDParteAgrupado"))
                '    'f.Add(New IsNullFilterItem("IDParteAgrupadoMat", False)) 'TODO => SÓLO ESCOGER LOS QUE SE GENEREN EN ESTE BUCLE

                '    'Dim dttMat As DataTable = New ObraMaterialControl().Filter(f) 'cogemos los mat generados para este trabajo
                '    Dim dttMat As DataTable = New DataEngine().Filter("NegBdgActualizacionStocksPartesAgrupados", f)
                '    ProcessServer.ExecuteTask(Of DataTable)(AddressOf ActualizacionStocksParteAgrupado, dttMat, services)
                'End If
            Next

        End If
        Return True
    End Function

    <Task()> Private Function GenerarDatosConcepto(ByVal drExcel As DataRow, ByVal Tasa As Integer, ByVal services As ServiceProvider) As DataTable
        Dim dttAux As DataTable = New DataTable
        dttAux.Columns.Add("IDOperario", GetType(String))
        dttAux.Columns.Add("IDHora", GetType(String))
        dttAux.Columns.Add("QHoras", GetType(String))
        dttAux.Columns.Add("TasaA", GetType(String))
        dttAux.Columns.Add("IDParteAgrupadoMod", GetType(String))

        Dim dtrAux As DataRow = dttAux.NewRow
        dtrAux("IDOperario") = drExcel("IDOperario")
        dtrAux("IDHora") = drExcel("IDHora")
        dtrAux("QHoras") = drExcel("HorasRealMod")
        dtrAux("TasaA") = Tasa

        dttAux.Rows.Add(dtrAux)
        Return dttAux
    End Function
    <Task()> Private Function ObtenerTasa(ByVal drExcel As DataRow, ByVal services As ServiceProvider) As BusinessData
        Dim HC As New HoraCategoria
        Dim op As New Operario
        Dim dtOP As DataTable = op.SelOnPrimaryKey(drExcel("IDOperario"))
        Dim cat As String = dtOP.Rows(0)("IDCategoria")
        Dim StData As New BusinessData()
        StData.Add("IDCategoria", cat)
        StData.Add("IDHora", drExcel("IDHora"))
        StData.Add("Fecha", drExcel("Fecha"))
        StData.Add("TasaRealModA", 0)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf HoraCategoria.ObtenerPrecioHoraCategoria, StData, services)
        Return StData
    End Function

    <Serializable()> Public Class stValidarGenerarTrabajo
        Public IDObra As Integer
        Public IDTipoObra As Integer
        Public IDTipoTrabajo As Integer
        Public IDSubtipoTrabajo As Integer
        Public Fecha As Date

        Public Sub New(ByVal IDObra As Integer, ByVal IDTipoObra As Integer, ByVal IDTipoTrabajo As Integer, ByVal IDSubtipoTrabajo As Integer, ByVal fecha As Date)
            Me.IDObra = IDObra
            Me.IDTipoObra = IDTipoObra
            Me.IDTipoTrabajo = IDTipoTrabajo
            Me.IDSubtipoTrabajo = IDSubtipoTrabajo
            Me.Fecha = fecha
        End Sub

    End Class

    <Task()> Public Shared Function ValidarGenerarTrabajo(ByVal data As stValidarGenerarTrabajo, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New IsNullFilterItem("IDTrabajoPadre", False))
        f.Add(New NumberFilterItem("IDObra", data.IDObra))
        f.Add(New NumberFilterItem("IDSubtipoTrabajo", data.IDSubtipoTrabajo))
        f.Add(New NumberFilterItem("IDTipoObra", data.IDTipoObra))
        f.Add(New NumberFilterItem("IDTipoTrabajo", data.IDTipoTrabajo))

        Dim dttTrabajos As DataTable = New ObraTrabajo().Filter(f)

        Return True
    End Function

    'TODO - Cambiar parámetro entrada a objeto tipificado
    <Serializable()> Public Class stGenerarTrabajo
        Public IDObra As Integer
        Public IDTipoObra As String
        Public IDTipoTrabajo As String
        Public IDSubtipoTrabajo As String
        Public IDTrabajo As Integer
        Public CodTrabajo As String
        Public Fecha As Date

        Public Sub New(ByVal IDObra As Integer, ByVal IDTipoObra As String, ByVal IDTipoTrabajo As String, ByVal IDSubtipoTrabajo As String, ByVal IDTrabajo As Integer, ByVal CodTrabajo As String, ByVal fecha As Date)
            Me.IDObra = IDObra
            Me.IDTipoObra = IDTipoObra
            Me.IDTipoTrabajo = IDTipoTrabajo
            Me.IDSubtipoTrabajo = IDSubtipoTrabajo
            Me.IDTrabajo = IDTrabajo
            Me.CodTrabajo = CodTrabajo
            Me.Fecha = fecha
        End Sub

    End Class
    <Task()> Public Shared Function GenerarTrabajoExcel(ByVal data As stGenerarTrabajo, ByVal services As ServiceProvider) As Integer
        Dim OT As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraTrabajo")

        Dim dtOT As DataTable = OT.AddNew
        Dim drOT As DataRow = dtOT.NewRow
        drOT("IDTrabajo") = AdminData.GetAutoNumeric
        drOT("IDObra") = data.IDObra
        drOT("IDTipoObra") = data.IDTipoObra
        drOT("IDTipoTrabajo") = data.IDTipoTrabajo
        drOT("IDSubTipoTrabajo") = data.IDSubtipoTrabajo

        drOT = OT.ApplyBusinessRule("IDSubTipoTrabajo", drOT("IDSubTipoTrabajo"), drOT, New BusinessData)
        If Length(data.IDTrabajo) > 0 Then
            drOT("DescTrabajo") = CDate(data.Fecha).ToString & " - " & drOT("DescTrabajo")

            Dim dataCodTrabajo As New dataGenerarCodTrabajo(data.IDObra, data.IDTrabajo, data.CodTrabajo, True)
            drOT("CodTrabajo") = ProcessServer.ExecuteTask(Of dataGenerarCodTrabajo, String)(AddressOf GenerarCodTrabajo, dataCodTrabajo, services)
            drOT("FechaInicio") = data.Fecha
            drOT("FechaFin") = data.Fecha
        End If
        data.CodTrabajo = drOT("CodTrabajo")
        drOT("IDTrabajoPadre") = data.IDTrabajo
        drOT("Facturable") = True
        drOT("QPrev") = 1
        drOT("Estado") = enumotEstado.otPendiente
        Dim datainversion As New ObraSubtipoTrabajo.stObtenerTrabajoInversion
        datainversion.IDTipoObra = drOT("IDTipoObra")
        datainversion.IDTipoTrabajo = drOT("IdTipoTrabajo")
        datainversion.IDSubtipoTrabajo = drOT("IDSubtipoTrabajo")
        datainversion.FechaInicioTrabajo = data.Fecha
        datainversion.FechaFinTrabajo = data.Fecha
        ProcessServer.ExecuteTask(Of ObraSubtipoTrabajo.stObtenerTrabajoInversion, Boolean) _
                                        (AddressOf ObraSubtipoTrabajo.ObtenerTrabajoEsInversion, datainversion, services)

        dtOT.Rows.Add(drOT)
        OT.Update(dtOT)

        Return drOT("IDTrabajo")
    End Function

    <Task()> Public Shared Function GenerarTrabajo(ByVal data As DataRow, ByVal services As ServiceProvider) As Integer
        Dim OT As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraTrabajo")

        Dim dtOT As DataTable = OT.AddNew
        Dim drOT As DataRow = dtOT.NewRow
        drOT("IDTrabajo") = AdminData.GetAutoNumeric
        drOT("IDObra") = data("IDObra")
        drOT("IDTipoObra") = data("IDTipoObra")
        drOT("IDTipoTrabajo") = data("IDTipoTrabajo")
        drOT("IDSubTipoTrabajo") = data("IDSubTipoTrabajo")

        drOT = OT.ApplyBusinessRule("IDSubTipoTrabajo", drOT("IDSubTipoTrabajo"), drOT, New BusinessData)
        If Length(data("IDTrabajo")) > 0 Then
            drOT("DescTrabajo") = CDate(data("FechaInicio")).ToString & " - " & drOT("DescTrabajo")

            Dim dataCodTrabajo As New dataGenerarCodTrabajo(data("IDObra"), data("IDTrabajo"), data("CodTrabajo"), True)
            drOT("CodTrabajo") = ProcessServer.ExecuteTask(Of dataGenerarCodTrabajo, String)(AddressOf GenerarCodTrabajo, dataCodTrabajo, services)
            drOT("FechaInicio") = data("FechaInicio")
            drOT("FechaFin") = data("FechaInicio")
        End If
        data("CodTrabajo") = drOT("CodTrabajo")
        drOT("IDTrabajoPadre") = data("IDTrabajo")
        drOT("Facturable") = True
        drOT("QPrev") = 1
        drOT("Estado") = enumotEstado.otPendiente
        Dim datainversion As New ObraSubtipoTrabajo.stObtenerTrabajoInversion
        datainversion.IDTipoObra = drOT("IDTipoObra")
        datainversion.IDTipoTrabajo = drOT("IdTipoTrabajo")
        datainversion.IDSubtipoTrabajo = drOT("IDSubtipoTrabajo")
        datainversion.FechaInicioTrabajo = data("FechaInicio")
        datainversion.FechaFinTrabajo = data("FechaInicio")
        ProcessServer.ExecuteTask(Of ObraSubtipoTrabajo.stObtenerTrabajoInversion, Boolean) _
                                        (AddressOf ObraSubtipoTrabajo.ObtenerTrabajoEsInversion, datainversion, services)

        dtOT.Rows.Add(drOT)
        OT.Update(dtOT)

        Return drOT("IDTrabajo")
    End Function

    <Serializable()> _
    Public Class dataGenerarCodTrabajo
        Public IDObra As String
        Public IDTrabajoPadre As String
        Public CodTrabajoPadre As String
        Public PorNivel As Boolean

        Public Sub New(ByVal IDObra As String, ByVal IDTrabajoPadre As String, ByVal CodTrabajoPadre As String, Optional ByVal PorNivel As Boolean = False)
            Me.IDObra = IDObra
            Me.IDTrabajoPadre = IDTrabajoPadre
            Me.CodTrabajoPadre = CodTrabajoPadre
            Me.PorNivel = PorNivel
        End Sub
    End Class
    <Task()> Public Shared Function GenerarCodTrabajo(ByVal data As dataGenerarCodTrabajo, ByVal services As ServiceProvider) As String
        Dim dtTrabajos As DataTable = New BE.DataEngine().Filter("tbObraTrabajo", New NumberFilterItem("IDObra", data.IDObra))
        Dim dvTrabajos As New DataView(dtTrabajos.Copy)
        Dim strRowFilter As String = dvTrabajos.RowFilter
        If Not dvTrabajos Is Nothing Then
            Dim i As Integer = 1
            Dim CodTrabajoAux As String = data.CodTrabajoPadre
            Dim CodTrabajo As String = CodTrabajoAux & "." & dvTrabajos.Count + i

            Dim f As New Filter
            If data.PorNivel Then
                If data.PorNivel Then
                    f.Add(New NumberFilterItem("IDTrabajoPadre", data.IDTrabajoPadre))
                Else
                    f.Add(New StringFilterItem("CodTrabajoPadre", CodTrabajoAux))
                End If
                CodTrabajo = CodTrabajoAux
            Else
                f.Add(New IsNullFilterItem("IDTrabajoPadre", True))
            End If

            dvTrabajos.RowFilter = f.Compose(New AdoFilterComposer)
            If dvTrabajos.Count <> 0 Then
                Dim NotExist As Boolean
                i = dvTrabajos.Count
                Do
                    i = i + 1
                    f.Clear()
                    f.Add(New StringFilterItem("CodTrabajo", CodTrabajoAux & "." & i))
                    dvTrabajos.RowFilter = f.Compose(New AdoFilterComposer)
                    If dvTrabajos.Count = 0 Then
                        NotExist = True
                    End If
                Loop Until NotExist
            End If
            CodTrabajo = CodTrabajo & "." & i
            dvTrabajos.RowFilter = strRowFilter

            Return CodTrabajo
        End If
        Return String.Empty
    End Function

    <Task()> Public Shared Sub TratarParteAgrupadoLinea(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtLinea As DataTable = Nothing
        If Length(data("IDParteAgrupadoLinea")) = 0 Then
            dtLinea = New ObraParteAgrupadoLinea().AddNewForm
            dtLinea.Rows(0)("IDParteAgrupadoLinea") = AdminData.GetAutoNumeric
            dtLinea.Rows(0)("IDParteAgrupado") = data("IDParteAgrupado")
            dtLinea.Rows(0)("IDTrabajo") = data("IDTrabajo")
        Else
            dtLinea = New ObraParteAgrupadoLinea().SelOnPrimaryKey(data("IDParteAgrupadoLinea"))
            dtLinea.Rows(0)("IDTrabajo") = data("IDTrabajoGenerado")
        End If
        dtLinea.Rows(0)("Cantidad") = data("Cantidad")
        dtLinea.Rows(0)("Porcentaje") = data("Porcentaje")
        dtLinea.Rows(0)("IDTrabajoGenerado") = data("IDTrabajoGenerado")
        ObraParteAgrupadoLinea.UpdateTable(dtLinea)
    End Sub

    <Serializable()> _
    Public Class dataGenerarControl
        Public IDObra As Integer
        Public IDTrabajo As Integer
        Public IDTrabajoPadre As Integer
        Public Porcentaje As Double
        Public IDParteAgrupado As Guid
        Public dtConcepto As DataTable
        Public FechaParte As Date

        Public Sub New(ByVal IDObra As Integer, ByVal IDTrabajo As Integer, ByVal IDTrabajoPadre As Integer, ByVal Porcentaje As Double, ByVal IDParteAgrupado As Guid, ByVal FechaParte As Date)
            Me.IDObra = IDObra
            Me.IDTrabajo = IDTrabajo
            Me.IDTrabajoPadre = IDTrabajoPadre
            Me.Porcentaje = Porcentaje
            Me.IDParteAgrupado = IDParteAgrupado
            Me.FechaParte = FechaParte
        End Sub
    End Class

    <Serializable()> _
    Public Class dataGenerarMaterial
        Public IDObra As String
        Public IDTrabajo As String
        Public IDParteAgrupadoMat As Guid
        Public IDArticulo As String
        Public DescArticulo As String
        Public Lote As String
        Public Cantidad As Double
        Public Porcentaje As Double
        Public FechaParte As DateTime
        Public IDLineaFactura As String '?
        Public IDLineaAlbaran As String '?

        Public Sub New(ByVal IDObra As String, ByVal IDTrabajo As String, ByVal IDParteAgrupadoMat As Guid, ByVal IDArticulo As String, _
                         ByVal DescArticulo As String, ByVal Lote As String, ByVal Cantidad As Double, ByVal Porcentaje As Double, ByVal FechaParte As DateTime)
            Me.IDObra = IDObra
            Me.IDTrabajo = IDTrabajo
            Me.IDParteAgrupadoMat = IDParteAgrupadoMat
            Me.IDArticulo = IDArticulo
            Me.DescArticulo = DescArticulo
            Me.Lote = Lote
            Me.Cantidad = Cantidad
            Me.Porcentaje = Porcentaje
            Me.FechaParte = FechaParte
        End Sub

    End Class
    <Task()> Public Shared Sub GenerarObraMaterialControl(ByVal data As dataGenerarControl, ByVal services As ServiceProvider)
        If Not data.dtConcepto Is Nothing AndAlso data.dtConcepto.Rows.Count > 0 Then
            Dim OMLote As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraParteAgrupadoMaterialLote")
            For Each drConcepto As DataRow In data.dtConcepto.Rows
                Dim dttLotes As DataTable = OMLote.Filter(New GuidFilterItem("IDParteAgrupadoMat", drConcepto("IDParteAgrupadoMat")))
                If Not dttLotes Is Nothing AndAlso dttLotes.Rows.Count > 0 Then
                    For Each dtrLote As DataRow In dttLotes.Rows
                        Dim dataMat As New dataGenerarMaterial(data.IDObra, data.IDTrabajo, drConcepto("IDParteAgrupadoMat"), drConcepto("IDArticulo"), drConcepto("DescArticulo"), _
                                                            dtrLote("Lote"), dtrLote("QInterna"), data.Porcentaje, data.FechaParte)
                        GenerarLineaObraMaterialControl(dataMat, services)
                    Next
                Else
                    Dim dataMat As New dataGenerarMaterial(data.IDObra, data.IDTrabajo, drConcepto("IDParteAgrupadoMat"), drConcepto("IDArticulo"), drConcepto("DescArticulo"), _
                                                           String.Empty, drConcepto("Cantidad"), data.Porcentaje, data.FechaParte)
                    GenerarLineaObraMaterialControl(dataMat, services)
                End If
            Next
        End If
    End Sub

    <Task()> Private Shared Sub GenerarLineaObraMaterialControl(ByVal data As dataGenerarMaterial, ByVal services As ServiceProvider)
        Dim OMC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraMaterialControl")
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDObra", data.IDObra))
        f.Add(New NumberFilterItem("IDTrabajo", data.IDTrabajo))
        If Length(data.IDParteAgrupadoMat) > 0 Then
            f.Add(New GuidFilterItem("IDParteAgrupadoMat", data.IDParteAgrupadoMat))
        Else
            f.Add(New NoRowsFilterItem)
        End If
        Dim dtMatControl As DataTable = OMC.Filter(f)
        If dtMatControl.Rows.Count = 0 Then
            dtMatControl = OMC.AddNewForm
            dtMatControl.Rows(0)("IDLineaMaterialControl") = AdminData.GetAutoNumeric
            dtMatControl.Rows(0)("IDObra") = data.IDObra
            dtMatControl.Rows(0)("IDTrabajo") = data.IDTrabajo
        End If
        dtMatControl.Rows(0)("IDMaterial") = data.IDArticulo
        OMC.ApplyBusinessRule("IDMaterial", dtMatControl.Rows(0)("IDMaterial"), dtMatControl.Rows(0), New BusinessData)
        dtMatControl.Rows(0)("DescMaterial") = data.DescArticulo
        dtMatControl.Rows(0)("QReal") = data.Cantidad * data.Porcentaje / 100
        OMC.ApplyBusinessRule("QReal", dtMatControl.Rows(0)("QReal"), dtMatControl.Rows(0), New BusinessData)

        dtMatControl.Rows(0)("LotePrevisto") = data.Lote


        dtMatControl.Rows(0)("Fecha") = data.FechaParte
        If Length(data.IDLineaFactura) > 0 Then dtMatControl.Rows(0)("IDLinFactura") = data.IDLineaFactura
        If Length(data.IDLineaAlbaran) > 0 Then dtMatControl.Rows(0)("IDLineaAlbaran") = data.IDLineaAlbaran
        dtMatControl.Rows(0)("IDParteAgrupadoMat") = data.IDParteAgrupadoMat

        OMC.Update(dtMatControl)
    End Sub

    <Task()> Public Shared Sub GenerarObraModControl(ByVal data As dataGenerarControl, ByVal services As ServiceProvider)
        If Not data.dtConcepto Is Nothing AndAlso data.dtConcepto.Rows.Count > 0 Then
            Dim OMC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraModControl")
            For Each drConcepto As DataRow In data.dtConcepto.Rows
                Dim f As New Filter
                f.Add(New NumberFilterItem("IDObra", data.IDObra))
                f.Add(New NumberFilterItem("IDTrabajo", data.IDTrabajo))
                If Not drConcepto.Table.Columns.Contains("IDParteAgrupadoMod") OrElse Length(drConcepto("IDParteAgrupadoMod")) = 0 Then
                    f.Add(New NoRowsFilterItem)
                Else
                    f.Add(New GuidFilterItem("IDParteAgrupadoMod", drConcepto("IDParteAgrupadoMod")))
                End If
                Dim dtModControl As DataTable = OMC.Filter(f)
                If dtModControl.Rows.Count = 0 Then
                    dtModControl = OMC.AddNewForm
                    dtModControl.Rows(0)("IDLineaMODControl") = AdminData.GetAutoNumeric
                    dtModControl.Rows(0)("IDObra") = data.IDObra
                    dtModControl.Rows(0)("IDTrabajo") = data.IDTrabajo
                End If

                dtModControl.Rows(0)("IDOperario") = drConcepto("IDOperario")
                OMC.ApplyBusinessRule("IDOperario", dtModControl.Rows(0)("IDOperario"), dtModControl.Rows(0), New BusinessData)

                dtModControl.Rows(0)("IDHora") = drConcepto("IDHora")
                dtModControl.Rows(0)("FechaInicio") = data.FechaParte
                dtModControl.Rows(0)("HorasRealMOD") = (drConcepto("QHoras") * data.Porcentaje) / 100
                dtModControl.Rows(0)("TasaRealModA") = drConcepto("TasaA")
                OMC.ApplyBusinessRule("TasaRealModA", dtModControl.Rows(0)("TasaRealModA"), dtModControl.Rows(0), New BusinessData)

                If Length(drConcepto("IDParteAgrupadoMod")) > 0 Then dtModControl.Rows(0)("IDParteAgrupadoMod") = drConcepto("IDParteAgrupadoMod")

                OMC.Update(dtModControl)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub GenerarObraCentrosControl(ByVal data As dataGenerarControl, ByVal services As ServiceProvider)
        If Not data.dtConcepto Is Nothing AndAlso data.dtConcepto.Rows.Count > 0 Then
            Dim OCC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCentroControl")
            For Each drConcepto As DataRow In data.dtConcepto.Rows
                Dim f As New Filter
                f.Add(New NumberFilterItem("IDObra", data.IDObra))
                f.Add(New NumberFilterItem("IDTrabajo", data.IDTrabajo))
                If Length(drConcepto("IDParteAgrupadoCentro")) > 0 Then
                    f.Add(New GuidFilterItem("IDParteAgrupadoCentro", drConcepto("IDParteAgrupadoCentro")))
                Else
                    f.Add(New NoRowsFilterItem)
                End If
                Dim dtCentroControl As DataTable = OCC.Filter(f)
                If dtCentroControl.Rows.Count = 0 Then
                    dtCentroControl = OCC.AddNewForm
                    dtCentroControl.Rows(0)("IDLineaCentroControl") = AdminData.GetAutoNumeric
                    dtCentroControl.Rows(0)("IDObra") = data.IDObra
                    dtCentroControl.Rows(0)("IDTrabajo") = data.IDTrabajo
                End If

                dtCentroControl.Rows(0)("IDCentro") = drConcepto("IDCentro")
                dtCentroControl.Rows(0)("DescCentro") = drConcepto("DescCentro")
                dtCentroControl.Rows(0)("FechaInicio") = data.FechaParte
                dtCentroControl.Rows(0)("HorasRealCentro") = (drConcepto("QHoras") * data.Porcentaje) / 100
                dtCentroControl.Rows(0)("TasaRealCentroA") = drConcepto("TasaA")
                OCC.ApplyBusinessRule("TasaRealCentroA", dtCentroControl.Rows(0)("TasaRealCentroA"), dtCentroControl.Rows(0), New BusinessData)

                dtCentroControl.Rows(0)("IDParteAgrupadoCentro") = drConcepto("IDParteAgrupadoCentro")

                OCC.Update(dtCentroControl)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub GenerarObraVariosControl(ByVal data As dataGenerarControl, ByVal services As ServiceProvider)
        If Not data.dtConcepto Is Nothing AndAlso data.dtConcepto.Rows.Count > 0 Then
            Dim OVC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraVariosControl")
            For Each drConcepto As DataRow In data.dtConcepto.Rows
                Dim f As New Filter
                f.Add(New NumberFilterItem("IDObra", data.IDObra))
                f.Add(New NumberFilterItem("IDTrabajo", data.IDTrabajo))
                If Length(drConcepto("IDParteAgrupadoVarios")) > 0 Then
                    f.Add(New GuidFilterItem("IDParteAgrupadoVarios", drConcepto("IDParteAgrupadoVarios")))
                Else
                    f.Add(New NoRowsFilterItem)
                End If
                Dim dtVariosControl As DataTable = OVC.Filter(f)
                If dtVariosControl.Rows.Count = 0 Then
                    dtVariosControl = OVC.AddNewForm
                    dtVariosControl.Rows(0)("IDLineaVariosControl") = AdminData.GetAutoNumeric
                    dtVariosControl.Rows(0)("IDObra") = data.IDObra
                    dtVariosControl.Rows(0)("IDTrabajo") = data.IDTrabajo
                End If
                dtVariosControl.Rows(0)("IDVarios") = drConcepto("IDVarios")
                dtVariosControl.Rows(0)("DescVarios") = drConcepto("DescVarios")
                dtVariosControl.Rows(0)("Fecha") = data.FechaParte
                dtVariosControl.Rows(0)("ImpRealVariosA") = (drConcepto("ImporteA") * data.Porcentaje) / 100
                OVC.ApplyBusinessRule("ImpRealVariosA", dtVariosControl.Rows(0)("ImpRealVariosA"), dtVariosControl.Rows(0), New BusinessData)
                dtVariosControl.Rows(0)("IDParteAgrupadoVarios") = drConcepto("IDParteAgrupadoVarios")

                OVC.Update(dtVariosControl)
            Next
        End If
    End Sub

#End Region

#Region "BusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDTipoObra", AddressOf CambioTipoTrabajo)
        Obrl.Add("IDTipoTrabajo", AddressOf CambioTipoTrabajo)
        Obrl.Add("IDSubTipoTrabajo", AddressOf CambioTipoTrabajo)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioTipoTrabajo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        'por si acaso está vacío, de primeras se lo ponemos
        If Length(data.Current("IDParteAgrupado")) = 0 Then data.Current("IDParteAgrupado") = Guid.NewGuid
        If Length(data.Current("IDTipoObra")) > 0 AndAlso Length(data.Current("IDTipoTrabajo")) > 0 AndAlso Length(data.Current("IDSubTipoTrabajo")) > 0 Then
            Dim taskData As stObtenerDatosParteAgrupado = New stObtenerDatosParteAgrupado()
            taskData.IDParteAgrupado = data.Current("IDParteAgrupado")
            taskData.IDTipoObra = data.Current("IDTipoObra")
            taskData.IDTipoTrabajo = data.Current("IDTipoTrabajo")
            taskData.IDSubTipoTrabajo = data.Current("IDSubTipoTrabajo")

            Dim taskResult As stDatosParteAgrupado = ProcessServer.ExecuteTask(Of stObtenerDatosParteAgrupado, stDatosParteAgrupado)(AddressOf ObtenerDatosParteAgrupado, taskData, services)
            If data.Current.Contains("Materiales") Then data.Current("Materiales") = taskResult.DatosMat Else data.Current.Add("Materiales", taskResult.DatosMat) ' data.Context.Add("Materiales", taskResult.DatosMat)
            If data.Current.Contains("Mod") Then data.Current("Mod") = taskResult.DatosMod Else data.Current.Add("Mod", taskResult.DatosMod)
            If data.Current.Contains("Centros") Then data.Current("Centros") = taskResult.DatosCentros Else data.Current.Add("Centros", taskResult.DatosCentros)
            If data.Current.Contains("Varios") Then data.Current("Varios") = taskResult.DatosVarios Else data.Current.Add("Varios", taskResult.DatosVarios)

            If data.Current.Contains("IDAnalisis") Then
                Dim dt As DataTable = New ObraSubtipoTrabajo().SelOnPrimaryKey(data.Current("IDTipoObra"), data.Current("IDTipoTrabajo"), data.Current("IDSubTipoTrabajo"))
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    data.Current("IDAnalisis") = dt.Rows(0)("IDAnalisis") & String.Empty
                End If
            End If
        End If
    End Sub

#End Region

#Region "Obtener registros"

    <Serializable()> Public Class stObtenerDatosParteAgrupado
        Public IDParteAgrupado As Guid
        Public IDTipoObra As String
        Public IDTipoTrabajo As String
        Public IDSubTipoTrabajo As String
    End Class

    <Serializable()> Public Class stDatosParteAgrupado
        Public DatosMat As DataTable
        Public DatosMod As DataTable
        Public DatosCentros As DataTable
        Public DatosVarios As DataTable
    End Class

    <Task()> Public Shared Function ObtenerDatosParteAgrupado(ByVal data As stObtenerDatosParteAgrupado, ByVal services As ServiceProvider) As stDatosParteAgrupado
        Dim resultData As New stDatosParteAgrupado
        If data Is Nothing OrElse Length(data.IDTipoObra) = 0 Then Return resultData

        Dim dtrTipoProyecto As DataRow = New ObraTipo().GetItemRow(data.IDTipoObra)
        If dtrTipoProyecto Is Nothing OrElse Length(dtrTipoProyecto("IDObraModelo")) = 0 Then Return resultData

        Dim f As New Filter
        f.Add("IDObra", dtrTipoProyecto("IDObraModelo"))
        f.Add("IDTipoObra", data.IDTipoObra)
        f.Add("IDTipoTrabajo", data.IDTipoTrabajo)
        f.Add("IDSubTipoTrabajo", data.IDSubTipoTrabajo)
        f.Add(New IsNullFilterItem("IDTrabajoPadre", True))

        Dim dttTrabajo As DataTable = New ObraTrabajo().Filter(f)
        If (dttTrabajo Is Nothing OrElse dttTrabajo.Rows.Count = 0) Then Return resultData

        f.Clear()
        f.Add("IDTrabajo", dttTrabajo.Rows(0)("IDTrabajo"))

        Dim dataaux As New stObtenerDatosParteInfo(data.IDParteAgrupado, New ObraMaterialControl().Filter(f), New DataTable())
        dataaux = ProcessServer.ExecuteTask(Of stObtenerDatosParteInfo, stObtenerDatosParteInfo)(AddressOf ObtenerDatosParteAgrupadoMateriales, dataaux, services)
        resultData.DatosMat = dataaux.Destino

        dataaux.Origen = New ObraMODControl().Filter(f)
        dataaux = ProcessServer.ExecuteTask(Of stObtenerDatosParteInfo, stObtenerDatosParteInfo)(AddressOf ObtenerDatosParteAgrupadoMod, dataaux, services)
        resultData.DatosMod = dataaux.Destino


        dataaux.Origen = New ObraCentroControl().Filter(f)
        dataaux = ProcessServer.ExecuteTask(Of stObtenerDatosParteInfo, stObtenerDatosParteInfo)(AddressOf ObtenerDatosParteAgrupadoCentro, dataaux, services)
        resultData.DatosCentros = dataaux.Destino

        dataaux.Origen = New ObraVariosControl().Filter(f)
        dataaux = ProcessServer.ExecuteTask(Of stObtenerDatosParteInfo, stObtenerDatosParteInfo)(AddressOf ObtenerDatosParteAgrupadoVarios, dataaux, services)
        resultData.DatosVarios = dataaux.Destino

        Return resultData
    End Function

    <Serializable()> Public Class stObtenerDatosParteInfo
        Public IDParteAgrupado As Guid
        Public Origen As DataTable
        Public Destino As DataTable
        Public CopiarCantidad As Boolean

        Public Sub New(ByVal IDParteAgrupado As Guid, ByVal dttOrigen As DataTable, ByVal dttDestino As DataTable)
            Me.IDParteAgrupado = IDParteAgrupado
            Me.Origen = dttOrigen
            Me.Destino = dttDestino
            CopiarCantidad = New BdgParametro().VolcarCantidadPartesAgrupados()
        End Sub
    End Class

    <Task()> Private Shared Function ObtenerDatosParteAgrupadoMateriales(ByVal data As stObtenerDatosParteInfo, ByVal services As ServiceProvider) As stObtenerDatosParteInfo
        Dim bsnArt As New Articulo()
        Dim bsnOPAM As New ObraParteAgrupadoMaterial
        data.Destino = New DataEngine().Filter("vBdgMntoObraParteAgrupadoMaterial", New NoRowsFilterItem)

        For Each dtrOrigen As DataRow In data.Origen.Rows
            Dim dtrNueva As DataRow = data.Destino.NewRow
            dtrNueva("IDParteAgrupadoMat") = Guid.NewGuid
            dtrNueva("IDParteAgrupado") = data.IDParteAgrupado
            dtrNueva("IDArticulo") = dtrOrigen("IDMaterial")
            dtrNueva = bsnOPAM.ApplyBusinessRule("IDArticulo", dtrNueva("IDArticulo"), dtrNueva, New BusinessData)
            dtrNueva("Cantidad") = 0
            If data.CopiarCantidad Then dtrNueva("Cantidad") = dtrOrigen("QReal")
            dtrNueva = bsnOPAM.ApplyBusinessRule("Cantidad", dtrNueva("Cantidad"), dtrNueva, New BusinessData)

            If Not data.Destino.Columns.Contains("GestionStockPorLotes") Then data.Destino.Columns.Add("GestionStockPorLotes", GetType(Boolean))
            Dim dtrArticulo As DataRow = bsnArt.GetItemRow(dtrNueva("IDArticulo"))
            dtrNueva("DescArticulo") = dtrArticulo("DescArticulo")
            dtrNueva("GestionStockPorLotes") = dtrArticulo("GestionStockPorLotes")
            data.Destino.Rows.Add(dtrNueva)
        Next
        Return data
    End Function

    <Task()> Private Shared Function ObtenerDatosParteAgrupadoMod(ByVal data As stObtenerDatosParteInfo, ByVal services As ServiceProvider) As stObtenerDatosParteInfo
        Dim bsnOPAM As New ObraParteAgrupadoMod
        data.Destino = New DataEngine().Filter("vBdgMntoObraParteAgrupadoMOD", New NoRowsFilterItem)

        For Each dtrOrigen As DataRow In data.Origen.Rows
            Dim dtrNueva As DataRow = data.Destino.NewRow

            dtrNueva("IDParteAgrupadoMod") = Guid.NewGuid
            dtrNueva("IDParteAgrupado") = data.IDParteAgrupado
            dtrNueva("IDOperario") = dtrOrigen("IDOperario")
            dtrNueva = bsnOPAM.ApplyBusinessRule("IDOperario", dtrOrigen("IDOperario"), dtrNueva, New BusinessData)
            dtrNueva("QHoras") = 0
            If data.CopiarCantidad Then dtrNueva("QHoras") = dtrOrigen("HorasRealMod")
            dtrOrigen = bsnOPAM.ApplyBusinessRule("QHoras", dtrNueva("QHoras"), dtrNueva, New BusinessData)
            Dim Operarios As EntityInfoCache(Of OperarioInfo) = services.GetService(Of EntityInfoCache(Of OperarioInfo))()
            Dim Operario As OperarioInfo = Operarios.GetEntity(dtrOrigen("IDOperario"))
            If Length(Operario.IDCategoria) > 0 Then
                dtrNueva("IDCategoria") = Operario.IDCategoria
                Dim IDHora As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf HoraCategoria.GetHoraPredeterminada, Operario.IDCategoria, services)
                If Length(IDHora) > 0 Then
                    dtrNueva("IDHora") = IDHora
                    dtrNueva("DescHora") = New Hora().GetItemRow(IDHora)("DescHora")
                    dtrOrigen = bsnOPAM.ApplyBusinessRule("IDHora", dtrNueva("IDHora"), dtrNueva, New BusinessData)
                End If
            End If

            data.Destino.Rows.Add(dtrNueva)
        Next
        Return data
    End Function

    <Task()> Private Shared Function ObtenerDatosParteAgrupadoCentro(ByVal data As stObtenerDatosParteInfo, ByVal services As ServiceProvider) As stObtenerDatosParteInfo
        Dim bsnOPAC As New ObraParteAgrupadoCentro
        data.Destino = New DataEngine().Filter("vBdgMntoObraParteAgrupadoCentro", New NoRowsFilterItem)

        For Each dtrOrigen As DataRow In data.Origen.Rows
            Dim dtrNueva As DataRow = data.Destino.NewRow
            dtrNueva("IDParteAgrupadoCentro") = Guid.NewGuid
            dtrNueva("IDParteAgrupado") = data.IDParteAgrupado
            dtrNueva("IDCentro") = dtrOrigen("IDCentro")
            dtrNueva = bsnOPAC.ApplyBusinessRule("IDCentro", dtrOrigen("IDCentro"), dtrNueva, New BusinessData)
            dtrNueva("QHoras") = 0
            If data.CopiarCantidad Then dtrNueva("QHoras") = dtrOrigen("HorasRealCentro")
            dtrOrigen = bsnOPAC.ApplyBusinessRule("QHoras", dtrNueva("QHoras"), dtrNueva, New BusinessData)
            data.Destino.Rows.Add(dtrNueva)
        Next
        Return data
    End Function

    <Task()> Private Shared Function ObtenerDatosParteAgrupadoVarios(ByVal data As stObtenerDatosParteInfo, ByVal services As ServiceProvider) As stObtenerDatosParteInfo
        Dim bsnOPAV As New ObraParteAgrupadoVarios
        data.Destino = New DataEngine().Filter("vBdgMntoObraParteAgrupadoVarios", New NoRowsFilterItem)

        For Each dtrOrigen As DataRow In data.Origen.Rows
            Dim dtrNueva As DataRow = data.Destino.NewRow
            dtrNueva("IDParteAgrupadoVarios") = Guid.NewGuid
            dtrNueva("IDParteAgrupado") = data.IDParteAgrupado
            dtrNueva("IDVarios") = dtrOrigen("IDVarios")
            dtrNueva("DescVarios") = dtrOrigen("DescVarios")
            dtrNueva("ImporteA") = 0
            If data.CopiarCantidad Then dtrNueva("ImporteA") = dtrOrigen("ImpRealVariosA")
            dtrOrigen = bsnOPAV.ApplyBusinessRule("IDVarios", dtrOrigen("IDVarios"), dtrNueva, New BusinessData)
            data.Destino.Rows.Add(dtrNueva)
        Next
        Return data
    End Function

#End Region

#Region " Task públicas "

    <Serializable()> _
      Public Class dataGetLineasAgrupacion
        Public IDParteAgrupado As Guid
        Public Fecha As Date
        Public IDFinca As Guid
        Public IDFincaHija As Guid
        Public IDTipoObra As String
        Public IDTipoTrabajo As String
        Public IDSubTipoTrabajo As String
        Public Planificado As enumBoolean

        Public Sub New(ByVal IDParteAgrupado As Guid, ByVal Fecha As Date, ByVal IDTipoObra As String, ByVal IDTipoTrabajo As String, ByVal IDSubTipoTrabajo As String, ByVal Planificado As enumBoolean)
            Me.IDTipoObra = IDTipoObra
            Me.IDTipoTrabajo = IDTipoTrabajo
            Me.IDParteAgrupado = IDParteAgrupado
            Me.Fecha = Fecha
            Me.IDSubTipoTrabajo = IDSubTipoTrabajo
            Me.Planificado = Planificado
        End Sub
    End Class
    <Task()> Public Shared Function GetLineasParteAgrupado(ByVal data As dataGetLineasAgrupacion, ByVal services As ServiceProvider) As DataTable
        Dim f As New Filter
        If Length(data.IDParteAgrupado) > 0 Then
            f.Add(New GuidFilterItem("IDParteAgrupado", data.IDParteAgrupado))
        Else
            f.Add(New NoRowsFilterItem)
        End If
        Dim dt As DataTable = New BE.DataEngine().Filter("vBdgMntoObraParteAgrupadoLineas", f)
        'dt.Columns.Add("ID", GetType(Integer))
        'dt.Columns.Add("Marca", GetType(Boolean))
        dt.Columns.Add("RepartoManual", GetType(Boolean))
        Dim dtLineas As DataTable = dt.Clone
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                dtLineas.Rows.Add(dr.ItemArray)
            Next
        End If
        dtLineas.AcceptChanges()
        Return dtLineas
    End Function

    <Serializable()> Public Class dataCopiarParteAgrupado
        Public IDParteAgrupadoOrigen As Guid
        Public IDParteAgrupadoDestino As Guid
        Public MensajeError As String

        Public Sub New(ByVal idOrigen As Guid)
            IDParteAgrupadoOrigen = idOrigen
        End Sub
    End Class

    <Task()> Public Shared Function CopiarParteAgrupado(ByVal data As dataCopiarParteAgrupado, ByVal services As ServiceProvider) As dataCopiarParteAgrupado 'podría devolver todos los datatables pero no necesario?
        If data Is Nothing OrElse Length(data.IDParteAgrupadoOrigen) = 0 OrElse data.IDParteAgrupadoOrigen = Guid.Empty Then Return data

        Try
            AdminData.BeginTx()
            'copiar cabecera
            Dim dttCabecera As DataTable = ProcessServer.ExecuteTask(Of Guid, DataTable)(AddressOf CopiarParteAgrupadoCabecera, data.IDParteAgrupadoOrigen, services)
            If dttCabecera Is Nothing Then Return data

            data.IDParteAgrupadoDestino = dttCabecera.Rows(0)("IDParteAgrupado")

            'copia líneas
            Dim dataLineas As New stCopiaParteAgrupadoHijo(data.IDParteAgrupadoOrigen, data.IDParteAgrupadoDestino)
            Dim dttLineas As DataTable = ProcessServer.ExecuteTask(Of stCopiaParteAgrupadoHijo, DataTable)(AddressOf CopiaParteAgrupadoLineas, dataLineas, services)

            'copia imputaciones
            Dim dttMat As DataTable = ProcessServer.ExecuteTask(Of stCopiaParteAgrupadoHijo, DataTable)(AddressOf CopiaParteAgrupadoMateriales, dataLineas, services)
            Dim dttMod As DataTable = ProcessServer.ExecuteTask(Of stCopiaParteAgrupadoHijo, DataTable)(AddressOf CopiaParteAgrupadoMOD, dataLineas, services)
            Dim dttCentros As DataTable = ProcessServer.ExecuteTask(Of stCopiaParteAgrupadoHijo, DataTable)(AddressOf CopiaParteAgrupadoCentros, dataLineas, services)
            Dim dttVarios As DataTable = ProcessServer.ExecuteTask(Of stCopiaParteAgrupadoHijo, DataTable)(AddressOf CopiaParteAgrupadoVarios, dataLineas, services)

            AdminData.CommitTx()
        Catch ex As Exception
            data.MensajeError = ex.Message
            AdminData.RollBackTx()
        End Try
        Return data
    End Function

    <Task()> Public Shared Function CopiarParteAgrupadoCabecera(ByVal data As Guid, ByVal services As ServiceProvider) As DataTable
        If Length(data) = 0 OrElse data = Guid.Empty Then Return Nothing
        Dim bsnOPA As ObraParteAgrupado = New ObraParteAgrupado()
        Dim origen As DataTable = bsnOPA.SelOnPrimaryKey(data)
        If (origen Is Nothing OrElse origen.Rows.Count = 0) Then Return Nothing

        Dim str(1) As String
        str(0) = "IDParteAgrupado" 'campos auditoria?
        str(1) = "NParteAgrupado" 'campos auditoria?

        Dim dataCopia As New stCopiaDataTable(origen, str)
        Dim resultData As stCopiaDataTable = ProcessServer.ExecuteTask(Of stCopiaDataTable, stCopiaDataTable)(AddressOf CopiarDatatable, dataCopia, services)
        If resultData Is Nothing OrElse resultData.Destino Is Nothing OrElse resultData.Destino.Rows.Count = 0 Then Return Nothing

        resultData.Destino.Rows(0)("IDParteAgrupado") = Guid.NewGuid
        Dim IDContador As String = ProcessServer.ExecuteTask(Of ContadorEntidad, String)(AddressOf CentroGestion.GetContadorPredeterminadoCGestionUsuario, ContadorEntidad.AlbaranVenta, services)
        resultData.Destino.Rows(0)("NParteAgrupado") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, IDContador, services)
        bsnOPA.Update(resultData.Destino)

        Return resultData.Destino
    End Function

    <Serializable()> Public Class stCopiaParteAgrupadoHijo
        Public ParteAgrupadoOrigen As Guid
        Public ParteAgrupadoDestino As Guid

        Public CopiarImputaciones As Boolean

        Public Sub New(ByVal idOrigen As Guid, ByVal idDestino As Guid)
            Me.ParteAgrupadoOrigen = idOrigen
            Me.ParteAgrupadoDestino = idDestino
            Me.CopiarImputaciones = New BdgParametro().CopiarImputacionesDuplicarParte
        End Sub
    End Class
    <Task()> Public Shared Function CopiaParteAgrupadoLineas(ByVal data As stCopiaParteAgrupadoHijo, ByVal services As ServiceProvider) As DataTable
        If data Is Nothing OrElse Nz(data.ParteAgrupadoOrigen, Guid.Empty) = Guid.Empty OrElse Nz(data.ParteAgrupadoDestino, Guid.Empty) = Guid.Empty Then Return Nothing

        Dim bsnOPAL As New ObraParteAgrupadoLinea
        Dim dttOrigen As DataTable = bsnOPAL.Filter(New GuidFilterItem("IDParteAgrupado", data.ParteAgrupadoOrigen))
        Dim str(3) As String
        str(0) = "IDParteAgrupado"
        str(1) = "IDParteAgrupadoLinea"
        str(2) = "CerrarLinea"
        str(3) = "IDAnalisisCabecera"
        Dim dataCopia As New stCopiaDataTable(dttOrigen, str)
        Dim dataCopiaResult As stCopiaDataTable = ProcessServer.ExecuteTask(Of stCopiaDataTable, stCopiaDataTable)(AddressOf CopiarDatatable, dataCopia, services)
        If dataCopiaResult Is Nothing OrElse dataCopiaResult.Destino Is Nothing OrElse dataCopiaResult.Destino.Rows.Count = 0 Then Return Nothing

        For Each dtr As DataRow In dataCopiaResult.Destino.Rows
            dtr("IDParteAgrupadoLinea") = Guid.NewGuid
            dtr("IDParteAgrupado") = data.ParteAgrupadoDestino
        Next
        bsnOPAL.Update(dataCopiaResult.Destino)
        Return dataCopiaResult.Destino
    End Function
    '<Task()> Public Shared Function CopiaParteAgrupadoMateriales(ByVal data As stCopiaParteAgrupadoHijo, ByVal services As ServiceProvider) As DataTable
    '    If data Is Nothing OrElse Nz(data.ParteAgrupadoOrigen, Guid.Empty) = Guid.Empty OrElse Nz(data.ParteAgrupadoDestino, Guid.Empty) = Guid.Empty Then Return Nothing
    '    Dim bsnOPAM As New ObraParteAgrupadoMaterial
    '    Dim dttOrigen As DataTable = bsnOPAM.Filter(New GuidFilterItem("IDParteAgrupado", data.ParteAgrupadoOrigen))
    '    Dim str(1) As String
    '    str(0) = "IDParteAgrupado"
    '    str(1) = "IDParteAgrupadoMat"
    '    Dim dataCopia As New stCopiaDataTable(dttOrigen, str)
    '    Dim dataCopiaResult As stCopiaDataTable = ProcessServer.ExecuteTask(Of stCopiaDataTable, stCopiaDataTable)(AddressOf CopiarDatatable, dataCopia, services)
    '    If dataCopiaResult Is Nothing OrElse dataCopiaResult.Destino Is Nothing OrElse dataCopiaResult.Destino.Rows.Count = 0 Then Return Nothing

    '    For Each dtr As DataRow In dataCopiaResult.Destino.Rows
    '        dtr("IDParteAgrupadoMat") = Guid.NewGuid
    '        dtr("IDParteAgrupado") = data.ParteAgrupadoDestino
    '    Next
    '    bsnOPAM.Update(dataCopiaResult.Destino)
    '    Return dataCopiaResult.Destino
    'End Function
    <Task()> Public Shared Function CopiaParteAgrupadoMOD(ByVal data As stCopiaParteAgrupadoHijo, ByVal services As ServiceProvider)
        If data Is Nothing OrElse Nz(data.ParteAgrupadoOrigen, Guid.Empty) = Guid.Empty OrElse Nz(data.ParteAgrupadoDestino, Guid.Empty) = Guid.Empty Then Return Nothing
        Dim bsnOPAM As New ObraParteAgrupadoMod
        Dim dttOrigen As DataTable = bsnOPAM.Filter(New GuidFilterItem("IDParteAgrupado", data.ParteAgrupadoOrigen))
        Dim str(1) As String
        str(0) = "IDParteAgrupado"
        str(1) = "IDParteAgrupadoMod"
        Dim dataCopia As New stCopiaDataTable(dttOrigen, str)
        Dim dataCopiaResult As stCopiaDataTable = ProcessServer.ExecuteTask(Of stCopiaDataTable, stCopiaDataTable)(AddressOf CopiarDatatable, dataCopia, services)
        If dataCopiaResult Is Nothing OrElse dataCopiaResult.Destino Is Nothing OrElse dataCopiaResult.Destino.Rows.Count = 0 Then Return Nothing

        For Each dtr As DataRow In dataCopiaResult.Destino.Rows
            dtr("IDParteAgrupadoMod") = Guid.NewGuid
            dtr("IDParteAgrupado") = data.ParteAgrupadoDestino
            If (Not data.CopiarImputaciones) Then
                dtr("QHoras") = 0
                dtr("ImporteA") = 0
            End If
        Next
        bsnOPAM.Update(dataCopiaResult.Destino)
        Return dataCopiaResult.Destino
    End Function
    <Task()> Public Shared Function CopiaParteAgrupadoCentros(ByVal data As stCopiaParteAgrupadoHijo, ByVal services As ServiceProvider)
        If data Is Nothing OrElse Nz(data.ParteAgrupadoOrigen, Guid.Empty) = Guid.Empty OrElse Nz(data.ParteAgrupadoDestino, Guid.Empty) = Guid.Empty Then Return Nothing
        Dim bsnOPAC As New ObraParteAgrupadoCentro
        Dim dttOrigen As DataTable = bsnOPAC.Filter(New GuidFilterItem("IDParteAgrupado", data.ParteAgrupadoOrigen))
        Dim str(1) As String
        str(0) = "IDParteAgrupado"
        str(1) = "IDParteAgrupadoCentro"
        Dim dataCopia As New stCopiaDataTable(dttOrigen, str)
        Dim dataCopiaResult As stCopiaDataTable = ProcessServer.ExecuteTask(Of stCopiaDataTable, stCopiaDataTable)(AddressOf CopiarDatatable, dataCopia, services)
        If dataCopiaResult Is Nothing OrElse dataCopiaResult.Destino Is Nothing OrElse dataCopiaResult.Destino.Rows.Count = 0 Then Return Nothing

        For Each dtr As DataRow In dataCopiaResult.Destino.Rows
            dtr("IDParteAgrupadoCentro") = Guid.NewGuid
            dtr("IDParteAgrupado") = data.ParteAgrupadoDestino
            If (Not data.CopiarImputaciones) Then
                dtr("ImporteA") = 0
                dtr("QHoras") = 0
            End If
        Next
        bsnOPAC.Update(dataCopiaResult.Destino)
        Return dataCopiaResult.Destino
    End Function
    <Task()> Public Shared Function CopiaParteAgrupadoVarios(ByVal data As stCopiaParteAgrupadoHijo, ByVal services As ServiceProvider)
        If data Is Nothing OrElse Nz(data.ParteAgrupadoOrigen, Guid.Empty) = Guid.Empty OrElse Nz(data.ParteAgrupadoDestino, Guid.Empty) = Guid.Empty Then Return Nothing
        Dim bsnOPAV As New ObraParteAgrupadoVarios
        Dim dttOrigen As DataTable = bsnOPAV.Filter(New GuidFilterItem("IDParteAgrupado", data.ParteAgrupadoOrigen))
        Dim str(1) As String
        str(0) = "IDParteAgrupado"
        str(1) = "IDParteAgrupadoVarios"
        Dim dataCopia As New stCopiaDataTable(dttOrigen, str)
        Dim dataCopiaResult As stCopiaDataTable = ProcessServer.ExecuteTask(Of stCopiaDataTable, stCopiaDataTable)(AddressOf CopiarDatatable, dataCopia, services)
        If dataCopiaResult Is Nothing OrElse dataCopiaResult.Destino Is Nothing OrElse dataCopiaResult.Destino.Rows.Count = 0 Then Return Nothing

        For Each dtr As DataRow In dataCopiaResult.Destino.Rows
            dtr("IDParteAgrupadoVarios") = Guid.NewGuid
            dtr("IDParteAgrupado") = data.ParteAgrupadoDestino
            If (Not data.CopiarImputaciones) Then
                dtr("ImporteA") = 0
            End If
        Next
        bsnOPAV.Update(dataCopiaResult.Destino)
        Return dataCopiaResult.Destino
    End Function
    <Task()> Public Shared Function CopiaParteAgrupadoMateriales(ByVal data As stCopiaParteAgrupadoHijo, ByVal services As ServiceProvider)
        If data Is Nothing OrElse Nz(data.ParteAgrupadoOrigen, Guid.Empty) = Guid.Empty OrElse Nz(data.ParteAgrupadoDestino, Guid.Empty) = Guid.Empty Then Return Nothing
        Dim bsnOPAM As New ObraParteAgrupadoMaterial
        Dim bsnOPAML As New ObraParteAgrupadoMaterialLote
        Dim dttMateriales As DataTable = bsnOPAM.Filter(New GuidFilterItem("IDParteAgrupado", data.ParteAgrupadoOrigen))
        Dim dttDestino As DataTable = dttMateriales.Clone
        Dim dttLotesDestino As DataTable
        If Not dttMateriales Is Nothing AndAlso dttMateriales.Rows.Count > 0 Then
            For Each dtrMat As DataRow In dttMateriales.Rows
                'copiar material
                dttDestino.Rows.Add(dtrMat.ItemArray)
                If (Not data.CopiarImputaciones) Then dttDestino.Rows(dttDestino.Rows.Count - 1)("Cantidad") = 0
                dttDestino.Rows(dttDestino.Rows.Count - 1)("IDParteAgrupadoMat") = Guid.NewGuid
                dttDestino.Rows(dttDestino.Rows.Count - 1)("IDParteAgrupado") = data.ParteAgrupadoDestino

                'copiar sus lotes
                Dim f As New Filter
                f.Add("IDParteAgrupadoMat", dtrMat("IDParteAgrupadoMat"))
                Dim dttLotes As DataTable = bsnOPAML.Filter(f)
                If Not dttLotes Is Nothing AndAlso dttLotes.Rows.Count > 0 Then
                    If dttLotesDestino Is Nothing Then dttLotesDestino = dttLotes.Clone
                    For Each dtrLot As DataRow In dttLotes.Rows
                        dttLotesDestino.Rows.Add(dtrLot.ItemArray)
                        dttLotesDestino.Rows(dttLotesDestino.Rows.Count - 1)("IDParteAgrupadoMatLote") = Guid.NewGuid
                        dttLotesDestino.Rows(dttLotesDestino.Rows.Count - 1)("IDParteAgrupadoMat") = dttDestino.Rows(dttDestino.Rows.Count - 1)("IDParteAgrupadoMat")
                    Next
                End If
            Next
        End If
        If Not dttDestino Is Nothing AndAlso dttDestino.Rows.Count > 0 Then bsnOPAM.Update(dttDestino)
        If Not dttLotesDestino Is Nothing AndAlso dttLotesDestino.Rows.Count > 0 Then bsnOPAML.Update(dttLotesDestino)

    End Function

    'a lo mejor existe ya una? otro sitio?
    <Serializable()> Public Class stCopiaDataTable
        Public Origen As DataTable
        Public Destino As DataTable
        Public ColumnasExcluidas(-1) As String

        Public Sub New(ByVal origen As DataTable)
            Me.Origen = origen
        End Sub
        Public Sub New(ByVal origen As DataTable, ByVal columnasExcluidas() As String)
            Me.New(origen)
            Me.ColumnasExcluidas = columnasExcluidas
        End Sub

        Public Sub New(ByVal origen As DataTable, ByVal destino As DataTable)
            Me.New(origen)
            Me.Destino = destino
        End Sub

        Public Sub New(ByVal origen As DataTable, ByVal destino As DataTable, ByVal columnasExcluidas() As String)
            Me.New(origen, destino)
            Me.ColumnasExcluidas = columnasExcluidas
        End Sub

    End Class
    <Task()> Public Shared Function CopiarDatatable(ByVal data As stCopiaDataTable, ByVal services As ServiceProvider) As stCopiaDataTable
        If data Is Nothing OrElse data.Origen Is Nothing OrElse data.Origen.Rows.Count = 0 Then Return Nothing

        If data.Destino Is Nothing Then
            data.Destino = data.Origen.Clone
        End If
        For Each dtrOrigen As DataRow In data.Origen.Rows
            Dim dtrDestino As DataRow = data.Destino.NewRow
            For Each dtc As DataColumn In data.Destino.Columns
                If data.ColumnasExcluidas Is Nothing OrElse data.ColumnasExcluidas.Length = 0 OrElse Array.IndexOf(data.ColumnasExcluidas, dtc.ColumnName) = -1 Then
                    If data.Origen.Columns.Contains(dtc.ColumnName) AndAlso Length(dtrOrigen(dtc.ColumnName)) > 0 Then
                        dtrDestino(dtc.ColumnName) = dtrOrigen(dtc.ColumnName)
                    End If
                End If
            Next
            data.Destino.Rows.Add(dtrDestino)
        Next
        Return data
    End Function

    <Task()> Public Shared Function ChequearPartes(ByVal data As DataTable, ByVal services As ServiceProvider) As DataTable
        'data.Columns.Add("Error", GetType(Boolean))
        For Each dr As DataRow In data.Select
            Dim f As New Filter()
            Dim op As New Filter()
            Dim tr As New Filter()
            Dim h As New Filter()
            Dim Finca As New BdgFinca()
            Dim Operario As New Operario()
            Dim Trabajo As New ObraTrabajo()
            Dim Hora As New Hora()
            f.Add(New StringFilterItem("CFinca", dr("CFinca")))
            Dim dtfinca As DataTable = Finca.Filter(f)
            If dtfinca.Rows.Count > 0 Then
                op.Add(New StringFilterItem("IDOperario", dr("IDOperario")))
                Dim dtop As DataTable = Operario.Filter(op)
                If dtop.Rows.Count > 0 Then
                    If Not dtfinca.Rows(0)("IDObraCampaña") Is DBNull.Value Then
                        Dim strIDObra As String = dtfinca.Rows(0)("IDObraCampaña") 'TODO - IDObra / IDObraCampaña en función de si es o no inversión
                        tr.Add(New StringFilterItem("IDObra", strIDObra))
                        tr.Add(New StringFilterItem("CodTrabajo", dr("CodTrabajo")))
                        Dim dttrabajo As DataTable = Trabajo.Filter(tr)
                        If dttrabajo.Rows.Count > 0 Then
                            h.Add(New StringFilterItem("IDHora", dr("IDHora")))
                            Dim dthora As DataTable = Hora.Filter(h)
                            If dthora.Rows.Count > 0 Then
                                dr("error") = False
                                dr("Mensaje") = String.Empty
                            Else
                                dr("error") = True
                                dr("Mensaje") = "El tipo de hora no es correcto"
                            End If
                        Else
                            dr("error") = True
                            dr("Mensaje") = "El trabajo no existe"
                        End If
                    Else
                        dr("error") = True
                        dr("Mensaje") = "La Finca no tiene asignada una Obra"

                    End If

                Else
                    dr("error") = True
                    dr("Mensaje") = "El Operario no existe"
                End If
            Else
                dr("error") = True
                dr("Mensaje") = "La Finca no existe"
            End If

        Next
        Return data
    End Function

#Region " Partes materiales desde Excel "
    <Task()> Public Shared Function ChequearPartesMaterial(ByVal data As DataTable, ByVal services As ServiceProvider) As DataTable
        'data.Columns.Add("Error", GetType(Boolean))
        For Each dr As DataRow In data.Select
            Dim f As New Filter()
            Dim op As New Filter()
            Dim tr As New Filter()
            Dim h As New Filter()
            Dim Finca As New BdgFinca()
            Dim Operario As New Operario()
            Dim Trabajo As New ObraTrabajo()
            Dim Material As New Articulo()
            f.Add(New StringFilterItem("CFinca", dr("CFinca")))
            Dim dtfinca As DataTable = Finca.Filter(f)
            If dtfinca.Rows.Count > 0 Then

                If Not dtfinca.Rows(0)("IDObraCampaña") Is DBNull.Value Then
                    Dim strIDObra As String = dtfinca.Rows(0)("IDObraCampaña") 'TODO - IDObra / IDObraCampaña en función de si es o no inversión
                    tr.Add(New StringFilterItem("IDObra", strIDObra))
                    tr.Add(New StringFilterItem("CodTrabajo", dr("CodTrabajo")))
                    Dim dttrabajo As DataTable = Trabajo.Filter(tr)
                    If dttrabajo.Rows.Count > 0 Then
                        h.Add(New StringFilterItem("IDArticulo", dr("IDProducto")))
                        Dim dtMaterial As DataTable = Material.Filter(h)
                        If dtMaterial.Rows.Count > 0 Then
                            dr("error") = False
                            dr("Mensaje") = String.Empty
                        Else
                            dr("error") = True
                            dr("Mensaje") = "El producto no es correcto"
                        End If
                    Else
                        dr("error") = True
                        dr("Mensaje") = "El trabajo no existe"
                    End If
                Else
                    dr("error") = True
                    dr("Mensaje") = "La Finca no tiene asignada una Obra"

                End If


            Else
                dr("error") = True
                dr("Mensaje") = "La Finca no existe"
            End If

        Next
        Return data
    End Function
    <Serializable()> _
    Public Class dataDatosImportacionMat
        Public CFinca As String
        Public Fecha As Date
        Public CodTrabajo As String
        Public IDProducto As String
        Public Cantidad As Integer
        Public dtExcel As DataTable

        Public Sub New(ByVal dtExcel As DataTable)
            Me.dtExcel = dtExcel
        End Sub

    End Class

    <Task()> Public Shared Function GenerarPartesMaterialExcel(ByVal data As dataDatosImportacionMat, ByVal services As ServiceProvider) As Boolean
        If Not data.dtExcel Is Nothing AndAlso data.dtExcel.Rows.Count > 0 Then
            AdminData.BeginTx()
            Dim bsnObraTrabajo As New ObraTrabajo
            Dim IDObras As New Dictionary(Of Integer, Integer)
            For Each drExcel As DataRow In data.dtExcel.Select
                '01. OBTENER IDOBRACAMPAÑA
                Dim f As New Filter
                f.Add("CFinca", FilterOperator.Equal, drExcel("CFinca").ToString)
                Dim bsnFinca As New BdgFinca
                Dim dtfinca As DataTable = bsnFinca.Filter(f)
                Dim dtObra As DataTable = New BE.DataEngine().Filter("vBdgNegDatosFincaObraCampaña", New GuidFilterItem("IDFinca", dtfinca.Rows(0)("IDFinca")))
                If dtObra Is Nothing OrElse dtObra.Rows.Count = 0 Then Return False
                If Not IDObras.ContainsKey(dtObra.Rows(0)("IDObra")) Then
                    IDObras.Add(dtObra.Rows(0)("IDObra"), dtObra.Rows(0)("IDObra"))
                End If
                '02. COMPROBAR QUE EL TRABAJO TIENE UN TRABAJO HIJO ASIGNADO
                '02.1 OBTENER EL IDTRABAJO QUE NOS HAN ENVIADO 
                f.Clear()
                f.Add("CodTrabajo", drExcel("CodTrabajo").ToString)
                f.Add("IDObra", dtObra.Rows(0)("IDObra"))
                Dim dtTrabajo As DataTable = bsnObraTrabajo.Filter(f)
                If dtTrabajo Is Nothing OrElse dtTrabajo.Rows.Count = 0 Then Return False

                '02.2 MIRAR SI ALGÚN TRABAJO TIENE ESE COMO PADRE
                'TODO => OJO CON TRABAJOS CERRADOS
                f.Clear()
                f.Add("IDTrabajoPadre", dtTrabajo.Rows(0)("IDTrabajo"))
                Dim dtTrabajoImputacion As DataTable = bsnObraTrabajo.Filter(f)
                Dim IDTrabajoGenerado As Integer
                If dtTrabajoImputacion Is Nothing OrElse dtTrabajoImputacion.Rows.Count = 0 Then
                    '03. SINO, CREAR UNO ' TODO
                    Dim generatrabajo As New ObraParteAgrupado.stGenerarTrabajo(dtObra.Rows(0)("IDObra"), dtTrabajo.Rows(0)("IDTipoObra") & String.Empty, dtTrabajo.Rows(0)("IDTipoTrabajo") & String.Empty, _
                                            dtTrabajo.Rows(0)("IDSubTipoTrabajo") & String.Empty, dtTrabajo.Rows(0)("IDTrabajo"), drExcel("CodTrabajo").ToString, drExcel("Fecha"))
                    IDTrabajoGenerado = ProcessServer.ExecuteTask(Of ObraParteAgrupado.stGenerarTrabajo, Integer)(AddressOf ObraParteAgrupado.GenerarTrabajoExcel, generatrabajo, services)
                Else
                    IDTrabajoGenerado = dtTrabajo.Rows(0)("IDTrabajo")
                End If

                '04. PREPARAMOS DATOS PARA METER EN TBOBRAMATERIALCONTROL
                Dim dataControl As New ObraParteAgrupado.dataGenerarControl(dtObra.Rows(0)("IDObra"), IDTrabajoGenerado, dtTrabajo.Rows(0)("IDTrabajo"), 100, Guid.Empty, drExcel("Fecha"))
                dataControl.dtConcepto = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf GenerarDatosConceptoMat, drExcel, services)
                ProcessServer.ExecuteTask(Of ObraParteAgrupado.dataGenerarControl)(AddressOf ObraParteAgrupado.GenerarObraMaterialControl, dataControl, services)
            Next

            '05. RECALCULAR OBRA
            For Each IDObra As Integer In IDObras.Values
                ProcessServer.ExecuteTask(Of Integer)(AddressOf ObraCabecera.RecalcularObra, IDObra, services)
            Next
        End If
        Return True
    End Function


    <Task()> Public Shared Function GenerarDatosConceptoMat(ByVal drExcel As DataRow, ByVal services As ServiceProvider) As DataTable
        Dim dttAux As DataTable = New DataTable
        dttAux.Columns.Add("IDArticulo", GetType(String))
        dttAux.Columns.Add("DescArticulo", GetType(String))
        dttAux.Columns.Add("Cantidad", GetType(String))
        dttAux.Columns.Add("IDAlmacen", GetType(String))
        dttAux.Columns.Add("IDParteAgrupadoMat", GetType(String))

        Dim dtArticulo As DataTable = New Articulo().SelOnPrimaryKey(drExcel("IDProducto"))

        Dim dtrAux As DataRow = dttAux.NewRow
        If dtArticulo.Rows.Count > 0 Then
            dtrAux("DescArticulo") = dtArticulo.Rows(0)("DescArticulo")
        End If
        dtrAux("IDArticulo") = drExcel("IDProducto")
        dtrAux("Cantidad") = drExcel("Cantidad")
        dttAux.Rows.Add(dtrAux)
        Return dttAux
    End Function

#End Region

    <Serializable()> _
    Public Class DataGetAnalisis
        Public IDAnalisis As String = String.Empty
        Public DescAnalisis As String = String.Empty
        Public Estado As Integer?

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDAnalisis As String, ByVal DescAnalisis As String, ByVal Estado As Integer)
            Me.IDAnalisis = IDAnalisis
            Me.DescAnalisis = DescAnalisis
            Me.Estado = Estado
        End Sub
    End Class

    <Task()> Public Shared Function GetAnalisis(ByVal IDAnalisisCabecera As Guid, ByVal services As ServiceProvider) As DataGetAnalisis
        Dim StData As New DataGetAnalisis
        If Length(IDAnalisisCabecera) > 0 AndAlso IDAnalisisCabecera <> Guid.Empty Then
            Dim DtBdgAnalisCab As DataTable = New BdgAnalisisCabecera().Filter(New GuidFilterItem("IDAnalisisCabecera", IDAnalisisCabecera))
            If Not DtBdgAnalisCab Is Nothing AndAlso DtBdgAnalisCab.Rows.Count > 0 Then
                StData.IDAnalisis = DtBdgAnalisCab.Rows(0)("IDAnalisis")
                StData.Estado = DtBdgAnalisCab.Rows(0)("Estado")
                Dim DtAnalis As DataTable = New BdgAnalisis().Filter(New FilterItem("IDAnalisis", FilterOperator.Equal, DtBdgAnalisCab.Rows(0)("IDAnalisis")))
                If Not DtAnalis Is Nothing AndAlso DtAnalis.Rows.Count > 0 Then
                    StData.DescAnalisis = DtAnalis.Rows(0)("DescAnalisis")
                End If
            End If
        End If
        Return StData
    End Function

#End Region

End Class

Public Enum enumBdgEstadoParteAgrupado
    Pendiente
    TrabajoGenerado
End Enum

Public Enum enumBdgOrigenParteAgrupado
    PartesTrabajoFinca
    MaterialesTrabajoFinca
End Enum

Public Enum enumBdgEstadoAnalisisGen
    Solicitado = 0
    Terminado = 1
    Anulado = 2
End Enum

Public Enum enumBdgTipoAnalisisCabGen
    Finca = 0
    Observatorio = 1
End Enum