Public Class BdgVinoAnalisis

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVinoAnalisis"

#End Region

#Region " RegisterAddnewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data(_VA.Fecha) = Date.Now
        Dim datacont As New Contador.DatosDefaultCounterValue(data, "BdgVinoAnalisis", _VA.NVinoAnalisis)
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, datacont, services)
    End Sub

#End Region

#Region " RegisterAddnewTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDVino")) = 0 Then ApplicationService.GenerateError("El código de vino es obligatorio.")
        If Length(data("IDAnalisis")) = 0 Then ApplicationService.GenerateError("El código de análisis es obligatorio.")
        If Length(data("Fecha")) = 0 Then ApplicationService.GenerateError("La fecha de análisis es obligatoria.")
    End Sub

#End Region

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarNVinoAnalisis)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If IsDBNull(data("IDVinoAnalisis")) Then data("IDVinoAnalisis") = Guid.NewGuid
    End Sub

    <Task()> Public Shared Sub AsignarNVinoAnalisis(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf GetNVinoAnalisis, New DataRowPropertyAccessor(data), services)
        End If
    End Sub

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("IDContador", AddressOf CambiaContador)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambiaContador(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf GetNVinoAnalisis, data.Current, services)
    End Sub

    <Task()> Public Shared Sub GetNVinoAnalisis(ByVal Current As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(Current("IDContador")) > 0 Then
            Current("NVinoAnalisis") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, Current("IDContador"), services)
        End If
    End Sub

#End Region


#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StCrearVinoAnalisisMasivo
        Public IDAnalisis As String
        Public NOperacion As String
        Public Fecha As Date
        Public Data As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDAnalisis As String, ByVal NOperacion As String, ByVal Fecha As Date, ByVal Data As DataTable)
            Me.IDAnalisis = IDAnalisis
            Me.NOperacion = NOperacion
            Me.Fecha = Fecha
            Me.Data = Data
        End Sub
    End Class

    <Task()> Public Shared Sub CrearVinoAnalisisMasivo(ByVal data As StCrearVinoAnalisisMasivo, ByVal services As ServiceProvider)
        If Not data.Data Is Nothing AndAlso data.Data.Rows.Count > 0 Then
            Dim IDVino As Guid
            Dim ClsVinoAn As New BdgVinoAnalisis
            Dim dtVA As DataTable = ClsVinoAn.AddNew
            Dim IDVinoAnalisis As Guid

            Dim oVV As New BdgVinoVariable
            Dim dtVV As DataTable = oVV.AddNew

            If Not data.Data.Columns.Contains(_VV.IDVinoAnalisis) Then
                data.Data.Columns.Add(_VV.IDVinoAnalisis, GetType(Guid))
            End If

            For Each oRw As DataRow In data.Data.Select(Nothing, _VA.IDVino)
                If Not IDVino.Equals(oRw(_VA.IDVino)) Then
                    IDVino = oRw(_VA.IDVino)
                    IDVinoAnalisis = Guid.NewGuid
                    Dim rwVA As DataRow = New BdgVinoAnalisis().AddNewForm.Rows(0)
                    Dim dt As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, "BdgVinoAnalisis", services)
                    If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                        Dim N As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, dt.Rows(0)("IDContador"), services)
                    Else
                        ApplicationService.GenerateError("La entidad BdgVinoAnalisis no tiene contador asignado.")
                    End If
                    rwVA(_VA.IDVinoAnalisis) = IDVinoAnalisis
                    rwVA(_VA.IDVino) = IDVino
                    rwVA(_VA.IDAnalisis) = data.IDAnalisis
                    rwVA(_VA.Fecha) = data.Fecha
                    rwVA(_VA.NOperacion) = data.NOperacion
                    rwVA(_VA.CodigoFoss) = oRw(_VA.CodigoFoss)
                    rwVA(_VA.Observaciones) = oRw(_VA.Observaciones)
                    dtVA.Rows.Add(rwVA.ItemArray)
                End If

                'Líneas (tbBdgVinoVariable)
                Dim drVV As DataRow = dtVV.NewRow
                drVV(_VV.IDVariable) = oRw(_VV.IDVariable)
                drVV(_VV.IDVinoAnalisis) = IDVinoAnalisis
                drVV(_VV.Valor) = oRw(_VV.Valor)
                drVV(_VV.ValorNumerico) = oRw(_VV.ValorNumerico)
                drVV(_VV.Orden) = oRw(_VV.Orden)
                dtVV.Rows.Add(drVV)
            Next
            BusinessHelper.UpdateTable(dtVA)
            oVV.Update(dtVV)
        End If
    End Sub



    <Serializable()> _
    Public Class DataCrearAnalisisVino
        Public Fecha As DateTime
        Public IDContador As String
        Public IDAnalisis As String

        Public NOperacion As String
        Public CodigoFoss As String
        Public Observaciones As String
        Public IDFactura As Integer?

        Public dtVariables As DataTable
        Public IDVino As Guid

        Public IDDeposito As String
        Public Lote As String

        Public ProcesoMasivo As Boolean

        Public Sub New(ByVal Fecha As DateTime, ByVal IDAnalisis As String, ByVal dtVariables As DataTable, ByVal IDVino As Guid, Optional ByVal IDContador As String = "")
            Me.Fecha = Fecha
            Me.IDAnalisis = IDAnalisis
            Me.dtVariables = dtVariables
            Me.IDVino = IDVino
            Me.IDContador = IDContador
        End Sub

        Public Sub New(ByVal Fecha As DateTime, ByVal IDAnalisis As String, ByVal dtVariables As DataTable, ByVal IDDeposito As String, ByVal Lote As String, Optional ByVal IDContador As String = "")
            Me.Fecha = Fecha
            Me.IDAnalisis = IDAnalisis
            Me.dtVariables = dtVariables
            Me.IDDeposito = IDDeposito
            Me.Lote = Lote
            Me.IDContador = IDContador
        End Sub
    End Class

    <Serializable()> _
    Public Class DataCrearAnalisisVinoResult
        Public IDVinoAnalisis As Guid
        Public NVinoAnalisis As String
        Public dtVinoAnalisis As DataTable
        Public dtVinoAnalisisVariable As DataTable
    End Class

    <Task()> Public Shared Function CrearAnalisisVino(ByVal data As DataCrearAnalisisVino, ByVal services As ServiceProvider) As DataCrearAnalisisVinoResult
        Dim va As New BdgVinoAnalisis
        Dim vv As New BdgVinoVariable

        Dim dtVinoAnalisis As DataTable = va.AddNew
        Dim dtVinoAnalisisVariable As DataTable = vv.AddNew

        Dim IDVinoAnalisis As Guid = Guid.NewGuid
        Dim drNewAnalisis As DataRow = dtVinoAnalisis.NewRow
        drNewAnalisis("IDVinoAnalisis") = IDVinoAnalisis
        If Length(data.IDContador) > 0 Then
            drNewAnalisis("IDContador") = data.IDContador
        Else
            Dim dtContadores As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, "BdgVinoAnalisis", services)
            If dtContadores Is Nothing AndAlso dtContadores.Rows.Count > 0 Then
                drNewAnalisis("IDContador") = dtContadores.Rows(0)("IDContador")
            Else
                ApplicationService.GenerateError("La entidad BdgVinoAnalisis no tiene contador asignado.")
            End If
        End If

        drNewAnalisis("Fecha") = data.Fecha
        drNewAnalisis("IDAnalisis") = data.IDAnalisis
        If Length(data.NOperacion) > 0 Then drNewAnalisis("NOperacion") = data.NOperacion
        If Length(data.CodigoFoss) > 0 Then drNewAnalisis("CodigoFoss") = data.CodigoFoss
        If Length(data.Observaciones) > 0 Then drNewAnalisis("Observaciones") = data.Observaciones
        If Not data.IDFactura Is Nothing Then drNewAnalisis("IDFactura") = CInt(data.IDFactura)

        Dim IDVino As Guid
        If data.IDVino.Equals(Guid.Empty) Then
            If Length(data.IDDeposito) > 0 AndAlso Length(data.Lote) > 0 AndAlso Nz(data.Fecha, cnMinDate) = cnMinDate Then
                Dim datBuscaVino As New DataSearchVino(data.IDDeposito, data.Lote, data.Fecha)
                IDVino = ProcessServer.ExecuteTask(Of DataSearchVino, Guid)(AddressOf SearchVino, datBuscaVino, services)
            End If
        Else
            IDVino = data.IDVino
        End If

        If Not IDVino.Equals(Guid.Empty) Then drNewAnalisis("IDVino") = IDVino
        dtVinoAnalisis.Rows.Add(drNewAnalisis)

        Dim AnalisisVacio As Boolean
        If data.dtVariables Is Nothing Then
            AnalisisVacio = True
            data.dtVariables = New BdgAnalisisVariable().Filter(New FilterItem("IDAnalisis", data.IDAnalisis))
        End If
        For Each drVar As DataRow In data.dtVariables.Select
            Dim drVV As DataRow = dtVinoAnalisisVariable.NewRow
            drVV(_VV.IDVariable) = drVar(_VV.IDVariable)
            drVV(_VV.IDVinoAnalisis) = IDVinoAnalisis
            If Not AnalisisVacio Then
                drVV(_VV.Valor) = drVar(_VV.Valor)
                drVV(_VV.ValorNumerico) = drVar(_VV.ValorNumerico)
            End If
            drVV(_VV.Orden) = drVar(_VV.Orden)
            dtVinoAnalisisVariable.Rows.Add(drVV)
        Next


        If Not data.ProcesoMasivo Then
            AdminData.BeginTx()
            va.Update(dtVinoAnalisis)
            vv.Update(dtVinoAnalisisVariable)
            AdminData.CommitTx(True)
        End If

        Dim rslt As New DataCrearAnalisisVinoResult
        If Not data.ProcesoMasivo Then
            rslt.IDVinoAnalisis = IDVinoAnalisis
            rslt.NVinoAnalisis = dtVinoAnalisis.Rows(0)("NVinoAnalisis") & String.Empty
        Else
            rslt.dtVinoAnalisis = dtVinoAnalisis
            rslt.dtVinoAnalisisVariable = dtVinoAnalisisVariable
        End If
        Return rslt
    End Function


    <Serializable()> _
    Public Class StCopiarVinoAnalisis
        Public IDVinoAnalisis As Guid
        Public IDVino As Guid
        Public Fecha As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVinoAnalisis As Guid, ByVal IDVino As Guid, ByVal Fecha As Date)
            Me.IDVinoAnalisis = IDVinoAnalisis
            Me.IDVino = IDVino
            Me.Fecha = Fecha
        End Sub
    End Class

    <Task()> Public Shared Function CopiarVinoAnalisis(ByVal data As StCopiarVinoAnalisis, ByVal services As ServiceProvider) As DataTable
        Dim ClsVinoAn As New BdgVinoAnalisis
        Dim drVinoAnalisis As DataRow = ClsVinoAn.GetItemRow(data.IDVinoAnalisis)
        Dim dtVinoAnalisisNew As DataTable = ClsVinoAn.AddNewForm
        dtVinoAnalisisNew.Rows(0)("IDVino") = data.IDVino
        dtVinoAnalisisNew.Rows(0)("IDAnalisis") = drVinoAnalisis("IDAnalisis")
        dtVinoAnalisisNew.Rows(0)("Fecha") = data.Fecha
        dtVinoAnalisisNew.Rows(0)("NOperacion") = drVinoAnalisis("NOperacion")
        dtVinoAnalisisNew.Rows(0)("Observaciones") = drVinoAnalisis("Observaciones")
        dtVinoAnalisisNew.Rows(0)("NVinoAnalisis") = System.DBNull.Value
        dtVinoAnalisisNew = ClsVinoAn.Update(dtVinoAnalisisNew)

        Dim vv As New BdgVinoVariable
        Dim f As New Filter
        f.Add(New GuidFilterItem("IDVinoAnalisis", drVinoAnalisis("IDVinoAnalisis")))
        Dim dtVinoVariable As DataTable = vv.Filter(f)
        If Not dtVinoVariable Is Nothing AndAlso dtVinoVariable.Rows.Count > 0 Then
            Dim dtVinoVariableNew As DataTable = dtVinoVariable.Clone
            For Each drVinoVariable As DataRow In dtVinoVariable.Rows
                Dim dr As DataRow = dtVinoVariable.NewRow()
                dr("IDVinoAnalisis") = dtVinoAnalisisNew.Rows(0)("IDVinoAnalisis")
                dr("IDVariable") = drVinoVariable("IDVariable")
                dr("Valor") = drVinoVariable("Valor")
                dr("ValorNumerico") = drVinoVariable("ValorNumerico")
                dr("Orden") = drVinoVariable("Orden")
                dtVinoVariableNew.Rows.Add(dr.ItemArray)
            Next
            vv.Update(dtVinoVariableNew)
        End If
        Return dtVinoAnalisisNew
    End Function

    <Task()> Public Shared Function GetAnalisisVino(ByVal IDVino As Guid, ByVal services As ServiceProvider) As DataTable
        Return AdminData.Execute("spBdgUltimoAnalisisVino", False, IDVino.ToString)
    End Function

    <Serializable()> _
    Public Class StGetAnalisisTraza
        Public IDAnalisis As String
        Public Filtro As Filter
        Public Articulo As String
        Public Deposito As String
        Public UltimoAnalisis As Boolean
        Public FechaVinoDesde As Date
        Public FechaVinoHasta As Date
        Public IDEstadoVino As String
        Public TipoDeposito As Integer
        Public IDNave As String
        Public IDBarrica As String
        Public IDMadera As String
        Public IDTipo As String
        Public IDFamilia As String
        Public IDSubFamilia As String
        Public Lote As String
        Public IDVino As Guid

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDAnalisis As String, ByVal Filtro As Filter, Optional ByVal Articulo As String = "", Optional ByVal Deposito As String = "", _
                       Optional ByVal UltimoAnalisis As Boolean = False, Optional ByVal FechaVinoDesde As Date = cnMinDate, Optional ByVal FechaVinoHasta As Date = cnMinDate, _
                       Optional ByVal IDEstadoVino As String = "", Optional ByVal TipoDeposito As Integer = -1, Optional ByVal IDNave As String = "", _
                       Optional ByVal IDBarrica As String = "", Optional ByVal IDMadera As String = "", Optional ByVal IDTipo As String = "", _
                       Optional ByVal IDFamilia As String = "", Optional ByVal IDSubFamilia As String = "", Optional ByVal Lote As String = "")
            Me.IDAnalisis = IDAnalisis
            Me.Filtro = Filtro
            Me.Articulo = Articulo
            Me.Deposito = Deposito
            Me.UltimoAnalisis = UltimoAnalisis
            Me.FechaVinoDesde = FechaVinoDesde
            Me.FechaVinoHasta = FechaVinoHasta
            Me.IDEstadoVino = IDEstadoVino
            Me.TipoDeposito = TipoDeposito
            Me.IDNave = IDNave
            Me.IDBarrica = IDBarrica
            Me.IDMadera = IDMadera
            Me.IDTipo = IDTipo
            Me.IDFamilia = IDFamilia
            Me.IDSubFamilia = IDSubFamilia
            Me.Lote = Lote
        End Sub
    End Class

    <Task()> Public Shared Function getAnalisisTrazabilidad(ByVal data As StGetAnalisisTraza, ByVal services As ServiceProvider) As DataTable
        Dim fltr As New Filter
        If Length(data.Articulo) > 0 Then fltr.Add(New StringFilterItem(_V.IDArticulo, data.Articulo))
        If Length(data.Deposito) > 0 Then fltr.Add(New StringFilterItem(_V.IDDeposito, data.Deposito))
        If data.FechaVinoDesde <> cnMinDate Then fltr.Add(New DateFilterItem("FechaVino", FilterOperator.GreaterThanOrEqual, data.FechaVinoDesde))
        If data.FechaVinoHasta <> cnMinDate Then fltr.Add(New DateFilterItem("FechaVino", FilterOperator.LessThanOrEqual, data.FechaVinoHasta))
        If Length(data.IDEstadoVino) > 0 Then fltr.Add(New StringFilterItem(_V.IDEstadoVino, data.IDEstadoVino))
        If data.TipoDeposito >= 0 Then fltr.Add(New NumberFilterItem(_D.TipoDeposito, data.TipoDeposito))
        If Length(data.IDNave) > 0 Then fltr.Add(New StringFilterItem(_D.IDNave, data.IDNave))
        If Length(data.IDBarrica) > 0 Then fltr.Add(New StringFilterItem(_V.IDBarrica, data.IDBarrica))
        If Length(data.IDMadera) > 0 Then fltr.Add(New StringFilterItem("IDMadera", data.IDMadera))
        If Length(data.IDTipo) > 0 Then fltr.Add(New StringFilterItem("IDTipo", data.IDTipo))
        If Length(data.IDFamilia) > 0 Then fltr.Add(New StringFilterItem("IDFamilia", data.IDFamilia))
        If Length(data.IDSubFamilia) > 0 Then fltr.Add(New StringFilterItem("IDSubfamilia", data.IDSubFamilia))
        If Length(data.Lote) > 0 Then fltr.Add(New StringFilterItem(_V.Lote, data.Lote))
        If Not data.IDVino.Equals(Guid.Empty) Then fltr.Add(New GuidFilterItem("IDVino", data.IDVino))

        Dim dtVinos As DataTable = New BE.DataEngine().Filter("vBdgExplosionVino1a1", fltr, "IDVino")
        If dtVinos Is Nothing OrElse dtVinos.Rows.Count = 0 Then Exit Function

        Dim dt1a1 As DataTable
        Dim dtUltimoAnalisis As DataTable
        Dim filPadre As New Filter(FilterUnionOperator.Or)
        Dim filHijo As New Filter(FilterUnionOperator.Or)

        If data.UltimoAnalisis Then
            dt1a1 = New DataTable
            dt1a1.Columns.Add("VinoPadre", GetType(Guid))
            dt1a1.Columns.Add("VinoHijo", GetType(Guid))
            dt1a1.Columns.Add("Nivel", GetType(Integer))

            Dim strB As Text.StringBuilder = New Text.StringBuilder
            strB.Append("create table #vars (IDVino uniqueidentifier, IDVariable varchar(10), Valor varchar(10), ValorNumerico numeric(23,8), FechaAnalisis datetime)")
            strB.Append(vbCrLf)
            For Each oRw As DataRow In dtVinos.Rows
                Dim IDVino As Guid = oRw(_V.IDVino)

                dt1a1.Rows.Add(New Object() {IDVino, IDVino, 0})

                filHijo.Add(New GuidFilterItem(_V.IDVino, IDVino))
                filPadre.Add(New GuidFilterItem(_V.IDVino, IDVino))

                strB.Append("insert into #vars exec spBdgUltimoAnalisisVino " & Quoted(IDVino.ToString))
                strB.Append(vbCrLf)
            Next
            strB.Append("select * from #vars")

            Dim SqlCmd As Common.DbCommand = AdminData.GetCommand
            SqlCmd.CommandText = strB.ToString
            dtUltimoAnalisis = AdminData.Execute(SqlCmd, ExecuteCommand.ExecuteReader)
        Else
            Dim strB As Text.StringBuilder = New Text.StringBuilder
            strB.Append("create table #exp (VinoPadre uniqueidentifier not null, VinoHijo uniqueidentifier not null, Nivel Int)")
            strB.Append(vbCrLf)

            For Each oRw As DataRow In dtVinos.Rows
                Dim IDVino As Guid = oRw(_V.IDVino)

                'filHijo.Add(New GuidFilterItem(_V.IDVino, oRw("VinoHijo")))
                'filPadre.Add(New GuidFilterItem(_V.IDVino, oRw("VinoPadre")))

                strB.Append("insert into #exp exec spBdgExplosionVino1a1 " & Quoted(IDVino.ToString))
                strB.Append(vbCrLf)
            Next

            strB.Append("select * from #exp")
            Dim SqlCmd As Common.DbCommand = AdminData.GetCommand
            SqlCmd.CommandText = strB.ToString
            dt1a1 = AdminData.Execute(SqlCmd, ExecuteCommand.ExecuteReader)
            For Each oRw As DataRow In dt1a1.Rows
                filHijo.Add(New GuidFilterItem(_V.IDVino, oRw("VinoHijo")))
                filPadre.Add(New GuidFilterItem(_V.IDVino, oRw("VinoPadre")))
            Next
        End If

        data.Filtro.Add(filHijo)

        'Recupera los analisis de los Vinos Hijos
        Dim StGet As New StGetAnalisis(data.IDAnalisis, data.Filtro, True, data.UltimoAnalisis)
        Dim dtAnaVinoHijo As DataTable = ProcessServer.ExecuteTask(Of StGetAnalisis, DataTable)(AddressOf GetAnalisis, StGet, services)

        If data.UltimoAnalisis Then
            Dim dvUltimoAnalisis As New DataView(dtUltimoAnalisis)

            dvUltimoAnalisis.Sort = _V.IDVino
            For Each oRw As DataRow In dtAnaVinoHijo.Rows
                For Each oRwv As DataRowView In dvUltimoAnalisis.FindRows(oRw(_V.IDVino))
                    If dtAnaVinoHijo.Columns.Contains(oRwv("IDVariable")) Then
                        oRw(oRwv("IDVariable")) = oRwv("Valor")
                        oRw(oRwv("IDVariable") & "_N") = oRwv("ValorNumerico")
                    End If
                Next

                If Length(oRw("FechaAnalisis")) = 0 Then
                    Dim f As New Filter
                    f.Add(New GuidFilterItem("IDVino", oRw(_V.IDVino)))
                    dtUltimoAnalisis.DefaultView.RowFilter = f.Compose(New AdoFilterComposer)
                    dtUltimoAnalisis.DefaultView.Sort = "FechaAnalisis DESC"
                    If dtUltimoAnalisis.DefaultView.Count > 0 Then
                        oRw("FechaAnalisis") = dtUltimoAnalisis.DefaultView(0).Row("FechaAnalisis")
                    End If
                    dtUltimoAnalisis.DefaultView.RowFilter = ""
                End If
            Next
        End If

        'Recupera los datos del Vino Padre (articulo, deposito...)
        Dim dtDatosVinoPadre As DataTable = AdminData.GetData("frmBdgCIAnalisisVinoDetalleDatos", filPadre)
        Dim dtDatosVinoHijo As DataTable = AdminData.GetData("frmBdgCIAnalisisVinoDetalleDatos", filHijo)

        Dim StAnalisis As New StTbAnalisisTrazabilidad(data.IDAnalisis, dtAnaVinoHijo, dtDatosVinoPadre, dtDatosVinoHijo)
        Dim dt As DataTable = ProcessServer.ExecuteTask(Of StTbAnalisisTrazabilidad, DataTable)(AddressOf TbAnalisisTrazabilidad, StAnalisis, services)

        Dim StUnir As New StUnirTablas(dt1a1, dtDatosVinoPadre, dtAnaVinoHijo, dt, dtDatosVinoHijo)
        ProcessServer.ExecuteTask(Of StUnirTablas)(AddressOf UnirTablas, StUnir, services)

        Return dt
    End Function

    <Serializable()> _
    Public Class StUnirTablas
        Public Dt1a1 As DataTable
        Public DtDatosVinoPadre As DataTable
        Public DtAnaVinoHijo As DataTable
        Public Dt As DataTable
        Public DtDatosVinoHijo As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal Dt1a1 As DataTable, ByVal DtDatosVinoPadre As DataTable, ByVal DtAnaVinoHijo As DataTable, ByVal Dt As DataTable, ByVal DtDatosVinoHijo As DataTable)
            Me.Dt1a1 = Dt1a1
            Me.DtDatosVinoPadre = DtDatosVinoPadre
            Me.DtAnaVinoHijo = DtAnaVinoHijo
            Me.Dt = Dt
            Me.DtDatosVinoHijo = DtDatosVinoHijo
        End Sub
    End Class

    <Task()> Public Shared Sub UnirTablas(ByVal data As StUnirTablas, ByVal services As ServiceProvider)
        'Para cada vino explosionado

        Dim VinosExplosionados As List(Of DataRow) = (From c In data.Dt1a1 Select c).ToList
        If VinosExplosionados Is Nothing OrElse VinosExplosionados.Count = 0 Then Exit Sub
        For Each dr1a1 As DataRow In VinosExplosionados
            'Dim filtroPadre As New Filter
            ' Dim filtroHijo As New Filter
            'filtroHijo.Add("IdVino", FilterOperator.Equal, dr1a1("VinoHijo"), FilterType.Guid)
            'filtroPadre.Add("IdVino", FilterOperator.Equal, dr1a1("VinoPadre"), FilterType.Guid)

            Dim Cols As List(Of Object)
            'Si un hijo tiene mas de un analisis se debe de crear una linea por cada Analisis que tenga
            'Sino debe de sacar la linea igualmente.
            Dim AnadaVinoHijo As List(Of DataRow)
            AnadaVinoHijo = (From c In data.DtAnaVinoHijo Where Not c.IsNull("IdVino") AndAlso c("IdVino") = dr1a1("VinoHijo") Select c).ToList
            If AnadaVinoHijo Is Nothing OrElse AnadaVinoHijo.Count = 0 Then 'NO hay análisis de los hijos y se introduce un registro con las variables en blanco
                Dim dr As DataRow = data.Dt.NewRow
                dr("IdVinoPadre") = dr1a1("VinoPadre")
                dr("IdVinoHijo") = dr1a1("VinoHijo")
                dr("Nivel") = dr1a1("Nivel")

                If Not data.DtDatosVinoPadre Is Nothing Then
                    Dim VinoPadre As List(Of DataRow) = (From c In data.DtDatosVinoPadre Where Not c.IsNull("IdVino") AndAlso c("IdVino") = dr1a1("VinoPadre") Select c).ToList
                    If Not VinoPadre Is Nothing AndAlso VinoPadre.Count > 0 Then
                        For Each drVinoPadre As DataRow In VinoPadre 'data.DtDatosVinoPadre.Select(filtroPadre.Compose(New AdoFilterComposer))
                            Cols = (From c In data.DtDatosVinoPadre.Columns Where CType(c, DataColumn).ColumnName <> "IdVino" Select c).ToList
                            For Each column As DataColumn In Cols 'data.DtDatosVinoPadre.Columns
                                ' If Not column.ColumnName = "IdVino" Then
                                dr(column.ColumnName & "Actual") = drVinoPadre(column.ColumnName)
                                ' End If
                            Next
                        Next
                    End If
                End If

                If Not data.DtDatosVinoHijo Is Nothing Then
                    Dim VinoHijo As List(Of DataRow) = (From c In data.DtDatosVinoHijo Where Not c.IsNull("IdVino") AndAlso c("IdVino") = dr1a1("VinoHijo") Select c).ToList
                    If Not VinoHijo Is Nothing AndAlso VinoHijo.Count > 0 Then

                        For Each drVinoHijo As DataRow In VinoHijo 'data.DtDatosVinoHijo.Select(filtroHijo.Compose(New AdoFilterComposer))
                            Cols = (From c In data.DtDatosVinoHijo.Columns Where CType(c, DataColumn).ColumnName <> "IdVino" Select c).ToList
                            For Each column As DataColumn In Cols 'data.DtDatosVinoHijo.Columns
                                'If Not column.ColumnName = "IdVino" Then
                                dr(column.ColumnName & "Anterior") = drVinoHijo(column.ColumnName)
                                ' End If
                            Next
                        Next
                    End If

                End If

                'Añadir las nuevas lineas al datatable que vamos a devolver
                data.Dt.Rows.Add(dr)
            Else 'SI hay análisis de los hijos y se introduce un registro por cada análisis.
                For Each drAnalisis As DataRow In AnadaVinoHijo
                    'Crear una nueva linea
                    Dim dr As DataRow = data.Dt.NewRow
                    dr("IdVinoPadre") = dr1a1("VinoPadre")
                    dr("IdVinoHijo") = dr1a1("VinoHijo")
                    dr("Nivel") = dr1a1("Nivel")

                    'Los nombres de las columnas van a ser los mismos en la tb que en dtAnaVinoHijo
                    Cols = (From c In data.DtAnaVinoHijo.Columns Where CType(c, DataColumn).ColumnName <> "IdVino" Select c).ToList
                    For Each column As DataColumn In Cols 'data.DtAnaVinoHijo.Columns
                        ' If Not column.ColumnName = "IdVino" Then
                        dr(column.ColumnName) = drAnalisis(column.ColumnName)
                        'End If
                    Next

                    If Not data.DtDatosVinoPadre Is Nothing Then
                        Dim VinoPadre As List(Of DataRow) = (From c In data.DtDatosVinoPadre Where Not c.IsNull("IdVino") AndAlso c("IdVino") = dr1a1("VinoPadre") Select c).ToList
                        If Not VinoPadre Is Nothing AndAlso VinoPadre.Count > 0 Then
                            For Each drVinoPadre As DataRow In VinoPadre 'data.DtDatosVinoPadre.Select(filtroPadre.Compose(New AdoFilterComposer))
                                Cols = (From c In data.DtDatosVinoPadre.Columns Where CType(c, DataColumn).ColumnName <> "IdVino" Select c).ToList
                                For Each column As DataColumn In Cols 'data.DtDatosVinoPadre.Columns
                                    ' If Not column.ColumnName = "IdVino" Then
                                    dr(column.ColumnName & "Actual") = drVinoPadre(column.ColumnName)
                                    ' End If
                                Next
                            Next
                        End If
                    End If

                    If Not data.DtDatosVinoHijo Is Nothing Then
                        Dim VinoHijo As List(Of DataRow) = (From c In data.DtDatosVinoHijo Where Not c.IsNull("IdVino") AndAlso c("IdVino") = dr1a1("VinoHijo") Select c).ToList
                        If Not VinoHijo Is Nothing AndAlso VinoHijo.Count > 0 Then

                            For Each drVinoHijo As DataRow In VinoHijo 'data.DtDatosVinoHijo.Select(filtroHijo.Compose(New AdoFilterComposer))
                                Cols = (From c In data.DtDatosVinoHijo.Columns Where CType(c, DataColumn).ColumnName <> "IdVino" Select c).ToList
                                For Each column As DataColumn In Cols 'data.DtDatosVinoHijo.Columns
                                    'If Not column.ColumnName = "IdVino" Then
                                    dr(column.ColumnName & "Anterior") = drVinoHijo(column.ColumnName)
                                    ' End If
                                Next
                            Next
                        End If

                    End If
                    'Añadir las nuevas lineas al datatable que vamos a devolver
                    data.Dt.Rows.Add(dr)
                Next
            End If
        Next
    End Sub

    <Serializable()> _
    Public Class StTbAnalisisTrazabilidad
        Public IDAnalisis As String
        Public DtAh As DataTable
        Public DtVp As DataTable
        Public DtVh As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDAnalisis As String, ByVal DtAh As DataTable, ByVal DtVp As DataTable, ByVal DtVh As DataTable)
            Me.IDAnalisis = IDAnalisis
            Me.DtAh = DtAh
            Me.DtVp = DtVp
            Me.DtVh = DtVh
        End Sub
    End Class

    <Task()> Public Shared Function TbAnalisisTrazabilidad(ByVal data As StTbAnalisisTrazabilidad, ByVal services As ServiceProvider) As DataTable
        Dim dt As New DataTable

        'Columnas del procedimiento almacenado
        dt.Columns.Add("IDVinoPadre", GetType(Guid))
        dt.Columns.Add("IDVinoHijo", GetType(Guid))

        'Columnas del vino padre o actual
        Dim Cols As List(Of Object)
        If Not data.DtVp Is Nothing Then
            Cols = (From c In data.DtVp.Columns Where CType(c, DataColumn).ColumnName <> "IDVino" Select c).ToList
            For Each column As DataColumn In Cols 'data.DtVp.Columns
                ' If Not column.ColumnName = "IdVino" Then
                dt.Columns.Add(column.ColumnName & "Actual", column.DataType)
                ' End If
            Next
        End If

        'Columnas del vino hijo o anterior
        dt.Columns.Add("Nivel", GetType(Integer))
        If Not data.DtVh Is Nothing Then
            Cols = (From c In data.DtVh.Columns Where CType(c, DataColumn).ColumnName <> "IDVino" Select c).ToList
            For Each column As DataColumn In Cols 'data.DtVh.Columns
                ' If Not column.ColumnName = "IdVino" Then
                dt.Columns.Add(column.ColumnName & "Anterior", column.DataType)
                ' End If
            Next
        End If

        'Columnas relacionadas con el análisis
        If Not data.DtAh Is Nothing Then
            Cols = (From c In data.DtAh.Columns Where CType(c, DataColumn).ColumnName <> "IDVino" Select c).ToList
            For Each column As DataColumn In Cols 'data.DtAh.Columns
                '  If Not column.ColumnName = "IdVino" Then
                dt.Columns.Add(column.ColumnName, column.DataType)
                ' End If
            Next
        End If

        Return dt
    End Function

    <Serializable()> _
    Public Class StGetAnalisis
        Public IDAnalisis As String
        Public Filtro As Filter
        Public ConTrazabilidad As Boolean
        Public UltimoAnalisis As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDAnalisis As String, ByVal Filtro As Filter, Optional ByVal ConTrazabilidad As Boolean = False, _
                       Optional ByVal UltimoAnalisis As Boolean = False)
            Me.IDAnalisis = IDAnalisis
            Me.Filtro = Filtro
            Me.ConTrazabilidad = ConTrazabilidad
            Me.UltimoAnalisis = UltimoAnalisis
        End Sub
    End Class

    <Task()> Public Shared Function GetAnalisis(ByVal data As StGetAnalisis, ByVal services As ServiceProvider) As DataTable
        Dim strView As String = "frmBdgCIAnalisisVinoDetalle"
        If data.ConTrazabilidad Then strView = "frmBdgCIAnalisisVinoDetalleTraza"
        Dim BEDateEngine As New DataEngine
        If data.UltimoAnalisis Then data.Filtro.Add(New BooleanFilterItem("EsUltimoAnalisisDepArt", data.UltimoAnalisis))
        Dim dt As DataTable = BEDateEngine.Filter(strView, data.Filtro)


        '1. Definir en el datatable resultado todas las columnas para las variables
        Dim dtVariables As DataTable = New BdgAnalisisVariable().Filter(New StringFilterItem("IDAnalisis", data.IDAnalisis))

        Dim fVariables As New Filter(FilterUnionOperator.Or)
        If Not dtVariables Is Nothing AndAlso dtVariables.Rows.Count > 0 Then
            Dim ListaVariables As List(Of DataRow) = (From c In dtVariables Order By c("Orden") Select c).ToList
            For Each dcV As DataRow In ListaVariables
                If Not dt.Columns.Contains(dcV("IDVariable")) Then
                    dt.Columns.Add(dcV("IDVariable"), GetType(String))
                    dt.Columns.Add(dcV("IDVariable") & "_N", GetType(Double))

                    fVariables.Add("IDVariable", dcV("IDVariable"))
                End If
            Next
        End If


        If Not data.UltimoAnalisis Then
            'ahora las actualizamos
            Dim bsnBVV As New BdgVinoVariable
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then

                Dim RegsAnalisis As List(Of DataRow) = (From c In dt Select c).ToList
                For Each dr As DataRow In RegsAnalisis
                    Dim f As New Filter
                    f.Add("IDVinoAnalisis", dr("IDVinoAnalisis"))
                    f.Add(fVariables)
                    Dim dtResult As DataTable = bsnBVV.Filter(f)
                    If Not dtResult Is Nothing AndAlso dtResult.Rows.Count > 0 Then
                        Dim ValoresVariables As List(Of DataRow) = (From c In dtResult Select c).ToList

                        For Each drResult As DataRow In ValoresVariables
                            dr(drResult("IDVariable")) = drResult("Valor")
                            dr(drResult("IDVariable") & "_N") = drResult("ValorNumerico")
                        Next
                    End If
                Next
            End If

        End If
        Return dt
    End Function

    <Task()> Public Shared Sub CrearVinoAnalisisDeposito(ByVal data As DataTable, ByVal services As ServiceProvider)
        If Not data Is Nothing AndAlso data.Rows.Count > 0 Then
            Dim DtVinoAn As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CreateDataVinoAnalisis, Nothing, services)
            For Each DrAn As DataRow In data.Select
                DtVinoAn.Rows.Clear()
                Dim DtAn As DataTable = New BdgAnalisisVariable().Filter(New FilterItem("IDAnalisis", FilterOperator.Equal, DrAn("IDAnalisis")))
                For Each DrVar As DataRow In DtAn.Select
                    If Length(DrAn("Lote")) > 0 AndAlso Length(DrAn("IDDeposito")) > 0 AndAlso Length(DrAn("FechaHora")) > 0 Then
                        Dim StSearch As New DataSearchVino(DrAn("IDDeposito"), DrAn("Lote"), DrAn("FechaHora"))
                        Dim DrNew As DataRow = DtVinoAn.NewRow
                        DrNew("IDVino") = ProcessServer.ExecuteTask(Of DataSearchVino, Guid)(AddressOf SearchVino, StSearch, services)
                        DrNew("IDAnalisis") = DrAn("IDAnalisis")
                        DrNew("Fecha") = DrAn("FechaHora")
                        DrNew("IDVariable") = DrVar("IDVariable")
                        DrNew("Orden") = DrVar("Orden")
                        If IsNumeric(DrAn(DrVar("IDVariable"))) Then
                            DrNew("ValorNumerico") = Nz(DrAn(DrVar("IDVariable")), 0)
                            DrNew("Valor") = Nz(DrAn(DrVar("IDVariable")), String.Empty)
                        Else : DrNew("Valor") = Nz(DrAn(DrVar("IDVariable")), String.Empty)
                        End If
                        DtVinoAn.Rows.Add(DrNew)
                    End If
                Next
                DtVinoAn.AcceptChanges()
                If Not DtVinoAn Is Nothing AndAlso DtVinoAn.Rows.Count > 0 Then
                    Dim StData As New StCrearVinoAnalisisMasivo(DrAn("IDAnalisis"), String.Empty, DrAn("FechaHora"), DtVinoAn)
                    ProcessServer.ExecuteTask(Of StCrearVinoAnalisisMasivo)(AddressOf CrearVinoAnalisisMasivo, StData, services)
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Function CreateDataVinoAnalisis(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim DtVinoAnalisis As New DataTable
        DtVinoAnalisis.Columns.Add("IDVino", GetType(Guid))
        DtVinoAnalisis.Columns.Add("IDVinoAnalisis", GetType(Guid))
        DtVinoAnalisis.Columns.Add("IDAnalisis", GetType(String))
        DtVinoAnalisis.Columns.Add("Fecha", GetType(DateTime))
        DtVinoAnalisis.Columns.Add("NOperacion", GetType(String))
        DtVinoAnalisis.Columns.Add("CodigoFoss", GetType(String))
        DtVinoAnalisis.Columns.Add("Observaciones", GetType(String))
        DtVinoAnalisis.Columns.Add("IDVariable", GetType(String))
        DtVinoAnalisis.Columns.Add("Valor", GetType(String))
        DtVinoAnalisis.Columns.Add("ValorNumerico", GetType(Double))
        DtVinoAnalisis.Columns.Add("Orden", GetType(Integer))
        Return DtVinoAnalisis
    End Function

    <Serializable()> _
    Public Class DataSearchVino
        Public IDDeposito As String
        Public Lote As String
        Public Fecha As DateTime

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDDeposito As String, ByVal Lote As String, ByVal Fecha As DateTime)
            Me.IDDeposito = IDDeposito
            Me.Lote = Lote
            Me.Fecha = Fecha
        End Sub
    End Class

    <Task()> Public Shared Function SearchVino(ByVal data As DataSearchVino, ByVal services As ServiceProvider) As Guid
        Dim f As New Filter
        f.Add("IDDeposito", FilterOperator.Equal, data.IDDeposito)
        f.Add("Lote", FilterOperator.Equal, data.Lote)
        Dim datDepositosFecha As New BdgVino.StDepositosEnFecha(data.Fecha, f, False)
        Dim dtVinosActivos As DataTable = ProcessServer.ExecuteTask(Of BdgVino.StDepositosEnFecha, DataTable)(AddressOf BdgVino.DepositosEnFecha, datDepositosFecha, services)
        If Not dtVinosActivos Is Nothing AndAlso dtVinosActivos.Rows.Count > 0 Then
            Return dtVinosActivos.Select(String.Empty, "FechaVino DESC")(0)("IDVino")
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _VA
    Public Const IDVinoAnalisis As String = "IDVinoAnalisis"
    Public Const NVinoAnalisis As String = "NVinoAnalisis"
    Public Const IDVino As String = "IDVino"
    Public Const IDAnalisis As String = "IDAnalisis"
    Public Const Fecha As String = "Fecha"
    Public Const NOperacion As String = "NOperacion"
    Public Const CodigoFoss As String = "CodigoFoss"
    Public Const Observaciones As String = "Observaciones"
End Class