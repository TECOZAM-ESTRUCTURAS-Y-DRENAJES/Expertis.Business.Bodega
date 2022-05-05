Public MustInherit Class DocumentoBdgOperacion
    Inherits Solmicro.Expertis.Business.General.Document

    Public Cabecera As OperCab


#Region " Entidades "

    Public MustOverride Function ClaseOperacion() As BdgClaseOperacion
    Public MustOverride Function EntidadCabecera() As String
    Public MustOverride Function EntidadOperacionVino() As String
    Public MustOverride Function EntidadOperacionMaterial() As String
    Public MustOverride Function EntidadOperacionMaterialLote() As String
    Public MustOverride Function EntidadOperacionMOD() As String
    Public MustOverride Function EntidadOperacionCentro() As String
    Public MustOverride Function EntidadOperacionVarios() As String

    Public MustOverride Function EntidadOperacionVinoMaterial() As String
    Public MustOverride Function EntidadOperacionVinoMaterialLote() As String
    Public MustOverride Function EntidadOperacionVinoMOD() As String
    Public MustOverride Function EntidadOperacionVinoCentro() As String
    Public MustOverride Function EntidadOperacionVinoVarios() As String

#End Region

#Region " Datatables "

    Public ReadOnly Property dtOperacionVino() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionVino)
        End Get
    End Property


    Public ReadOnly Property dtOperacionMaterial() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionMaterial)
        End Get
    End Property

    Public ReadOnly Property dtOperacionMaterialLote() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionMaterialLote)
        End Get
    End Property

    Public ReadOnly Property dtOperacionMOD() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionMOD)
        End Get
    End Property

    Public ReadOnly Property dtOperacionCentro() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionCentro)
        End Get
    End Property

    Public ReadOnly Property dtOperacionVarios() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionVarios)
        End Get
    End Property



    Public ReadOnly Property dtOperacionVinoMaterial() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionVinoMaterial)
        End Get
    End Property

    Public ReadOnly Property dtOperacionVinoMaterialLote() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionVinoMaterialLote)
        End Get
    End Property

    Public ReadOnly Property dtOperacionVinoMOD() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionVinoMOD)
        End Get
    End Property

    Public ReadOnly Property dtOperacionVinoCentro() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionVinoCentro)
        End Get
    End Property

    Public ReadOnly Property dtOperacionVinoVarios() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionVinoVarios)
        End Get
    End Property

#End Region

#Region "PrimaryKeys"

    Public MustOverride Function PrimaryKeyCab() As String()

    Public MustOverride Function PrimaryKeyCentro() As String()

    Public MustOverride Function PrimaryKeyMaterial() As String()

    Public MustOverride Function PrimaryKeyMaterialLote() As String()

    Public MustOverride Function PrimaryKeyMOD() As String()

    Public MustOverride Function PrimaryKeyVarios() As String()

    Public MustOverride Function PrimaryKeyVino() As String()

    Public MustOverride Function PrimaryKeyVinoCentro() As String()

    Public MustOverride Function PrimaryKeyVinoMaterial() As String()

    Public MustOverride Function PrimaryKeyVinoMaterialLote() As String()

    Public MustOverride Function PrimaryKeyVinoMOD() As String()

    Public MustOverride Function PrimaryKeyVinoVarios() As String()


#End Region

#Region " Propiedades Nombres de Campos "

    Public ReadOnly Property FieldNOperacion() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "NOperacionPlan"
                Case BdgClaseOperacion.Real
                    Return "NOperacion"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldIDLineaMaterialGlobal() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "IDOperacionPlanMaterial"
                Case BdgClaseOperacion.Real
                    Return "IDOperacionMaterial"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldIDLineaMaterialGlobalLote() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "IDOperacionPlanMaterialLote"
                Case BdgClaseOperacion.Real
                    Return "IDOperacionMaterialLote"
            End Select
        End Get
    End Property


    Public ReadOnly Property FieldIDLineaMODGlobal() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "IDOperacionPlanMOD"
                Case BdgClaseOperacion.Real
                    Return "IDOperacionMOD"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldIDLineaCentroGlobal() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "IDOperacionPlanCentro"
                Case BdgClaseOperacion.Real
                    Return "IDOperacionCentro"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldIDLineaVariosGlobal() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "IDOperacionPlanVarios"
                Case BdgClaseOperacion.Real
                    Return "IDOperacionVarios"
            End Select
        End Get
    End Property


    Public ReadOnly Property FieldIDLineaMaterialLineas() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "IDOperacionVinoPlanMaterial"
                Case BdgClaseOperacion.Real
                    Return "IDVinoMaterial"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldIDLineaMaterialLineasLote() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "IDOperacionVinoPlanMaterialLote"
                Case BdgClaseOperacion.Real
                    Return "IDVinoMaterialLote"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldIDLineaMODLineas() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "IDOperacionVinoPlanMOD"
                Case BdgClaseOperacion.Real
                    Return "IDVinoMOD"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldIDLineaCentroLineas() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "IDOperacionVinoPlanCentro"
                Case BdgClaseOperacion.Real
                    Return "IDVinoCentro"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldIDLineaVariosLineas() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "IDOperacionVinoPlanVarios"
                Case BdgClaseOperacion.Real
                    Return "IDVinoVarios"
            End Select
        End Get
    End Property


    Public ReadOnly Property FieldIDLineaOperacionVino() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "IDLineaOperacionVinoPlan"
                Case BdgClaseOperacion.Real
                    Return "IDVino"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldChkImputacionGlobalMaterial() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "ImputacionGlobalMat"
                Case BdgClaseOperacion.Real
                    Return "ImputacionRealMaterial"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldChkImputacionGlobalMOD() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "ImputacionGlobalMOD"
                Case BdgClaseOperacion.Real
                    Return "ImputacionRealMOD"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldChkImputacionGlobalCentro() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "ImputacionGlobalCentro"
                Case BdgClaseOperacion.Real
                    Return "ImputacionRealCentro"
            End Select
        End Get
    End Property

    Public ReadOnly Property FieldChkImputacionGlobalVarios() As String
        Get
            Select Case Me.ClaseOperacion
                Case BdgClaseOperacion.Prevista
                    Return "ImputacionGlobalVarios"
                Case BdgClaseOperacion.Real
                    Return "ImputacionRealVarios"
            End Select
        End Get
    End Property

#End Region


#Region "Creación de instancias"

    Public Sub New(ByVal cab As OperCab, ByVal services As ServiceProvider)
        LoadEntityHeader(True, Cab)


        LoadEntitiesChilds(True)
    End Sub

    Public Sub New(ByVal UpdtCtx As UpdatePackage)
        LoadEntityHeader(False, Nothing, UpdtCtx)
        LoadEntitiesChilds(False, UpdtCtx)
    End Sub

    Public Sub New(ByVal ParamArray PrimaryKey() As Object)
        LoadEntityHeader(False, Nothing, Nothing, PrimaryKey)
        LoadEntitiesChilds(False)
    End Sub

    Protected Overridable Sub LoadEntityHeader(ByVal AddNew As Boolean, Optional ByVal Cab As OperCab = Nothing, Optional ByVal UpdtCtx As UpdatePackage = Nothing, Optional ByVal PrimaryKey() As Object = Nothing)
        If AddNew Then
            Dim oBusinessEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadCabecera)
            Dim dtCabeceras As DataTable = oBusinessEntity.AddNewForm
            'dtCabeceras.Rows.Add(dtCabeceras.NewRow)
            MyBase.AddHeader(EntidadCabecera, dtCabeceras)

            Me.Cabecera = Cab
        ElseIf UpdtCtx Is Nothing Then
            Dim oBusinessEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadCabecera)
            Dim dtCabeceras As DataTable = oBusinessEntity.SelOnPrimaryKey(PrimaryKey)
            MyBase.AddHeader(Me.GetType.Name, dtCabeceras)
        Else
            MyBase.AddHeader(Me.GetType.Name, UpdtCtx(EntidadCabecera).First)
            'Dim PKCabecera() As String = PrimaryKeyCab()
            'MergeData(UpdtCtx, EntidadCabecera, PKCabecera, PKCabecera, True)
        End If
        LoadTipoOperacion()
    End Sub

    Protected Overridable Sub LoadEntitiesChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
        LoadEntityChild(AddNew, EntidadOperacionVino, UpdtCtx)
        LoadEntityChild(AddNew, EntidadOperacionCentro, UpdtCtx)
        LoadEntityChild(AddNew, EntidadOperacionMaterial, UpdtCtx)
        LoadEntityChild(AddNew, EntidadOperacionMOD, UpdtCtx)
        LoadEntityChild(AddNew, EntidadOperacionVarios, UpdtCtx)
        LoadEntitiesGrandChilds(AddNew, UpdtCtx)
        'LoadEntitiesGrandChildsOperacionVino(AddNew, UpdtCtx)
    End Sub

    Public Overridable Sub LoadEntityChild(ByVal AddNew As Boolean, ByVal Entidad As String, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
        Dim oEntidad As BusinessHelper = BusinessHelper.CreateBusinessObject(Entidad)
        Dim Dt As DataTable
        If AddNew Then
            Dt = oEntidad.AddNew
        ElseIf UpdtCtx Is Nothing Then
            Dim PKCabecera() As String = PrimaryKeyCab()
            Dt = oEntidad.Filter(New FilterItem(PKCabecera(0), HeaderRow(PKCabecera(0))))
        Else
            Dim dtCabecera As DataTable = UpdtCtx(EntidadCabecera).First
            If Not dtCabecera Is Nothing AndAlso dtCabecera.Rows.Count > 0 AndAlso dtCabecera.Rows(0).RowState = DataRowState.Added Then
                Dim PKCabecera() As String = PrimaryKeyCab()
                Dt = MergeDataHeaderNew(UpdtCtx, Entidad, PKCabecera, PKCabecera, False)
            Else
                Dim PKCabecera() As String = PrimaryKeyCab()
                Dt = MergeData(UpdtCtx, Entidad, PKCabecera, PKCabecera, False)
            End If

        End If
        Me.Add(oEntidad.GetType.Name, Dt)


        Select Case Entidad
            Case EntidadOperacionMaterial()
                Dim dtEntidadLote As DataTable
                Dim oEntidadLote As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadOperacionMaterialLote)
                If AddNew Then
                    dtEntidadLote = oEntidadLote.AddNew
                ElseIf UpdtCtx Is Nothing Then
                    Dim filLote As New Filter
                    If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
                        Dim Ids(Dt.Rows.Count - 1) As Object
                        For i As Integer = 0 To Dt.Rows.Count - 1
                            Ids(i) = Dt.Rows(i)(Me.FieldIDLineaMaterialGlobal)
                        Next

                        filLote.Add(New InListFilterItem(Me.FieldIDLineaMaterialGlobal, Ids, FilterType.Guid))
                    End If
                    If filLote.Count = 0 Then filLote.Add(New NoRowsFilterItem)
                    dtEntidadLote = oEntidadLote.Filter(filLote)
                Else
                    Dim filLote As New Filter
                    If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
                        Dim Ids(Dt.Rows.Count - 1) As Object
                        For i As Integer = 0 To Dt.Rows.Count - 1
                            Ids(i) = Dt.Rows(i)(Me.FieldIDLineaMaterialGlobal)
                        Next

                        filLote.Add(New InListFilterItem(Me.FieldIDLineaMaterialGlobal, Ids, FilterType.Guid))
                    End If
                    If filLote.Count = 0 Then filLote.Add(New NoRowsFilterItem)
                    dtEntidadLote = oEntidadLote.Filter(filLote)

                    Dim lst As List(Of UpdatePackageItem) = (From c As UpdatePackageItem In UpdtCtx.ToList Where c.EntityName = EntidadOperacionMaterialLote() Select c).ToList
                    If Not lst Is Nothing AndAlso lst.Count > 0 Then
                        Dim dtEntidadLoteUpPkg As DataTable = lst(0).Data

                        For Each drLote As DataRow In dtEntidadLoteUpPkg.Rows
                            If drLote.RowState = DataRowState.Added Then
                                dtEntidadLote.ImportRow(drLote)
                            ElseIf drLote.RowState = DataRowState.Modified Then
                                Dim LotesModificar As List(Of DataRow) = (From c In dtEntidadLote _
                                                                          Where Not c.IsNull(Me.FieldIDLineaMaterialGlobalLote) AndAlso c(Me.FieldIDLineaMaterialGlobalLote) = drLote(Me.FieldIDLineaMaterialGlobalLote)).ToList()
                                If Not LotesModificar Is Nothing AndAlso LotesModificar.Count > 0 Then
                                    For Each drLoteDel As DataRow In LotesModificar
                                        dtEntidadLote.Rows.Remove(drLoteDel)
                                    Next
                                End If

                                dtEntidadLote.ImportRow(drLote)
                            End If
                        Next

                        UpdtCtx.Remove(lst(0))
                    End If
                End If
                Me.Add(oEntidadLote.GetType.Name, dtEntidadLote)
        End Select
    End Sub

    Public Overridable Sub LoadEntitiesGrandChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
        Dim PKLineas() As String = PrimaryKeyVino()
        Dim ids(dtOperacionVino.Rows.Count - 1) As Object
        For i As Integer = 0 To dtOperacionVino.Rows.Count - 1
            ids(i) = dtOperacionVino.Rows(i)(PKLineas(0))
        Next
        If ids.Length = 0 Then ids = New Object() {New Guid}
        LoadEntityGrandChild(PKLineas, ids, AddNew, Me.EntidadOperacionVinoCentro, UpdtCtx)
        LoadEntityGrandChild(PKLineas, ids, AddNew, Me.EntidadOperacionVinoMaterial, UpdtCtx)
        LoadEntityGrandChild(PKLineas, ids, AddNew, Me.EntidadOperacionVinoMOD, UpdtCtx)
        LoadEntityGrandChild(PKLineas, ids, AddNew, Me.EntidadOperacionVinoVarios, UpdtCtx)
    End Sub

    Protected Overridable Sub LoadEntityGrandChild(ByVal PKLineas() As String, ByVal Ids() As Object, ByVal AddNew As Boolean, ByVal Entidad As String, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
        Dim f As New Filter
        If Not Ids Is Nothing AndAlso Ids.Count > 0 Then f.Add(New InListFilterItem(PKLineas(0), Ids, FilterType.Guid))
        Dim oEntidad As BusinessHelper = BusinessHelper.CreateBusinessObject(Entidad)
        Dim dtEntidad As DataTable
        If AddNew Then
            dtEntidad = oEntidad.AddNew
        ElseIf UpdtCtx Is Nothing Then
            If f.Count = 0 Then f.Add(New NoRowsFilterItem)
            dtEntidad = oEntidad.Filter(f)
        Else
            Dim dtCabecera As DataTable = UpdtCtx(EntidadCabecera).First
            If Not dtCabecera Is Nothing AndAlso dtCabecera.Rows.Count > 0 AndAlso dtCabecera.Rows(0).RowState = DataRowState.Added Then
                Dim PKCabecera() As String = PrimaryKeyCab()
                dtEntidad = MergeDataHeaderNew(UpdtCtx, Entidad, PKCabecera, PKCabecera, False)
            Else
                Dim PKCabecera() As String = PrimaryKeyCab()
                dtEntidad = MergeData(UpdtCtx, Entidad, PKCabecera, PKCabecera, False)
            End If
        End If
        Me.Add(oEntidad.GetType.Name, dtEntidad)
    End Sub

    'Public Overridable Sub LoadEntitiesGrandChildsOperacionVino(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
    '    Dim LstPKS As New List(Of DataPKS)
    '    If Not Me.GetOperacionVinoDestino Is Nothing Then
    '        For Each Dr As DataRow In Me.GetOperacionVinoDestino
    '            LstPKS.Add(New DataPKS(Dr("NOperacion"), Dr("IDVino")))
    '        Next
    '    End If

    '    '    Dim PKLineas() As String = PrimaryKeyVino()
    '    '    Dim ids(dtOperacionVino.Rows.Count - 1) As Object
    '    '    For i As Integer = 0 To dtOperacionVino.Rows.Count - 1
    '    '        ids(i) = dtOperacionVino.Rows(i)(PKLineas(0))
    '    '    Next

    '    LoadEntityGrandChild(LstPKS, AddNew, EntidadOperacionVinoCentro, UpdtCtx)
    '    LoadEntityGrandChild(LstPKS, AddNew, EntidadOperacionVinoMaterial, UpdtCtx)
    '    LoadEntityGrandChild(LstPKS, AddNew, EntidadOperacionVinoMOD, UpdtCtx)
    '    LoadEntityGrandChild(LstPKS, AddNew, EntidadOperacionVinoVarios, UpdtCtx)
    '    LoadEntityGrandChild(LstPKS, AddNew, EntidadOperacionVinoAnalisis, UpdtCtx)
    'End Sub

#End Region


#Region "Funcionalidad Pública"

#Region "Datos Tipo Operación"

    Public Overridable Sub LoadTipoOperacion()
        If Length(Me.HeaderRow("IDTipoOperacion")) > 0 Then
            Dim DtTipo As DataTable = New BdgTipoOperacion().SelOnPrimaryKey(Me.HeaderRow("IDTipoOperacion"))
            GetTipoOperTipoMov = DtTipo.Rows(0)("TipoMovimiento")
            GetTipoOperPermitirMerma = DtTipo.Rows(0)("PermitirMerma")
            GetTipoOperMermaMax = DtTipo.Rows(0)("PorcMermaMaxima")
            GetTipoOperRequiereOF = DtTipo.Rows(0)("RequiereOF")
            GetTipoOperEstadoVino = Nz(DtTipo.Rows(0)("IDEstadoVino"), String.Empty)
            Dim f As New Filter
            f.Add(New StringFilterItem("IDTipoOperacion", Me.HeaderRow("IDTipoOperacion")))
            Dim dtReparto As DataTable = New BdgTipoOperacionRepartoCoste().Filter(f)
            If Not dtReparto Is Nothing AndAlso dtReparto.Rows.Count > 0 Then
                GetUsarRepartoTipoOperacion = True
            Else
                GetUsarRepartoTipoOperacion = False
            End If
        End If
    End Sub

    Private MIntTipoMov As Integer?
    Public Property GetTipoOperTipoMov() As Integer
        Get
            Return MIntTipoMov
        End Get
        Set(ByVal value As Integer)
            MIntTipoMov = value
        End Set
    End Property

    Private MBlnPermitirMerma As Boolean?
    Public Property GetTipoOperPermitirMerma() As Boolean
        Get
            Return MBlnPermitirMerma
        End Get
        Set(ByVal value As Boolean)
            MBlnPermitirMerma = value
        End Set
    End Property

    Private MDblPorcMermaMax As Double?
    Public Property GetTipoOperMermaMax() As Double
        Get
            Return MDblPorcMermaMax
        End Get
        Set(ByVal value As Double)
            MDblPorcMermaMax = value
        End Set
    End Property

    Private MBlnRequiereOF As Boolean
    Public Property GetTipoOperRequiereOF() As Boolean
        Get
            Return MBlnRequiereOF
        End Get
        Set(ByVal value As Boolean)
            MBlnRequiereOF = value
        End Set
    End Property

    Private MIDEstadoVino As String
    Public Property GetTipoOperEstadoVino() As String
        Get
            Return MIDEstadoVino
        End Get
        Set(ByVal value As String)
            MIDEstadoVino = value
        End Set
    End Property

    Private MUsarRepartoTipoOperacion As Boolean
    Public Property GetUsarRepartoTipoOperacion() As Boolean
        Get
            Return MUsarRepartoTipoOperacion
        End Get
        Set(ByVal value As Boolean)
            MUsarRepartoTipoOperacion = value
        End Set
    End Property

#End Region

#Region "Datos Origenes / Destinos"

    Public Overridable Function GetOperacionVinoDestino() As DataRow()
        If Not Me.dtOperacionVino Is Nothing AndAlso Me.dtOperacionVino.Rows.Count > 0 Then
            Return Me.dtOperacionVino.Select("Destino = 1", Nothing)
        Else : Return Nothing
        End If
    End Function

    Public Overridable Function GetOperacionVinoOrigen() As DataRow()
        If Not Me.dtOperacionVino Is Nothing AndAlso Me.dtOperacionVino.Rows.Count > 0 Then
            Return Me.dtOperacionVino.Select("Destino = 0", Nothing)
        Else : Return Nothing
        End If
    End Function

#End Region

#End Region


#Region " MergeData especial para cuando estamos creando la operación "

    Public Function MergeDataHeaderNew(ByVal updtCtx As UpdatePackage, _
                            ByVal businessEntity As String, _
                            ByVal primaryFields() As String, _
                            ByVal secondaryFields() As String, _
                            ByVal autonumeric As Boolean) As DataTable

        Dim oBusinessEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(businessEntity)
        '//NOTA: La diferencia con el MergeData del document, es que éste devuelve no devuelve filas de la BBDD.
        Dim dtBusinessData As DataTable = oBusinessEntity.Filter(New NoRowsFilterItem)

        For i As Integer = updtCtx.Count - 1 To 0 Step -1
            Dim oUD As UpdatePackageItem = updtCtx(i)
            Select Case oUD.EntityName
                Case businessEntity
                    Dim dtKeys As DataTable = oBusinessEntity.PrimaryKeyTable

                    'TODO sacar keys de aqui
                    Dim Keys(dtKeys.Columns.Count - 1) As String
                    For j As Integer = 0 To dtKeys.Columns.Count - 1
                        Keys(j) = dtKeys.Columns(j).ColumnName
                    Next

                    Dim dvBusinessData As New DataView(dtBusinessData, Nothing, Strings.Join(Keys, ", "), DataViewRowState.CurrentRows)
                    For Each oRw As DataRow In oUD.Data.Rows
                        If oRw.RowState = DataRowState.Modified Then
                            Dim KeyValues(Keys.Length - 1) As Object
                            For j As Integer = 0 To Keys.Length - 1
                                KeyValues(j) = oRw(Keys(j))
                            Next

                            Dim idx As Integer = dvBusinessData.Find(KeyValues)
                            If idx >= 0 Then
                                dtBusinessData.Rows.Remove(dvBusinessData(idx).Row)
                                dtBusinessData.ImportRow(oRw)
                            End If
                        ElseIf oRw.RowState = DataRowState.Added Then
                            If autonumeric Then
                                If Length(oRw(Keys(0))) = 0 Then oRw(Keys(0)) = DAL.AdminData.GetAutoNumeric
                            End If

                            '//Nos aseguramos la relación entre cabecera y lineas
                            For j As Integer = 0 To primaryFields.Length - 1
                                oRw(secondaryFields(j)) = HeaderRow(primaryFields(j))
                            Next

                            dtBusinessData.ImportRow(oRw)
                        End If
                    Next
                    updtCtx.Remove(oUD)
            End Select
        Next

        Return dtBusinessData
    End Function

#End Region

End Class
