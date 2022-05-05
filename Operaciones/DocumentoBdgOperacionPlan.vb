Public Class DocumentoBdgOperacionPlan
    Inherits DocumentoBdgOperacion

#Region "Datos Documento BdgOperacionPlan"

#Region "Entidades"

    Public Overrides Function ClaseOperacion() As BusinessEnum.BdgClaseOperacion
        Return BdgClaseOperacion.Prevista
    End Function

    Public Overrides Function EntidadCabecera() As String
        Return GetType(BdgOperacionPlan).Name
    End Function

    Public Overrides Function EntidadOperacionCentro() As String
        Return GetType(BdgOperacionPlanCentro).Name
    End Function

    Public Overrides Function EntidadOperacionMaterial() As String
        Return GetType(BdgOperacionPlanMaterial).Name
    End Function

    Public Overrides Function EntidadOperacionMaterialLote() As String
        Return GetType(BdgOperacionPlanMaterialLote).Name
    End Function

    Public Overrides Function EntidadOperacionMOD() As String
        Return GetType(BdgOperacionPlanMOD).Name
    End Function

    Public Overrides Function EntidadOperacionVarios() As String
        Return GetType(BdgOperacionPlanVarios).Name
    End Function

    Public Overrides Function EntidadOperacionVino() As String
        Return GetType(BdgOperacionVinoPlan).Name
    End Function

    Public Overrides Function EntidadOperacionVinoCentro() As String
        Return GetType(BdgOperacionVinoPlanCentro).Name
    End Function

    Public Overrides Function EntidadOperacionVinoMaterial() As String
        Return GetType(BdgOperacionVinoPlanMaterial).Name
    End Function

    Public Overrides Function EntidadOperacionVinoMaterialLote() As String
        Return GetType(BdgOperacionVinoPlanMaterialLote).Name
    End Function

    Public Overrides Function EntidadOperacionVinoMOD() As String
        Return GetType(BdgOperacionVinoPlanMOD).Name
    End Function

    Public Overrides Function EntidadOperacionVinoVarios() As String
        Return GetType(BdgOperacionVinoPlanVarios).Name
    End Function

#End Region

#Region "PrimaryKeys"

    Public Overrides Function PrimaryKeyCab() As String()
        Return New String() {"NOperacionPlan"}
    End Function

    Public Overrides Function PrimaryKeyCentro() As String()
        Return New String() {"IDOperacionPlanCentro"}
    End Function

    Public Overrides Function PrimaryKeyMaterial() As String()
        Return New String() {"IDOperacionPlanMaterial"}
    End Function

    Public Overrides Function PrimaryKeyMaterialLote() As String()
        Return New String() {"IDOperacionPlanMaterialLote"}
    End Function

    Public Overrides Function PrimaryKeyMOD() As String()
        Return New String() {"IDOperacionMOD"}
    End Function

    Public Overrides Function PrimaryKeyVarios() As String()
        Return New String() {"IDOperacionPlanVarios"}
    End Function

    Public Overrides Function PrimaryKeyVino() As String()
        Return New String() {"IDLineaOperacionVinoPlan"}
    End Function

    Public Overrides Function PrimaryKeyVinoCentro() As String()
        Return New String() {"IDOperacionVinoPlanCentro"}
    End Function

    Public Overrides Function PrimaryKeyVinoMaterial() As String()
        Return New String() {"IDOperacionVinoPlanMaterial"}
    End Function

    Public Overrides Function PrimaryKeyVinoMaterialLote() As String()
        Return New String() {"IDOperacionVinoPlanMaterialLote"}
    End Function

    Public Overrides Function PrimaryKeyVinoMOD() As String()
        Return New String() {"IDOperacionVinoPlanMOD"}
    End Function

    Public Overrides Function PrimaryKeyVinoVarios() As String()
        Return New String() {"IDOperacionVinoPlanVarios"}
    End Function

#End Region

#End Region

#Region "Creación de instancias"

    '//New a utilizar en los procesos de creación de elementos de tipo cabecera/lineas
    Public Sub New(ByVal Cabecera As OperCab, ByVal services As ServiceProvider)
        MyBase.New(Cabecera, services)
    End Sub
    '//New a utilizar desde presentación (utilizado por el motor para realizar las actualizaciones de los elementos cabecera/lineas)
    Public Sub New(ByVal UpdtCtx As Engine.BE.UpdatePackage)
        MyBase.New(UpdtCtx)
    End Sub
    '//New utilizado para obtener un Documento almacenado en la BBDD.
    Public Sub New(ByVal ParamArray PrimaryKey() As Object)
        MyBase.New(PrimaryKey)
    End Sub

    'Public Sub New(ByVal UpdtCtx As UpdatePackage)
    '    LoadEntityHeader(False, UpdtCtx)
    '    LoadEntitiesChilds(False, UpdtCtx)
    'End Sub

    'Public Sub New(ByVal ParamArray PrimaryKey() As Object)
    '    LoadEntityHeader(False, Nothing, PrimaryKey)
    '    LoadEntitiesChilds(False)
    'End Sub

    'Protected Overridable Sub LoadEntityHeader(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing, Optional ByVal PrimaryKey() As Object = Nothing)
    '    If AddNew Then
    '        Dim oBusinessEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadCabecera)
    '        Dim dtCabeceras As DataTable = oBusinessEntity.AddNew
    '        dtCabeceras.Rows.Add(dtCabeceras.NewRow)
    '        MyBase.AddHeader(EntidadCabecera, dtCabeceras)
    '    ElseIf UpdtCtx Is Nothing Then
    '        Dim oBusinessEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadCabecera)
    '        Dim dtCabeceras As DataTable = oBusinessEntity.SelOnPrimaryKey(PrimaryKey)
    '        MyBase.AddHeader(Me.GetType.Name, dtCabeceras)
    '    Else
    '        MyBase.AddHeader(Me.GetType.Name, UpdtCtx(EntidadCabecera).First)
    '        'Dim PKCabecera() As String = PrimaryKeyCab()
    '        'MergeData(UpdtCtx, EntidadCabecera, PKCabecera, PKCabecera, True)
    '    End If
    'End Sub

    'Protected Overridable Sub LoadEntitiesChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
    '    LoadEntityChild(AddNew, EntidadOperacionVino, UpdtCtx)
    '    LoadEntityChild(AddNew, EntidadCentro, UpdtCtx)
    '    LoadEntityChild(AddNew, EntidadMaterial, UpdtCtx)
    '    LoadEntityChild(AddNew, EntidadMOD, UpdtCtx)
    '    LoadEntityChild(AddNew, EntidadVarios, UpdtCtx)
    '    LoadEntitiesGrandChilds(AddNew, UpdtCtx)
    'End Sub

    'Protected Overridable Sub LoadEntityChild(ByVal AddNew As Boolean, ByVal Entidad As String, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
    '    Dim oEntidad As BusinessHelper = BusinessHelper.CreateBusinessObject(Entidad)
    '    Dim Dt As DataTable
    '    If AddNew Then
    '        Dt = oEntidad.AddNew
    '    ElseIf UpdtCtx Is Nothing Then
    '        Dim PKCabecera() As String = PrimaryKeyCab()
    '        Dt = oEntidad.Filter(New FilterItem(PKCabecera(0), HeaderRow(PKCabecera(0))))
    '    Else
    '        Dim dtCabecera As DataTable = UpdtCtx(EntidadCabecera).First
    '        If Not dtCabecera Is Nothing AndAlso dtCabecera.Rows.Count > 0 AndAlso dtCabecera.Rows(0).RowState = DataRowState.Added Then
    '            Dim PKCabecera() As String = PrimaryKeyCab()
    '            Dt = MergeDataHeaderNew(UpdtCtx, Entidad, PKCabecera, PKCabecera, False)
    '        Else
    '            Dim PKCabecera() As String = PrimaryKeyCab()
    '            Dt = MergeData(UpdtCtx, Entidad, PKCabecera, PKCabecera, False)
    '        End If
    '    End If
    '    Me.Add(oEntidad.GetType.Name, Dt)
    'End Sub

    'Protected Overridable Sub LoadEntitiesGrandChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
    '    Dim PKLineas() As String = PrimaryKeyVino()
    '    Dim ids(dtOperacionVino.Rows.Count - 1) As Object
    '    For i As Integer = 0 To dtOperacionVino.Rows.Count - 1
    '        ids(i) = dtOperacionVino.Rows(i)(PKLineas(0))
    '    Next
    '    If ids.Length = 0 Then ids = New Object() {New Guid}
    '    LoadEntityGrandChild(PKLineas, ids, AddNew, EntidadVinoCentro, UpdtCtx)
    '    LoadEntityGrandChild(PKLineas, ids, AddNew, EntidadVinoMaterial, UpdtCtx)
    '    LoadEntityGrandChild(PKLineas, ids, AddNew, EntidadVinoMOD, UpdtCtx)
    '    LoadEntityGrandChild(PKLineas, ids, AddNew, EntidadVinoVarios, UpdtCtx)
    'End Sub

    'Protected Overridable Sub LoadEntityGrandChild(ByVal PKLineas() As String, ByVal Ids() As Object, ByVal AddNew As Boolean, ByVal Entidad As String, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
    '    Dim f As New Filter
    '    f.Add(New InListFilterItem(PKLineas(0), Ids, FilterType.Guid))
    '    Dim oEntidad As BusinessHelper = BusinessHelper.CreateBusinessObject(Entidad)
    '    Dim dtEntidad As DataTable
    '    If AddNew Then
    '        dtEntidad = oEntidad.AddNew
    '    ElseIf UpdtCtx Is Nothing Then
    '        dtEntidad = oEntidad.Filter(f)
    '    Else
    '        Dim dtCabecera As DataTable = UpdtCtx(EntidadCabecera).First
    '        If Not dtCabecera Is Nothing AndAlso dtCabecera.Rows.Count > 0 AndAlso dtCabecera.Rows(0).RowState = DataRowState.Added Then
    '            Dim PKCabecera() As String = PrimaryKeyCab()
    '            dtEntidad = MergeDataHeaderNew(UpdtCtx, Entidad, PKCabecera, PKCabecera, False)
    '        Else
    '            Dim PKCabecera() As String = PrimaryKeyCab()
    '            dtEntidad = MergeData(UpdtCtx, Entidad, PKCabecera, PKCabecera, False)
    '        End If
    '    End If
    '    Me.Add(oEntidad.GetType.Name, dtEntidad)
    'End Sub

#End Region



    '#Region " MergeData especial para cuando estamos creando la operación "

    '    Public Function MergeDataHeaderNew(ByVal updtCtx As UpdatePackage, _
    '                            ByVal businessEntity As String, _
    '                            ByVal primaryFields() As String, _
    '                            ByVal secondaryFields() As String, _
    '                            ByVal autonumeric As Boolean) As DataTable

    '        Dim oBusinessEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(businessEntity)
    '        '//NOTA: La diferencia con el MergeData del document, es que éste devuelve no devuelve filas de la BBDD.
    '        Dim dtBusinessData As DataTable = oBusinessEntity.Filter(New NoRowsFilterItem)

    '        For i As Integer = updtCtx.Count - 1 To 0 Step -1
    '            Dim oUD As UpdatePackageItem = updtCtx(i)
    '            Select Case oUD.EntityName
    '                Case businessEntity
    '                    Dim dtKeys As DataTable = oBusinessEntity.PrimaryKeyTable

    '                    'TODO sacar keys de aqui
    '                    Dim Keys(dtKeys.Columns.Count - 1) As String
    '                    For j As Integer = 0 To dtKeys.Columns.Count - 1
    '                        Keys(j) = dtKeys.Columns(j).ColumnName
    '                    Next

    '                    Dim dvBusinessData As New DataView(dtBusinessData, Nothing, Strings.Join(Keys, ", "), DataViewRowState.CurrentRows)
    '                    For Each oRw As DataRow In oUD.Data.Rows
    '                        If oRw.RowState = DataRowState.Modified Then
    '                            Dim KeyValues(Keys.Length - 1) As Object
    '                            For j As Integer = 0 To Keys.Length - 1
    '                                KeyValues(j) = oRw(Keys(j))
    '                            Next

    '                            Dim idx As Integer = dvBusinessData.Find(KeyValues)
    '                            If idx >= 0 Then
    '                                dtBusinessData.Rows.Remove(dvBusinessData(idx).Row)
    '                                dtBusinessData.ImportRow(oRw)
    '                            End If
    '                        ElseIf oRw.RowState = DataRowState.Added Then
    '                            If autonumeric Then
    '                                If Length(oRw(Keys(0))) = 0 Then oRw(Keys(0)) = DAL.AdminData.GetAutoNumeric
    '                            End If

    '                            '//Nos aseguramos la relación entre cabecera y lineas
    '                            For j As Integer = 0 To primaryFields.Length - 1
    '                                oRw(secondaryFields(j)) = HeaderRow(primaryFields(j))
    '                            Next

    '                            dtBusinessData.ImportRow(oRw)
    '                        End If
    '                    Next
    '                    updtCtx.Remove(oUD)
    '            End Select
    '        Next

    '        Return dtBusinessData
    '    End Function

    '#End Region

End Class