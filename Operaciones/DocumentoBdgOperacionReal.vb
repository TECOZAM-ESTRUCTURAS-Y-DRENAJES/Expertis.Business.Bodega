Public Class DocumentoBdgOperacionReal
    Inherits DocumentoBdgOperacion

#Region "Datos Documento BdgOperacion"

#Region "Entidades"

    Public Overrides Function ClaseOperacion() As BusinessEnum.BdgClaseOperacion
        Return BdgClaseOperacion.Real
    End Function

    Public Overrides Function EntidadCabecera() As String
        Return GetType(BdgOperacion).Name
    End Function

    Public Overrides Function EntidadOperacionCentro() As String
        Return GetType(BdgOperacionCentro).Name
    End Function

    Public Overrides Function EntidadOperacionMaterial() As String
        Return GetType(BdgOperacionMaterial).Name
    End Function

    Public Overrides Function EntidadOperacionMaterialLote() As String
        Return GetType(BdgOperacionMaterialLote).Name
    End Function

    Public Overrides Function EntidadOperacionMOD() As String
        Return GetType(BdgOperacionMOD).Name
    End Function

    Public Overrides Function EntidadOperacionVarios() As String
        Return GetType(BdgOperacionVarios).Name
    End Function

    Public Overrides Function EntidadOperacionVino() As String
        Return GetType(BdgOperacionVino).Name
    End Function

    Public Overrides Function EntidadOperacionVinoCentro() As String
        Return GetType(BdgVinoCentro).Name
    End Function

    Public Overrides Function EntidadOperacionVinoMaterial() As String
        Return GetType(BdgVinoMaterial).Name
    End Function

    Public Overrides Function EntidadOperacionVinoMaterialLote() As String
        Return GetType(BdgVinoMaterialLote).Name
    End Function

    Public Overrides Function EntidadOperacionVinoMOD() As String
        Return GetType(BdgVinoMod).Name
    End Function

    Public Overrides Function EntidadOperacionVinoVarios() As String
        Return GetType(BdgVinoVarios).Name
    End Function

    Public Function EntidadOperacionVinoAnalisis() As String
        Return GetType(BdgVinoAnalisis).Name
    End Function

    Public Function EntidadOperacionVinoAnalisisVariable() As String
        Return GetType(BdgVinoVariable).Name
    End Function

#End Region

#Region "Datatables"

    Public ReadOnly Property dtOperacionVinoAnalisis() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionVinoAnalisis)
        End Get
    End Property

    Public ReadOnly Property dtOperacionVinoAnalisisVariable() As DataTable
        Get
            Return MyBase.Item(EntidadOperacionVinoAnalisisVariable)
        End Get
    End Property

#End Region

#Region "PrimaryKeys"

    Public Overrides Function PrimaryKeyCab() As String()
        Return New String() {"NOperacion"}
    End Function

    Public Overrides Function PrimaryKeyCentro() As String()
        Return New String() {"IDOperacionCentro"}
    End Function

    Public Overrides Function PrimaryKeyMaterial() As String()
        Return New String() {"IDOperacionMaterial"}
    End Function

    Public Overrides Function PrimaryKeyMaterialLote() As String()
        Return New String() {"IDOperacionMaterialLote"}
    End Function

    Public Overrides Function PrimaryKeyMOD() As String()
        Return New String() {"IDOperacionMOD"}
    End Function

    Public Overrides Function PrimaryKeyVarios() As String()
        Return New String() {"IDOperacionVarios"}
    End Function

    Public Overrides Function PrimaryKeyVino() As String()
        Return New String() {"NOperacion", "IDVino"}
    End Function

    Public Overrides Function PrimaryKeyVinoCentro() As String()
        Return New String() {"IDVinoCentro"}
    End Function

    Public Overrides Function PrimaryKeyVinoMaterial() As String()
        Return New String() {"IDVinoMaterial"}
    End Function

    Public Overrides Function PrimaryKeyVinoMaterialLote() As String()
        Return New String() {"IDVinoMaterialLote"}
    End Function

    Public Overrides Function PrimaryKeyVinoMOD() As String()
        Return New String() {"IDVinoMOD"}
    End Function

    Public Overrides Function PrimaryKeyVinoVarios() As String()
        Return New String() {"IDVinoVarios"}
    End Function

    Public Function PrimaryKeyVinoAnalisis() As String()
        Return New String() {"IDVinoAnalisis"}
    End Function

    Public Function PrimaryKeyVinoVariable() As String()
        Return New String() {"IDVinoAnalisis", "IDVariable"}
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


    Public Overrides Sub LoadEntitiesGrandChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
        Dim LstPKS As New List(Of DataPKS)
        If Not Me.GetOperacionVinoDestino Is Nothing Then
            For Each Dr As DataRow In Me.GetOperacionVinoDestino
                LstPKS.Add(New DataPKS(Dr("NOperacion"), Dr("IDVino")))
            Next
        End If
        LoadEntityGrandChild(LstPKS, AddNew, EntidadOperacionVinoCentro, UpdtCtx)
        LoadEntityGrandChild(LstPKS, AddNew, EntidadOperacionVinoMaterial, UpdtCtx)
        LoadEntityGrandChild(LstPKS, AddNew, EntidadOperacionVinoMOD, UpdtCtx)
        LoadEntityGrandChild(LstPKS, AddNew, EntidadOperacionVinoVarios, UpdtCtx)
        LoadEntityGrandChild(LstPKS, AddNew, EntidadOperacionVinoAnalisis, UpdtCtx)
    End Sub

    <Serializable()> _
    Public Class DataPKS
        Public NOperacion As String
        Public IDVino As Guid

        Public Sub New()
        End Sub
        Public Sub New(ByVal NOperacion As String, ByVal IDVino As Guid)
            Me.NOperacion = NOperacion
            Me.IDVino = IDVino
        End Sub
    End Class

    Public Sub LoadEntityGrandChild(ByVal Ids As List(Of DataPKS), ByVal AddNew As Boolean, ByVal Entidad As String, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
        Dim oEntidad As BusinessHelper = BusinessHelper.CreateBusinessObject(Entidad)
        Dim dtEntidad As DataTable
        If AddNew Then
            dtEntidad = oEntidad.AddNew
        ElseIf UpdtCtx Is Nothing Then
            Dim FilR As New Filter(FilterUnionOperator.Or)
            For Each StPK As DataPKS In Ids
                Dim FilPK As New Filter
                FilPK.Add("NOperacion", FilterOperator.Equal, StPK.NOperacion)
                FilPK.Add("IDVino", FilterOperator.Equal, StPK.IDVino)
                FilR.Add(FilPK)
            Next
            If FilR.Count = 0 Then FilR.Add(New NoRowsFilterItem)
            dtEntidad = oEntidad.Filter(FilR)
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
        Select Case Entidad
            Case EntidadOperacionVinoMaterial()
                'Filtro de unión de datos entre el material y lote
                Dim dtEntidadLote As DataTable
                Dim oEntidadLote As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadOperacionVinoMaterialLote)
                If AddNew Then
                    dtEntidadLote = oEntidadLote.AddNew
                ElseIf UpdtCtx Is Nothing Then
                    Dim filLote As New Filter
                    If Not dtEntidad Is Nothing AndAlso dtEntidad.Rows.Count > 0 Then
                        Dim IdsMat(dtEntidad.Rows.Count - 1) As Object
                        For i As Integer = 0 To dtEntidad.Rows.Count - 1
                            IdsMat(i) = dtEntidad.Rows(i)("IDVinoMaterial")
                        Next

                        filLote.Add(New InListFilterItem("IDVinoMaterial", IdsMat, FilterType.Guid))
                    End If
                    If filLote.Count = 0 Then filLote.Add(New NoRowsFilterItem)
                    dtEntidadLote = oEntidadLote.Filter(filLote)
                Else
                    Dim filLote As New Filter
                    If Not dtEntidad Is Nothing AndAlso dtEntidad.Rows.Count > 0 Then
                        Dim IdsMat(dtEntidad.Rows.Count - 1) As Object
                        For i As Integer = 0 To dtEntidad.Rows.Count - 1
                            IdsMat(i) = dtEntidad.Rows(i)("IDVinoMaterial")
                        Next

                        filLote.Add(New InListFilterItem("IDVinoMaterial", IdsMat, FilterType.Guid))
                    End If
                    If filLote.Count = 0 Then filLote.Add(New NoRowsFilterItem)
                    dtEntidadLote = oEntidadLote.Filter(filLote)

                    Dim lst As List(Of UpdatePackageItem) = (From c As UpdatePackageItem In UpdtCtx.ToList Where c.EntityName = EntidadOperacionVinoMaterialLote() Select c).ToList
                    If Not lst Is Nothing AndAlso lst.Count > 0 Then
                        Dim dtEntidadLoteUpPkg As DataTable = lst(0).Data
                        For Each drLote As DataRow In dtEntidadLoteUpPkg.Rows
                            If drLote.RowState = DataRowState.Added Then
                                dtEntidadLote.ImportRow(drLote)
                            ElseIf drLote.RowState = DataRowState.Modified Then
                                Dim LotesModificar As List(Of DataRow) = (From c In dtEntidadLote _
                                                                          Where Not c.IsNull("IDVinoMaterialLote") AndAlso c("IDVinoMaterialLote") = drLote("IDVinoMaterialLote")).ToList()
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
            Case EntidadOperacionVinoAnalisis()
                Dim DtEntidadVariable As DataTable
                Dim oEntidadVariable As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadOperacionVinoAnalisisVariable)
                If AddNew Then
                    DtEntidadVariable = oEntidadVariable.AddNew
                ElseIf UpdtCtx Is Nothing Then
                    Dim FilAnalisis As New Filter
                    If Not dtEntidad Is Nothing AndAlso dtEntidad.Rows.Count > 0 Then
                        Dim IdsAnalisis(dtEntidad.Rows.Count - 1) As Object
                        For i As Integer = 0 To dtEntidad.Rows.Count - 1
                            IdsAnalisis(i) = dtEntidad.Rows(i)("IDVinoAnalisis")
                        Next
                        FilAnalisis.Add(New InListFilterItem("IDVinoAnalisis", IdsAnalisis, FilterType.Guid))
                    End If
                    If FilAnalisis.Count = 0 Then FilAnalisis.Add(New NoRowsFilterItem)
                    DtEntidadVariable = oEntidadVariable.Filter(FilAnalisis)
                Else

                    Dim FilAnalisis As New Filter
                    If Not dtEntidad Is Nothing AndAlso dtEntidad.Rows.Count > 0 Then
                        Dim IdsAnalisis(dtEntidad.Rows.Count - 1) As Object
                        For i As Integer = 0 To dtEntidad.Rows.Count - 1
                            IdsAnalisis(i) = dtEntidad.Rows(i)("IDVinoAnalisis")
                        Next
                        FilAnalisis.Add(New InListFilterItem("IDVinoAnalisis", IdsAnalisis, FilterType.Guid))
                    End If
                    If FilAnalisis.Count = 0 Then FilAnalisis.Add(New NoRowsFilterItem)
                    DtEntidadVariable = oEntidadVariable.Filter(FilAnalisis)

                    Dim lst As List(Of UpdatePackageItem) = (From c As UpdatePackageItem In UpdtCtx.ToList Where c.EntityName = EntidadOperacionVinoAnalisisVariable() Select c).ToList
                    If Not lst Is Nothing AndAlso lst.Count > 0 Then
                        Dim dtEntidadVariableUpPkg As DataTable = lst(0).Data
                        For Each drVariable As DataRow In dtEntidadVariableUpPkg.Rows
                            If drVariable.RowState = DataRowState.Added Then
                                DtEntidadVariable.ImportRow(drVariable)
                            ElseIf drVariable.RowState = DataRowState.Modified Then
                                Dim VariablesModificar As List(Of DataRow) = (From c In DtEntidadVariable _
                                                                              Where Not c.IsNull("IDAnalisis") AndAlso Not c.IsNull("IDVariable") AndAlso _
                                                                                     c("IDAnalisis") = drVariable("IDAnalisis") AndAlso _
                                                                                     c("IDVariable") = drVariable("IDVariable")).ToList()
                                If Not VariablesModificar Is Nothing AndAlso VariablesModificar.Count > 0 Then
                                    For Each drVariableDel As DataRow In VariablesModificar
                                        DtEntidadVariable.Rows.Remove(drVariableDel)
                                    Next
                                End If

                                DtEntidadVariable.ImportRow(drVariable)
                            End If
                        Next
                        UpdtCtx.Remove(lst(0))
                    End If
                End If
                Me.Add(oEntidadVariable.GetType.Name, DtEntidadVariable)
        End Select
    End Sub

#End Region


End Class
