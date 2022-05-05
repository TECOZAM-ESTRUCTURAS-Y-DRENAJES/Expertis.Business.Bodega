Public Class DocumentoEntradaVino
    Inherits Document

#Region "Datos Documento Entrada Vino"

    'Datatables de Entrada Vino
    Public Fincas As DataTable
    Public Depositos As DataTable
    Public Variedades As DataTable
    Public Analisis As DataTable

    Public Function EntidadCabecera() As String
        Return GetType(BdgEntradaVino).Name
    End Function

    Public Function EntidadFincas() As String
        Return GetType(BdgEntradaVinoFinca).Name
    End Function

    Public Function EntidadDepositos() As String
        Return GetType(BdgEntradaVinoDeposito).Name
    End Function

    Public Function EntidadVariedades() As String
        Return GetType(BdgEntradaVinoVariedad).Name
    End Function

    Public Function EntidadAnalisis() As String
        Return GetType(BdgEntradaVinoAnalisis).Name
    End Function

    Public ReadOnly Property dtFincas() As DataTable
        Get
            Return MyBase.Item(EntidadFincas)
        End Get
    End Property

    Public ReadOnly Property dtDepositos() As DataTable
        Get
            Return MyBase.Item(EntidadDepositos)
        End Get
    End Property

    Public ReadOnly Property dtVariedades() As DataTable
        Get
            Return MyBase.Item(EntidadVariedades)
        End Get
    End Property

    Public ReadOnly Property dtAnalisis() As DataTable
        Get
            Return MyBase.Item(EntidadAnalisis)
        End Get
    End Property

#End Region

#Region "PrimaryKeys de Entrada Vino"

    Public Function PrimaryKeyCab() As String()
        Return New String() {"NEntrada"}
    End Function

    Public Function PrimaryKeyFincas() As String()
        Return New String() {"NEntrada, IDFinca"}
    End Function

    Public Function PrimaryKeyDepositos() As String()
        Return New String() {"NEntrada, IDDeposito"}
    End Function

    Public Function PrimaryKeyVariedades() As String()
        Return New String() {"NEntrada, IDVariedad"}
    End Function

    Public Function PrimaryKeyAnalisis() As String()
        Return New String() {"NEntrada, IDVariable"}
    End Function

#End Region

#Region " Creación de instancias "

    '//New a utilizar desde presentación (utilizado por el motor para realizar las actualizaciones de los elementos cabecera/lineas)
    Public Sub New(ByVal UpdtCtx As UpdatePackage)
        '// al crearse la bola desde el updatepackage, se debieran eliminar los conjuntos de datos que se van a tratar
        '// y dejar que el motor trate el resto
        LoadEntityHeader(False, UpdtCtx)
        LoadEntitiesChilds(False, UpdtCtx)
    End Sub

    '//New utilizado para obtener un Documento alamacenado en la BBDD.
    Public Sub New(ByVal ParamArray PrimaryKey() As Object)
        LoadEntityHeader(False, Nothing, PrimaryKey)
        LoadEntitiesChilds(False)
    End Sub

    Protected Overridable Sub LoadEntityHeader(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing, Optional ByVal PrimaryKey() As Object = Nothing)
        If AddNew Then '//New de Procesos
            Dim oBusinessEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadCabecera)
            Dim dtCabeceras As DataTable = oBusinessEntity.AddNew

            dtCabeceras.Rows.Add(dtCabeceras.NewRow)
            MyBase.AddHeader(EntidadCabecera, dtCabeceras)     '//Creamos el HeaderRow
        ElseIf UpdtCtx Is Nothing Then  '//New de PrimaryKey
            Dim oBusinessEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadCabecera)
            Dim dtCabeceras As DataTable = oBusinessEntity.SelOnPrimaryKey(PrimaryKey)
            MyBase.AddHeader(Me.GetType.Name, dtCabeceras)
        Else '//New del formulario
            MyBase.AddHeader(Me.GetType.Name, UpdtCtx(EntidadCabecera).First)
            Dim PKCabecera() As String = PrimaryKeyCab()
            MergeData(UpdtCtx, EntidadCabecera, PKCabecera, PKCabecera, True)
        End If
    End Sub

    Protected Overridable Sub LoadEntitiesChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
        LoadEntityChild(AddNew, EntidadFincas, UpdtCtx)
        LoadEntityChild(AddNew, EntidadDepositos, UpdtCtx)
        LoadEntityChild(AddNew, EntidadVariedades, UpdtCtx)
        LoadEntityChild(AddNew, EntidadAnalisis, UpdtCtx)
    End Sub

    Protected Overridable Sub LoadEntityChild(ByVal AddNew As Boolean, ByVal Entidad As String, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
        Dim oEntidad As BusinessHelper = BusinessHelper.CreateBusinessObject(Entidad)
        Dim Dt As DataTable
        If AddNew Then  '//New de Procesos
            Dt = oEntidad.AddNew
        ElseIf UpdtCtx Is Nothing Then  '//New de PrimaryKey
            Dim PKCabecera() As String = PrimaryKeyCab()
            Dt = oEntidad.Filter(New FilterItem(PKCabecera(0), HeaderRow(PKCabecera(0))))
        Else  '//New del formulario
            '// al crearse la bola desde el updatepackage, se debieran eliminar los conjuntos de datos que se van a tratar
            '// y dejar que el motor trate el resto  (MergeData)
            Dim PKCabecera() As String = PrimaryKeyCab()
            Dt = MergeData(UpdtCtx, Entidad, PKCabecera, PKCabecera, True)
        End If
        Me.Add(oEntidad.GetType.Name, Dt)
    End Sub

#End Region

End Class