
Public Class BdgDocumentoEntrada
    Inherits Document

#Region "CTOR"

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

#End Region

#Region "Entidades"

    Public Function EntidadCabecera() As String
        Return GetType(BdgEntrada).Name
    End Function

    Public Function EntidadDepositos() As String
        Return GetType(BdgEntradaDeposito).Name
    End Function

    Public Function EntidadAnalisis() As String
        Return GetType(BdgEntradaAnalisis).Name
    End Function

    Public Function EntidadCartillista() As String
        Return GetType(BdgEntradaCartillista).Name
    End Function

    Public Function EntidadFinca() As String
        Return GetType(BdgEntradaFinca).Name
    End Function

    Public Function EntidadProveedor() As String
        Return GetType(BdgEntradaProveedor).Name
    End Function

    Public Function EntidadVariedad() As String
        Return GetType(BdgEntradaVariedad).Name
    End Function

    Public Function EntidadFacturacion() As String
        Return GetType(BdgEntradaFacturacion).Name
    End Function

#End Region

#Region "PK"

    Public Function PrimaryKeyCab() As String()
        Return New String() {"IDEntrada"}
    End Function

    Public Function PrimaryKeyAnalisis() As String()
        Return New String() {"IDEntrada", "IDVariable"}
    End Function

    Public Function PrimaryKeyCartillista() As String()
        Return New String() {"IDEntrada", "IDCartillista"}
    End Function

    Public Function PrimaryKeyDeposito() As String()
        Return New String() {"IDEntrada", "IDDeposito"}
    End Function

    Public Function PrimaryKeyFinca() As String()
        Return New String() {"IDEntrada", "IDFinca"}
    End Function

    Public Function PrimaryKeyProveedor() As String()
        Return New String() {"IDEntrada", "IDProveedor"}
    End Function

    Public Function PrimaryKeyVariedad() As String()
        Return New String() {"IDEntrada", "IDVariedad"}
    End Function

    Public Function PrimaryKeyFacturacion() As String()
        Return New String() {"IDEntradaFacturacion"}
    End Function
#End Region

#Region "DataTable"

    Public ReadOnly Property dtCabecera() As DataTable
        Get
            Return MyBase.Item(EntidadCabecera)
        End Get
    End Property

    Public ReadOnly Property dtDepositos() As DataTable
        Get
            Return MyBase.Item(EntidadDepositos)
        End Get
    End Property

    Public ReadOnly Property dtCartillistas() As DataTable
        Get
            Return MyBase.Item(EntidadCartillista)
        End Get
    End Property

    Public ReadOnly Property dtFincas() As DataTable
        Get
            Return MyBase.Item(EntidadFinca)
        End Get
    End Property

    Public ReadOnly Property dtProveedores() As DataTable
        Get
            Return MyBase.Item(EntidadProveedor)
        End Get
    End Property

    Public ReadOnly Property dtVariedades() As DataTable
        Get
            Return MyBase.Item(EntidadVariedad)
        End Get
    End Property

    Public ReadOnly Property dtAnalisis() As DataTable
        Get
            Return MyBase.Item(EntidadAnalisis)
        End Get
    End Property

    Public ReadOnly Property dtFacturacion() As DataTable
        Get
            Return MyBase.Item(EntidadFacturacion)
        End Get
    End Property

#End Region

#Region "Load"

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
        LoadEntityChild(AddNew, EntidadAnalisis, UpdtCtx)
        LoadEntityChild(AddNew, EntidadCartillista, UpdtCtx)
        LoadEntityChild(AddNew, EntidadDepositos, UpdtCtx)
        LoadEntityChild(AddNew, EntidadFinca, UpdtCtx)
        LoadEntityChild(AddNew, EntidadProveedor, UpdtCtx)
        LoadEntityChild(AddNew, EntidadVariedad, UpdtCtx)
        LoadEntityChild(AddNew, EntidadFacturacion, UpdtCtx)
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

            Dt = MergeData(UpdtCtx, Entidad, PKCabecera, PKCabecera, Entidad <> EntidadFacturacion())
        End If
        Me.Add(oEntidad.GetType.Name, Dt)
    End Sub

#End Region

End Class