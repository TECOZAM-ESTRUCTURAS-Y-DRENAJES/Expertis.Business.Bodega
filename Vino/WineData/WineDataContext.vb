
Imports System.Collections.Generic
Imports System.Data.Linq
Imports System.Data.Linq.Mapping
Imports System.Reflection


Partial Public Class WineDataContext

    <[Function](Name:="spBdgVinoEstructura")> _
    <ResultType(GetType(Vino))> _
    <ResultType(GetType(VinoEstructura))> _
    <ResultType(GetType(VinoTratamientos))> _
    <ResultType(GetType(VinoDetalle))> _
    Public Function GetWineStructure(<Parameter(Name:="IDVino", DbType:="UniqueIdentifier")> ByVal IDVino As Guid, _
                                     <Parameter(Name:="pFecha", DbType:="datetime")> ByVal Fecha As Date, _
                                     <Parameter(Name:="VerTratamientos", DbType:="bit")> ByVal VerTratamientos As Boolean, _
                                     <Parameter(Name:="CosteElaboracion", DbType:="bit")> ByVal CosteElaboracion As Boolean, _
                                     <Parameter(Name:="CosteEstanciaNave", DbType:="bit")> ByVal CosteEstanciaNave As Boolean, _
                                     <Parameter(Name:="CosteInicial", DbType:="bit")> ByVal CosteInicial As Boolean, _
                                     <Parameter(Name:="VerDetalleOperaciones", DbType:="bit")> ByVal VerDetalleOperaciones As Boolean, _
                                     <Parameter(Name:="VerDetalleAnalisis", DbType:="bit")> ByVal VerDetalleAnalisis As Boolean, _
                                     <Parameter(Name:="VerDetalleEntradasUva", DbType:="bit")> ByVal VerDetalleEntradasUva As Boolean, _
                                     <Parameter(Name:="VerDetalleEntradasVino", DbType:="bit")> ByVal VerDetalleEntradasVino As Boolean, _
                                     <Parameter(Name:="ConsiderarNoTrazar", DbType:="bit")> ByVal ConsiderarNoTrazar As Boolean, _
                                     <Parameter(Name:="ConsiderarCosteInicial", DbType:="bit")> ByVal ConsiderarCosteInicial As Boolean, _
                                     <Parameter(Name:="IDQueryCoste", DbType:="int")> ByVal IDQueryCoste As Integer) As IMultipleResults
        Dim result As IExecuteResult = Me.ExecuteMethodCall(Me, DirectCast(MethodInfo.GetCurrentMethod(), MethodInfo), IDVino, Fecha, VerTratamientos, CosteElaboracion, CosteEstanciaNave, CosteInicial, _
                                                            VerDetalleOperaciones, VerDetalleAnalisis, VerDetalleEntradasUva, VerDetalleEntradasVino, ConsiderarNoTrazar, ConsiderarCosteInicial, IDQueryCoste)
        Return DirectCast(result.ReturnValue, IMultipleResults)
    End Function

    <[Function](Name:="spBdgVinoOrigenes")> _
    <ResultType(GetType(OrigenVariedades))> _
    <ResultType(GetType(OrigenFincas))> _
    <ResultType(GetType(OrigenCompras))> _
    <ResultType(GetType(OrigenAniadas))> _
    <ResultType(GetType(OrigenEntradasUVA))> _
    Public Function VinoOrigen(<Parameter(Name:="idQuery", DbType:="Int")> ByVal idQuery As System.Nullable(Of Integer), _
                               <Parameter(Name:="QDepositoVino", DbType:="numeric(23,8)")> ByVal QDepositoVino As Nullable(Of Decimal)) As IMultipleResults
        Dim result As IExecuteResult = Me.ExecuteMethodCall(Me, DirectCast(MethodInfo.GetCurrentMethod(), MethodInfo), idQuery, QDepositoVino)
        Return DirectCast(result.ReturnValue, IMultipleResults)
    End Function

    <[Function](Name:="spBdgVinoEstructuraImplosion")> _
   <ResultType(GetType(Vino))> _
   <ResultType(GetType(VinoEstructura))> _
   <ResultType(GetType(VinoTrazabilidadSalidas))> _
   <ResultType(GetType(VinoTrazabilidadStock))> _
   <ResultType(GetType(VinoTrazabilidadLotes))> _
    Public Function GetWineStructureImplosion(<Parameter(Name:="IDVino", DbType:="UniqueIdentifier")> ByVal IDVino As Guid, _
                                     <Parameter(Name:="pFecha", DbType:="datetime")> ByVal Fecha As Date, _
                                     <Parameter(Name:="VerEstructura", DbType:="bit")> ByVal VerEstructura As Boolean, _
                                     <Parameter(Name:="VerSalidas", DbType:="bit")> ByVal VerSalidas As Boolean, _
                                     <Parameter(Name:="VerStock", DbType:="bit")> ByVal VerStock As Boolean, _
                                     <Parameter(Name:="VerLotes", DbType:="bit")> ByVal VerLotes As Boolean, _
                                     <Parameter(Name:="IDQueryCoste", DbType:="int")> ByVal IDQueryCoste As Integer) As IMultipleResults
        Dim result As IExecuteResult = Me.ExecuteMethodCall(Me, DirectCast(MethodInfo.GetCurrentMethod(), MethodInfo), IDVino, Fecha, VerEstructura, VerSalidas, VerStock, VerLotes, IDQueryCoste)
        Return DirectCast(result.ReturnValue, IMultipleResults)
    End Function

    <[Function](Name:="spBdgVinoExplosionArbol")> _
  <ResultType(GetType(VinoExplosionArbol))> _
  <ResultType(GetType(VinoExplosionArbolDetalle))> _
   Public Function GetExplosionDetalladaArbol( _
                                    <Parameter(Name:="pVino", DbType:="UniqueIdentifier")> ByVal IDVino As Guid, _
                                    <Parameter(Name:="pNivelMaximo", DbType:="int")> Optional ByVal NivelMaximo As Integer = 100, _
                                    <Parameter(Name:="pVerMateriales", DbType:="bit")> Optional ByVal VerMateriales As Boolean = False, _
                                    <Parameter(Name:="pVerCosteElaboracion", DbType:="bit")> Optional ByVal VerVerCosteElaboracion As Boolean = False, _
                                    <Parameter(Name:="pVerCosteEstanciaNave", DbType:="bit")> Optional ByVal VerCosteEstanciaNave As Boolean = False, _
                                    <Parameter(Name:="pVerCosteInicial", DbType:="bit")> Optional ByVal VerCosteInicial As Boolean = False, _
                                    <Parameter(Name:="pVerDetalleOperaciones", DbType:="bit")> Optional ByVal VerDetalleOperaciones As Boolean = True, _
                                    <Parameter(Name:="pVerDetalleAnalisis", DbType:="bit")> Optional ByVal VerDetalleAnalisis As Boolean = True, _
                                    <Parameter(Name:="pVerDetalleEntradasUva", DbType:="bit")> Optional ByVal VerDetalleEntradasUva As Boolean = True, _
                                    <Parameter(Name:="pVerDetalleEntradasVino", DbType:="bit")> Optional ByVal VerDetalleEntradasVino As Boolean = True, _
                                    <Parameter(Name:="pVerDetalleSalidasAV", DbType:="bit")> Optional ByVal VerDetalleSalidasAV As Boolean = True, _
                                    <Parameter(Name:="pVerDetalleOtrasEntradas", DbType:="bit")> Optional ByVal VerDetalleOtrasEntradas As Boolean = True, _
                                    <Parameter(Name:="pVerDetalleAjustes", DbType:="bit")> Optional ByVal VerDetalleAjustes As Boolean = True, _
                                    <Parameter(Name:="pVerDetalleOfs", DbType:="bit")> Optional ByVal VerDetalleOfs As Boolean = True, _
                                    <Parameter(Name:="pVerDetalleTransferencias", DbType:="bit")> Optional ByVal VerDetalleTransferencias As Boolean = True, _
                                    <Parameter(Name:="pVerDetalleInventarios", DbType:="bit")> Optional ByVal VerDetalleInventarios As Boolean = True, _
                                    <Parameter(Name:="pVerDetalleOtrasSalidas", DbType:="bit")> Optional ByVal VerDetalleOtrasSalidas As Boolean = True, _
                                    <Parameter(Name:="pVerPorcentajes", DbType:="bit")> Optional ByVal VerPorcentajes As Boolean = False) As IMultipleResults
        If NivelMaximo <= 0 Then NivelMaximo = 100
        Dim result As IExecuteResult = Me.ExecuteMethodCall(Me, DirectCast(MethodInfo.GetCurrentMethod(), MethodInfo), IDVino, NivelMaximo, _
                                                            VerMateriales, VerVerCosteElaboracion, VerCosteEstanciaNave, VerCosteInicial, _
                                                            VerDetalleOperaciones, VerDetalleAnalisis, VerDetalleEntradasUva, VerDetalleEntradasVino, _
                                                            VerDetalleSalidasAV, VerDetalleOtrasEntradas, VerDetalleAjustes, VerDetalleOfs, _
                                                            VerDetalleTransferencias, VerDetalleInventarios, VerDetalleOtrasSalidas, VerPorcentajes)
        Return DirectCast(result.ReturnValue, IMultipleResults)
    End Function

End Class

<Serializable()> _
 Public Class Vino

    Private _IDVino As System.Guid

    Private _Fecha As Date

    Private _IDDeposito As String

    Private _IDEstadoVino As String

    Private _Origen As Integer

    Private _TipoDeposito As Integer

    Private _IDArticulo As String

    Private _NOperacion As String

    Private _Ficticio As Boolean

    Private _Grupo As Boolean

    Private _IDUdMedida As String

    Private _CosteVariable As Decimal

    Private _CosteFiscal As Decimal

    Private _CosteTotal As Decimal

    Private _Lote As String

    Private _IDAlmacen As String

    Private _IDBarrica As String

    Private _DiasDeposito As Integer

    Private _DiasBarrica As Integer

    Private _DiasBotellero As Integer

    Private _QTotal As Double

    Private _CosteUnitarioInicialA As Decimal

    Private _FechaCosteUnitarioInicial As System.Nullable(Of Date)

    Private _QTotMermaCoste As Double

    Private _QTotMermaSinCoste As Double


    <Column(Storage:="_IDVino", DbType:="UniqueIdentifier NOT NULL", IsPrimaryKey:=True)> _
 Public Property IDVino() As System.Guid
        Get
            Return Me._IDVino
        End Get
        Set(ByVal value As System.Guid)
            If ((Me._IDVino = value) _
               = False) Then
                Me._IDVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_Fecha", DbType:="DateTime NOT NULL")> _
    Public Property Fecha() As Date
        Get
            Return Me._Fecha
        End Get
        Set(ByVal value As Date)
            If ((Me._Fecha = value) _
               = False) Then
                Me._Fecha = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDDeposito", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDDeposito() As String
        Get
            Return Me._IDDeposito
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDDeposito, value) = False) Then
                Me._IDDeposito = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDEstadoVino", DbType:="NVarChar(10)")> _
    Public Property IDEstadoVino() As String
        Get
            Return Me._IDEstadoVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDEstadoVino, value) = False) Then
                Me._IDEstadoVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_Origen", DbType:="Int NOT NULL")> _
    Public Property Origen() As Integer
        Get
            Return Me._Origen
        End Get
        Set(ByVal value As Integer)
            If ((Me._Origen = value) _
               = False) Then
                Me._Origen = value
            End If
        End Set
    End Property

    <Column(Storage:="_TipoDeposito", DbType:="Int NOT NULL")> _
    Public Property TipoDeposito() As Integer
        Get
            Return Me._TipoDeposito
        End Get
        Set(ByVal value As Integer)
            If ((Me._TipoDeposito = value) _
               = False) Then
                Me._TipoDeposito = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticulo", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticulo() As String
        Get
            Return Me._IDArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticulo, value) = False) Then
                Me._IDArticulo = value
            End If
        End Set
    End Property

    <Column(Storage:="_NOperacion", DbType:="NVarChar(10)")> _
    Public Property NOperacion() As String
        Get
            Return Me._NOperacion
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NOperacion, value) = False) Then
                Me._NOperacion = value
            End If
        End Set
    End Property

    <Column(Storage:="_Ficticio", DbType:="Bit NOT NULL")> _
    Public Property Ficticio() As Boolean
        Get
            Return Me._Ficticio
        End Get
        Set(ByVal value As Boolean)
            If ((Me._Ficticio = value) _
               = False) Then
                Me._Ficticio = value
            End If
        End Set
    End Property

    <Column(Storage:="_Grupo", DbType:="Bit NOT NULL")> _
    Public Property Grupo() As Boolean
        Get
            Return Me._Grupo
        End Get
        Set(ByVal value As Boolean)
            If ((Me._Grupo = value) _
               = False) Then
                Me._Grupo = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDUdMedida", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDUdMedida() As String
        Get
            Return Me._IDUdMedida
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDUdMedida, value) = False) Then
                Me._IDUdMedida = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteVariable", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteVariable() As Decimal
        Get
            Return Me._CosteVariable
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteVariable = value) _
               = False) Then
                Me._CosteVariable = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteFiscal", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteFiscal() As Decimal
        Get
            Return Me._CosteFiscal
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteFiscal = value) _
               = False) Then
                Me._CosteFiscal = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteTotal", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteTotal() As Decimal
        Get
            Return Me._CosteTotal
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteTotal = value) _
               = False) Then
                Me._CosteTotal = value
            End If
        End Set
    End Property

    <Column(Storage:="_Lote", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property Lote() As String
        Get
            Return Me._Lote
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Lote, value) = False) Then
                Me._Lote = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlmacen", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAlmacen() As String
        Get
            Return Me._IDAlmacen
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacen, value) = False) Then
                Me._IDAlmacen = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDBarrica", DbType:="NVarChar(10)")> _
    Public Property IDBarrica() As String
        Get
            Return Me._IDBarrica
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDBarrica, value) = False) Then
                Me._IDBarrica = value
            End If
        End Set
    End Property

    <Column(Storage:="_DiasDeposito", DbType:="Int NOT NULL")> _
    Public Property DiasDeposito() As Integer
        Get
            Return Me._DiasDeposito
        End Get
        Set(ByVal value As Integer)
            If ((Me._DiasDeposito = value) _
               = False) Then
                Me._DiasDeposito = value
            End If
        End Set
    End Property

    <Column(Storage:="_DiasBarrica", DbType:="Int NOT NULL")> _
    Public Property DiasBarrica() As Integer
        Get
            Return Me._DiasBarrica
        End Get
        Set(ByVal value As Integer)
            If ((Me._DiasBarrica = value) _
               = False) Then
                Me._DiasBarrica = value
            End If
        End Set
    End Property

    <Column(Storage:="_DiasBotellero", DbType:="Int NOT NULL")> _
    Public Property DiasBotellero() As Integer
        Get
            Return Me._DiasBotellero
        End Get
        Set(ByVal value As Integer)
            If ((Me._DiasBotellero = value) _
               = False) Then
                Me._DiasBotellero = value
            End If
        End Set
    End Property

    <Column(Storage:="_QTotal", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property QTotal() As Double
        Get
            Return Me._QTotal
        End Get
        Set(ByVal value As Double)
            If ((Me._QTotal = value) _
               = False) Then
                Me._QTotal = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUnitarioInicialA", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteUnitarioInicialA() As Decimal
        Get
            Return Me._CosteUnitarioInicialA
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteUnitarioInicialA = value) _
               = False) Then
                Me._CosteUnitarioInicialA = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaCosteUnitarioInicial", DbType:="DateTime")> _
    Public Property FechaCosteUnitarioInicial() As System.Nullable(Of Date)
        Get
            Return Me._FechaCosteUnitarioInicial
        End Get
        Set(ByVal value As System.Nullable(Of Date))
            If (Me._FechaCosteUnitarioInicial.Equals(value) = False) Then
                Me._FechaCosteUnitarioInicial = value
            End If
        End Set
    End Property

    <Column(Storage:="_QTotMermaCoste", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property QTotMermaCoste() As Double
        Get
            Return Me._QTotMermaCoste
        End Get
        Set(ByVal value As Double)
            If ((Me._QTotMermaCoste = value) _
               = False) Then
                Me._QTotMermaCoste = value
            End If
        End Set
    End Property

    <Column(Storage:="_QTotMermaSinCoste", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property QTotMermaSinCoste() As Double
        Get
            Return Me._QTotMermaSinCoste
        End Get
        Set(ByVal value As Double)
            If ((Me._QTotMermaSinCoste = value) _
               = False) Then
                Me._QTotMermaSinCoste = value
            End If
        End Set
    End Property

    Private _links As New List(Of VinoEstructura)()
    Public Property Links() As IList(Of VinoEstructura)
        Get
            Return _links
        End Get
        Set(ByVal value As IList(Of VinoEstructura))
            If Not value Is Nothing Then
                _links = value
            End If
        End Set
    End Property

    Private m_PorcentajeOrigen As Double
    Public Property PorcentajeOrigen() As Double
        Get
            Return m_PorcentajeOrigen
        End Get
        Set(ByVal value As Double)
            m_PorcentajeOrigen = value
        End Set
    End Property

    Private m_PorcentajeOrigenEstructura As Double
    Public Property PorcentajeOrigenEstructura() As Double
        Get
            Return m_PorcentajeOrigenEstructura
        End Get
        Set(ByVal value As Double)
            m_PorcentajeOrigenEstructura = value
        End Set
    End Property


    Private m_PorcentajeCoste As Double
    Public Property PorcentajeCoste() As Double
        Get
            Return m_PorcentajeCoste
        End Get
        Set(ByVal value As Double)
            m_PorcentajeCoste = value
        End Set
    End Property

    <Column(Name:="QDepositoVino", DbType:="Numeric(23,8)", CanBeNull:=True)> _
    Private m_QDepositoVino As Double?
    Public Property QDepositoVino() As Double?
        Get
            Return m_QDepositoVino
        End Get
        Set(ByVal value As Double?)
            m_QDepositoVino = value
        End Set
    End Property

    Private m_IsLeaf As Boolean
    Public Property IsLeaf() As Boolean
        Get
            Return m_IsLeaf
        End Get
        Set(ByVal value As Boolean)
            m_IsLeaf = value
        End Set
    End Property

    Private _level As Integer = Integer.MinValue
    Public Property Level() As Integer
        Get
            Return _level
        End Get
        Set(ByVal value As Integer)
            If _level <= value Then
                If BdgExplosionVino.ActivarTest Then
                    If BdgExplosionVino.TestStack.Contains(Me) Then
                        Throw New Exception("El vino " + Quoted(Me.IDVino.ToString) + " está buclado. Artículo:" + Me.IDArticulo + " IDDeposito:" + Me.IDDeposito + " Lote:" + Me.Lote + " Nivel1:" + CStr(Me.Level) + " Nivel2:" + CStr(value))
                    End If
                End If
                _level = value

                If BdgExplosionVino.ActivarTest Then BdgExplosionVino.TestStack.Push(Me)

                'Se extiende el nivel por el arbol de descendientes
                If Not Me.Links Is Nothing Then
                    For i As Integer = 0 To Links.Count - 1
                        Links(i).Child.Level = _level + 1
                    Next
                    If BdgExplosionVino.ActivarTest Then BdgExplosionVino.TestStack.Pop()
                End If
            End If
        End Set
    End Property


    Private m_Orden As Integer = Integer.MinValue
    Public Property Orden() As Integer
        Get
            Return m_Orden
        End Get
        Set(ByVal value As Integer)
            m_Orden = value
        End Set
    End Property


    <Column(Name:="TieneCosteElaboracion", DbType:="bit", CanBeNull:=False)> _
    Private m_TieneCosteElaboracion As Boolean
    Public Property TieneCosteElaboracion() As Boolean
        Get
            Return m_TieneCosteElaboracion
        End Get
        Set(ByVal value As Boolean)
            m_TieneCosteElaboracion = value
        End Set
    End Property

    <Column(Name:="TieneCosteEstanciaNave", DbType:="bit", CanBeNull:=False)> _
    Private m_TieneCosteEstanciaNave As Boolean
    Public Property TieneCosteEstanciaNave() As Boolean
        Get
            Return m_TieneCosteEstanciaNave
        End Get
        Set(ByVal value As Boolean)
            m_TieneCosteEstanciaNave = value
        End Set
    End Property

    <Column(Name:="TieneCosteInicial", DbType:="bit", CanBeNull:=False)> _
    Private m_TieneCosteInicial As Boolean
    Public Property TieneCosteInicial() As Boolean
        Get
            Return m_TieneCosteInicial
        End Get
        Set(ByVal value As Boolean)
            m_TieneCosteInicial = value
        End Set
    End Property

    <Column(Name:="IDUDLitros", DbType:="NVarchar(10)", CanBeNull:=True)> _
    Private m_IDUDLitros As String
    Public Property IDUDLitros() As String
        Get
            Return m_IDUDLitros
        End Get
        Set(ByVal value As String)
            m_IDUDLitros = value
        End Set
    End Property

    <Column(Name:="QUdLitros", DbType:="Numeric(23,8)", CanBeNull:=True)> _
  Private m_QUdLitros As Double
    Public Property QUdLitros() As Double
        Get
            Return m_QUdLitros
        End Get
        Set(ByVal value As Double)
            m_QUdLitros = value
        End Set
    End Property

    <Column(Name:="IDUdDeposito", DbType:="NVarchar(10)", CanBeNull:=True)> _
    Private m_IDUdDeposito As String
    Public Property IDUdDeposito() As String
        Get
            Return m_IDUdDeposito
        End Get
        Set(ByVal value As String)
            m_IDUdDeposito = value
        End Set
    End Property

    <Column(Name:="QDepositoVino", DbType:="Numeric(23,8)", CanBeNull:=False)> _
    Private m_QUdDeposito As Double
    Public Property QUdDeposito() As Double
        Get
            Return m_QUdDeposito
        End Get
        Set(ByVal value As Double)
            m_QUdDeposito = value
        End Set
    End Property


    <Column(Name:="_EnDeposito", DbType:="bit", CanBeNull:=False)> _
    Private _EnDeposito As Boolean
    Public Property EnDeposito() As Boolean
        Get
            Return _EnDeposito
        End Get
        Set(ByVal value As Boolean)
            _EnDeposito = value
        End Set
    End Property


    <Column(Name:="DescArticulo", DbType:="NVarchar(300)", CanBeNull:=False)> _
    Private m_DescArticulo As String
    Public Property DescArticulo() As String
        Get
            Return m_DescArticulo
        End Get
        Set(ByVal value As String)
            m_DescArticulo = value
        End Set
    End Property

    Private mlstCosteVino As New Dictionary(Of Date, VinoCoste)
    Public Property CosteVinoFecha() As Dictionary(Of Date, VinoCoste)
        Get
            Return mlstCosteVino
        End Get
        Set(ByVal value As Dictionary(Of Date, VinoCoste))
            mlstCosteVino = value
        End Set
    End Property

End Class

<Serializable()> _
Public Class VinoCoste
   

    Private m_CantidadCoste As Double
    Public Property CantidadCoste() As Double
        Get
            Return m_CantidadCoste
        End Get
        Set(ByVal value As Double)
            m_CantidadCoste = value
        End Set
    End Property

    Private m_FechaFinCoste As Date?
    Public Property FechaFinCoste() As Date?
        Get
            Return m_FechaFinCoste
        End Get
        Set(ByVal value As Date?)
            m_FechaFinCoste = value
        End Set
    End Property
End Class

<Serializable()> _
Public Class VinoEstructura

    Private _IDVino As System.Guid

    Private _IDVinoComponente As System.Guid

    Private _Cantidad As Double

    Private _Merma As Double

    Private _Factor As Double

    Private _Operacion As String

    Private _NOperacion As String

    Private _blCrecesAcumulanCoste As Boolean

    Private _blUsarRepartoTipoOperacion As Boolean

    Private _PorcentajeRepartoTipoOperacion As Double

    Private _CantidadRepartoTipoOperacion As Double

    <Column(Storage:="_IDVino", DbType:="UniqueIdentifier NOT NULL", IsPrimaryKey:=True)> _
    Public Property IDVino() As System.Guid
        Get
            Return Me._IDVino
        End Get
        Set(ByVal value As System.Guid)
            If ((Me._IDVino = value) _
               = False) Then
                Me._IDVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDVinoComponente", DbType:="UniqueIdentifier NOT NULL", IsPrimaryKey:=True)> _
    Public Property IDVinoComponente() As System.Guid
        Get
            Return Me._IDVinoComponente
        End Get
        Set(ByVal value As System.Guid)
            If ((Me._IDVinoComponente = value) _
               = False) Then
                Me._IDVinoComponente = value
            End If
        End Set
    End Property

    <Column(Storage:="_Cantidad", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Cantidad() As Double
        Get
            Return Me._Cantidad
        End Get
        Set(ByVal value As Double)
            If ((Me._Cantidad = value) _
               = False) Then
                Me._Cantidad = value
            End If
        End Set
    End Property

    <Column(Storage:="_Merma", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Merma() As Double
        Get
            Return Me._Merma
        End Get
        Set(ByVal value As Double)
            If ((Me._Merma = value) _
               = False) Then
                Me._Merma = value
            End If
        End Set
    End Property

    <Column(Storage:="_Factor", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Factor() As Double
        Get
            Return Me._Factor
        End Get
        Set(ByVal value As Double)
            If ((Me._Factor = value) _
               = False) Then
                Me._Factor = value
            End If
        End Set
    End Property

    <Column(Storage:="_Operacion", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False, IsPrimaryKey:=True)> _
    Public Property Operacion() As String
        Get
            Return Me._Operacion
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Operacion, value) = False) Then
                Me._Operacion = value
            End If
        End Set
    End Property

    <Column(Storage:="_NOperacion", DbType:="NVarChar(10)")> _
    Public Property NOperacion() As String
        Get
            Return Me._NOperacion
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NOperacion, value) = False) Then
                Me._NOperacion = value
            End If
        End Set
    End Property

    <Column(Storage:="_blCrecesAcumulanCoste", DbType:="Bit NOT NULL")> _
   Public Property CrecesAcumulanCoste() As Boolean
        Get
            Return Me._blCrecesAcumulanCoste
        End Get
        Set(ByVal value As Boolean)
            If ((Me._blCrecesAcumulanCoste = value) _
               = False) Then
                Me._blCrecesAcumulanCoste = value
            End If
        End Set
    End Property

    <Column(Storage:="_blUsarRepartoTipoOperacion", DbType:="Bit NOT NULL")> _
   Public Property UsarRepartoTipoOperacion() As Boolean
        Get
            Return Me._blUsarRepartoTipoOperacion
        End Get
        Set(ByVal value As Boolean)
            If ((Me._blUsarRepartoTipoOperacion = value) _
               = False) Then
                Me._blUsarRepartoTipoOperacion = value
            End If
        End Set
    End Property

    <Column(Storage:="_PorcentajeRepartoTipoOperacion", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property PorcentajeRepartoTipoOperacion() As Double
        Get
            Return Me._PorcentajeRepartoTipoOperacion
        End Get
        Set(ByVal value As Double)
            If ((Me._PorcentajeRepartoTipoOperacion = value) _
               = False) Then
                Me._PorcentajeRepartoTipoOperacion = value
            End If
        End Set
    End Property

    <Column(Storage:="_CantidadRepartoTipoOperacion", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CantidadRepartoTipoOperacion() As Double
        Get
            Return Me._CantidadRepartoTipoOperacion
        End Get
        Set(ByVal value As Double)
            If ((Me._CantidadRepartoTipoOperacion = value) _
               = False) Then
                Me._CantidadRepartoTipoOperacion = value
            End If
        End Set
    End Property

    Public Property Child() As Vino
        Get
            Return m_Child
        End Get
        Set(ByVal value As Vino)
            m_Child = value
        End Set
    End Property
    Private m_Child As Vino
    Public Property Parent() As Vino
        Get
            Return m_Parent
        End Get
        Set(ByVal value As Vino)
            m_Parent = value
        End Set
    End Property
    Private m_Parent As Vino
End Class

<Serializable()> _
Public Class OrigenVariedades

    Private _IDVariedad As String

    Private _DescVariedad As String

    Private _Prctj As Nullable(Of Decimal)

    Private _Cantidad As Nullable(Of Decimal)

    Public Sub New()
    End Sub

    <Column(Storage:="_IDVariedad", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDVariedad() As String
        Get
            Return Me._IDVariedad
        End Get
        Set(ByVal value As String)
            If (Me._IDVariedad <> value) Then
                Me._IDVariedad = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescVariedad", DbType:="NVarChar(100) NOT NULL", CanBeNull:=False)> _
    Public Property DescVariedad() As String
        Get
            Return Me._DescVariedad
        End Get
        Set(ByVal value As String)
            If (Me._DescVariedad <> value) Then
                Me._DescVariedad = value
            End If
        End Set
    End Property

    <Column(Storage:="_Prctj", DbType:="Decimal(38,8)")> _
    Public Property Prctj() As System.Nullable(Of Decimal)
        Get
            Return Me._Prctj
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Prctj <> value) Then
                Me._Prctj = value
            End If
        End Set
    End Property

    <Column(Storage:="_Cantidad", DbType:="Decimal(38,8)")> _
   Public Property Cantidad() As System.Nullable(Of Decimal)
        Get
            Return Me._Cantidad
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Cantidad <> value) Then
                Me._Cantidad = value
            End If
        End Set
    End Property
End Class

<Serializable()> _
Public Class OrigenFincas

    Private _IDFinca As String

    Private _CFinca As String

    Private _DescFinca As String

    Private _Prctj As Nullable(Of Decimal)

    Private _Cantidad As Nullable(Of Decimal)

    Public Sub New()
    End Sub

    <Column(Storage:="_IDFinca", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDFinca() As String
        Get
            Return Me._IDFinca
        End Get
        Set(ByVal value As String)
            If (Me._IDFinca <> value) Then
                Me._IDFinca = value
            End If
        End Set
    End Property

    <Column(Storage:="_CFinca", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
   Public Property CFinca() As String
        Get
            Return Me._CFinca
        End Get
        Set(ByVal value As String)
            If (Me._CFinca <> value) Then
                Me._CFinca = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescFinca", DbType:="NVarChar(100) NOT NULL", CanBeNull:=False)> _
    Public Property DescFinca() As String
        Get
            Return Me._DescFinca
        End Get
        Set(ByVal value As String)
            If (Me._DescFinca <> value) Then
                Me._DescFinca = value
            End If
        End Set
    End Property

    <Column(Storage:="_Prctj", DbType:="Decimal(38,8)")> _
    Public Property Prctj() As System.Nullable(Of Decimal)
        Get
            Return Me._Prctj
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Prctj <> value) Then
                Me._Prctj = value
            End If
        End Set
    End Property


    <Column(Storage:="_Cantidad", DbType:="Decimal(38,8)")> _
    Public Property Cantidad() As System.Nullable(Of Decimal)
        Get
            Return Me._Cantidad
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Cantidad <> value) Then
                Me._Cantidad = value
            End If
        End Set
    End Property

End Class

<Serializable()> _
Public Class OrigenCompras
    Private _IDArticulo As String
    Private _DescArticulo As String
    Private _Prctj As Nullable(Of Decimal)
    Private _Cantidad As Nullable(Of Decimal)
    Private _NEntrada As Integer

    <Column(Storage:="_IDArticulo")> _
    Public Property IDArticulo() As String
        Get
            Return _IDArticulo
        End Get
        Set(ByVal value As String)
            _IDArticulo = value
        End Set
    End Property

    <Column(Storage:="_DescArticulo")> _
    Public Property DescArticulo() As String
        Get
            Return _DescArticulo
        End Get
        Set(ByVal value As String)
            _DescArticulo = value
        End Set
    End Property

    <Column(Storage:="_Prctj")> _
    Public Property Prctj() As Nullable(Of Decimal)
        Get
            Return _Prctj
        End Get
        Set(ByVal value As Nullable(Of Decimal))
            _Prctj = value
        End Set
    End Property

    <Column(Storage:="_Cantidad", DbType:="Decimal(38,8)")> _
  Public Property Cantidad() As System.Nullable(Of Decimal)
        Get
            Return Me._Cantidad
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Cantidad <> value) Then
                Me._Cantidad = value
            End If
        End Set
    End Property

    <Column(Storage:="_NEntrada")> _
    Public Property NEntrada() As Integer
        Get
            Return _NEntrada
        End Get
        Set(ByVal value As Integer)
            _NEntrada = value
        End Set
    End Property
End Class

<Serializable()> _
Public Class OrigenAniadas

    Private _IDAnada As String

    Private _DescAnada As String

    Private _Prctj As Nullable(Of Decimal)

    Private _Cantidad As Nullable(Of Decimal)

    Public Sub New()
    End Sub

    <Column(Storage:="_IDAnada", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAnada() As String
        Get
            Return Me._IDAnada
        End Get
        Set(ByVal value As String)
            If (Me._IDAnada <> value) Then
                Me._IDAnada = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescAnada", DbType:="NVarChar(100) NOT NULL", CanBeNull:=False)> _
    Public Property DescAnada() As String
        Get
            Return Me._DescAnada
        End Get
        Set(ByVal value As String)
            If (Me._DescAnada <> value) Then
                Me._DescAnada = value
            End If
        End Set
    End Property

    <Column(Storage:="_Prctj", DbType:="Decimal(38,8)")> _
    Public Property Prctj() As System.Nullable(Of Decimal)
        Get
            Return Me._Prctj
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Prctj <> value) Then
                Me._Prctj = value
            End If
        End Set
    End Property

    <Column(Storage:="_Cantidad", DbType:="Decimal(38,8)")> _
   Public Property Cantidad() As System.Nullable(Of Decimal)
        Get
            Return Me._Cantidad
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Cantidad <> value) Then
                Me._Cantidad = value
            End If
        End Set
    End Property
End Class

<Serializable()> _
Public Class OrigenEntradasUVA

    Private _Fecha As Date
    <Column(Storage:="_Fecha", DbType:="DateTime NOT NULL")> _
   Public Property Fecha() As Date
        Get
            Return Me._Fecha
        End Get
        Set(ByVal value As Date)
            If ((Me._Fecha = value) _
               = False) Then
                Me._Fecha = value
            End If
        End Set
    End Property

    Private _IDEntrada As Nullable(Of Integer)
    <Column(Storage:="_IDEntrada", DbType:="Int")> _
    Public Property IDEntrada() As System.Nullable(Of Integer)
        Get
            Return Me._IDEntrada
        End Get
        Set(ByVal value As System.Nullable(Of Integer))
            If (Me._IDEntrada <> value) Then
                Me._IDEntrada = value
            End If
        End Set
    End Property

    Private _NEntrada As String
    <Column(Name:="NEntrada", DbType:="NVarchar(25)", CanBeNull:=True)> _
    Public Property NEntrada() As String
        Get
            Return _NEntrada
        End Get
        Set(ByVal value As String)
            If (Me._NEntrada <> value) Then
                Me._NEntrada = value
            End If
        End Set
    End Property

    Private _Cantidad As Nullable(Of Decimal)
    <Column(Storage:="_Cantidad", DbType:="Decimal(38,8)")> _
    Public Property Cantidad() As System.Nullable(Of Decimal)
        Get
            Return Me._Cantidad
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Cantidad <> value) Then
                Me._Cantidad = value
            End If
        End Set
    End Property

    Private _Prctj As Nullable(Of Decimal)
    <Column(Storage:="_Prctj", DbType:="Decimal(38,8)")> _
   Public Property Prctj() As System.Nullable(Of Decimal)
        Get
            Return Me._Prctj
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Prctj <> value) Then
                Me._Prctj = value
            End If
        End Set
    End Property

    Private _DescVariedad As String
    <Column(Name:="DescVariedad", DbType:="NVarchar(100)", CanBeNull:=True)> _
    Public Property DescVariedad() As String
        Get
            Return _DescVariedad
        End Get
        Set(ByVal value As String)
            If (Me._DescVariedad <> value) Then
                Me._DescVariedad = value
            End If
        End Set
    End Property

    Private _DescFinca As String
    <Column(Name:="DescFinca", DbType:="NVarchar(100)", CanBeNull:=True)> _
    Public Property DescFinca() As String
        Get
            Return _DescFinca
        End Get
        Set(ByVal value As String)
            If (Me._DescFinca <> value) Then
                Me._DescFinca = value
            End If
        End Set
    End Property

    Private _IDCartillista As String
    <Column(Name:="IDCartillista", DbType:="NVarchar(25)", CanBeNull:=True)> _
    Public Property IDCartillista() As String
        Get
            Return _IDCartillista
        End Get
        Set(ByVal value As String)
            If (Me._IDCartillista <> value) Then
                Me._IDCartillista = value
            End If
        End Set
    End Property

    Private _DescCartillista As String
    <Column(Name:="DescCartillista", DbType:="NVarchar(100)", CanBeNull:=True)> _
    Public Property DescCartillista() As String
        Get
            Return _DescCartillista
        End Get
        Set(ByVal value As String)
            If (Me._DescCartillista <> value) Then
                Me._DescCartillista = value
            End If
        End Set
    End Property

End Class

<Serializable()> _
Public Class DepositoVino

    Private _IDVino As System.Guid

    Private _IDDeposito As String

    Private _Cantidad As Decimal

    <Column(Storage:="_IDVino", DbType:="UniqueIdentifier NOT NULL", IsPrimaryKey:=True)> _
    Public Property IDVino() As System.Guid
        Get
            Return Me._IDVino
        End Get
        Set(ByVal value As System.Guid)
            If ((Me._IDVino = value) _
               = False) Then
                Me._IDVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDDeposito", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDDeposito() As String
        Get
            Return Me._IDDeposito
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDDeposito, value) = False) Then
                Me._IDDeposito = value
            End If
        End Set
    End Property

    <Column(Storage:="_Cantidad", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Cantidad() As Decimal
        Get
            Return Me._Cantidad
        End Get
        Set(ByVal value As Decimal)
            If ((Me._Cantidad = value) _
               = False) Then
                Me._Cantidad = value
            End If
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return String.Format("{0}: {1:F2}", IDDeposito, Cantidad)
    End Function
End Class

<Serializable()> _
Public Class VinoTratamientos

    Private _IDVino As System.Guid

    Private _IDArticulo As String

    Private _DescArticulo As String

    Private _Fecha As Date

    Private _NOperacion As String

    Private _Cantidad As System.Nullable(Of Decimal)

    Private _Lote As String

    Private _IDUdInterna As String

    Public Sub New()
        MyBase.New()
    End Sub

    <Column(Storage:="_IDVino", DbType:="UniqueIdentifier NOT NULL")> _
    Public Property IDVino() As System.Guid
        Get
            Return Me._IDVino
        End Get
        Set(ByVal value As System.Guid)
            If ((Me._IDVino = value) _
               = False) Then
                Me._IDVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticulo", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticulo() As String
        Get
            Return Me._IDArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticulo, value) = False) Then
                Me._IDArticulo = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescArticulo", DbType:="NVarChar(300) NOT NULL", CanBeNull:=False)> _
    Public Property DescArticulo() As String
        Get
            Return Me._DescArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescArticulo, value) = False) Then
                Me._DescArticulo = value
            End If
        End Set
    End Property

    <Column(Storage:="_Fecha", DbType:="DateTime NOT NULL")> _
    Public Property Fecha() As Date
        Get
            Return Me._Fecha
        End Get
        Set(ByVal value As Date)
            If ((Me._Fecha = value) _
               = False) Then
                Me._Fecha = value
            End If
        End Set
    End Property

    <Column(Storage:="_NOperacion", DbType:="NVarChar(10)")> _
    Public Property NOperacion() As String
        Get
            Return Me._NOperacion
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NOperacion, value) = False) Then
                Me._NOperacion = value
            End If
        End Set
    End Property

    <Column(Storage:="_Cantidad", DbType:="Decimal(23,8)")> _
    Public Property Cantidad() As System.Nullable(Of Decimal)
        Get
            Return Me._Cantidad
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Cantidad.Equals(value) = False) Then
                Me._Cantidad = value
            End If
        End Set
    End Property

    <Column(Storage:="_Lote", DbType:="NVarChar(25)")> _
    Public Property Lote() As String
        Get
            Return Me._Lote
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Lote, value) = False) Then
                Me._Lote = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDUdInterna", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDUdInterna() As String
        Get
            Return Me._IDUdInterna
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDUdInterna, value) = False) Then
                Me._IDUdInterna = value
            End If
        End Set
    End Property
End Class

<Serializable()> _
Public Class VinoDetalle


    <Column(Storage:="_IDVino", DbType:="uniqueidentifier NOT NULL", CanBeNull:=False)> _
    Private _IDVino As Guid
    Public Property IDVino() As Guid
        Get
            Return Me._IDVino
        End Get
        Set(ByVal value As Guid)
            If Not value.Equals(Guid.Empty) Then
                Me._IDVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_TipoNodo", DbType:="int")> _
    Private _TipoNodo As Integer
    Public Property TipoNodo() As Integer
        Get
            Return Me._TipoNodo
        End Get
        Set(ByVal value As Integer)
            Me._TipoNodo = value
        End Set
    End Property

    <Column(Storage:="_Fecha", DbType:="datetime")> _
    Private _Fecha As Date
    Public Property Fecha() As Date
        Get
            Return Me._Fecha
        End Get
        Set(ByVal value As Date)
            Me._Fecha = value
        End Set
    End Property


    <Column(Storage:="_NOperacion", DbType:="nvarchar(10)")> _
    Private _NOperacion As String
    Public Property NOperacion() As String
        Get
            Return Me._NOperacion
        End Get
        Set(ByVal value As String)
            Me._NOperacion = value
        End Set
    End Property

    <Column(Storage:="_IDNodo", DbType:="nvarchar(50)")> _
    Private _IDNodo As String
    Public Property IDNodo() As String
        Get
            Return Me._IDNodo
        End Get
        Set(ByVal value As String)
            Me._IDNodo = value
        End Set
    End Property

    <Column(Storage:="_DescNodo", DbType:="nvarchar(500)")> _
    Private _DescNodo As String
    Public Property DescNodo() As String
        Get
            Return Me._DescNodo
        End Get
        Set(ByVal value As String)
            Me._DescNodo = value
        End Set
    End Property


    <Column(Storage:="_IDDocumentoInt", DbType:="int")> _
   Private _IDDocumentoInt As Integer
    Public Property IDDocumentoInt() As Integer
        Get
            Return Me._IDDocumentoInt
        End Get
        Set(ByVal value As Integer)
            Me._IDDocumentoInt = value
        End Set
    End Property

    <Column(Storage:="_IDDocumentoGuid", DbType:="uniqueidentifier")> _
    Private _IDDocumentoGuid As Guid
    Public Property IDDocumentoGuid() As Guid
        Get
            Return Me._IDDocumentoGuid
        End Get
        Set(ByVal value As Guid)
            If Not value.Equals(Guid.Empty) Then
                Me._IDDocumentoGuid = value
            End If
        End Set
    End Property


    <Column(Storage:="_NDocumento", DbType:="nvarchar(50)")> _
    Private _NDocumento As String
    Public Property NDocumento() As String
        Get
            Return Me._NDocumento
        End Get
        Set(ByVal value As String)
            Me._NDocumento = value
        End Set
    End Property

End Class


<Serializable()> _
Public Class VinoTrazabilidadSalidas
    Implements IEquatable(Of VinoTrazabilidadSalidas)

    Private _Traza As System.Nullable(Of System.Guid)

    Private _IDMovimiento As Integer

    Private _IDArticulo As String

    Private _IDAlmacen As String

    Private _Cantidad As Decimal

    Private _Acumulado As Decimal

    Private _FechaDocumento As Date

    Private _Documento As String

    Private _DescArticulo As String

    Private _IDTipoMovimiento As Integer

    Private _IDCliente As String

    Private _DescCliente As String

    Private _NAlbaran As String

    Private _IDDocumento As System.Nullable(Of Integer)

    Private _Lote As String

    Private _Ubicacion As String

    Private _TelefonoCliente As String

    Private _FaxCliente As String

    Private _EmailCliente As String

    Private _IDAlbaran As Integer

    Private _IDLineaMovimiento As Integer

    Private _IDDireccion As Integer

    Private _RazonSocialEnvio As String

    Private _DireccionEnvio As String

    Private _CodPostalEnvio As String

    Private _PoblacionEnvio As String

    Private _ProvinciaEnvio As String

    Private _DescPaisEnvio As String

    Private _TelefonoEnvio As String

    Private _FaxEnvio As String

    Private _EmailEnvio As String

    Private _NombreContacto As String

    Private _TelefonoContacto As String

    Private _FaxContacto As String

    Private _EmailContacto As String

    Private _EtiquetaContenida As Boolean


    Public Sub New()
        MyBase.New()
    End Sub

    <Column(Storage:="_Traza", DbType:="UniqueIdentifier")> _
    Public Property Traza() As System.Nullable(Of System.Guid)
        Get
            Return Me._Traza
        End Get
        Set(ByVal value As System.Nullable(Of System.Guid))
            If (Me._Traza.Equals(value) = False) Then
                Me._Traza = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDMovimiento", DbType:="Int NOT NULL")> _
    Public Property IDMovimiento() As Integer
        Get
            Return Me._IDMovimiento
        End Get
        Set(ByVal value As Integer)
            If ((Me._IDMovimiento = value) _
               = False) Then
                Me._IDMovimiento = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticulo", DbType:="NVarChar(25)")> _
    Public Property IDArticulo() As String
        Get
            Return Me._IDArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticulo, value) = False) Then
                Me._IDArticulo = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlmacen", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAlmacen() As String
        Get
            Return Me._IDAlmacen
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacen, value) = False) Then
                Me._IDAlmacen = value
            End If
        End Set
    End Property

    <Column(Storage:="_Cantidad", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Cantidad() As Decimal
        Get
            Return Me._Cantidad
        End Get
        Set(ByVal value As Decimal)
            If ((Me._Cantidad = value) _
               = False) Then
                Me._Cantidad = value
            End If
        End Set
    End Property

    <Column(Storage:="_Acumulado", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Acumulado() As Decimal
        Get
            Return Me._Acumulado
        End Get
        Set(ByVal value As Decimal)
            If ((Me._Acumulado = value) _
               = False) Then
                Me._Acumulado = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaDocumento", DbType:="DateTime NOT NULL")> _
    Public Property FechaDocumento() As Date
        Get
            Return Me._FechaDocumento
        End Get
        Set(ByVal value As Date)
            If ((Me._FechaDocumento = value) _
               = False) Then
                Me._FechaDocumento = value
            End If
        End Set
    End Property

    <Column(Storage:="_Documento", DbType:="NVarChar(50)")> _
    Public Property Documento() As String
        Get
            Return Me._Documento
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Documento, value) = False) Then
                Me._Documento = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescArticulo", DbType:="NVarChar(300) NOT NULL", CanBeNull:=False)> _
    Public Property DescArticulo() As String
        Get
            Return Me._DescArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescArticulo, value) = False) Then
                Me._DescArticulo = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDTipoMovimiento", DbType:="Int NOT NULL")> _
    Public Property IDTipoMovimiento() As Integer
        Get
            Return Me._IDTipoMovimiento
        End Get
        Set(ByVal value As Integer)
            If ((Me._IDTipoMovimiento = value) _
               = False) Then
                Me._IDTipoMovimiento = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDCliente", DbType:="NVarChar(25)")> _
    Public Property IDCliente() As String
        Get
            Return Me._IDCliente
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDCliente, value) = False) Then
                Me._IDCliente = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescCliente", DbType:="NVarChar(300)")> _
    Public Property DescCliente() As String
        Get
            Return Me._DescCliente
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescCliente, value) = False) Then
                Me._DescCliente = value
            End If
        End Set
    End Property

    <Column(Storage:="_NAlbaran", DbType:="NVarChar(25)")> _
    Public Property NAlbaran() As String
        Get
            Return Me._NAlbaran
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NAlbaran, value) = False) Then
                Me._NAlbaran = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDDocumento", DbType:="Int", CanBeNull:=True)> _
    Public Property IDDocumento() As System.Nullable(Of Integer)
        Get
            Return Me._IDDocumento
        End Get
        Set(ByVal value As System.Nullable(Of Integer))
            If (Me._IDDocumento.Equals(value) = False) Then
                Me._IDDocumento = value
            End If
        End Set
    End Property

    <Column(Storage:="_Lote", DbType:="NVarChar(25)")> _
    Public Property Lote() As String
        Get
            Return Me._Lote
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Lote, value) = False) Then
                Me._Lote = value
            End If
        End Set
    End Property

    <Column(Storage:="_Ubicacion", DbType:="NVarChar(25)")> _
    Public Property Ubicacion() As String
        Get
            Return Me._Ubicacion
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Ubicacion, value) = False) Then
                Me._Ubicacion = value
            End If
        End Set
    End Property

    <Column(Storage:="_TelefonoCliente", DbType:="NVarChar(25)")> _
    Public Property TelefonoCliente() As String
        Get
            Return Me._TelefonoCliente
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._TelefonoCliente, value) = False) Then
                Me._TelefonoCliente = value
            End If
        End Set
    End Property

    <Column(Storage:="_FaxCliente", DbType:="NVarChar(25)")> _
    Public Property FaxCliente() As String
        Get
            Return Me._FaxCliente
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._FaxCliente, value) = False) Then
                Me._FaxCliente = value
            End If
        End Set
    End Property

    <Column(Storage:="_EmailCliente", DbType:="NVarChar(100)")> _
    Public Property EmailCliente() As String
        Get
            Return Me._EmailCliente
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._EmailCliente, value) = False) Then
                Me._EmailCliente = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlbaran", DbType:="Int", CanBeNull:=True)> _
    Public Property IDAlbaran() As System.Nullable(Of Integer)
        Get
            Return Me._IDAlbaran
        End Get
        Set(ByVal value As System.Nullable(Of Integer))
            If (Me._IDAlbaran.Equals(value) = False) Then
                Me._IDAlbaran = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDLineaMovimiento", DbType:="Int", CanBeNull:=True)> _
   Public Property IDLineaMovimiento() As System.Nullable(Of Integer)
        Get
            Return Me._IDLineaMovimiento
        End Get
        Set(ByVal value As System.Nullable(Of Integer))
            If (Me._IDLineaMovimiento.Equals(value) = False) Then
                Me._IDLineaMovimiento = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDDireccion", DbType:="Int", CanBeNull:=True)> _
    Public Property IDDireccion() As System.Nullable(Of Integer)
        Get
            Return Me._IDDireccion
        End Get
        Set(ByVal value As System.Nullable(Of Integer))
            If (Me._IDDireccion.Equals(value) = False) Then
                Me._IDDireccion = value
            End If
        End Set
    End Property

    <Column(Storage:="_RazonSocialEnvio", DbType:="NVarChar(300)")> _
    Public Property RazonSocialEnvio() As String
        Get
            Return Me._RazonSocialEnvio
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._RazonSocialEnvio, value) = False) Then
                Me._RazonSocialEnvio = value
            End If
        End Set
    End Property

    <Column(Storage:="_DireccionEnvio", DbType:="NVarChar(100)")> _
    Public Property DireccionEnvio() As String
        Get
            Return Me._DireccionEnvio
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DireccionEnvio, value) = False) Then
                Me._DireccionEnvio = value
            End If
        End Set
    End Property

    <Column(Storage:="_CodPostalEnvio", DbType:="NVarChar(25)")> _
    Public Property CodPostalEnvio() As String
        Get
            Return Me._CodPostalEnvio
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._CodPostalEnvio, value) = False) Then
                Me._CodPostalEnvio = value
            End If
        End Set
    End Property

    <Column(Storage:="_PoblacionEnvio", DbType:="NVarChar(100)")> _
    Public Property PoblacionEnvio() As String
        Get
            Return Me._PoblacionEnvio
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._PoblacionEnvio, value) = False) Then
                Me._PoblacionEnvio = value
            End If
        End Set
    End Property

    <Column(Storage:="_ProvinciaEnvio", DbType:="NVarChar(100)")> _
    Public Property ProvinciaEnvio() As String
        Get
            Return Me._ProvinciaEnvio
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._ProvinciaEnvio, value) = False) Then
                Me._ProvinciaEnvio = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescPaisEnvio", DbType:="NVarChar(100)")> _
    Public Property DescPaisEnvio() As String
        Get
            Return Me._DescPaisEnvio
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescPaisEnvio, value) = False) Then
                Me._DescPaisEnvio = value
            End If
        End Set
    End Property

    <Column(Storage:="_TelefonoEnvio", DbType:="NVarChar(25)")> _
    Public Property TelefonoEnvio() As String
        Get
            Return Me._TelefonoEnvio
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._TelefonoEnvio, value) = False) Then
                Me._TelefonoEnvio = value
            End If
        End Set
    End Property

    <Column(Storage:="_FaxEnvio", DbType:="NVarChar(25)")> _
    Public Property FaxEnvio() As String
        Get
            Return Me._FaxEnvio
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._FaxEnvio, value) = False) Then
                Me._FaxEnvio = value
            End If
        End Set
    End Property

    <Column(Storage:="_EmailEnvio", DbType:="NVarChar(100)")> _
    Public Property EmailEnvio() As String
        Get
            Return Me._EmailEnvio
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._EmailEnvio, value) = False) Then
                Me._EmailEnvio = value
            End If
        End Set
    End Property

    <Column(Storage:="_NombreContacto", DbType:="NVarChar(100)")> _
    Public Property NombreContacto() As String
        Get
            Return Me._NombreContacto
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NombreContacto, value) = False) Then
                Me._NombreContacto = value
            End If
        End Set
    End Property

    <Column(Storage:="_TelefonoContacto", DbType:="NVarChar(25)")> _
    Public Property TelefonoContacto() As String
        Get
            Return Me._TelefonoContacto
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._TelefonoContacto, value) = False) Then
                Me._TelefonoContacto = value
            End If
        End Set
    End Property

    <Column(Storage:="_FaxContacto", DbType:="NVarChar(25)")> _
    Public Property FaxContacto() As String
        Get
            Return Me._FaxContacto
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._FaxContacto, value) = False) Then
                Me._FaxContacto = value
            End If
        End Set
    End Property

    <Column(Storage:="_EmailContacto", DbType:="NVarChar(100)")> _
    Public Property EmailContacto() As String
        Get
            Return Me._EmailContacto
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._EmailContacto, value) = False) Then
                Me._EmailContacto = value
            End If
        End Set
    End Property

    Public Property EtiquetaContenida() As Boolean
        Get
            Return Me._EtiquetaContenida
        End Get
        Set(ByVal value As Boolean)
            If ((Me._EtiquetaContenida = value) = False) Then
                Me._EtiquetaContenida = value
            End If
        End Set
    End Property



    Public Function Equals(ByVal o As VinoTrazabilidadSalidas) As Boolean Implements System.IEquatable(Of Solmicro.Expertis.Business.Bodega.VinoTrazabilidadSalidas).Equals
        If Me.Traza.Equals(o.Traza) AndAlso Me.IDMovimiento = o.IDMovimiento AndAlso Me.NAlbaran = o.NAlbaran AndAlso Me.FechaDocumento = o.FechaDocumento AndAlso Me.IDAlmacen = o.IDAlmacen AndAlso Me.Cantidad = o.Cantidad Then
            Return True
        End If
    End Function

End Class

<Serializable()> _
Public Class VinoTrazabilidadStock
    Implements IEquatable(Of VinoTrazabilidadStock)

    Private _Traza As System.Nullable(Of System.Guid)

    Private _IDArticulo As String

    Private _DescArticulo As String

    Private _IDAlmacen As String

    Private _Lote As String

    Private _StockFisico As Decimal

    Private _Ubicacion As String

    Private _IDUdInterna As String

    Public Sub New()
        MyBase.New()
    End Sub

    <Column(Storage:="_Traza", DbType:="UniqueIdentifier")> _
    Public Property Traza() As System.Nullable(Of System.Guid)
        Get
            Return Me._Traza
        End Get
        Set(ByVal value As System.Nullable(Of System.Guid))
            If (Me._Traza.Equals(value) = False) Then
                Me._Traza = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticulo", DbType:="VarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticulo() As String
        Get
            Return Me._IDArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticulo, value) = False) Then
                Me._IDArticulo = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescArticulo", DbType:="VarChar(300) NOT NULL", CanBeNull:=False)> _
    Public Property DescArticulo() As String
        Get
            Return Me._DescArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescArticulo, value) = False) Then
                Me._DescArticulo = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlmacen", DbType:="VarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAlmacen() As String
        Get
            Return Me._IDAlmacen
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacen, value) = False) Then
                Me._IDAlmacen = value
            End If
        End Set
    End Property

    <Column(Storage:="_Lote", DbType:="VarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property Lote() As String
        Get
            Return Me._Lote
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Lote, value) = False) Then
                Me._Lote = value
            End If
        End Set
    End Property

    <Column(Storage:="_StockFisico", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property StockFisico() As Decimal
        Get
            Return Me._StockFisico
        End Get
        Set(ByVal value As Decimal)
            If ((Me._StockFisico = value) _
               = False) Then
                Me._StockFisico = value
            End If
        End Set
    End Property

    <Column(Storage:="_Ubicacion", DbType:="VarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property Ubicacion() As String
        Get
            Return Me._Ubicacion
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Ubicacion, value) = False) Then
                Me._Ubicacion = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDUdInterna", DbType:="VarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDUdInterna() As String
        Get
            Return Me._IDUdInterna
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDUdInterna, value) = False) Then
                Me._IDUdInterna = value
            End If
        End Set
    End Property


    Public Function Equals(ByVal o As VinoTrazabilidadStock) As Boolean Implements System.IEquatable(Of Solmicro.Expertis.Business.Bodega.VinoTrazabilidadStock).Equals
        If Me.Traza.Equals(o.Traza) AndAlso Me.IDArticulo = o.IDArticulo AndAlso Me.IDAlmacen = o.IDAlmacen AndAlso Me.Lote = o.Lote AndAlso Me.Ubicacion = o.Ubicacion AndAlso Me.IDUdInterna = o.IDUdInterna AndAlso Me.StockFisico = o.StockFisico Then
            Return True
        End If
    End Function

End Class

<Serializable()> _
Public Class VinoTrazabilidadLotes
    Implements IEquatable(Of VinoTrazabilidadLotes)

    Private _IDArticulo As String

    Private _DescArticulo As String

    Private _Lote As String

    Private _Material As String

    Private _DescMaterial As String

    Private _LoteMaterial As String

    Private _Fecha As Date

    Private _NOperacion As String

    Private _IDAlbaran As Integer

    Private _NAlbaran As String

    Private _FechaAlbaran As Date

    Private _IDProveedor As String

    Private _DescProveedor As String

    Private _Telefono As String

    Private _Fax As String

    Private _EMail As String

    Private _Direccion As String

    Private _CodPostal As String

    Private _Poblacion As String

    Private _Provincia As String

    Private _DescPais As String

    Public Sub New()
        MyBase.New()
    End Sub

    <Column(Storage:="_IDArticulo", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticulo() As String
        Get
            Return Me._IDArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticulo, value) = False) Then
                Me._IDArticulo = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescArticulo", DbType:="NVarChar(300) NOT NULL", CanBeNull:=False)> _
    Public Property DescArticulo() As String
        Get
            Return Me._DescArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescArticulo, value) = False) Then
                Me._DescArticulo = value
            End If
        End Set
    End Property

    <Column(Storage:="_Lote", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property Lote() As String
        Get
            Return Me._Lote
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Lote, value) = False) Then
                Me._Lote = value
            End If
        End Set
    End Property

    <Column(Storage:="_Material", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property Material() As String
        Get
            Return Me._Material
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Material, value) = False) Then
                Me._Material = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescMaterial", DbType:="NVarChar(300) NOT NULL", CanBeNull:=False)> _
    Public Property DescMaterial() As String
        Get
            Return Me._DescMaterial
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescMaterial, value) = False) Then
                Me._DescMaterial = value
            End If
        End Set
    End Property

    <Column(Storage:="_LoteMaterial", DbType:="NVarChar(25)")> _
    Public Property LoteMaterial() As String
        Get
            Return Me._LoteMaterial
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._LoteMaterial, value) = False) Then
                Me._LoteMaterial = value
            End If
        End Set
    End Property

    <Column(Storage:="_Fecha", DbType:="DateTime NOT NULL")> _
    Public Property Fecha() As Date
        Get
            Return Me._Fecha
        End Get
        Set(ByVal value As Date)
            If ((Me._Fecha = value) _
               = False) Then
                Me._Fecha = value
            End If
        End Set
    End Property

    <Column(Storage:="_NOperacion", DbType:="NVarChar(10)")> _
    Public Property NOperacion() As String
        Get
            Return Me._NOperacion
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NOperacion, value) = False) Then
                Me._NOperacion = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlbaran", DbType:="Int", CanBeNull:=True)> _
Public Property IDAlbaran() As System.Nullable(Of Integer)
        Get
            Return Me._IDAlbaran
        End Get
        Set(ByVal value As System.Nullable(Of Integer))
            If (Me._IDAlbaran.Equals(value) = False) Then
                Me._IDAlbaran = value
            End If
        End Set
    End Property

    <Column(Storage:="_NAlbaran", DbType:="NVarChar(25)")> _
    Public Property NAlbaran() As String
        Get
            Return Me._NAlbaran
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NAlbaran, value) = False) Then
                Me._NAlbaran = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaAlbaran", DbType:="DateTime NOT NULL")> _
    Public Property FechaAlbaran() As Date
        Get
            Return Me._FechaAlbaran
        End Get
        Set(ByVal value As Date)
            If ((Me._FechaAlbaran = value) = False) Then
                Me._FechaAlbaran = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDProveedor", DbType:="NVarChar(25)")> _
    Public Property IDProveedor() As String
        Get
            Return Me._IDProveedor
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDProveedor, value) = False) Then
                Me._IDProveedor = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescProveedor", DbType:="NVarChar(300)")> _
    Public Property DescProveedor() As String
        Get
            Return Me._DescProveedor
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescProveedor, value) = False) Then
                Me._DescProveedor = value
            End If
        End Set
    End Property

    <Column(Storage:="_Telefono", DbType:="NVarChar(25)")> _
    Public Property Telefono() As String
        Get
            Return Me._Telefono
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Telefono, value) = False) Then
                Me._Telefono = value
            End If
        End Set
    End Property

    <Column(Storage:="_Fax", DbType:="NVarChar(25)")> _
    Public Property Fax() As String
        Get
            Return Me._Fax
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Fax, value) = False) Then
                Me._Fax = value
            End If
        End Set
    End Property

    <Column(Storage:="_EMail", DbType:="NVarChar(100)")> _
    Public Property EMail() As String
        Get
            Return Me._EMail
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._EMail, value) = False) Then
                Me._EMail = value
            End If
        End Set
    End Property

    <Column(Storage:="_Direccion", DbType:="NVarChar(100)")> _
    Public Property Direccion() As String
        Get
            Return Me._Direccion
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Direccion, value) = False) Then
                Me._Direccion = value
            End If
        End Set
    End Property

    <Column(Storage:="_CodPostal", DbType:="NVarChar(25)")> _
        Public Property CodPostal() As String
        Get
            Return Me._CodPostal
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._CodPostal, value) = False) Then
                Me._CodPostal = value
            End If
        End Set
    End Property

    <Column(Storage:="_Poblacion", DbType:="NVarChar(100)")> _
        Public Property Poblacion() As String
        Get
            Return Me._Poblacion
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Poblacion, value) = False) Then
                Me._Poblacion = value
            End If
        End Set
    End Property

    <Column(Storage:="_Provincia", DbType:="NVarChar(100)")> _
    Public Property Provincia() As String
        Get
            Return Me._Provincia
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Provincia, value) = False) Then
                Me._Provincia = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescPais", DbType:="NVarChar(100)")> _
    Public Property DescPais() As String
        Get
            Return Me._DescPais
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescPais, value) = False) Then
                Me._DescPais = value
            End If
        End Set
    End Property

    Public Function Equals(ByVal o As VinoTrazabilidadLotes) As Boolean Implements System.IEquatable(Of Solmicro.Expertis.Business.Bodega.VinoTrazabilidadLotes).Equals
        If Me.IDArticulo = o.IDArticulo AndAlso Me.Lote = o.Lote AndAlso Me.Material = o.Material AndAlso Me.LoteMaterial = o.LoteMaterial AndAlso Me.Fecha = o.Fecha AndAlso Me.NOperacion = o.NOperacion Then
            Return True
        End If
    End Function

End Class






<Serializable()> _
Public Class VinoExplosionArbol
    Implements IEquatable(Of VinoExplosionArbol)


    Private _ID As Integer
    <Column(Storage:="_ID", DbType:="Int NOT NULL")> _
   Public Property ID() As Integer
        Get
            Return Me._ID
        End Get
        Set(ByVal value As Integer)
            If (Me._ID.Equals(value) = False) Then
                Me._ID = value
            End If
        End Set
    End Property


    Private _IDPdr As Integer
    <Column(Storage:="_IDPdr", DbType:="Int NOT NULL")> _
   Public Property IDPdr() As Integer
        Get
            Return Me._IDPdr
        End Get
        Set(ByVal value As Integer)
            If (Me._IDPdr.Equals(value) = False) Then
                Me._IDPdr = value
            End If
        End Set
    End Property


    Private _IDVinoPadre As System.Nullable(Of Guid)
    <Column(Storage:="_IDVinoPadre", DbType:="UniqueIdentifier NULL")> _
    Public Property IDVinoPadre() As System.Nullable(Of Guid)
        Get
            Return Me._IDVinoPadre
        End Get
        Set(ByVal value As System.Nullable(Of Guid))
            If (Me._IDVinoPadre Is Nothing OrElse (Me._IDVinoPadre = value) = False) Then
                Me._IDVinoPadre = value
            End If
        End Set
    End Property

    Private _IDVino As Guid
    <Column(Storage:="_IDVino", DbType:="UniqueIdentifier NOT NULL")> _
   Public Property IDVino() As System.Guid
        Get
            Return Me._IDVino
        End Get
        Set(ByVal value As System.Guid)
            If ((Me._IDVino = value) = False) Then
                Me._IDVino = value
            End If
        End Set
    End Property

    Private _Fecha As Date
    <Column(Storage:="_Fecha", DbType:="DateTime NOT NULL")> _
      Public Property Fecha() As Date
        Get
            Return Me._Fecha
        End Get
        Set(ByVal value As Date)
            If ((Me._Fecha = value) = False) Then
                Me._Fecha = value
            End If
        End Set
    End Property

    Private _Cantidad As Decimal
    <Column(Storage:="_Cantidad", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Cantidad() As Decimal
        Get
            Return Me._Cantidad
        End Get
        Set(ByVal value As Decimal)
            If ((Me._Cantidad = value) = False) Then
                Me._Cantidad = value
            End If
        End Set
    End Property
  
    Private _Prctj As System.Nullable(Of Decimal)
    <Column(Storage:="_Prctj", DbType:="Decimal(23,8) NULL")> _
    Public Property Prctj() As System.Nullable(Of Decimal)
        Get
            Return Me._Prctj
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Prctj Is Nothing OrElse (Me._Prctj = value) = False) Then
                Me._Prctj = value
            End If
        End Set
    End Property

    Private _QTotPadre As System.Nullable(Of Decimal)
    <Column(Storage:="_QTotPadre", DbType:="Decimal(23,8) NULL")> _
    Public Property QTotPadre() As System.Nullable(Of Decimal)
        Get
            Return Me._QTotPadre
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._QTotPadre Is Nothing OrElse (Me._QTotPadre = value) = False) Then
                Me._QTotPadre = value
            End If
        End Set
    End Property

    Private _QTot As System.Nullable(Of Decimal)
    <Column(Storage:="_QTot", DbType:="Decimal(23,8) NULL")> _
    Public Property QTot() As System.Nullable(Of Decimal)
        Get
            Return Me._QTot
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._QTot Is Nothing OrElse (Me._QTot = value) = False) Then
                Me._QTot = value
            End If
        End Set
    End Property

    Private _Factor As System.Nullable(Of Decimal)
    <Column(Storage:="_Factor", DbType:="Decimal(23,8) NULL")> _
    Public Property Factor() As System.Nullable(Of Decimal)
        Get
            Return Me._Factor
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Factor Is Nothing OrElse (Me._Factor = value) = False) Then
                Me._Factor = value
            End If
        End Set
    End Property

    Private _IDDeposito As String
    <Column(Storage:="_IDDeposito", DbType:="NVarChar(25)")> _
   Public Property IDDeposito() As String
        Get
            Return Me._IDDeposito
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDDeposito, value) = False) Then
                Me._IDDeposito = value
            End If
        End Set
    End Property
    
    Private _TipoDeposito As Integer
    <Column(Storage:="_TipoDeposito", DbType:="Int")> _
   Public Property TipoDeposito() As Integer
        Get
            Return Me._TipoDeposito
        End Get
        Set(ByVal value As Integer)
            If (Me._TipoDeposito.Equals(value) = False) Then
                Me._TipoDeposito = value
            End If
        End Set
    End Property

    Private _IDArticulo As String
    <Column(Storage:="_IDArticulo", DbType:="NVarChar(25)")> _
   Public Property IDArticulo() As String
        Get
            Return Me._IDArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticulo, value) = False) Then
                Me._IDArticulo = value
            End If
        End Set
    End Property

    Private _Lote As String
    <Column(Storage:="_Lote", DbType:="NVarChar(25)")> _
    Public Property Lote() As String
        Get
            Return Me._Lote
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Lote, value) = False) Then
                Me._Lote = value
            End If
        End Set
    End Property

    Private _Origen As Integer
    <Column(Storage:="_Origen", DbType:="Int")> _
   Public Property Origen() As Integer
        Get
            Return Me._Origen
        End Get
        Set(ByVal value As Integer)
            If (Me._Origen.Equals(value) = False) Then
                Me._Origen = value
            End If
        End Set
    End Property

    Private _NOperacion As String
    <Column(Storage:="_NOperacion", DbType:="NVarChar(10)")> _
    Public Property NOperacion() As String
        Get
            Return Me._NOperacion
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NOperacion, value) = False) Then
                Me._NOperacion = value
            End If
        End Set
    End Property

    Private _NAlbaran As String
    <Column(Storage:="_NAlbaran", DbType:="NVarChar(10)")> _
   Public Property NAlbaran() As String
        Get
            Return Me._NAlbaran
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NAlbaran, value) = False) Then
                Me._NAlbaran = value
            End If
        End Set
    End Property

    Private _IDAlbaran As System.Nullable(Of Integer)
    <Column(Storage:="_IDAlbaran", DbType:="Int NULL")> _
   Public Property IDAlbaran() As System.Nullable(Of Integer)
        Get
            Return Me._IDAlbaran
        End Get
        Set(ByVal value As System.Nullable(Of Integer))
            If (Me._IDAlbaran Is Nothing OrElse Me._IDAlbaran.Equals(value) = False) Then
                Me._IDAlbaran = value
            End If
        End Set
    End Property

    Private _QUdArticulo As System.Nullable(Of Decimal)
    <Column(Storage:="_QUdArticulo", DbType:="Decimal(23,8) NULL")> _
    Public Property QUdArticulo() As System.Nullable(Of Decimal)
        Get
            Return Me._QUdArticulo
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._QUdArticulo Is Nothing OrElse (Me._QUdArticulo = value) = False) Then
                Me._QUdArticulo = value
            End If
        End Set
    End Property

    Private _IDUDArticulo As String
    <Column(Storage:="_IDUDArticulo", DbType:="NVarChar(25) NULL")> _
    Public Property IDUDArticulo() As String
        Get
            Return Me._IDUDArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDUDArticulo, value) = False) Then
                Me._IDUDArticulo = value
            End If
        End Set
    End Property

    Private _QUdLitros As System.Nullable(Of Decimal)
    <Column(Storage:="_QUdLitros", DbType:="Decimal(23,8) NULL")> _
    Public Property QUdLitros() As System.Nullable(Of Decimal)
        Get
            Return Me._QUdLitros
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._QUdLitros Is Nothing OrElse (Me._QUdLitros = value) = False) Then
                Me._QUdLitros = value
            End If
        End Set
    End Property

    Private _IDUDLitros As String
    <Column(Storage:="_IDUDLitros", DbType:="NVarChar(25) NULL")> _
    Public Property IDUDLitros() As String
        Get
            Return Me._IDUDLitros
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDUDLitros, value) = False) Then
                Me._IDUDLitros = value
            End If
        End Set
    End Property

    Private _QUdDeposito As System.Nullable(Of Decimal)
    <Column(Storage:="_QUdDeposito", DbType:="Decimal(23,8) NULL")> _
    Public Property QUdDeposito() As System.Nullable(Of Decimal)
        Get
            Return Me._QUdDeposito
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._QUdDeposito Is Nothing OrElse (Me._QUdDeposito = value) = False) Then
                Me._QUdDeposito = value
            End If
        End Set
    End Property

    Private _IDUdDeposito As String
    <Column(Storage:="_IDUdDeposito", DbType:="NVarChar(25) NULL")> _
    Public Property IDUdDeposito() As String
        Get
            Return Me._IDUdDeposito
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDUdDeposito, value) = False) Then
                Me._IDUdDeposito = value
            End If
        End Set
    End Property

    Private _TieneCosteElaboracion As Boolean
    <Column(Storage:="_TieneCosteElaboracion", DbType:="Bit NOT NULL")> _
    Public Property TieneCosteElaboracion() As Boolean
        Get
            Return Me._TieneCosteElaboracion
        End Get
        Set(ByVal value As Boolean)
            If ((Me._TieneCosteElaboracion = value) = False) Then
                Me._TieneCosteElaboracion = value
            End If
        End Set
    End Property

    Private _TieneCosteEstanciaNave As Boolean
    <Column(Storage:="_TieneCosteEstanciaNave", DbType:="Bit NOT NULL")> _
    Public Property TieneCosteEstanciaNave() As Boolean
        Get
            Return Me._TieneCosteEstanciaNave
        End Get
        Set(ByVal value As Boolean)
            If ((Me._TieneCosteEstanciaNave = value) = False) Then
                Me._TieneCosteEstanciaNave = value
            End If
        End Set
    End Property

    Private _TieneCosteInicial As Boolean
    <Column(Storage:="_TieneCosteInicial", DbType:="Bit NOT NULL")> _
    Public Property TieneCosteInicial() As Boolean
        Get
            Return Me._TieneCosteInicial
        End Get
        Set(ByVal value As Boolean)
            If ((Me._TieneCosteInicial = value) = False) Then
                Me._TieneCosteInicial = value
            End If
        End Set
    End Property

    Private _FechaCosteUnitarioInicial As System.Nullable(Of Date)
    <Column(Storage:="_FechaCosteUnitarioInicial", DbType:="DateTime NULL")> _
   Public Property FechaCosteUnitarioInicial() As System.Nullable(Of Date)
        Get
            Return Me._FechaCosteUnitarioInicial
        End Get
        Set(ByVal value As System.Nullable(Of Date))
            If (Me._FechaCosteUnitarioInicial Is Nothing OrElse (Me._FechaCosteUnitarioInicial = value) = False) Then
                Me._FechaCosteUnitarioInicial = value
            End If
        End Set
    End Property

    Private _DescArticulo As String
    <Column(Storage:="_DescArticulo", DbType:="NVarChar(50) NULL")> _
    Public Property DescArticulo() As String
        Get
            Return Me._DescArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescArticulo, value) = False) Then
                Me._DescArticulo = value
            End If
        End Set
    End Property

    Private _IDAlmacen As String
    <Column(Storage:="_IDAlmacen", DbType:="NVarChar(10)")> _
   Public Property IDAlmacen() As String
        Get
            Return Me._IDAlmacen
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacen, value) = False) Then
                Me._IDAlmacen = value
            End If
        End Set
    End Property

    Public Function Equals(ByVal o As VinoExplosionArbol) As Boolean Implements System.IEquatable(Of Solmicro.Expertis.Business.Bodega.VinoExplosionArbol).Equals
        'If Me.ID.Equals(o.ID) AndAlso Me.IDPdr = o.IDArticulo AndAlso Me.IDAlmacen = o.IDAlmacen AndAlso Me.Lote = o.Lote AndAlso Me.Ubicacion = o.Ubicacion AndAlso Me.IDUdInterna = o.IDUdInterna AndAlso Me.StockFisico = o.StockFisico Then
        '    Return True
        'End If
    End Function

End Class

<Serializable()> _
Public Class VinoExplosionArbolDetalle
    Implements IEquatable(Of VinoExplosionArbolDetalle)

    Private _IDVino As Guid
    <Column(Storage:="_IDVino", DbType:="UniqueIdentifier NOT NULL")> _
    Public Property IDVino() As System.Guid
        Get
            Return Me._IDVino
        End Get
        Set(ByVal value As System.Guid)
            If ((Me._IDVino = value) = False) Then
                Me._IDVino = value
            End If
        End Set
    End Property


    Private _TipoNodo As System.Nullable(Of Integer)
    <Column(Storage:="_TipoNodo", DbType:="Int", CanBeNull:=True)> _
    Public Property TipoNodo() As System.Nullable(Of Integer)
        Get
            Return Me._TipoNodo
        End Get
        Set(ByVal value As System.Nullable(Of Integer))
            If (Me._TipoNodo.Equals(value) = False) Then
                Me._TipoNodo = value
            End If
        End Set
    End Property


    Private _Fecha As System.Nullable(Of Date)
    <Column(Storage:="_Fecha", DbType:="DateTime NULL")> _
   Public Property Fecha() As System.Nullable(Of Date)
        Get
            Return Me._Fecha
        End Get
        Set(ByVal value As System.Nullable(Of Date))
            If ((Me._Fecha = value) = False) Then
                Me._Fecha = value
            End If
        End Set
    End Property


    Private _NOperacion As String
    <Column(Storage:="_NOperacion", DbType:="NVarChar(10)")> _
    Public Property NOperacion() As String
        Get
            Return Me._NOperacion
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NOperacion, value) = False) Then
                Me._NOperacion = value
            End If
        End Set
    End Property


    Private _IDNodo As String
    <Column(Storage:="_IDNodo", DbType:="NVarChar(25)")> _
    Public Property IDNodo() As String
        Get
            Return Me._IDNodo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDNodo, value) = False) Then
                Me._IDNodo = value
            End If
        End Set
    End Property

    Private _DescNodo As String
    <Column(Storage:="_DescNodo", DbType:="NVarChar(500)")> _
    Public Property DescNodo() As String
        Get
            Return Me._DescNodo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescNodo, value) = False) Then
                Me._DescNodo = value
            End If
        End Set
    End Property

    Private _IDDocumentoInt As System.Nullable(Of Integer)
    <Column(Storage:="_IDDocumentoInt", DbType:="Int NULL")> _
    Public Property IDDocumentoInt() As System.Nullable(Of Integer)
        Get
            Return Me._IDDocumentoInt
        End Get
        Set(ByVal value As System.Nullable(Of Integer))
            If ((Me._IDDocumentoInt = value) = False) Then
                Me._IDDocumentoInt = value
            End If
        End Set
    End Property

    Private _IDDocumentoGuid As System.Nullable(Of Guid)
    <Column(Storage:="_IDDocumentoGuid", DbType:="UniqueIdentifier NULL")> _
    Public Property IDDocumentoGuid() As System.Nullable(Of Guid)
        Get
            Return Me._IDDocumentoGuid
        End Get
        Set(ByVal value As System.Nullable(Of Guid))
            If ((Me._IDDocumentoGuid = value) = False) Then
                Me._IDDocumentoGuid = value
            End If
        End Set
    End Property

    Private _NDocumento As String
    <Column(Storage:="_NDocumento", DbType:="NVarChar(50)")> _
    Public Property NDocumento() As String
        Get
            Return Me._NDocumento
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NDocumento, value) = False) Then
                Me._NDocumento = value
            End If
        End Set
    End Property

    Private _IDUdInterna As String
    <Column(Storage:="_IDUdInterna", DbType:="NVarChar(10) NULL")> _
    Public Property IDUdInterna() As String
        Get
            Return Me._IDUdInterna
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDUdInterna, value) = False) Then
                Me._IDUdInterna = value
            End If
        End Set
    End Property

    Private _Cantidad As System.Nullable(Of Decimal)
    <Column(Storage:="_Cantidad", DbType:="Decimal(23,8) NULL")> _
    Public Property Cantidad() As System.Nullable(Of Decimal)
        Get
            Return Me._Cantidad
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If ((Me._Cantidad = value) = False) Then
                Me._Cantidad = value
            End If
        End Set
    End Property


    Private _Lote As String
    <Column(Storage:="_Lote", DbType:="NVarChar(25) NULL")> _
    Public Property Lote() As String
        Get
            Return Me._Lote
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._Lote, value) = False) Then
                Me._Lote = value
            End If
        End Set
    End Property

    Private _Merma As System.Nullable(Of Decimal)
    <Column(Storage:="_Merma", DbType:="Decimal(23,8) NULL")> _
    Public Property Merma() As System.Nullable(Of Decimal)
        Get
            Return Me._Merma
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If ((Me._Merma = value) = False) Then
                Me._Merma = value
            End If
        End Set
    End Property


    Public Function Equals(ByVal o As VinoExplosionArbolDetalle) As Boolean Implements System.IEquatable(Of Solmicro.Expertis.Business.Bodega.VinoExplosionArbolDetalle).Equals
        'If Me.Traza.Equals(o.Traza) AndAlso Me.IDArticulo = o.IDArticulo AndAlso Me.IDAlmacen = o.IDAlmacen AndAlso Me.Lote = o.Lote AndAlso Me.Ubicacion = o.Ubicacion AndAlso Me.IDUdInterna = o.IDUdInterna AndAlso Me.StockFisico = o.StockFisico Then
        '    Return True
        'End If
    End Function

End Class
