Imports System.Data.Linq.Mapping

Partial Public Class CostWineDataContext

    <[Function](Name:="spBdgCosteVinoArticulo")> _
    Public Overridable Function GetVinosArticulo(<Parameter(Name:="pArticulo", DbType:="nvarchar (25)")> ByVal IDArticulo As String, _
                                    <Parameter(Name:="pFechaHasta", DbType:="datetime")> ByVal FechaHasta As Date) As ISingleResult(Of VinoArticulo)
        Dim method As MethodInfo = DirectCast(MethodInfo.GetCurrentMethod(), MethodInfo)
        Dim result As IExecuteResult = Me.ExecuteMethodCall(Me, method, IDArticulo, FechaHasta)
        Return DirectCast(result.ReturnValue, ISingleResult(Of VinoArticulo))
    End Function

    <[Function](Name:="spBdgCosteVinoArticuloFabrica")> _
    Public Overridable Function GetVinosArticuloFabrica(<Parameter(Name:="pArticulo", DbType:="nvarchar (25)")> ByVal IDArticulo As String, _
                                            <Parameter(Name:="pFechaDesde", DbType:="datetime")> ByVal FechaDesde As Date, _
                                            <Parameter(Name:="pFechaHasta", DbType:="datetime")> ByVal FechaHasta As Date) As ISingleResult(Of VinoArticulo)
        Dim method As MethodInfo = DirectCast(MethodInfo.GetCurrentMethod(), MethodInfo)
        Dim result As IExecuteResult = Me.ExecuteMethodCall(Me, method, IDArticulo, FechaDesde, FechaHasta)
        Return DirectCast(result.ReturnValue, ISingleResult(Of VinoArticulo))
    End Function

    <[Function](Name:="spBdgCosteVinoArticuloExplosion")> _
     <ResultType(GetType(VinoCosteTotal))> _
       Public Overridable Function GetVinoCosteTotales(<Parameter(Name:="idQuery", DbType:="int")> ByVal IDQuery As Integer, _
                                           <Parameter(Name:="pVino", DbType:="UniqueIdentifier")> ByVal IDVino As Guid, _
                                           <Parameter(Name:="pDevolverSoloTotales", DbType:="bit")> ByVal DevolverSoloTotales As Boolean, _
                                           <Parameter(Name:="pFechaDesde", DbType:="datetime")> ByVal FechaDesde As Date, _
                                           <Parameter(Name:="pFechaHasta", DbType:="datetime")> ByVal FechaHasta As Date, _
                                           <Parameter(Name:="ConsiderarNoTrazar", DbType:="bit")> ByVal ConsiderarNoTrazar As Boolean, _
                                           <Parameter(Name:="ConsiderarCosteInicial", DbType:="bit")> ByVal ConsiderarCosteInicial As Boolean) As ISingleResult(Of VinoCosteTotal)
        Dim result As IExecuteResult = Me.ExecuteMethodCall(Me, DirectCast(MethodInfo.GetCurrentMethod(), MethodInfo), IDQuery, IDVino, True, FechaDesde, FechaHasta, ConsiderarNoTrazar, ConsiderarCosteInicial)
        Return DirectCast(result.ReturnValue, ISingleResult(Of VinoCosteTotal))
    End Function

    <[Function](Name:="spBdgCosteVinoArticuloExplosion")> _
      <ResultType(GetType(VinoCosteMateriales))> _
      <ResultType(GetType(VinoCosteMOD))> _
      <ResultType(GetType(VinoCosteCentros))> _
      <ResultType(GetType(VinoCosteVarios))> _
      <ResultType(GetType(VinoCosteCompras))> _
      <ResultType(GetType(VinoCosteTasas))> _
      <ResultType(GetType(VinoCosteVendimiaElaboracion))> _
      <ResultType(GetType(VinoCosteEntradaUVA))> _
      <ResultType(GetType(VinoCosteEstanciaEnNave))> _
      <ResultType(GetType(VinoCosteInicial))> _
      <ResultType(GetType(VinoCosteTotal))> _
     Public Overridable Function GetVinoCosteDetalle(<Parameter(Name:="idQuery", DbType:="int")> ByVal IDQuery As Integer, _
                                         <Parameter(Name:="pVino", DbType:="UniqueIdentifier")> ByVal IDVino As Guid, _
                                         <Parameter(Name:="pDevolverSoloTotales", DbType:="bit")> ByVal DevolverSoloTotales As Boolean, _
                                         <Parameter(Name:="pFechaDesde", DbType:="datetime")> ByVal FechaDesde As Date, _
                                         <Parameter(Name:="pFechaHasta", DbType:="datetime")> ByVal FechaHasta As Date, _
                                         <Parameter(Name:="ConsiderarNoTrazar", DbType:="bit")> ByVal ConsiderarNoTrazar As Boolean, _
                                         <Parameter(Name:="ConsiderarCosteInicial", DbType:="bit")> ByVal ConsiderarCosteInicial As Boolean) As IMultipleResults
        Dim result As IExecuteResult = Me.ExecuteMethodCall(Me, DirectCast(MethodInfo.GetCurrentMethod(), MethodInfo), IDQuery, IDVino, False, FechaDesde, FechaHasta, ConsiderarNoTrazar, ConsiderarCosteInicial)
        Return DirectCast(result.ReturnValue, IMultipleResults)
    End Function
  
End Class


<Serializable()> _
Public Class VinoCosteMateriales

    Private _IDDeposito As String

    Private _LoteVino As String

    Private _IDAlmacenVino As String

    Private _FechaVino As Date

    Private _IDArticuloVino As String

    Private _DescArticulo As String

    Private _IDVino As System.Guid

    Private _NOperacion As String

    Private _IDArticulo As String

    Private _Fecha As System.Nullable(Of Date)

    Private _Cantidad As System.Nullable(Of Decimal)

    Private _Merma As System.Nullable(Of Decimal)

    Private _Precio As System.Nullable(Of Decimal)

    Private _PorcentajeAcumulado As System.Nullable(Of Decimal)

    Private _Coste As System.Nullable(Of Decimal)

    Private _CosteMerma As System.Nullable(Of Decimal)

    Private _CosteUnit As System.Nullable(Of Decimal)

    Private _CosteUnitMerma As System.Nullable(Of Decimal)

    Public Sub New()
        MyBase.New()
    End Sub

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

    <Column(Storage:="_LoteVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property LoteVino() As String
        Get
            Return Me._LoteVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._LoteVino, value) = False) Then
                Me._LoteVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlmacenVino", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAlmacenVino() As String
        Get
            Return Me._IDAlmacenVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacenVino, value) = False) Then
                Me._IDAlmacenVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaVino", DbType:="DateTime NOT NULL")> _
    Public Property FechaVino() As Date
        Get
            Return Me._FechaVino
        End Get
        Set(ByVal value As Date)
            If ((Me._FechaVino = value) _
               = False) Then
                Me._FechaVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticuloVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticuloVino() As String
        Get
            Return Me._IDArticuloVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticuloVino, value) = False) Then
                Me._IDArticuloVino = value
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

    <Column(Storage:="_NOperacion", DbType:="NVarChar(50)")> _
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

    <Column(Storage:="_IDArticulo", DbType:="NVarChar(50)")> _
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

    <Column(Storage:="_Fecha", DbType:="DateTime")> _
    Public Property Fecha() As System.Nullable(Of Date)
        Get
            Return Me._Fecha
        End Get
        Set(ByVal value As System.Nullable(Of Date))
            If (Me._Fecha.Equals(value) = False) Then
                Me._Fecha = value
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

    <Column(Storage:="_Merma", DbType:="Decimal(23,8)")> _
    Public Property Merma() As System.Nullable(Of Decimal)
        Get
            Return Me._Merma
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Merma.Equals(value) = False) Then
                Me._Merma = value
            End If
        End Set
    End Property

    <Column(Storage:="_Precio", DbType:="Decimal(23,8)")> _
    Public Property Precio() As System.Nullable(Of Decimal)
        Get
            Return Me._Precio
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Precio.Equals(value) = False) Then
                Me._Precio = value
            End If
        End Set
    End Property

    <Column(Storage:="_PorcentajeAcumulado", DbType:="Decimal(23,8)")> _
    Public Property PorcentajeAcumulado() As System.Nullable(Of Decimal)
        Get
            Return Me._PorcentajeAcumulado
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._PorcentajeAcumulado.Equals(value) = False) Then
                Me._PorcentajeAcumulado = value
            End If
        End Set
    End Property

    <Column(Storage:="_Coste", DbType:="Decimal(23,8)")> _
    Public Property Coste() As System.Nullable(Of Decimal)
        Get
            Return Me._Coste
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Coste.Equals(value) = False) Then
                Me._Coste = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteMerma", DbType:="Decimal(23,8)")> _
    Public Property CosteMerma() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteMerma
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteMerma.Equals(value) = False) Then
                Me._CosteMerma = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteUnit.Equals(value) = False) Then
                Me._CosteUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUnitMerma", DbType:="Decimal(23,8)")> _
    Public Property CosteUnitMerma() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteUnitMerma
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteUnitMerma.Equals(value) = False) Then
                Me._CosteUnitMerma = value
            End If
        End Set
    End Property
End Class


<Serializable()> _
Public Class VinoCosteVendimiaElaboracion

    Private _Nodo As Integer

    Private _NodoPadre As Integer

    Private _IDDeposito As String

    Private _IDArticuloVino As String

    Private _LoteVino As String

    Private _IDAlmacen As String

    Private _FechaVino As Date

    Private _DescArticulo As String

    Private _IDVino As System.Guid

    Private _IDArticulo As String

    Private _Fecha As System.Nullable(Of Date)

    Private _Precio As System.Nullable(Of Decimal)

    Private _Cantidad As System.Nullable(Of Decimal)

    Private _PorcentajeAcumulado As System.Nullable(Of Decimal)

    Private _Coste As System.Nullable(Of Decimal)

    Private _CosteUnit As System.Nullable(Of Decimal)

    Private _CantidadAcum As System.Nullable(Of Decimal)

    Private _QTotal As Decimal

    Private _DifQElab As System.Nullable(Of Decimal)

    Public Sub New()
        MyBase.New()
    End Sub

    <Column(Storage:="_Nodo", DbType:="Int NOT NULL")> _
    Public Property Nodo() As Integer
        Get
            Return Me._Nodo
        End Get
        Set(ByVal value As Integer)
            If ((Me._Nodo = value) _
               = False) Then
                Me._Nodo = value
            End If
        End Set
    End Property

    <Column(Storage:="_NodoPadre", DbType:="Int NOT NULL")> _
    Public Property NodoPadre() As Integer
        Get
            Return Me._NodoPadre
        End Get
        Set(ByVal value As Integer)
            If ((Me._NodoPadre = value) _
               = False) Then
                Me._NodoPadre = value
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

    <Column(Storage:="_IDArticuloVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticuloVino() As String
        Get
            Return Me._IDArticuloVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticuloVino, value) = False) Then
                Me._IDArticuloVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_LoteVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property LoteVino() As String
        Get
            Return Me._LoteVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._LoteVino, value) = False) Then
                Me._LoteVino = value
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

    <Column(Storage:="_FechaVino", DbType:="DateTime NOT NULL")> _
    Public Property FechaVino() As Date
        Get
            Return Me._FechaVino
        End Get
        Set(ByVal value As Date)
            If ((Me._FechaVino = value) _
               = False) Then
                Me._FechaVino = value
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

    <Column(Storage:="_IDArticulo", DbType:="NVarChar(50)")> _
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

    <Column(Storage:="_Fecha", DbType:="DateTime")> _
    Public Property Fecha() As System.Nullable(Of Date)
        Get
            Return Me._Fecha
        End Get
        Set(ByVal value As System.Nullable(Of Date))
            If (Me._Fecha.Equals(value) = False) Then
                Me._Fecha = value
            End If
        End Set
    End Property

    <Column(Storage:="_Precio", DbType:="Decimal(23,8)")> _
    Public Property Precio() As System.Nullable(Of Decimal)
        Get
            Return Me._Precio
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Precio.Equals(value) = False) Then
                Me._Precio = value
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

    <Column(Storage:="_PorcentajeAcumulado", DbType:="Decimal(23,8)")> _
    Public Property PorcentajeAcumulado() As System.Nullable(Of Decimal)
        Get
            Return Me._PorcentajeAcumulado
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._PorcentajeAcumulado.Equals(value) = False) Then
                Me._PorcentajeAcumulado = value
            End If
        End Set
    End Property

    <Column(Storage:="_Coste", DbType:="Decimal(23,8)")> _
    Public Property Coste() As System.Nullable(Of Decimal)
        Get
            Return Me._Coste
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Coste.Equals(value) = False) Then
                Me._Coste = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteUnit.Equals(value) = False) Then
                Me._CosteUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CantidadAcum", DbType:="Decimal(23,8)")> _
    Public Property CantidadAcum() As System.Nullable(Of Decimal)
        Get
            Return Me._CantidadAcum
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CantidadAcum.Equals(value) = False) Then
                Me._CantidadAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_QTotal", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property QTotal() As Decimal
        Get
            Return Me._QTotal
        End Get
        Set(ByVal value As Decimal)
            If ((Me._QTotal = value) _
               = False) Then
                Me._QTotal = value
            End If
        End Set
    End Property

    <Column(Storage:="_DifQElab", DbType:="Decimal(23,8)")> _
    Public Property DifQElab() As System.Nullable(Of Decimal)
        Get
            Return Me._DifQElab
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._DifQElab.Equals(value) = False) Then
                Me._DifQElab = value
            End If
        End Set
    End Property
End Class

<Serializable()> _
Public Class VinoCosteCentros

    Private _IDVino As System.Guid

    Private _IDDeposito As String

    Private _LoteVino As String

    Private _IDAlmacenVino As String

    Private _FechaVino As Date

    Private _IDArticuloVino As String

    Private _NOperacion As String

    Private _Fecha As Date

    Private _IDCentro As String

    Private _DescCentro As String

    Private _Tiempo As Decimal

    Private _UDTiempo As Integer

    Private _Tasa As Decimal

    Private _PorCantidad As Boolean

    Private _Cantidad As Decimal

    Private _IdUdMedidaArticulo As String

    Private _IdUdMedidaCentro As String

    Private _PorcentajeAcum As Decimal

    Private _Coste As System.Nullable(Of Decimal)

    Private _CosteUnit As System.Nullable(Of Decimal)

    Private _TiempoPorcen As System.Nullable(Of Decimal)

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

    <Column(Storage:="_LoteVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property LoteVino() As String
        Get
            Return Me._LoteVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._LoteVino, value) = False) Then
                Me._LoteVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlmacenVino", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAlmacenVino() As String
        Get
            Return Me._IDAlmacenVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacenVino, value) = False) Then
                Me._IDAlmacenVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaVino", DbType:="DateTime NOT NULL")> _
    Public Property FechaVino() As Date
        Get
            Return Me._FechaVino
        End Get
        Set(ByVal value As Date)
            If ((Me._FechaVino = value) _
               = False) Then
                Me._FechaVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticuloVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticuloVino() As String
        Get
            Return Me._IDArticuloVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticuloVino, value) = False) Then
                Me._IDArticuloVino = value
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

    <Column(Storage:="_IDCentro", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDCentro() As String
        Get
            Return Me._IDCentro
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDCentro, value) = False) Then
                Me._IDCentro = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescCentro", DbType:="NVarChar(100)")> _
    Public Property DescCentro() As String
        Get
            Return Me._DescCentro
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescCentro, value) = False) Then
                Me._DescCentro = value
            End If
        End Set
    End Property

    <Column(Storage:="_Tiempo", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Tiempo() As Decimal
        Get
            Return Me._Tiempo
        End Get
        Set(ByVal value As Decimal)
            If ((Me._Tiempo = value) _
               = False) Then
                Me._Tiempo = value
            End If
        End Set
    End Property

    <Column(Storage:="_UDTiempo", DbType:="Int NOT NULL")> _
    Public Property UDTiempo() As Integer
        Get
            Return Me._UDTiempo
        End Get
        Set(ByVal value As Integer)
            If ((Me._UDTiempo = value) _
               = False) Then
                Me._UDTiempo = value
            End If
        End Set
    End Property

    <Column(Storage:="_Tasa", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Tasa() As Decimal
        Get
            Return Me._Tasa
        End Get
        Set(ByVal value As Decimal)
            If ((Me._Tasa = value) _
               = False) Then
                Me._Tasa = value
            End If
        End Set
    End Property

    <Column(Storage:="_PorCantidad", DbType:="Bit NOT NULL")> _
    Public Property PorCantidad() As Boolean
        Get
            Return Me._PorCantidad
        End Get
        Set(ByVal value As Boolean)
            If ((Me._PorCantidad = value) _
               = False) Then
                Me._PorCantidad = value
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

    <Column(Storage:="_IdUdMedidaArticulo", DbType:="NVarChar(10)")> _
    Public Property IdUdMedidaArticulo() As String
        Get
            Return Me._IdUdMedidaArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IdUdMedidaArticulo, value) = False) Then
                Me._IdUdMedidaArticulo = value
            End If
        End Set
    End Property

    <Column(Storage:="_IdUdMedidaCentro", DbType:="NVarChar(10)")> _
    Public Property IdUdMedidaCentro() As String
        Get
            Return Me._IdUdMedidaCentro
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IdUdMedidaCentro, value) = False) Then
                Me._IdUdMedidaCentro = value
            End If
        End Set
    End Property

    <Column(Storage:="_PorcentajeAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property PorcentajeAcum() As Decimal
        Get
            Return Me._PorcentajeAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._PorcentajeAcum = value) _
               = False) Then
                Me._PorcentajeAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_Coste", DbType:="Decimal(23,8)")> _
    Public Property Coste() As System.Nullable(Of Decimal)
        Get
            Return Me._Coste
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Coste.Equals(value) = False) Then
                Me._Coste = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteUnit.Equals(value) = False) Then
                Me._CosteUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_TiempoPorcen", DbType:="Decimal(23,8)")> _
    Public Property TiempoPorcen() As System.Nullable(Of Decimal)
        Get
            Return Me._TiempoPorcen
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TiempoPorcen.Equals(value) = False) Then
                Me._TiempoPorcen = value
            End If
        End Set
    End Property
End Class

<Serializable()> _
Public Class VinoCosteMOD

    Private _IDVino As System.Guid

    Private _IDDeposito As String

    Private _LoteVino As String

    Private _IDAlmacenVino As String

    Private _FechaVino As Date

    Private _IDArticuloVino As String

    Private _NOperacion As String

    Private _Fecha As Date

    Private _IDOperario As String

    Private _DescOperario As String

    Private _Tiempo As Decimal

    Private _Tasa As Decimal

    Private _PorcentajeAcum As Decimal

    Private _TiempoAcum As System.Nullable(Of Decimal)

    Private _Coste As System.Nullable(Of Decimal)

    Private _CosteUnit As System.Nullable(Of Decimal)

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

    <Column(Storage:="_LoteVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property LoteVino() As String
        Get
            Return Me._LoteVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._LoteVino, value) = False) Then
                Me._LoteVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlmacenVino", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAlmacenVino() As String
        Get
            Return Me._IDAlmacenVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacenVino, value) = False) Then
                Me._IDAlmacenVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaVino", DbType:="DateTime NOT NULL")> _
    Public Property FechaVino() As Date
        Get
            Return Me._FechaVino
        End Get
        Set(ByVal value As Date)
            If ((Me._FechaVino = value) _
               = False) Then
                Me._FechaVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticuloVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticuloVino() As String
        Get
            Return Me._IDArticuloVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticuloVino, value) = False) Then
                Me._IDArticuloVino = value
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

    <Column(Storage:="_IDOperario", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDOperario() As String
        Get
            Return Me._IDOperario
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDOperario, value) = False) Then
                Me._IDOperario = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescOperario", DbType:="NVarChar(302)")> _
    Public Property DescOperario() As String
        Get
            Return Me._DescOperario
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescOperario, value) = False) Then
                Me._DescOperario = value
            End If
        End Set
    End Property

    <Column(Storage:="_Tiempo", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Tiempo() As Decimal
        Get
            Return Me._Tiempo
        End Get
        Set(ByVal value As Decimal)
            If ((Me._Tiempo = value) _
               = False) Then
                Me._Tiempo = value
            End If
        End Set
    End Property

    <Column(Storage:="_Tasa", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Tasa() As Decimal
        Get
            Return Me._Tasa
        End Get
        Set(ByVal value As Decimal)
            If ((Me._Tasa = value) _
               = False) Then
                Me._Tasa = value
            End If
        End Set
    End Property

    <Column(Storage:="_PorcentajeAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property PorcentajeAcum() As Decimal
        Get
            Return Me._PorcentajeAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._PorcentajeAcum = value) _
               = False) Then
                Me._PorcentajeAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_TiempoAcum", DbType:="Decimal(23,8)")> _
    Public Property TiempoAcum() As System.Nullable(Of Decimal)
        Get
            Return Me._TiempoAcum
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TiempoAcum.Equals(value) = False) Then
                Me._TiempoAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_Coste", DbType:="Decimal(23,8)")> _
    Public Property Coste() As System.Nullable(Of Decimal)
        Get
            Return Me._Coste
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Coste.Equals(value) = False) Then
                Me._Coste = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteUnit.Equals(value) = False) Then
                Me._CosteUnit = value
            End If
        End Set
    End Property
End Class

<Serializable()> _
Public Class VinoCosteCompras

    Private _IDVino As System.Guid

    Private _IDDeposito As String

    Private _LoteVino As String

    Private _IDAlmacenVino As String

    Private _FechaVino As Date

    Private _IDArticuloVino As String

    Private _NEntrada As Integer

    Private _Fecha As Date

    Private _PorcentajeAcumulado As Decimal

    Private _Importe As System.Nullable(Of Decimal)

    Private _Coste As System.Nullable(Of Decimal)

    Private _CosteUnit As System.Nullable(Of Decimal)

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

    <Column(Storage:="_LoteVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property LoteVino() As String
        Get
            Return Me._LoteVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._LoteVino, value) = False) Then
                Me._LoteVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlmacenVino", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAlmacenVino() As String
        Get
            Return Me._IDAlmacenVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacenVino, value) = False) Then
                Me._IDAlmacenVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaVino", DbType:="DateTime NOT NULL")> _
    Public Property FechaVino() As Date
        Get
            Return Me._FechaVino
        End Get
        Set(ByVal value As Date)
            If ((Me._FechaVino = value) _
               = False) Then
                Me._FechaVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticuloVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticuloVino() As String
        Get
            Return Me._IDArticuloVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticuloVino, value) = False) Then
                Me._IDArticuloVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_NEntrada", DbType:="Int NOT NULL")> _
    Public Property NEntrada() As Integer
        Get
            Return Me._NEntrada
        End Get
        Set(ByVal value As Integer)
            If ((Me._NEntrada = value) _
               = False) Then
                Me._NEntrada = value
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

    <Column(Storage:="_PorcentajeAcumulado", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property PorcentajeAcumulado() As Decimal
        Get
            Return _PorcentajeAcumulado
        End Get
        Set(ByVal value As Decimal)
            If ((Me._PorcentajeAcumulado = value) _
               = False) Then
                Me._PorcentajeAcumulado = value
            End If
        End Set
    End Property

    <Column(Storage:="_Importe", DbType:="Decimal(23,8)")> _
    Public Property Importe() As System.Nullable(Of Decimal)
        Get
            Return Me._Importe
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Importe.Equals(value) = False) Then
                Me._Importe = value
            End If
        End Set
    End Property

    <Column(Storage:="_Coste", DbType:="Decimal(23,8)")> _
    Public Property Coste() As System.Nullable(Of Decimal)
        Get
            Return Me._Coste
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Coste.Equals(value) = False) Then
                Me._Coste = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteUnit.Equals(value) = False) Then
                Me._CosteUnit = value
            End If
        End Set
    End Property
End Class

<Serializable()> _
Public Class VinoCosteEntradaUVA

    Private _IDVino As System.Nullable(Of System.Guid)

    Private _IDDeposito As String

    Private _LoteVino As String

    Private _IDAlmacenVino As String

    Private _FechaVino As Date

    Private _IDArticuloVino As String

    Private _IDEntrada As Integer

    Private _NEntrada As String

    Private _Fecha As Date

    Private _Vendimia As Integer

    Private _Cantidad As Decimal

    Private _CosteKgUva As Decimal

    Private _PorcentajeAcumulado As Decimal

    Private _Coste As System.Nullable(Of Decimal)

    Private _CosteUnit As System.Nullable(Of Decimal)

    Public Sub New()
        MyBase.New()
    End Sub

    <Column(Storage:="_IDVino", DbType:="UniqueIdentifier")> _
    Public Property IDVino() As System.Nullable(Of System.Guid)
        Get
            Return Me._IDVino
        End Get
        Set(ByVal value As System.Nullable(Of System.Guid))
            If (Me._IDVino.Equals(value) = False) Then
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

    <Column(Storage:="_LoteVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property LoteVino() As String
        Get
            Return Me._LoteVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._LoteVino, value) = False) Then
                Me._LoteVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlmacenVino", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAlmacenVino() As String
        Get
            Return Me._IDAlmacenVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacenVino, value) = False) Then
                Me._IDAlmacenVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaVino", DbType:="DateTime NOT NULL")> _
    Public Property FechaVino() As Date
        Get
            Return Me._FechaVino
        End Get
        Set(ByVal value As Date)
            If ((Me._FechaVino = value) _
               = False) Then
                Me._FechaVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticuloVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticuloVino() As String
        Get
            Return Me._IDArticuloVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticuloVino, value) = False) Then
                Me._IDArticuloVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDEntrada", DbType:="Int NOT NULL")> _
    Public Property IDEntrada() As Integer
        Get
            Return Me._IDEntrada
        End Get
        Set(ByVal value As Integer)
            If ((Me._IDEntrada = value) _
               = False) Then
                Me._IDEntrada = value
            End If
        End Set
    End Property

    <Column(Storage:="_NEntrada", DbType:="NVarChar(25)")> _
    Public Property NEntrada() As String
        Get
            Return Me._NEntrada
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._NEntrada, value) = False) Then
                Me._NEntrada = value
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

    <Column(Storage:="_Vendimia", DbType:="Int NOT NULL")> _
    Public Property Vendimia() As Integer
        Get
            Return Me._Vendimia
        End Get
        Set(ByVal value As Integer)
            If ((Me._Vendimia = value) _
               = False) Then
                Me._Vendimia = value
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

    <Column(Storage:="_CosteKgUva", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteKgUva() As Decimal
        Get
            Return Me._CosteKgUva
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteKgUva = value) _
               = False) Then
                Me._CosteKgUva = value
            End If
        End Set
    End Property

    <Column(Storage:="_PorcentajeAcumulado", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property PorcentajeAcumulado() As Decimal
        Get
            Return Me._PorcentajeAcumulado
        End Get
        Set(ByVal value As Decimal)
            If ((Me._PorcentajeAcumulado = value) _
               = False) Then
                Me._PorcentajeAcumulado = value
            End If
        End Set
    End Property

    <Column(Storage:="_Coste", DbType:="Decimal(23,8)")> _
    Public Property Coste() As System.Nullable(Of Decimal)
        Get
            Return Me._Coste
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Coste.Equals(value) = False) Then
                Me._Coste = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteUnit.Equals(value) = False) Then
                Me._CosteUnit = value
            End If
        End Set
    End Property
End Class

<Serializable()> _
Public Class VinoCosteEstanciaEnNave

    Private _idVino As System.Guid

    Private _IDDeposito As String

    Private _LoteVino As String

    Private _IDAlmacenVino As String

    Private _FechaVino As Date

    Private _IDArticuloVino As String

    Private _QTotal As Decimal

    Private _IDCentro As String

    Private _Tiempo As Decimal

    Private _TiempoDias As System.Nullable(Of Decimal)

    Private _FechaInicioCosteNave As System.Nullable(Of Date)

    Private _FechaFinCosteNave As System.Nullable(Of Date)

    Private _Tasa As Decimal

    Private _IdUdMedidaArticulo As String

    Private _PorCantidad As Boolean

    Private _IdExpCosteNave As Integer

    Private _CantidadCoste As System.Nullable(Of Decimal)

    Private _CantidadCosteLitros As System.Nullable(Of Decimal)

    Private _Importe As System.Nullable(Of Decimal)

    Private _IdUdMedidaCentro As String

    Private _FactorTasa As System.Nullable(Of Decimal)

    Private _AnioMesFechaInicio As String

    Private _CosteUnit As System.Nullable(Of Decimal)

    Public Sub New()
        MyBase.New()
    End Sub

    <Column(Storage:="_idVino", DbType:="UniqueIdentifier NOT NULL")> _
    Public Property idVino() As System.Guid
        Get
            Return Me._idVino
        End Get
        Set(ByVal value As System.Guid)
            If ((Me._idVino = value) _
               = False) Then
                Me._idVino = value
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

    <Column(Storage:="_LoteVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property LoteVino() As String
        Get
            Return Me._LoteVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._LoteVino, value) = False) Then
                Me._LoteVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlmacenVino", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAlmacenVino() As String
        Get
            Return Me._IDAlmacenVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacenVino, value) = False) Then
                Me._IDAlmacenVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaVino", DbType:="DateTime NOT NULL")> _
    Public Property FechaVino() As Date
        Get
            Return Me._FechaVino
        End Get
        Set(ByVal value As Date)
            If ((Me._FechaVino = value) _
               = False) Then
                Me._FechaVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticuloVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticuloVino() As String
        Get
            Return Me._IDArticuloVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticuloVino, value) = False) Then
                Me._IDArticuloVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_QTotal", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property QTotal() As Decimal
        Get
            Return Me._QTotal
        End Get
        Set(ByVal value As Decimal)
            If ((Me._QTotal = value) _
               = False) Then
                Me._QTotal = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDCentro", DbType:="NVarChar(25)")> _
    Public Property IDCentro() As String
        Get
            Return Me._IDCentro
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDCentro, value) = False) Then
                Me._IDCentro = value
            End If
        End Set
    End Property

    <Column(Storage:="_Tiempo", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Tiempo() As Decimal
        Get
            Return Me._Tiempo
        End Get
        Set(ByVal value As Decimal)
            If ((Me._Tiempo = value) _
               = False) Then
                Me._Tiempo = value
            End If
        End Set
    End Property

    <Column(Storage:="_TiempoDias", DbType:="Decimal(23,8)")> _
    Public Property TiempoDias() As System.Nullable(Of Decimal)
        Get
            Return Me._TiempoDias
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TiempoDias.Equals(value) = False) Then
                Me._TiempoDias = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaInicioCosteNave", DbType:="DateTime")> _
    Public Property FechaInicioCosteNave() As System.Nullable(Of Date)
        Get
            Return Me._FechaInicioCosteNave
        End Get
        Set(ByVal value As System.Nullable(Of Date))
            If (Me._FechaInicioCosteNave.Equals(value) = False) Then
                Me._FechaInicioCosteNave = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaFinCosteNave", DbType:="DateTime")> _
    Public Property FechaFinCosteNave() As System.Nullable(Of Date)
        Get
            Return Me._FechaFinCosteNave
        End Get
        Set(ByVal value As System.Nullable(Of Date))
            If (Me._FechaFinCosteNave.Equals(value) = False) Then
                Me._FechaFinCosteNave = value
            End If
        End Set
    End Property

    <Column(Storage:="_Tasa", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Tasa() As Decimal
        Get
            Return Me._Tasa
        End Get
        Set(ByVal value As Decimal)
            If ((Me._Tasa = value) _
               = False) Then
                Me._Tasa = value
            End If
        End Set
    End Property

    <Column(Storage:="_IdUdMedidaArticulo", DbType:="NVarChar(10)")> _
    Public Property IdUdMedidaArticulo() As String
        Get
            Return Me._IdUdMedidaArticulo
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IdUdMedidaArticulo, value) = False) Then
                Me._IdUdMedidaArticulo = value
            End If
        End Set
    End Property

    <Column(Storage:="_PorCantidad", DbType:="Bit NOT NULL")> _
    Public Property PorCantidad() As Boolean
        Get
            Return Me._PorCantidad
        End Get
        Set(ByVal value As Boolean)
            If ((Me._PorCantidad = value) _
               = False) Then
                Me._PorCantidad = value
            End If
        End Set
    End Property

    <Column(Storage:="_IdExpCosteNave", AutoSync:=AutoSync.Always, DbType:="Int NOT NULL IDENTITY", IsDbGenerated:=True)> _
    Public Property IdExpCosteNave() As Integer
        Get
            Return Me._IdExpCosteNave
        End Get
        Set(ByVal value As Integer)
            If ((Me._IdExpCosteNave = value) _
               = False) Then
                Me._IdExpCosteNave = value
            End If
        End Set
    End Property

    <Column(Storage:="_CantidadCoste", DbType:="Decimal(23,8)")> _
    Public Property CantidadCoste() As System.Nullable(Of Decimal)
        Get
            Return Me._CantidadCoste
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CantidadCoste.Equals(value) = False) Then
                Me._CantidadCoste = value
            End If
        End Set
    End Property

    <Column(Storage:="_CantidadCosteLitros", DbType:="Decimal(23,8)")> _
    Public Property CantidadCosteLitros() As System.Nullable(Of Decimal)
        Get
            Return Me._CantidadCosteLitros
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CantidadCosteLitros.Equals(value) = False) Then
                Me._CantidadCosteLitros = value
            End If
        End Set
    End Property

    <Column(Storage:="_Importe", DbType:="Decimal(23,8)")> _
    Public Property Importe() As System.Nullable(Of Decimal)
        Get
            Return Me._Importe
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Importe.Equals(value) = False) Then
                Me._Importe = value
            End If
        End Set
    End Property

    <Column(Storage:="_IdUdMedidaCentro", DbType:="NVarChar(10)")> _
    Public Property IdUdMedidaCentro() As String
        Get
            Return Me._IdUdMedidaCentro
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IdUdMedidaCentro, value) = False) Then
                Me._IdUdMedidaCentro = value
            End If
        End Set
    End Property

    <Column(Storage:="_FactorTasa", DbType:="Decimal(23,8)")> _
    Public Property FactorTasa() As System.Nullable(Of Decimal)
        Get
            Return Me._FactorTasa
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._FactorTasa.Equals(value) = False) Then
                Me._FactorTasa = value
            End If
        End Set
    End Property

    <Column(Storage:="_AnioMesFechaInicio", DbType:="NVarChar(10)")> _
    Public Property AnioMesFechaInicio() As String
        Get
            Return Me._AnioMesFechaInicio
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._AnioMesFechaInicio, value) = False) Then
                Me._AnioMesFechaInicio = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteUnit.Equals(value) = False) Then
                Me._CosteUnit = value
            End If
        End Set
    End Property

End Class

<Serializable()> _
Public Class VinoCosteTasas

    Private _OrigenTasa As String

    Private _IDVino As System.Guid

    Private _IdTasa As String

    Private _IDCentro As String

    Private _NOperacion As String

    Private _TasaFija As System.Nullable(Of Decimal)

    Private _TasaVariable As System.Nullable(Of Decimal)

    Private _TasaDirecta As System.Nullable(Of Decimal)

    Private _TasaInDirecta As System.Nullable(Of Decimal)

    Private _TasaFiscal As System.Nullable(Of Decimal)

    Private _TasaTotal As System.Nullable(Of Decimal)

    Private _TasaFijaUnit As System.Nullable(Of Decimal)

    Private _TasaVariableUnit As System.Nullable(Of Decimal)

    Private _TasaDirectaUnit As System.Nullable(Of Decimal)

    Private _TasaInDirectaUnit As System.Nullable(Of Decimal)

    Private _TasaFiscalUnit As System.Nullable(Of Decimal)

    Private _TasaTotalUnit As System.Nullable(Of Decimal)

    Public Sub New()
        MyBase.New()
    End Sub

    <Column(Storage:="_OrigenTasa", DbType:="VarChar(9) NOT NULL", CanBeNull:=False)> _
    Public Property OrigenTasa() As String
        Get
            Return Me._OrigenTasa
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._OrigenTasa, value) = False) Then
                Me._OrigenTasa = value
            End If
        End Set
    End Property

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

    <Column(Storage:="_IdTasa", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IdTasa() As String
        Get
            Return Me._IdTasa
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IdTasa, value) = False) Then
                Me._IdTasa = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDCentro", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDCentro() As String
        Get
            Return Me._IDCentro
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDCentro, value) = False) Then
                Me._IDCentro = value
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

    <Column(Storage:="_TasaFija", DbType:="Decimal(23,8)")> _
    Public Property TasaFija() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaFija
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaFija.Equals(value) = False) Then
                Me._TasaFija = value
            End If
        End Set
    End Property

    <Column(Storage:="_TasaVariable", DbType:="Decimal(23,8)")> _
    Public Property TasaVariable() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaVariable
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaVariable.Equals(value) = False) Then
                Me._TasaVariable = value
            End If
        End Set
    End Property

    <Column(Storage:="_TasaDirecta", DbType:="Decimal(23,8)")> _
    Public Property TasaDirecta() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaDirecta
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaDirecta.Equals(value) = False) Then
                Me._TasaDirecta = value
            End If
        End Set
    End Property

    <Column(Storage:="_TasaInDirecta", DbType:="Decimal(23,8)")> _
    Public Property TasaInDirecta() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaInDirecta
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaInDirecta.Equals(value) = False) Then
                Me._TasaInDirecta = value
            End If
        End Set
    End Property

    <Column(Storage:="_TasaFiscal", DbType:="Decimal(23,8)")> _
    Public Property TasaFiscal() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaFiscal
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaFiscal.Equals(value) = False) Then
                Me._TasaFiscal = value
            End If
        End Set
    End Property

    <Column(Storage:="_TasaTotal", DbType:="Decimal(23,8)")> _
    Public Property TasaTotal() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaTotal
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaTotal.Equals(value) = False) Then
                Me._TasaTotal = value
            End If
        End Set
    End Property

    <Column(Storage:="_TasaFijaUnit", DbType:="Decimal(23,8)")> _
    Public Property TasaFijaUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaFijaUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaFijaUnit.Equals(value) = False) Then
                Me._TasaFijaUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_TasaVariableUnit", DbType:="Decimal(23,8)")> _
    Public Property TasaVariableUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaVariableUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaVariableUnit.Equals(value) = False) Then
                Me._TasaVariableUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_TasaDirectaUnit", DbType:="Decimal(23,8)")> _
    Public Property TasaDirectaUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaDirectaUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaDirectaUnit.Equals(value) = False) Then
                Me._TasaDirectaUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_TasaInDirectaUnit", DbType:="Decimal(23,8)")> _
    Public Property TasaInDirectaUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaInDirectaUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaInDirectaUnit.Equals(value) = False) Then
                Me._TasaInDirectaUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_TasaFiscalUnit", DbType:="Decimal(23,8)")> _
    Public Property TasaFiscalUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaFiscalUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaFiscalUnit.Equals(value) = False) Then
                Me._TasaFiscalUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_TasaTotalUnit", DbType:="Decimal(23,8)")> _
    Public Property TasaTotalUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._TasaTotalUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._TasaTotalUnit.Equals(value) = False) Then
                Me._TasaTotalUnit = value
            End If
        End Set
    End Property
End Class

<Serializable()> _
Public Class VinoCosteVarios

    Private _IDVino As System.Nullable(Of System.Guid)

    Private _IDDeposito As String

    Private _LoteVino As String

    Private _IDAlmacenVino As String

    Private _FechaVino As Date

    Private _IDArticuloVino As String

    Private _NOperacion As String

    Private _Fecha As System.Nullable(Of Date)

    Private _IdVarios As String

    Private _DescVarios As String

    Private _Cantidad As Decimal

    Private _Tasa As Decimal

    Private _TipoCosteFV As Integer

    Private _TipoCosteDI As Integer

    Private _Fiscal As Boolean

    Private _PorcentajeAcumulado As Decimal

    Private _CosteFijo As System.Nullable(Of Decimal)

    Private _CosteVariable As System.Nullable(Of Decimal)

    Private _CosteDirecto As System.Nullable(Of Decimal)

    Private _CosteInDirecto As System.Nullable(Of Decimal)

    Private _CosteFiscal As System.Nullable(Of Decimal)

    Private _CosteTotal As System.Nullable(Of Decimal)

    Private _CosteFijoUnit As System.Nullable(Of Decimal)

    Private _CosteVariableUnit As System.Nullable(Of Decimal)

    Private _CosteDirectoUnit As System.Nullable(Of Decimal)

    Private _CosteInDirectoUnit As System.Nullable(Of Decimal)

    Private _CosteFiscalUnit As System.Nullable(Of Decimal)

    Private _CosteUnit As System.Nullable(Of Decimal)

    Public Sub New()
        MyBase.New()
    End Sub

    <Column(Storage:="_IDVino", DbType:="UniqueIdentifier")> _
    Public Property IDVino() As System.Nullable(Of System.Guid)
        Get
            Return Me._IDVino
        End Get
        Set(ByVal value As System.Nullable(Of System.Guid))
            If (Me._IDVino.Equals(value) = False) Then
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

    <Column(Storage:="_LoteVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property LoteVino() As String
        Get
            Return Me._LoteVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._LoteVino, value) = False) Then
                Me._LoteVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlmacenVino", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAlmacenVino() As String
        Get
            Return Me._IDAlmacenVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacenVino, value) = False) Then
                Me._IDAlmacenVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaVino", DbType:="DateTime NOT NULL")> _
    Public Property FechaVino() As Date
        Get
            Return Me._FechaVino
        End Get
        Set(ByVal value As Date)
            If ((Me._FechaVino = value) _
               = False) Then
                Me._FechaVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticuloVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticuloVino() As String
        Get
            Return Me._IDArticuloVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticuloVino, value) = False) Then
                Me._IDArticuloVino = value
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

    <Column(Storage:="_Fecha", DbType:="DateTime")> _
    Public Property Fecha() As System.Nullable(Of Date)
        Get
            Return Me._Fecha
        End Get
        Set(ByVal value As System.Nullable(Of Date))
            If (Me._Fecha.Equals(value) = False) Then
                Me._Fecha = value
            End If
        End Set
    End Property

    <Column(Storage:="_IdVarios", DbType:="NVarChar(10)")> _
    Public Property IdVarios() As String
        Get
            Return Me._IdVarios
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IdVarios, value) = False) Then
                Me._IdVarios = value
            End If
        End Set
    End Property

    <Column(Storage:="_DescVarios", DbType:="NVarChar(100)")> _
    Public Property DescVarios() As String
        Get
            Return Me._DescVarios
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._DescVarios, value) = False) Then
                Me._DescVarios = value
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

    <Column(Storage:="_Tasa", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property Tasa() As Decimal
        Get
            Return Me._Tasa
        End Get
        Set(ByVal value As Decimal)
            If ((Me._Tasa = value) _
               = False) Then
                Me._Tasa = value
            End If
        End Set
    End Property

    <Column(Storage:="_TipoCosteFV", DbType:="Int NOT NULL")> _
    Public Property TipoCosteFV() As Integer
        Get
            Return Me._TipoCosteFV
        End Get
        Set(ByVal value As Integer)
            If ((Me._TipoCosteFV = value) _
               = False) Then
                Me._TipoCosteFV = value
            End If
        End Set
    End Property

    <Column(Storage:="_TipoCosteDI", DbType:="Int NOT NULL")> _
    Public Property TipoCosteDI() As Integer
        Get
            Return Me._TipoCosteDI
        End Get
        Set(ByVal value As Integer)
            If ((Me._TipoCosteDI = value) _
               = False) Then
                Me._TipoCosteDI = value
            End If
        End Set
    End Property

    <Column(Storage:="_Fiscal", DbType:="Bit NOT NULL")> _
    Public Property Fiscal() As Boolean
        Get
            Return Me._Fiscal
        End Get
        Set(ByVal value As Boolean)
            If ((Me._Fiscal = value) _
               = False) Then
                Me._Fiscal = value
            End If
        End Set
    End Property

    <Column(Storage:="_PorcentajeAcumulado", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property PorcentajeAcumulado() As Decimal
        Get
            Return Me._PorcentajeAcumulado
        End Get
        Set(ByVal value As Decimal)
            If ((Me._PorcentajeAcumulado = value) _
               = False) Then
                Me._PorcentajeAcumulado = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteFijo", DbType:="Decimal(23,8)")> _
    Public Property CosteFijo() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteFijo
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteFijo.Equals(value) = False) Then
                Me._CosteFijo = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteVariable", DbType:="Decimal(23,8)")> _
    Public Property CosteVariable() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteVariable
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteVariable.Equals(value) = False) Then
                Me._CosteVariable = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteDirecto", DbType:="Decimal(23,8)")> _
    Public Property CosteDirecto() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteDirecto
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteDirecto.Equals(value) = False) Then
                Me._CosteDirecto = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteInDirecto", DbType:="Decimal(23,8)")> _
    Public Property CosteInDirecto() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteInDirecto
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteInDirecto.Equals(value) = False) Then
                Me._CosteInDirecto = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteFiscal", DbType:="Decimal(23,8)")> _
    Public Property CosteFiscal() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteFiscal
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteFiscal.Equals(value) = False) Then
                Me._CosteFiscal = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteTotal", DbType:="Decimal(23,8)")> _
    Public Property CosteTotal() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteTotal
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteTotal.Equals(value) = False) Then
                Me._CosteTotal = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteFijoUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteFijoUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteFijoUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteFijoUnit.Equals(value) = False) Then
                Me._CosteFijoUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteVariableUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteVariableUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteVariableUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteVariableUnit.Equals(value) = False) Then
                Me._CosteVariableUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteDirectoUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteDirectoUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteDirectoUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteDirectoUnit.Equals(value) = False) Then
                Me._CosteDirectoUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteInDirectoUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteInDirectoUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteInDirectoUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteInDirectoUnit.Equals(value) = False) Then
                Me._CosteInDirectoUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteFiscalUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteFiscalUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteFiscalUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteFiscalUnit.Equals(value) = False) Then
                Me._CosteFiscalUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteUnit.Equals(value) = False) Then
                Me._CosteUnit = value
            End If
        End Set
    End Property
End Class

<Serializable()> _
Public Class VinoCosteInicial

    Private _idVino As System.Guid

    Private _IDDeposito As String

    Private _LoteVino As String

    Private _IDAlmacenVino As String

    Private _FechaVino As Date

    Private _IDArticuloVino As String

    Private _FechaCosteUnitarioInicial As System.Nullable(Of Date)

    Private _CosteUnitarioInicialA As Decimal

    Private _Cantidad As Decimal

    Private _PorcentajeAcumulado As Decimal

    Private _Coste As System.Nullable(Of Decimal)

    Private _CosteUnit As System.Nullable(Of Decimal)

    Public Sub New()
        MyBase.New()
    End Sub

    <Column(Storage:="_idVino", DbType:="UniqueIdentifier NOT NULL")> _
    Public Property idVino() As System.Guid
        Get
            Return Me._idVino
        End Get
        Set(ByVal value As System.Guid)
            If ((Me._idVino = value) _
               = False) Then
                Me._idVino = value
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

    <Column(Storage:="_LoteVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property LoteVino() As String
        Get
            Return Me._LoteVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._LoteVino, value) = False) Then
                Me._LoteVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDAlmacenVino", DbType:="NVarChar(10) NOT NULL", CanBeNull:=False)> _
    Public Property IDAlmacenVino() As String
        Get
            Return Me._IDAlmacenVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDAlmacenVino, value) = False) Then
                Me._IDAlmacenVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_FechaVino", DbType:="DateTime NOT NULL")> _
    Public Property FechaVino() As Date
        Get
            Return Me._FechaVino
        End Get
        Set(ByVal value As Date)
            If ((Me._FechaVino = value) _
               = False) Then
                Me._FechaVino = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDArticuloVino", DbType:="NVarChar(25) NOT NULL", CanBeNull:=False)> _
    Public Property IDArticuloVino() As String
        Get
            Return Me._IDArticuloVino
        End Get
        Set(ByVal value As String)
            If (String.Equals(Me._IDArticuloVino, value) = False) Then
                Me._IDArticuloVino = value
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

    <Column(Storage:="_PorcentajeAcumulado", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property PorcentajeAcumulado() As Decimal
        Get
            Return Me._PorcentajeAcumulado
        End Get
        Set(ByVal value As Decimal)
            If ((Me._PorcentajeAcumulado = value) _
               = False) Then
                Me._PorcentajeAcumulado = value
            End If
        End Set
    End Property

    <Column(Storage:="_Coste", DbType:="Decimal(23,8)")> _
    Public Property Coste() As System.Nullable(Of Decimal)
        Get
            Return Me._Coste
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._Coste.Equals(value) = False) Then
                Me._Coste = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUnit", DbType:="Decimal(23,8)")> _
    Public Property CosteUnit() As System.Nullable(Of Decimal)
        Get
            Return Me._CosteUnit
        End Get
        Set(ByVal value As System.Nullable(Of Decimal))
            If (Me._CosteUnit.Equals(value) = False) Then
                Me._CosteUnit = value
            End If
        End Set
    End Property
End Class

<Serializable()> _
 Public Class VinoCosteTotal

    Private _id As Integer

    Private _IDVinoInicial As System.Guid

    Private _QTot As Decimal

    Private _CosteTotalAcum As Decimal

    Private _CosteMaterialesAcum As Decimal

    Private _CosteManoObraAcum As Decimal

    Private _CosteCentroAcum As Decimal

    Private _CosteEstanciaNaveAcum As Decimal

    Private _CosteVariosAcum As Decimal

    Private _CosteComprasAcum As Decimal

    Private _CosteVendimiaAcum As Decimal

    Private _CosteMaterialesConMermaAcum As Decimal

    Private _CosteMaterialesSinMermaAcum As Decimal

    Private _CosteUvaAcum As Decimal

    Private _CosteInicialAcum As Decimal

    Private _CosteTotalUnit As Decimal

    Private _CosteMaterialesUnit As Decimal

    Private _CosteManoObraUnit As Decimal

    Private _CosteCentroUnit As Decimal

    Private _CosteEstanciaNaveUnit As Decimal

    Private _CosteVariosUnit As Decimal

    Private _CosteComprasUnit As Decimal

    Private _CosteVendimiaUnit As Decimal

    Private _CosteMaterialesConMermaUnit As Decimal

    Private _CosteMaterialesSinMermaUnit As Decimal

    Private _CosteVendimiaConMermaUnit As Decimal

    Private _CosteVendimiaSinMermaUnit As Decimal

    Private _CosteUvaUnit As Decimal

    Private _CosteInicialUnit As Decimal

    Public Sub New()
        MyBase.New()
    End Sub

    '<Column(Storage:="_id", AutoSync:=AutoSync.OnInsert, DbType:="Int NOT NULL IDENTITY", IsPrimaryKey:=True)> _
    <Column(Storage:="_id", DbType:="Int NOT NULL IDENTITY", IsPrimaryKey:=True)> _
    Public Property id() As Integer
        Get
            Return Me._id
        End Get
        Set(ByVal value As Integer)
            If ((Me._id = value) = False) Then
                Me._id = value
            End If
        End Set
    End Property

    <Column(Storage:="_IDVinoInicial", DbType:="UniqueIdentifier NOT NULL")> _
    Public Property IDVinoInicial() As System.Guid
        Get
            Return Me._IDVinoInicial
        End Get
        Set(ByVal value As System.Guid)
            If ((Me._IDVinoInicial = value) = False) Then
                Me._IDVinoInicial = value
            End If
        End Set
    End Property

    <Column(Storage:="_QTot", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property QTot() As Decimal
        Get
            Return Me._QTot
        End Get
        Set(ByVal value As Decimal)
            If ((Me._QTot = value) = False) Then
                Me._QTot = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteTotalAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteTotalAcum() As Decimal
        Get
            Return Me._CosteTotalAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteTotalAcum = value) = False) Then
                Me._CosteTotalAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteMaterialesAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteMaterialesAcum() As Decimal
        Get
            Return Me._CosteMaterialesAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteMaterialesAcum = value) = False) Then
                Me._CosteMaterialesAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteManoObraAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteManoObraAcum() As Decimal
        Get
            Return Me._CosteManoObraAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteManoObraAcum = value) = False) Then
                Me._CosteManoObraAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteCentroAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteCentroAcum() As Decimal
        Get
            Return Me._CosteCentroAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteCentroAcum = value) = False) Then
                Me._CosteCentroAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteEstanciaNaveAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteEstanciaNaveAcum() As Decimal
        Get
            Return Me._CosteEstanciaNaveAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteEstanciaNaveAcum = value) = False) Then
                Me._CosteEstanciaNaveAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteVariosAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteVariosAcum() As Decimal
        Get
            Return Me._CosteVariosAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteVariosAcum = value) = False) Then
                Me._CosteVariosAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteComprasAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteComprasAcum() As Decimal
        Get
            Return Me._CosteComprasAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteComprasAcum = value) = False) Then
                Me._CosteComprasAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteVendimiaAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteVendimiaAcum() As Decimal
        Get
            Return Me._CosteVendimiaAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteVendimiaAcum = value) = False) Then
                Me._CosteVendimiaAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteMaterialesConMermaAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteMaterialesConMermaAcum() As Decimal
        Get
            Return Me._CosteMaterialesConMermaAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteMaterialesConMermaAcum = value) = False) Then
                Me._CosteMaterialesConMermaAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteMaterialesSinMermaAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteMaterialesSinMermaAcum() As Decimal
        Get
            Return Me._CosteMaterialesSinMermaAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteMaterialesSinMermaAcum = value) = False) Then
                Me._CosteMaterialesSinMermaAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUvaAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteUvaAcum() As Decimal
        Get
            Return Me._CosteUvaAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteUvaAcum = value) = False) Then
                Me._CosteUvaAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteInicialAcum", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteInicialAcum() As Decimal
        Get
            Return Me._CosteInicialAcum
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteInicialAcum = value) = False) Then
                Me._CosteInicialAcum = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteTotalUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteTotalUnit() As Decimal
        Get
            Return Me._CosteTotalUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteTotalUnit = value) = False) Then
                Me._CosteTotalUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteMaterialesUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteMaterialesUnit() As Decimal
        Get
            Return Me._CosteMaterialesUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteMaterialesUnit = value) = False) Then
                Me._CosteMaterialesUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteManoObraUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteManoObraUnit() As Decimal
        Get
            Return Me._CosteManoObraUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteManoObraUnit = value) = False) Then
                Me._CosteManoObraUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteCentroUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteCentroUnit() As Decimal
        Get
            Return Me._CosteCentroUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteCentroUnit = value) = False) Then
                Me._CosteCentroUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteEstanciaNaveUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteEstanciaNaveUnit() As Decimal
        Get
            Return Me._CosteEstanciaNaveUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteEstanciaNaveUnit = value) = False) Then
                Me._CosteEstanciaNaveUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteVariosUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteVariosUnit() As Decimal
        Get
            Return Me._CosteVariosUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteVariosUnit = value) = False) Then
                Me._CosteVariosUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteComprasUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteComprasUnit() As Decimal
        Get
            Return Me._CosteComprasUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteComprasUnit = value) = False) Then
                Me._CosteComprasUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteVendimiaUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteVendimiaUnit() As Decimal
        Get
            Return Me._CosteVendimiaUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteVendimiaUnit = value) = False) Then
                Me._CosteVendimiaUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteMaterialesConMermaUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteMaterialesConMermaUnit() As Decimal
        Get
            Return Me._CosteMaterialesConMermaUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteMaterialesConMermaUnit = value) = False) Then
                Me._CosteMaterialesConMermaUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteMaterialesSinMermaUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteMaterialesSinMermaUnit() As Decimal
        Get
            Return Me._CosteMaterialesSinMermaUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteMaterialesSinMermaUnit = value) = False) Then
                Me._CosteMaterialesSinMermaUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteVendimiaConMermaUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteVendimiaConMermaUnit() As Decimal
        Get
            Return Me._CosteVendimiaConMermaUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteVendimiaConMermaUnit = value) = False) Then
                Me._CosteVendimiaConMermaUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteVendimiaSinMermaUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteVendimiaSinMermaUnit() As Decimal
        Get
            Return Me._CosteVendimiaSinMermaUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteVendimiaSinMermaUnit = value) = False) Then
                Me._CosteVendimiaSinMermaUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteUvaUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteUvaUnit() As Decimal
        Get
            Return Me._CosteUvaUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteUvaUnit = value) = False) Then
                Me._CosteUvaUnit = value
            End If
        End Set
    End Property

    <Column(Storage:="_CosteInicialUnit", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property CosteInicialUnit() As Decimal
        Get
            Return Me._CosteInicialUnit
        End Get
        Set(ByVal value As Decimal)
            If ((Me._CosteInicialUnit = value) _
               = False) Then
                Me._CosteInicialUnit = value
            End If
        End Set
    End Property
End Class



<Serializable()> _
Public Class VinoArticulo

    Private _idVino As System.Guid

    Private _IDDeposito As String

    Private _Lote As String

    Private _IDAlmacen As String

    Private _Fecha As Date

    Private _QActual As Decimal

    Public Sub New()
        MyBase.New()
    End Sub

    <Column(Storage:="_idVino", DbType:="UniqueIdentifier NOT NULL")> _
    Public Property idVino() As System.Guid
        Get
            Return Me._idVino
        End Get
        Set(ByVal value As System.Guid)
            If ((Me._idVino = value) _
               = False) Then
                Me._idVino = value
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

   
    <Column(Storage:="_QActual", DbType:="Decimal(23,8) NOT NULL")> _
    Public Property QActual() As Decimal
        Get
            Return Me._QActual
        End Get
        Set(ByVal value As Decimal)
            If ((Me._QActual = value) _
               = False) Then
                Me._QActual = value
            End If
        End Set
    End Property

End Class