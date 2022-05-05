Public Class OperCab
    Public TipoOperacion As enumBdgOrigenOperacion
    Public Origen As OrigenOperacion

    Public IDTipoOperacion As String
    Public Fecha As Date

    Public Doc As DocumentoBdgOperacion

    Public Sub New(ByVal TipoOperacion As enumBdgOrigenOperacion, ByVal Origen As OrigenOperacion, ByVal IDTipoOperacion As String, ByVal Fecha As Date)
        Me.TipoOperacion = TipoOperacion
        Me.Origen = Origen
        Me.IDTipoOperacion = IDTipoOperacion
        Me.Fecha = Fecha
    End Sub

End Class

Public Class OperCabDepositos
    Inherits OperCab

    Public DepositosOp As List(Of DatosDeposito)
    Public Sub New(ByVal TipoOperacion As enumBdgOrigenOperacion, ByVal Origen As OrigenOperacion, ByVal DepositosOp As List(Of DatosDeposito), ByVal IDTipoOperacion As String, ByVal Fecha As Date)
        MyBase.New(TipoOperacion, Origen, IDTipoOperacion, Fecha)
        Me.DepositosOp = DepositosOp
    End Sub

End Class

Public Class OperCabPlanificadas
    Inherits OperCab

    Public IDOrigen As String
    Public Sub New(ByVal TipoOperacion As enumBdgOrigenOperacion, ByVal Origen As OrigenOperacion, ByVal IDOrigen As String, ByVal Fecha As Date)
        MyBase.New(TipoOperacion, Origen, String.Empty, Fecha)
        Me.IDOrigen = IDOrigen
    End Sub

End Class


Public Class OperCabOFs
    Inherits OperCab

    Public IDOrden As Integer
    Public NOrden As String
    Public CantidadOrden As Double

    Public Sub New(ByVal TipoOperacion As enumBdgOrigenOperacion, ByVal Origen As OrigenOperacion, ByVal IDOrden As Integer, ByVal NOrden As String, ByVal CantidadOrden As Double, ByVal IDTipoOperacion As String, ByVal Fecha As Date)
        MyBase.New(TipoOperacion, Origen, IDTipoOperacion, Fecha)
        Me.IDOrden = IDOrden
        Me.NOrden = NOrden
        Me.CantidadOrden = CantidadOrden
    End Sub

End Class


Public Class OperCabExp
    Inherits OperCab

    Public IDLineaPedido As Integer
    Public IDDeposito As String
    Public dtArtCompatibles As DataTable
  
    Public Sub New(ByVal TipoOperacion As enumBdgOrigenOperacion, ByVal Origen As OrigenOperacion, ByVal IDTipoOperacion As String, ByVal Fecha As Date, ByVal IDLineaPedido As Integer, ByVal IDDeposito As String, ByVal dtArtCompatibles As DataTable)
        MyBase.New(TipoOperacion, Origen, IDTipoOperacion, Fecha)
        Me.IDDeposito = IDDeposito
        Me.IDLineaPedido = IDLineaPedido
        Me.dtArtCompatibles = dtArtCompatibles
    End Sub

End Class