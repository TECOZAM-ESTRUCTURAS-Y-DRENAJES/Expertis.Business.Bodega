<Serializable()> _
Public Class DataPropuestaDatosDepositos
    Public IDTipoOperacion As String
    Public Fecha As Date

    Public Depositos As List(Of DatosDeposito)

    Public Sub New(ByVal IDTipoOperacion As String, ByVal Fecha As Date, ByVal Depositos As List(Of DatosDeposito))
        Me.IDTipoOperacion = IDTipoOperacion
        Me.Fecha = Fecha
        Me.Depositos = Depositos
    End Sub
End Class

<Serializable()> _
Public Class DatosDeposito
    Public IDDeposito As String
    Public IDVino As Guid
   
    Public dtRegistro As DataTable

    Public Destino As Boolean '//True - Destino, False - Origen

    Public Sub New(ByVal IDDeposito As String, ByVal IDVino As Guid, ByVal Destino As Boolean)
        Me.IDDeposito = IDDeposito
        Me.IDVino = IDVino
        Me.Destino = Destino
    End Sub
  
    Public Sub New(ByVal IDDeposito As String, ByVal dtRegistro As DataTable, ByVal Destino As Boolean)
        Me.IDDeposito = IDDeposito
        Me.IDVino = Guid.Empty
        Me.Destino = Destino
        Me.dtRegistro = dtRegistro
    End Sub

End Class


<Serializable()> _
Public Class DataPrcPropuestaOperacion
    Public TipoOperacion As enumBdgOrigenOperacion
    Public Origen As OrigenOperacion

    Public IDOrigen As List(Of String)
    Public OrigenesDepositos As List(Of DataPropuestaDatosDepositos)
    'Public OrigenesOFs As Dictionary(Of Integer, Double)
    Public IDTipoOperacion As String
    Public Fecha As Date

    Public Multiple As Boolean

    Public AutoCalcularVino As Boolean
    Public GuardarPropuesta As Boolean

    Protected Sub New()

    End Sub
    '//Desde la confirmación de Operación Planificada
    Public Sub New(ByVal TipoOperacion As enumBdgOrigenOperacion, ByVal Origen As OrigenOperacion, ByVal IDOrigen As List(Of String), Optional ByVal AutoCalcularVino As Boolean = True, Optional ByVal GuardarPropuesta As Boolean = False)
        Me.TipoOperacion = TipoOperacion
        Me.Origen = Origen
        Me.IDOrigen = IDOrigen
        Me.AutoCalcularVino = AutoCalcularVino
        Me.GuardarPropuesta = GuardarPropuesta
        Select Case Origen
            Case OrigenOperacion.OperacionPlanificada
                Me.Fecha = cnMinDate
            Case Else
                Me.Fecha = Now
        End Select
    End Sub

    '//Desde depósitos
    Public Sub New(ByVal TipoOperacion As enumBdgOrigenOperacion, ByVal Origen As OrigenOperacion, ByVal OrigenesDepositos As List(Of DataPropuestaDatosDepositos), Optional ByVal AutoCalcularVino As Boolean = True, Optional ByVal GuardarPropuesta As Boolean = False)
        Me.TipoOperacion = TipoOperacion
        Me.Origen = Origen
        Me.OrigenesDepositos = OrigenesDepositos
        Me.AutoCalcularVino = AutoCalcularVino
        Me.GuardarPropuesta = GuardarPropuesta
    End Sub

End Class

<Serializable()> _
Public Class DataPrcPropuestaOperacionExpediciones
    Inherits DataPrcPropuestaOperacion

    Public OrigenesExp As List(Of DataArtCompatiblesExp)

    Public Sub New(ByVal TipoOperacion As enumBdgOrigenOperacion, ByVal Origen As OrigenOperacion, ByVal OrigenesExp As List(Of DataArtCompatiblesExp), Optional ByVal AutoCalcularVino As Boolean = True, Optional ByVal GuardarPropuesta As Boolean = False)
        Me.TipoOperacion = TipoOperacion
        Me.Origen = Origen
        Me.OrigenesExp = OrigenesExp
        Me.AutoCalcularVino = AutoCalcularVino
        Me.GuardarPropuesta = GuardarPropuesta
        ' Me.IDTipoOperacion = IDTipoOperacion
    End Sub

End Class

<Serializable()> _
Public Class DataPrcPropuestaOperacionOFs
    Inherits DataPrcPropuestaOperacion

    Public OrigenesOFs As Dictionary(Of Integer, Double)
    Public TiposOperacionesOrdenes As Dictionary(Of Integer, String)

    Public Sub New(ByVal TipoOperacion As enumBdgOrigenOperacion, ByVal Origen As OrigenOperacion, ByVal OrigenesOFs As Dictionary(Of Integer, Double), ByVal IDTipoOperacion As String, ByVal Fecha As Date, Optional ByVal AutoCalcularVino As Boolean = True, Optional ByVal GuardarPropuesta As Boolean = False)
        Me.TipoOperacion = TipoOperacion
        Me.Origen = Origen
        Me.OrigenesOFs = OrigenesOFs
        Me.AutoCalcularVino = AutoCalcularVino
        Me.GuardarPropuesta = GuardarPropuesta
        Me.IDTipoOperacion = IDTipoOperacion
        Me.Fecha = Fecha
    End Sub

End Class




