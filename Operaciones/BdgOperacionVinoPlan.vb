Public Class BdgOperacionVinoPlan

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgOperacionVinoPlan"

#End Region

#Region "Eventos Entidad"

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim OBrl As New BusinessRules
        OBrl.Add("IDDeposito", AddressOf BdgGeneral.CambioIDDeposito)
        OBrl.Add("QDeposito", AddressOf BdgGeneral.CambioQDeposito)
        OBrl.Add("Cantidad", AddressOf BdgGeneral.CambioCantidad)
        OBrl.Add("Litros", AddressOf BdgGeneral.CambioLitros)
        OBrl.Add("IDArticulo", AddressOf BdgGeneral.CambioIDArticulo)
        OBrl.Add("IDOrden", AddressOf BdgGeneral.CambioIDOrden)
        Return OBrl
    End Function


#End Region

#End Region

#Region "Tareas Públicas"


    <Task()> Public Shared Function SelOnNOperacionPlan(ByVal NOperacionPlan As String, ByVal services As ServiceProvider) As DataTable
        Return New BdgOperacionVinoPlan().Filter(New StringFilterItem("NOperacionPlan", NOperacionPlan))
    End Function

#End Region

End Class

Public Class BdgOperacionVinoPlanInfo
    Inherits ClassEntityInfo

    Public IDLineaOperacionVinoPlan As Guid
    Public NOperacionPlan As String
    Public Destino As Boolean
    Public IDDeposito As String
    Public IDArticulo As String
    Public Lote As String
    Public Cantidad As Double
    Public Merma As Double
    Public QDeposito As Double
    Public Litros As Double
    Public IDEstructura As String
    Public IDBarrica As String
    Public IDEstadoVino As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Sub New(ByVal IDLineaOperacionVinoPlan As String)
        MyBase.New()
        Me.Fill(IDLineaOperacionVinoPlan)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dtOVP As DataTable = New BdgOperacionVinoPlan().Filter(New StringFilterItem("IDLineaOperacionVinoPlan", PrimaryKey(0)))
        If dtOVP.Rows.Count > 0 Then
            Me.Fill(dtOVP.Rows(0))
        Else
            ApplicationService.GenerateError("La operacion-vino | no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub
End Class

<Serializable()> _
Public Class stRepartirImputaciones
    Public dttLineasVinoOrigen As DataTable
    Public dttOrigen As DataTable
    Public dttDestino As DataTable
    Public Key As String
End Class
