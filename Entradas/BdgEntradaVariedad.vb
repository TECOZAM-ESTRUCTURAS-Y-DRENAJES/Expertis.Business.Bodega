Public Class BdgEntradaVariedad

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaVariedad"

#End Region

#Region "Eventos Entidad"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("Neto", "Cantidad")
        oBRL.Add("Porcentaje", AddressOf BdgProcesoEntrada.CambioPorcentaje)
        oBRL.Add("Cantidad", AddressOf BdgProcesoEntrada.CambioCantidad)
        Return oBRL
    End Function

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarVariedad)
    End Sub

    <Task()> Public Shared Sub ValidarVariedad(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtVariedad As DataTable = New BdgVariedad().SelOnPrimaryKey(data(_EV.IDVariedad))
        If dtVariedad Is Nothing OrElse dtVariedad.Rows.Count = 0 Then ApplicationService.GenerateError("La variedad | no existe.", data(_EV.IDVariedad))
    End Sub

#End Region

#Region "Funciones Públicos"

    <Serializable()> _
    Public Class StCrearEntradaVariedad
        Public IDEntrada As Integer
        Public IDVariedad As String
        Public Neto As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDEntrada As Integer, ByVal IDVariedad As String, ByVal Neto As Double)
            Me.IDEntrada = IDEntrada
            Me.IDVariedad = IDVariedad
            Me.Neto = Neto
        End Sub
    End Class

    <Task()> Public Shared Function CrearEntradaVariedad(ByVal data As StCrearEntradaVariedad, ByVal services As ServiceProvider) As DataTable
        Dim dtEV As DataTable = New BdgEntradaVariedad().AddNew
        If Length(data.IDVariedad) > 0 And data.IDEntrada > 0 Then
            Dim rwEV As DataRow = dtEV.NewRow
            rwEV(_EV.IDEntrada) = data.IDEntrada
            rwEV(_EV.IDVariedad) = data.IDVariedad
            rwEV(_EV.Neto) = data.Neto
            rwEV(_EV.Porcentaje) = 100
            dtEV.Rows.Add(rwEV)
            Return dtEV
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _EV
    Public Const IDEntrada As String = "IDEntrada"
    Public Const IDVariedad As String = "IDVariedad"
    Public Const Porcentaje As String = "Porcentaje"
    Public Const Neto As String = "Neto"
End Class