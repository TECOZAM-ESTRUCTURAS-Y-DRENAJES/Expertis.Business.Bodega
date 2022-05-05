Public Class BdgEntradaFinca

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaFinca"

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
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarFinca)
    End Sub

    <Task()> Public Shared Sub ValidarFinca(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtFinca As DataTable = New BdgFinca().SelOnPrimaryKey(data(_EF.IdFinca))
        If DtFinca Is Nothing Or DtFinca.Rows.Count = 0 Then ApplicationService.GenerateError("La finca | no existe.", data(_EF.IdFinca))
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StCrearEntradaFinca
        Public IDEntrada As Integer
        Public IDFinca As Guid
        Public Neto As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDEntrada As Integer, ByVal IDFinca As Guid, ByVal Neto As Double)
            Me.IDEntrada = IDEntrada
            Me.IDFinca = IDFinca
            Me.Neto = Neto
        End Sub
    End Class

    <Task()> Public Shared Function CrearEntradaFinca(ByVal data As StCrearEntradaFinca, ByVal services As ServiceProvider) As DataTable
        Dim dtEF As DataTable = New BdgEntradaFinca().AddNew
        If Length(data.IDFinca.ToString) > 0 AndAlso data.IDEntrada > 0 Then
            Dim rwEF As DataRow = dtEF.NewRow
            rwEF(_EF.IdEntrada) = data.IDEntrada
            rwEF(_EF.IdFinca) = data.IDFinca
            rwEF(_EF.Neto) = data.Neto
            rwEF(_EF.Porcentaje) = 100
            dtEF.Rows.Add(rwEF)
            Return dtEF
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _EF
    Public Const IdEntrada As String = "IdEntrada"
    Public Const IdFinca As String = "IdFinca"
    Public Const Porcentaje As String = "Porcentaje"
    Public Const Neto As String = "Neto"
End Class