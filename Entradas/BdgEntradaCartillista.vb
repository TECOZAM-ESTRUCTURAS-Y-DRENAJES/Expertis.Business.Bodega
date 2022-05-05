Public Class BdgEntradaCartillista

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaCartillista"

#End Region

#Region "Eventos Entidad"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("Declarado", "Cantidad")
        oBRL.Add("Porcentaje", AddressOf BdgProcesoEntrada.CambioPorcentaje)
        oBRL.Add("Cantidad", AddressOf BdgProcesoEntrada.CambioCantidad)
        oBRL.Add("NCartilla", AddressOf CambioNCartilla)

        Return Obrl
    End Function

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavesSecundarias)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_EC.IDEntrada)) = 0 Then ApplicationService.GenerateError("El Nº de Entrada no es válido.")
        If Length(data(_EC.IDCartillista)) = 0 Then ApplicationService.GenerateError("El Cartillista es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClavesSecundarias(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            ProcessServer.ExecuteTask(Of Integer)(AddressOf BdgEntrada.ValidatePrimaryKey, data(_EC.IDEntrada), services)
            ProcessServer.ExecuteTask(Of String)(AddressOf BdgCartillista.ValidatePrimaryKey, data(_EC.IDCartillista), services)
        End If
    End Sub
    <Task()> Public Shared Sub CambioNCartilla(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        data.Current("IDCartillista") = System.DBNull.Value
        data.Current("IDProveedor") = System.DBNull.Value
        data.Current("DescCartillista") = System.DBNull.Value

        If data.Context.ContainsKey("Vendimia") Then
            If Length(data.Current("NCartilla")) > 0 Then
                Dim objBdgEntrada As New Solmicro.Expertis.Business.Bodega.BdgEntrada
                Dim StObtener As New Business.Bodega.BdgEntrada.StObtenerDatos(data.Current("NCartilla"), data.Context("Vendimia"))
                Dim dtNCartilla As DataTable = ProcessServer.ExecuteTask(Of Business.Bodega.BdgEntrada.StObtenerDatos, DataTable)(AddressOf Business.Bodega.BdgEntrada.ObtenerDatosConNCartilla, StObtener, services)
                If Not IsNothing(dtNCartilla) AndAlso dtNCartilla.Rows.Count > 0 Then
                    data.Current("IDCartillista") = dtNCartilla.Rows(0)("IDCartillista")
                    data.Current("IDProveedor") = dtNCartilla.Rows(0)("IDProveedor")
                    data.Current("DescCartillista") = dtNCartilla.Rows(0)("DescCartillista")
                End If
            End If
        End If

    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StCrearEntradaCartillista
        Public IDEntrada As Integer
        Public IDCartillista As String
        Public Vendimia As Integer
        Public TipoVariedad As BdgTipoVariedad
        Public Neto As Double
        Public Talon As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDEntrada As Integer, ByVal IDCartillista As String, ByVal Vendimia As Integer, ByVal TipoVariedad As BdgTipoVariedad, ByVal Neto As Double, ByVal Talon As Integer)
            Me.IDEntrada = IDEntrada
            Me.IDCartillista = IDCartillista
            Me.Vendimia = Vendimia
            Me.TipoVariedad = TipoVariedad
            Me.Neto = Neto
            Me.Talon = Talon
        End Sub
    End Class

    <Task()> Public Shared Function CrearEntradaCartillista(ByVal data As StCrearEntradaCartillista, ByVal services As ServiceProvider) As DataTable
        Dim dtEC As DataTable = New BdgEntradaCartillista().AddNew
        Dim rwEC As DataRow = dtEC.NewRow
        rwEC(_EC.IDEntrada) = data.IDEntrada
        rwEC(_EC.IDCartillista) = data.IDCartillista
        rwEC(_EC.Vendimia) = data.Vendimia
        rwEC(_EC.TipoVariedad) = data.TipoVariedad
        rwEC(_EC.Porcentaje) = 100
        rwEC(_EC.Declarado) = data.Neto
        rwEC(_EC.Talon) = data.Talon
        dtEC.Rows.Add(rwEC)
        Return dtEC
    End Function

#End Region

End Class

<Serializable()> _
Public Class _EC
    Public Const IDEntrada As String = "IDEntrada"
    Public Const IDCartillista As String = "IDCartillista"
    Public Const Declarado As String = "Declarado"
    Public Const Porcentaje As String = "Porcentaje"
    Public Const Talon As String = "Talon"
    Public Const Vendimia As String = "Vendimia"
    Public Const TipoVariedad As String = "TipoVariedad"
    Public Const IDVariedad As String = "IDVariedad"
End Class