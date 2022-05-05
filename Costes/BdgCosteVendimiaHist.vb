Public Class BdgCosteVendimiaHist

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCosteVendimiaHist"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarEntrada)
    End Sub

    <Task()> Public Shared Sub ValidarEntrada(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim FilEntrada As New Filter
        FilEntrada.Add("Vendimia", FilterOperator.Equal, data("Vendimia"))
        FilEntrada.Add("IdDeposito", FilterOperator.Equal, data("IdDeposito"))
        Dim Dt As DataTable = New BdgCosteVendimiaHist().Filter(FilEntrada)
        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            ApplicationService.GenerateError("Ya existe esa entrada en el historico de coste vendimia.")
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StCrearCosteVendimiaHistInd
        Public Vendimia As Integer
        Public IDDeposito As String
        Public IDVinoMaterial As Guid
        Public IDVino As Guid
        Public IDArticuloVendimia As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal Vendimia As Integer, ByVal IDDeposito As String, ByVal IDVinoMaterial As Guid, ByVal IDVino As Guid, ByVal IDArticuloVendimia As String)
            Me.Vendimia = Vendimia
            Me.IDDeposito = IDDeposito
            Me.IDVinoMaterial = IDVinoMaterial
            Me.IDVino = IDVino
            Me.IDArticuloVendimia = IDArticuloVendimia
        End Sub
    End Class

    <Task()> Public Shared Sub CrearCosteVendimiaHistInd(ByVal data As StCrearCosteVendimiaHistInd, ByVal services As ServiceProvider)
        Dim ff As New Filter
        ff.Add("Vendimia", FilterOperator.Equal, data.Vendimia)
        ff.Add("IDDeposito", FilterOperator.Equal, data.IDDeposito)
        ff.Add("IDVino", FilterOperator.Equal, data.IDVino)
        ff.Add("IDArticulo", FilterOperator.Equal, data.IDArticuloVendimia)

        Dim dtLista As DataTable = New BE.DataEngine().Filter("vNegBdgBorrarCosteVinoMaterial", ff, , "Vendimia")
        If Not dtLista Is Nothing AndAlso dtLista.Rows.Count > 0 Then
            ApplicationService.GenerateError("El Depósito | para la Vendimia | ya tiene imputado el Coste de Elaboración (|).", data.IDDeposito, data.Vendimia, data.IDArticuloVendimia)
        End If

        Dim dtCVH As DataTable = New BdgCosteVendimiaHist().AddNew
        Dim rwCVH As DataRow = dtCVH.NewRow
        rwCVH(_CVH.Vendimia) = data.Vendimia
        rwCVH(_CVH.IdDeposito) = data.IDDeposito
        rwCVH(_CVH.IdVinoMaterial) = data.IDVinoMaterial
        dtCVH.Rows.Add(rwCVH)
        BusinessHelper.UpdateTable(dtCVH)
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _CVH
    Public Const Vendimia As String = "Vendimia"
    Public Const IdDeposito As String = "IdDeposito"
    Public Const IdVinoMaterial As String = "IdVinoMaterial"
End Class