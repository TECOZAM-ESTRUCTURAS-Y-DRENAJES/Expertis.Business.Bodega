Public Class BdgVinoVarios

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVinoVarios"

#End Region

#Region "Eventos Entidad"

#Region "Tareas GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim OBrl As New BusinessRules
        OBrl.Add("IDVarios", AddressOf BdgGeneral.CambioVarios)
        Return OBrl
    End Function

#End Region


#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDVarios")) = 0 Then ApplicationService.GenerateError("El Varios es un dato obligatorio.")
    End Sub

#End Region



#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StCrearVinoVarios
        Public Fecha As Date
        Public NOperacion As String
        Public Data As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal Fecha As Date, ByVal NOperacion As String, ByVal Data As DataTable)
            Me.Fecha = Fecha
            Me.NOperacion = NOperacion
            Me.Data = Data
        End Sub
    End Class

    <Task()> Public Shared Function CrearVinoVarios(ByVal data As StCrearVinoVarios, ByVal services As ServiceProvider)
        If Not data.Data Is Nothing AndAlso data.Data.Rows.Count > 0 Then
            Dim dtVVR As DataTable = New BdgVinoVarios().AddNew
            Dim oVar As Varios = New Varios
            For Each oRw As DataRow In data.Data.Rows
                Dim rwVVR As DataRow = dtVVR.NewRow
                rwVVR(_VVR.IdVino) = oRw(_VVR.IdVino)
                rwVVR(_VVR.IdVarios) = oRw(_VVR.IdVarios)
                rwVVR(_VVR.DescVarios) = oRw(_VVR.DescVarios)
                rwVVR(_VVR.Fecha) = data.Fecha
                rwVVR(_VVR.NOperacion) = data.NOperacion
                rwVVR(_VVR.Cantidad) = oRw(_VVR.Cantidad)
                rwVVR(_VVR.Tasa) = oRw(_VVR.Tasa)

                Dim rwVar As DataRow = oVar.GetItemRow(oRw(_VVR.IdVarios))
                rwVVR(_VVR.TipoCosteFV) = rwVar(_VVR.TipoCosteFV)
                rwVVR(_VVR.TipoCosteDI) = rwVar(_VVR.TipoCosteDI)
                rwVVR(_VVR.Fiscal) = rwVar(_VVR.Fiscal)
                rwVVR(_VVR.IDOperacionVarios) = oRw(_VVR.IDOperacionVarios)
                dtVVR.Rows.Add(rwVVR)
            Next
            BusinessHelper.UpdateTable(dtVVR)
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _VVR
    Public Const IdVinoVarios As String = "IdVinoVarios"
    Public Const IdVino As String = "IdVino"
    Public Const IdVarios As String = "IdVarios"
    Public Const Fecha As String = "Fecha"
    Public Const NOperacion As String = "NOperacion"
    Public Const Tasa As String = "Tasa"
    Public Const Cantidad As String = "Cantidad"
    Public Const DescVarios As String = "DescVarios"
    Public Const TipoCosteFV As String = "TipoCosteFV"
    Public Const TipoCosteDI As String = "TipoCosteDI"
    Public Const Fiscal As String = "Fiscal"
    Public Const IDOperacionVarios As String = "IDOperacionVarios"
End Class