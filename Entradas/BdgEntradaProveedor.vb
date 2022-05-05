Public Class BdgEntradaProveedor

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaProveedor"

#End Region

#Region "Eventos Entidad"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("Porcentaje", AddressOf CambioPorcentaje)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioPorcentaje(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim DblPctj As Double = 0
        If Length(data.Value) > 0 Then
            DblPctj = data.Value
        Else : ApplicationService.GenerateError("El campo Porcentaje debe ser numérico.")
        End If
        data.Current(_EF.Porcentaje) = DblPctj
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarProveedor)
    End Sub

    <Task()> Public Shared Sub ValidarProveedor(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtProveedor As DataTable = New Proveedor().SelOnPrimaryKey(data(_EP.IdProveedor))
        If dtProveedor Is Nothing OrElse dtProveedor.Rows.Count = 0 Then ApplicationService.GenerateError("El proveedor | no existe.", data(_EP.IdProveedor))
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StCrearEntradaProveedor
        Public IDEntrada As Integer
        Public IDCartillista As String
        Public IDFinca As Guid

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDEntrada As Integer, ByVal IDCartillista As String, ByVal IDFinca As Guid)
            Me.IDEntrada = IDEntrada
            Me.IDCartillista = IDCartillista
            Me.IDFinca = IDFinca
        End Sub
    End Class

    <Task()> Public Shared Function CrearEntradaProveedor(ByVal data As StCrearEntradaProveedor, ByVal services As ServiceProvider) As DataTable
        Dim dtEP As DataTable = New BdgEntradaProveedor().AddNew
        Dim dtFP As DataTable = New BdgFincaProveedor().Filter(New GuidFilterItem("IdFinca", data.IDFinca))
        If dtFP.Rows.Count > 0 Then
            For Each rwFP As DataRow In dtFP.Select
                Dim rwEP As DataRow = dtEP.NewRow
                rwEP(_EP.IdEntrada) = data.IDEntrada
                rwEP(_EP.IdProveedor) = rwFP(_FP.IdProveedor)
                rwEP(_EF.Porcentaje) = rwFP(_FP.Porcentaje)
                dtEP.Rows.Add(rwEP)
            Next
        Else
            If Length(data.IDCartillista) > 0 AndAlso data.IDEntrada > 0 Then
                Dim dtCar As DataTable = New BdgCartillista().SelOnPrimaryKey(data.IDCartillista)
                Dim IdProveedor As String
                If dtCar.Rows.Count > 0 Then IdProveedor = dtCar.Rows(0)("IdProveedor")
                Dim rwEP As DataRow = dtEP.NewRow
                rwEP(_EP.IdEntrada) = data.IDEntrada
                rwEP(_EP.IdProveedor) = IdProveedor
                rwEP(_EF.Porcentaje) = 100
                dtEP.Rows.Add(rwEP)
            End If
        End If
        Return dtEP
    End Function

#End Region

End Class

<Serializable()> _
Public Class _EP
    Public Const IdEntrada As String = "IdEntrada"
    Public Const IdProveedor As String = "IdProveedor"
    Public Const Porcentaje As String = "Porcentaje"
End Class
