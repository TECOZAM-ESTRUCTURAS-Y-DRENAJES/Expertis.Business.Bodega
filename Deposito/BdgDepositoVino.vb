Public Class BdgDepositoVino

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgDepositoVino"

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function SelOnIDDeposito(ByVal IDDeposito As String, ByVal services As ServiceProvider) As DataTable
        Return New BdgDepositoVino().Filter(New StringFilterItem(_DV.IDDeposito, FilterOperator.Equal, IDDeposito))
    End Function

    Public Overloads Function GetItemRow(ByVal IDVino As Guid) As DataRow
        Dim Services As New ServiceProvider
        Dim Dt As DataTable = New BdgDepositoVino().SelOnPrimaryKey(IDVino)
        If Dt Is Nothing OrElse Dt.Rows.Count = 0 Then
            Dim stError As String = "El vino ya no está en el depósito."
            Dim StMensaje As New BdgWorkClass.StMensajeError(IDVino, stError)
            stError = String.Format(ProcessServer.ExecuteTask(Of BdgWorkClass.StMensajeError, String)(AddressOf BdgWorkClass.MensajeErrorVino, StMensaje, services))
            ApplicationService.GenerateError(stError)
        Else : Return Dt.Rows(0)
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _DV
    Public Const IDVino As String = "IDVino"
    Public Const IDDeposito As String = "IDDeposito"
    Public Const Cantidad As String = "Cantidad"
End Class