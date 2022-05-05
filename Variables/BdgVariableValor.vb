Public Class BdgVariableValor

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVariableValor"

#End Region

#Region "Eventos Entidad"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("ValorNumerico", AddressOf CambioValorNumerico)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioValorNumerico(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current("Valor") = CDbl(data.Value).ToString
        End If
    End Sub

#End Region

End Class