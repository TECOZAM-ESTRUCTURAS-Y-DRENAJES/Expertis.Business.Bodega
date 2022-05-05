Public Class BdgFincaClonVariedad

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgFincaClonVariedad"

#End Region

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("Porcentaje", AddressOf CambioPorcentaje)
        Obrl.Add("Superficie", AddressOf CambioSuperficie)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioPorcentaje(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim DblSuperficie As Double = 0
        Dim DblPctj As Double = 0
        If Length(data.Context("Superficie")) > 0 Then DblSuperficie = data.Context("Superficie")
        If Length(data.Value) > 0 Then
            DblPctj = data.Value
        Else : ApplicationService.GenerateError("El campo Porcentaje debe ser numérico.")
        End If
        Dim dblQ As Double = DblSuperficie * DblPctj / 100
        data.Current("Porcentaje") = dblPctj
        data.Current("Superficie") = dblQ
    End Sub

    <Task()> Public Shared Sub CambioSuperficie(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim DblSuperficie As Double = 0
        Dim DblPctj As Double = 0
        Dim DblQ As Double = 0
        If Length(data.Context("Superficie")) > 0 Then DblSuperficie = data.Context("Superficie")
        If Length(data.Value) > 0 Then
            DblQ = data.Value
        Else : ApplicationService.GenerateError("El campo Superficie debe ser numérico.")
        End If
        If DblSuperficie > 0 Then DblPctj = DblQ / DblSuperficie * 100
        data.Current("Porcentaje") = DblPctj
        data.Current("Superficie") = DblQ
    End Sub

End Class
