Public Class BdgVinoCentroTasa

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVinoCentroTasa"

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function Coste(ByVal IDVino As Guid, ByVal services As ServiceProvider) As DataTable
        If Not IDVino.Equals(Guid.Empty) Then
            Return AdminData.Execute("spVinoCosteTasaExplosion", False, IDVino.ToString)
        End If
    End Function

    <Serializable()> _
    Public Class StCrearVinoCentroTasa
        Public Fecha As DateTime
        Public dtCentros As DataTable

        Public Sub New(ByVal Fecha As DateTime, ByVal dtCentros As DataTable)
            Me.Fecha = Fecha
            Me.dtCentros = dtCentros
        End Sub
    End Class

    <Task()> Public Shared Sub CrearVinoCentroTasa(ByVal data As StCrearVinoCentroTasa, ByVal services As ServiceProvider)
        Dim CT As New CentroTasa
        Dim dtVCT As DataTable = New BdgVinoCentroTasa().AddNew
        Dim VCT As New BdgVinoCentroTasa
        Dim FieldTasa As String
        If Not data.dtCentros Is Nothing AndAlso data.dtCentros.Rows.Count > 0 Then
            For Each drVinoCentro As DataRow In data.dtCentros.Rows
                FieldTasa = "EjecucionValorA"
                If data.dtCentros.Columns.Contains("IDIncidencia") AndAlso Length(drVinoCentro("IDIncidencia")) > 0 Then
                    FieldTasa = "PreparacionValorA"
                End If

                '//Eliminar los CentroTasa ya creados por modificacion
                Dim fIDVinoCentro As New Filter
                fIDVinoCentro.Add("IDVinoCentro", FilterOperator.Equal, drVinoCentro("IDVinoCentro"))
                Dim dtTasasDelete As DataTable = VCT.Filter(fIDVinoCentro)
                If Not IsNothing(dtTasasDelete) AndAlso dtTasasDelete.Rows.Count > 0 Then
                    VCT.Delete(dtTasasDelete)
                End If

                '//Cogemos las tasas del centro
                Dim fTasasVigentes As New Filter
                fTasasVigentes.Add("IDCentro", FilterOperator.Equal, drVinoCentro("IDCentro"))
                fTasasVigentes.Add("FechaDesde", FilterOperator.LessThanOrEqual, data.Fecha)
                fTasasVigentes.Add("FechaHasta", FilterOperator.GreaterThanOrEqual, data.Fecha)
                Dim dtMaestroCentroTasa As DataTable = CT.Filter(fTasasVigentes)
                If dtMaestroCentroTasa.Rows.Count > 0 Then
                    '//Acumulamos la tass del centro, para ver cual es el total
                    Dim TasaTotal As Double = 0
                    TasaTotal = (Aggregate c In dtMaestroCentroTasa Into Sum(CDbl(c(FieldTasa))))


                    '//Calcular el Porcentaje de cada una de las tasas
                    For Each drCT As DataRow In dtMaestroCentroTasa.Rows
                        Dim drVCT As DataRow = dtVCT.NewRow
                        drVCT("IDVinoCentro") = drVinoCentro("IDVinoCentro")
                        drVCT("IDTasa") = drCT("IDTasa")
                        If TasaTotal <> 0 Then
                            drVCT("Porcentaje") = (drCT(FieldTasa) / TasaTotal)
                        Else
                            drVCT("Porcentaje") = 0
                        End If
                        drVCT("TipoCosteFV") = drCT("TipoCosteFV")
                        drVCT("TipoCosteDI") = drCT("TipoCosteDI")
                        drVCT("Fiscal") = drCT("Fiscal")
                        dtVCT.Rows.Add(drVCT)
                    Next
                    VCT.Update(dtVCT)
                End If
            Next
        End If
    End Sub

#End Region

End Class