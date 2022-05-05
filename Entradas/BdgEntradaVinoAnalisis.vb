Public Class BdgEntradaVinoAnalisis

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaVinoAnalisis"

#End Region

#Region "Eventos Entidad"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDVariable", AddressOf CambioIDVariable)
        Obrl.Add("Valor", AddressOf CambioValor)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioIDVariable(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        Dim dtVar As DataTable = New BdgVariable().SelOnPrimaryKey(data.Current("IDVariable"))
        If dtVar.Rows.Count > 0 Then
            data.Current("Abreviatura") = dtVar.Rows(0)("Abreviatura")
            data.Current("DescVariable") = dtVar.Rows(0)("DescVariable")
            data.Current("TipoVariable") = dtVar.Rows(0)("TipoVariable")
            data.Current("Valor") = System.DBNull.Value
            data.Current("ValorNumerico") = System.DBNull.Value
            data.Current("Lista") = dtVar.Rows(0)("Lista")
            data.Current("UdMedida") = dtVar.Rows(0)("UdMedida")
            data.Current("Maximo") = dtVar.Rows(0)("Maximo")
            data.Current("Minimo") = dtVar.Rows(0)("Minimo")
            data.Current("ColorMaximo") = dtVar.Rows(0)("ColorMaximo")
            data.Current("ColorMinimo") = dtVar.Rows(0)("ColorMinimo")
            data.Current("NDecimales") = dtVar.Rows(0)("NDecimales")
        Else
            data.Current("Abreviatura") = System.DBNull.Value
            data.Current("DescVariable") = System.DBNull.Value
            data.Current("TipoVariable") = System.DBNull.Value
            data.Current("Valor") = System.DBNull.Value
            data.Current("ValorNumerico") = System.DBNull.Value
            data.Current("Lista") = False
            data.Current("UdMedida") = System.DBNull.Value
            data.Current("Maximo") = System.DBNull.Value
            data.Current("Minimo") = System.DBNull.Value
            data.Current("ColorMaximo") = System.DBNull.Value
            data.Current("ColorMinimo") = System.DBNull.Value
            data.Current("NDecimales") = 0
        End If
    End Sub

    <Task()> Public Shared Sub CambioValor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If data.Current("TipoVariable") = BdgTipoVariable.Numerica Then
            If Length(data.Current("Valor")) > 0 Then
                If IsNumeric(data.Current("Valor")) Then
                    Dim dblMax As Double = Nz(data.Current("Maximo"), 0)
                    Dim dblMin As Double = Nz(data.Current("Minimo"), 0)
                    Dim dblValor As Double = CDbl(data.Current("Valor"))
                    If dblMax <> 0 OrElse dblMin <> 0 Then
                        If Not ((dblMin <= dblValor) And (dblMax >= dblValor)) Then
                            ApplicationService.GenerateError("El valor de la variable no está dentro del intervalo establecido.")
                        Else
                            data.Current("ValorNumerico") = dblValor
                        End If
                    ElseIf dblMax = 0 AndAlso dblMin = 0 Then
                        data.Current("ValorNumerico") = dblValor
                    End If
                Else
                    ApplicationService.GenerateError("El campo Valor debe ser numérico.")
                End If
            Else
                data.Current("ValorNumerico") = System.DBNull.Value
            End If
        Else
            data.Current("ValorNumerico") = System.DBNull.Value
        End If
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _EVA
    Public Const NEntrada As String = "NEntrada"
    Public Const IDVariable As String = "IDVariable"
    Public Const Valor As String = "Valor"
    Public Const ValorNumerico As String = "ValorNumerico"
End Class