Public Class BdgEntradaAnalisis

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaAnalisis"

#End Region

#Region "Eventos Entidad"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("Valor", AddressOf CambioValor)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioValor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.current(_Vr.TipoVariable) = BdgTipoVariable.Numerica Then
            If Length(data.Value) > 0 Then
                Dim dblMax As Double = Nz(data.Current("Maximo"))
                Dim dblMin As Double = Nz(data.Current("Minimo"))

                Dim dblValor As Double = CDbl(data.Value)
                If dblMax <> 0 Or dblMin <> 0 Then
                    If Not ((dblMin <= dblValor) And (dblMax >= dblValor)) Then
                        ApplicationService.GenerateError("El valor de la variable no está dentro del intervalo establecido.")
                    End If
                End If
            ElseIf Length(data.Value) = 0 Then
            Else : ApplicationService.GenerateError("El campo Valor debe ser numérico.")
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub BorrarEntradaAnalisis(ByVal idEntradaAnalisis As Integer, ByVal services As ServiceProvider)
        Dim dtEntradaAnalisis As DataTable = New BE.DataEngine().Filter("tbBdgEntradaAnalisis", New FilterItem("IDEntrada", FilterOperator.Equal, idEntradaAnalisis, FilterType.Numeric))
        For Each drEntradaAnalisis As DataRow In dtEntradaAnalisis.Rows
            AdminData.DeleteData("BdgEntradaAnalisis", drEntradaAnalisis)
        Next
    End Sub

    <Serializable()> _
    Public Class StGetAnalisis
        Public IDAnalisis As String
        Public Filtro As Filter
        Public Detalle As enumDetalleEntradaUva

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDAnalisis As String, ByVal Filtro As Filter, ByVal Detalle As enumDetalleEntradaUva)
            Me.IDAnalisis = IDAnalisis
            Me.Filtro = Filtro
            Me.Detalle = Detalle
        End Sub
    End Class

    <Task()> Public Shared Function GetAnalisis(ByVal data As StGetAnalisis, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable
        Select Case data.Detalle
            Case enumDetalleEntradaUva.Cabecera
                dt = New DataEngine().Filter("frmBdgCIAnalisisEntradaUva", data.Filtro)
            Case enumDetalleEntradaUva.Cartillista
                dt = New DataEngine().Filter("frmBdgCIAnalisisEntradaUvaCartillista", data.Filtro)
            Case enumDetalleEntradaUva.Finca
                dt = New DataEngine().Filter("frmBdgCIAnalisisEntradaUvaFinca", data.Filtro)
        End Select
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Dim strSelect As String = String.Empty
            'ahora traemos todas las columnas correspondientes
            Dim dtVariables As DataTable = New BdgAnalisisVariable().Filter(New StringFilterItem("IDAnalisis", data.IDAnalisis))
            Dim fVariables As New Filter(FilterUnionOperator.Or)
            For Each dcV As DataRow In dtVariables.Select(Nothing, "Orden")
                If Not dt.Columns.Contains(dcV("IDVariable")) Then
                    dt.Columns.Add(dcV("IDVariable"), GetType(String))
                    dt.Columns.Add(dcV("IDVariable") & "_N", GetType(Double))
                    fVariables.Add("IDVariable", dcV("IDVariable"))
                    strSelect = strSelect + dcV("IDVariable") + ","
                End If
            Next
            strSelect = strSelect + "0"

            'ahora las actualizamos
            Dim bsnBEA As New BdgEntradaAnalisis
            For Each dr As DataRow In dt.Select
                Dim f As New Filter
                f.Add("IDEntrada", dr("IDEntrada"))
                f.Add(fVariables)
                Dim dtResult As DataTable = bsnBEA.Filter(f)
                For Each drResult As DataRow In dtResult.Select
                    dr(drResult("IDVariable")) = drResult("Valor")
                    dr(drResult("IDVariable") & "_N") = drResult("ValorNumerico")
                Next
            Next
        End If
        Return dt
    End Function

#End Region

End Class

<Serializable()> _
Public Class _EA
    Public Const IDEntrada As String = "IDEntrada"
    Public Const IDVariable As String = "IDVariable"
    Public Const Valor As String = "Valor"
    Public Const ValorNumerico As String = "ValorNumerico"
End Class