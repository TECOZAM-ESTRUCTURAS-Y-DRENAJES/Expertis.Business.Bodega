Public Class BdgEntradaVto

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaVto"

#End Region

#Region "Funciones Públicas"
    <Task()> Public Shared Sub NuevaEntradaVto(ByVal Data As DataEntradaVto, ByVal services As ServiceProvider)

        Dim dtrNewRow As DataRow = Data.dtEntradaVto.NewRow

        dtrNewRow(_EVto.IDEntrada) = Data.IDEntrada 'drvProvFact("IDEntrada")
        dtrNewRow(_EVto.IDVendimiaVto) = Data.IDVendimiaVto
        dtrNewRow(_EVto.IDProveedor) = Data.IDProveedor 'drvProvFact("IDProveedor")

        dtrNewRow(_EVto.TipoVariedad) = Data.TipoVariedad 'drvProvFact("TipoVariedad")
        dtrNewRow(_EVto.GradoEntrada) = Data.dblVariableValor
        dtrNewRow(_EVto.GradoTarifa) = Data.dblVariableValorTarifa
        dtrNewRow(_EVto.IDVariable) = Data.IDVariableTarifa 'strIDVariableTarifa
        If Len(Data.IDTarifa) Then dtrNewRow(_EVto.IDTarifa) = Data.IDTarifa

        dtrNewRow(_EVto.Precio) = Data.dblPrecioDeclarado
        dtrNewRow(_EVto.Kilos) = Data.dblCantidadDeclarado 'drvProvFact("CantidadDeclarado")
        dtrNewRow(_EVto.Importe) = Data.dblPrecioDeclarado * Data.dblCantidadDeclarado

        dtrNewRow(_EVto.PrecioExc) = Data.dblPrecioExcedente
        dtrNewRow(_EVto.KilosExc) = Data.dblCantidadExcedente 'drvProvFact("CantidadExcedente")
        dtrNewRow(_EVto.ImporteExc) = Data.dblPrecioExcedente * Data.dblCantidadExcedente

        dtrNewRow(_EVto.PrecioO) = Data.dblPrecioOrigen
        dtrNewRow(_EVto.KilosO) = Data.dblCantidadOrigen 'drvProvFact("CantidadOrigen")
        dtrNewRow(_EVto.ImporteO) = Data.dblPrecioOrigen * Data.dblCantidadOrigen

        dtrNewRow(_EVto.PrecioSO) = Data.dblPrecioSinOrigen
        dtrNewRow(_EVto.KilosSO) = Data.dblCantidadSinOrigen 'drvProvFact("CantidadSinOrigen")
        dtrNewRow(_EVto.ImporteSO) = Data.dblPrecioSinOrigen * Data.dblCantidadSinOrigen

        dtrNewRow(_EVto.KilosEntrada) = Data.dblCantidadDeclarado + Data.dblCantidadExcedente + Data.dblCantidadOrigen + Data.dblCantidadSinOrigen
        dtrNewRow(_EVto.ImporteEntrada) = dtrNewRow(_EVto.Importe) + dtrNewRow(_EVto.ImporteExc) + dtrNewRow(_EVto.ImporteO) + dtrNewRow(_EVto.ImporteSO)
        If dtrNewRow(_EVto.KilosEntrada) <> 0 Then dtrNewRow(_EVto.PrecioEntrada) = dtrNewRow(_EVto.ImporteEntrada) / dtrNewRow(_EVto.KilosEntrada)

        Data.dtEntradaVto.Rows.Add(dtrNewRow)
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _EVto
    Public Const IDEntradaVto As String = "IDEntradaVto"
    Public Const IDEntrada As String = "IDEntrada"
    Public Const IDVendimiaVto As String = "IDVendimiaVto"
    Public Const IDProveedor As String = "IDProveedor"

    Public Const TipoVariedad As String = "TipoVariedad"
    Public Const GradoEntrada As String = "GradoEntrada"
    Public Const GradoTarifa As String = "GradoTarifa"
    Public Const IDVariable As String = "IDVariable"
    Public Const IDTarifa As String = "IDTarifa"

    Public Const Precio As String = "Precio"
    Public Const Kilos As String = "Kilos"
    Public Const Importe As String = "Importe"

    Public Const PrecioExc As String = "PrecioExc"
    Public Const KilosExc As String = "KilosExc"
    Public Const ImporteExc As String = "ImporteExc"

    Public Const PrecioO As String = "PrecioO"
    Public Const KilosO As String = "KilosO"
    Public Const ImporteO As String = "ImporteO"

    Public Const PrecioSO As String = "PrecioSO"
    Public Const KilosSO As String = "KilosSO"
    Public Const ImporteSO As String = "ImporteSO"

    Public Const PrecioEntrada As String = "PrecioEntrada"
    Public Const KilosEntrada As String = "KilosEntrada"
    Public Const ImporteEntrada As String = "ImporteEntrada"

    Public Const IDLineaFactura As String = "IDLineaFactura"
    Public Const IDLineaFacturaExc As String = "IDLineaFacturaExc"
    Public Const IDLineaFacturaO As String = "IDLineaFacturaO"
    Public Const IDLineaFacturaSO As String = "IDLineaFacturaSO"
End Class

<Serializable()> _
Public Class DataEntradaVto
    Public dtEntradaVto As DataTable
    Public IDEntrada As Long = 0
    Public IDVendimiaVto As Guid
    Public IDProveedor As String = String.Empty
    Public TipoVariedad As Integer = 0
    Public dblVariableValor As Double = 0
    Public dblVariableValorTarifa As Double = 0
    Public IDVariableTarifa As String = String.Empty
    Public IDTarifa As String = String.Empty

    Public dblPrecioDeclarado As Double = 0
    Public dblCantidadDeclarado As Double = 0
    Public dblPrecioExcedente As Double = 0
    Public dblCantidadExcedente As Double = 0
    Public dblPrecioOrigen As Double = 0
    Public dblCantidadOrigen As Double = 0
    Public dblPrecioSinOrigen As Double = 0
    Public dblCantidadSinOrigen As Double = 0

    Public Sub New(ByVal dtEntradaVto As DataTable, ByVal IDEntrada As Long, ByVal IDVendimiaVto As Guid, _
                   ByVal IDProveedor As String, ByVal TipoVariedad As Integer, ByVal dblVariableValor As Double, _
                   ByVal dblVariableValorTarifa As Double, ByVal IDVariableTarifa As String, ByVal IDTarifa As String, _
                   ByVal dblPrecioDeclarado As Double, ByVal dblCantidadDeclarado As Double, ByVal dblPrecioExcedente As Double, _
                   ByVal dblCantidadExcedente As Double, ByVal dblPrecioOrigen As Double, ByVal dblCantidadOrigen As Double, _
                   ByVal dblPrecioSinOrigen As Double, ByVal dblCantidadSinOrigen As Double)
        Me.dtEntradaVto = dtEntradaVto
        Me.IDEntrada = IDEntrada
        Me.IDVendimiaVto = IDVendimiaVto
        Me.IDProveedor = IDProveedor
        Me.TipoVariedad = TipoVariedad
        Me.dblVariableValor = dblVariableValor
        Me.dblVariableValorTarifa = dblVariableValorTarifa
        Me.IDVariableTarifa = IDVariableTarifa
        Me.IDTarifa = IDTarifa

        Me.dblPrecioDeclarado = dblPrecioDeclarado
        Me.dblCantidadDeclarado = dblCantidadDeclarado
        Me.dblPrecioExcedente = dblPrecioExcedente
        Me.dblCantidadExcedente = dblCantidadExcedente
        Me.dblPrecioOrigen = dblPrecioOrigen
        Me.dblCantidadOrigen = dblCantidadOrigen
        Me.dblPrecioSinOrigen = dblPrecioSinOrigen
        Me.dblCantidadSinOrigen = dblCantidadSinOrigen
    End Sub
End Class
