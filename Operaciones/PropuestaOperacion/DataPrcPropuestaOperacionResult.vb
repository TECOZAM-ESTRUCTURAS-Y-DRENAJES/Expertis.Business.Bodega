<Serializable()> _
Public Class DataPrcPropuestaOperacionResult
    Public TipoOperacion As enumBdgOrigenOperacion

    Public OperacionCabecera As DataTable
    Public OperacionVinoOrigen As DataTable
    Public OperacionVinoDestino As DataTable

    Public OperacionMaterial As DataTable
    Public OperacionMaterialLotes As DataTable
    Public OperacionMOD As DataTable
    Public OperacionCentro As DataTable
    Public OperacionVarios As DataTable

    Public OperacionVinoMaterial As DataTable
    Public OperacionVinoMaterialLotes As DataTable
    Public OperacionVinoMOD As DataTable
    Public OperacionVinoCentro As DataTable
    Public OperacionVinoVarios As DataTable
    Public OperacionVinoAnalisis As DataTable
    Public OperacionVinoAnalisisVariable As DataTable

    Public Sub New(ByVal TipoOperacion As enumBdgOrigenOperacion)
        Me.TipoOperacion = TipoOperacion
    End Sub
End Class


<Serializable()> _
Public Class DataPrcPropuestaOperacionResultLog

    Public logPropuesta As LogProcess
    Public lstPropuestas As List(Of DataPrcPropuestaOperacionResult)

    Public Sub New(ByVal lstPropuestas As List(Of DataPrcPropuestaOperacionResult), ByVal logPropuesta As LogProcess)
        Me.lstPropuestas = lstPropuestas
        Me.logPropuesta = logPropuesta
    End Sub

End Class