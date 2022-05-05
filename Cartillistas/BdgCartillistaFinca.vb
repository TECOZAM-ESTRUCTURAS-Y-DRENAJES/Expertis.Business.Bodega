Public Class BdgCartillistaFinca

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCartillistaFinca"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf TratarMunicipiosDelete)
    End Sub

    <Task()> Public Shared Sub TratarMunicipiosDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim TipoVariedad As BdgTipoVariedad
        Dim dblSuperficie, dblIncMaximo, dblIncHaT, dblIncHaB, dblIncMaxT, dblIncMaxB As Double
        Dim strIDMunicipio As String
        Dim StCalc As New StDatosCalculoRendDesdeFinca(data("IDFinca"), data("Vendimia"))
        Dim udtDatosRdto As udtBdgDatosCalculoRendimiento = ProcessServer.ExecuteTask(Of StDatosCalculoRendDesdeFinca, udtBdgDatosCalculoRendimiento)(AddressOf DatosCalculoRendimientoDesdeFinca, StCalc, services)
        With udtDatosRdto
            dblSuperficie = .Superficie
            strIDMunicipio = .MunicipioFinca
            TipoVariedad = .TipoVariedad
        End With

        Select Case TipoVariedad
            Case BdgTipoVariedad.Tinta
                dblIncHaT = -dblSuperficie
            Case BdgTipoVariedad.Blanca
                dblIncHaB = -dblSuperficie
        End Select

        dblIncMaximo = -Nz(data("Maximo"), 0)

        Select Case TipoVariedad
            Case BdgTipoVariedad.Tinta
                dblIncMaxT = dblIncMaximo
            Case BdgTipoVariedad.Blanca
                dblIncMaxB = dblIncMaximo
        End Select

        Dim rwCartillista As DataRow = New BdgCartillista().GetItemRow(data("IDCartillista"))
        If rwCartillista("CupoCartillaPor") = BdgCupoCartillaPor.Finca Then
            Dim StPrepMunicipio As New BdgCartillista.StPrepararMunicipio(data("IDCartillista"), data("Vendimia"), strIDMunicipio, dblIncHaT, dblIncHaB, dblIncMaxT, dblIncMaxB)
            Dim dtMunicipio As DataTable = ProcessServer.ExecuteTask(Of BdgCartillista.StPrepararMunicipio, DataTable)(AddressOf BdgCartillista.PrepararMunicipio, StPrepMunicipio, services)
            If Not dtMunicipio Is Nothing Then
                If dtMunicipio.Rows.Count > 0 Then
                    Dim rwMunicipio As DataRow = dtMunicipio.Rows(0)
                    If rwMunicipio("HaT") <= 0 And rwMunicipio("HaB") <= 0 Then
                        rwMunicipio.Delete()
                    End If
                End If
            End If

            Dim StMunSob As New BdgCartillista.StMunicipiosSobrantes(data("IDCartillista"), data("Vendimia"))
            Dim DtMunicipiosAEliminar As DataTable = ProcessServer.ExecuteTask(Of BdgCartillista.StMunicipiosSobrantes, DataTable)(AddressOf BdgCartillista.MunicipiosSobrantes, StMunSob, services)

            BusinessHelper.UpdateTable(dtMunicipio)
            BusinessHelper.UpdateTable(DtMunicipiosAEliminar)
        End If
    End Sub

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("Rdto", AddressOf CambioRdto)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioRdto(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDCartillista")) > 0 Then
            If Not IsNumeric(data.Value) Then data.Value = 100
        End If
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarVendimia)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCartillista")) = 0 Then ApplicationService.GenerateError("El Cartillista es obligatorio.")
        If Length(data("IDFinca")) = 0 Then ApplicationService.GenerateError("La Finca es obligatoria.")
        If (New BdgParametro().GestionExplotadorObligatoria) Then
            If Length(data("IDExplotadorFincas")) = 0 Then ApplicationService.GenerateError("El Explotador es obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState <> DataRowState.Unchanged Then
            Dim IntVendimia As Integer = 0
            If data.RowState = DataRowState.Added Then
                If Length(data("Vendimia")) = 0 Then
                    IntVendimia = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf BdgVendimia.UltimaVendimia, New Object, services)
                Else : IntVendimia = data("Vendimia")
                End If
                Dim oFltr As New Filter
                oFltr.Add(New NumberFilterItem("Vendimia", IntVendimia))
                oFltr.Add(New GuidFilterItem("IDFinca", IntVendimia))
                Dim rcsAux As DataTable = New BdgCartillistaFinca().Filter(oFltr)
                If rcsAux.Rows.Count > 0 Then
                    ApplicationService.GenerateError("La finca '|' ya está asociada al cartillista | para la vendimia |.", data("IDFinca"), rcsAux.Rows(0)("IDCartillista"), IntVendimia)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarVendimia(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If Length(data("Vendimia")) = 0 Then ApplicationService.GenerateError("El valor asignado a la vendimia no es válido.")
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class udtBdgDatosCalculoRendimiento
        Public Superficie As Double
        Public TipoVariedad As BdgTipoVariedad
        Public MunicipioFinca As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal Superficie As Double, ByVal TipoVariedad As BdgTipoVariedad, ByVal MunicipioFinca As String)
            Me.Superficie = Superficie
            Me.TipoVariedad = TipoVariedad
            Me.MunicipioFinca = MunicipioFinca
        End Sub
    End Class

    <Serializable()> _
    Public Class StDatosCalculoRendDesdeFinca
        Public IDFinca As Guid
        Public Vendimia As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDFinca As Guid, ByVal Vendimia As Integer)
            Me.IDFinca = IDFinca
            Me.Vendimia = Vendimia
        End Sub
    End Class

    <Task()> Public Shared Function DatosCalculoRendimientoDesdeFinca(ByVal data As StDatosCalculoRendDesdeFinca, ByVal services As ServiceProvider) As udtBdgDatosCalculoRendimiento
        If data.Vendimia <> 0 AndAlso Not data.IDFinca.Equals(Guid.Empty) Then
            Dim rwFinca As DataRow = New BdgFinca().GetItemRow(data.IDFinca)
            Dim TipoVariedad As BdgTipoVariedad
            If Len(rwFinca("IDVariedad") & String.Empty) > 0 Then
                Dim rwVariedad As DataRow = New BdgVariedad().GetItemRow(rwFinca("IDVariedad") & String.Empty)
                TipoVariedad = rwVariedad("TipoVariedad")
            End If

            Dim StReturn As New udtBdgDatosCalculoRendimiento
            StReturn.Superficie = rwFinca("Superficie")
            StReturn.TipoVariedad = TipoVariedad
            StReturn.MunicipioFinca = rwFinca("IDMunicipio") & String.Empty
            Return StReturn
        End If
    End Function

    <Serializable()> _
    Public Class StValidate
        Public IDCartillista As String
        Public Vendimia As Integer
        Public IDFinca As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDCartillista As String, ByVal Vendimia As Integer, ByVal IDFinca As String)
            Me.IDCartillista = IDCartillista
            Me.Vendimia = Vendimia
            Me.IDFinca = IDFinca
        End Sub
    End Class

    <Task()> Public Shared Sub ValidateDuplicateKey(ByVal data As StValidate, ByVal services As ServiceProvider)
        Dim rcsAux As DataTable = New BdgCartillistaFinca().SelOnPrimaryKey(data.IDCartillista, data.Vendimia, data.IDFinca)
        If rcsAux.Rows.Count > 0 Then
            ApplicationService.GenerateError("El valor introducido ya existe en la Base de Datos")
        End If
    End Sub

#End Region

End Class