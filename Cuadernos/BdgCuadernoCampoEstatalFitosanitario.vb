Public Class BdgCuadernoCampoEstatalFitosanitario

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCuadernoCampoEstatalFitosanitario"

#End Region

#Region "Eventos Entidad"
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    Public Overrides Function GetBusinessRules() As Solmicro.Expertis.Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDFinca", AddressOf CambioIDFinca)
        oBrl.Add("CFinca", AddressOf CambioCFinca)
        oBrl.Add("IDVariedad", AddressOf CambioVariedad)
        oBrl.Add("IDArticulo", AddressOf CambioArticulo)
        oBrl.Add("CantidadArticulo", AddressOf CambioCantidadArticulo)
        oBrl.Add("CantidadTratada", AddressOf CambioCantidadArticulo)
        oBrl.Add("IDRegistroFitosanitario", AddressOf CambioIDRegistroFitosanitario)
        Return oBrl
    End Function
#End Region

#Region "Funciones Públicas"
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)

        Dim blnValidarFinca As Boolean = data("IDTipoRegistro") = BdgTipoFitosanitario.Parcela _
                                            OrElse data("IDTipoRegistro") = BdgTipoFitosanitario.ParcelaAsesorada _
                                            OrElse data("IDTipoRegistro") = BdgTipoFitosanitario.Semilla
        If blnValidarFinca Then
            If Length(data("IDFinca")) = 0 Then ApplicationService.GenerateError("La Finca es un dato obligatorio.")
            Dim dt As DataTable = New BdgFinca().SelOnPrimaryKey(data("IDFinca"))
            If Length(dt) = 0 Then ApplicationService.GenerateError("La Finca no existe.")
        End If

        If Length(data("Fecha")) = 0 Then ApplicationService.GenerateError("La Fecha es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDCuadernoFitosanitario") = Guid.NewGuid
    End Sub

    <Task()> Public Shared Sub CambioCFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalFitosanitario
            Dim f As New StringFilterItem("CFinca", FilterOperator.Equal, data.Value)
            Dim dt As DataTable = New BdgFinca().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("IDFinca") = dt.Rows(0)("IDFinca")
                data.Current("SuperficieViñedo") = dt.Rows(0)("SuperficieViñedo")
                data.Current("IDVariedad") = dt.Rows(0)("IDVariedad")
                data.Current = ClsBdg.ApplyBusinessRule("IDVariedad", data.Current("IDVariedad"), data.Current)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalFitosanitario
            Dim dttFinca As DataTable = New BdgFinca().SelOnPrimaryKey(data.Current("IDFinca"))
            If (Not dttFinca Is Nothing AndAlso dttFinca.Rows.Count > 0) Then
                With dttFinca
                    data.Current("CFinca") = .Rows(0)("CFinca")
                    data.Current = ClsBdg.ApplyBusinessRule("CFinca", data.Current("CFinca"), data.Current)
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDRegistroFitosanitario(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim dtRegFit As DataTable = New BdgRegistroFitosanitario().SelOnPrimaryKey(data.Current("IDRegistroFitosanitario"))
            If (Not dtRegFit Is Nothing AndAlso dtRegFit.Rows.Count > 0) Then
                With dtRegFit
                    If Length(data.Current("DescArticulo")) = 0 Then data.Current("DescArticulo") = .Rows(0)("NombreComercialFitosanitario")
                    Dim f As New StringFilterItem("IDRegistroFitosanitario", FilterOperator.Equal, data.Value)
                    Dim dtArt As DataTable = New Articulo().Filter(f)
                    If Not dtArt Is Nothing AndAlso dtArt.Rows.Count > 0 Then
                        data.Current("IDArticulo") = dtArt.Rows(0)("IDArticulo")
                    End If
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioVariedad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim f As New StringFilterItem("IDVariedad", FilterOperator.Equal, data.Value)
            Dim dt As DataTable = New BdgVariedad().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("IDEspecie") = dt.Rows(0)("IDEspecie")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalFitosanitario
            Dim dtArt As DataTable = New Articulo().SelOnPrimaryKey(data.Value)
            If (Not dtArt Is Nothing AndAlso dtArt.Rows.Count > 0) Then
                With dtArt
                    data.Current("DescArticulo") = .Rows(0)("DescArticulo")
                    data.Current("IDRegistroFitosanitario") = .Rows(0)("IDRegistroFitosanitario")
                    data.Current = ClsBdg.ApplyBusinessRule("IDRegistroFitosanitario", data.Current("IDRegistroFitosanitario"), data.Current)
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCantidadArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim dblCantidadArticulo As Double = 0
            Dim dblCantidadTratada As Double = 0
            Select Case data.ColumnName
                Case "CantidadArticulo"
                    dblCantidadArticulo = data.Value
                    dblCantidadTratada = Nz(data.Current("CantidadTratada"), 0)
                Case "CantidadTratada"
                    dblCantidadArticulo = Nz(data.Current("CantidadArticulo"), 0)
                    dblCantidadTratada = data.Value
            End Select
            If dblCantidadTratada > 0 Then
                data.Current("Dosis") = xRound(dblCantidadArticulo / dblCantidadTratada, 2)
            Else
                data.Current("Dosis") = 0
            End If
        Else
            data.Current("Dosis") = 0
        End If
    End Sub
#End Region

End Class

Public Enum BdgTipoFitosanitario
    Parcela = 0
    ParcelaAsesorada = 1
    Semilla = 2
    Postcosecha = 3
    Locales = 4
    Transporte = 5
End Enum