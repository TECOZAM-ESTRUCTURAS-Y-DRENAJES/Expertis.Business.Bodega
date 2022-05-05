Public Class BdgTarifa

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgTarifa"

#End Region

#Region "Estructuras"

    Public Enum enumBdgTipoTarifa
        Grado = 0
        Formula = 1
        PorVariables = 2
        PorKilogrado = 3
    End Enum

    <Serializable()> _
    Public Class udtPrecioTarifa
        Public Precio As Double
        Public GradoTarifa As Double

        Public Sub New()
        End Sub
    End Class

#End Region

#Region "Eventos Entidad"

    Public Overloads Function GetItemRow(ByVal IDTarifa As String) As DataRow
        Dim dt As DataTable = New BdgTarifa().SelOnPrimaryKey(IDTarifa)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe la Tarifa |", IDTarifa)
        Else : Return dt.Rows(0)
        End If
    End Function

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarScript)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_T.IDTarifa)) = 0 Then ApplicationService.GenerateError("La Tarifa es un dato obligatorio.")
        If Length(data(_T.DescTarifa)) = 0 Then ApplicationService.GenerateError("La Descripción es obligatoria.")
        If Length(data(_T.TipoTarifa)) = 0 Then ApplicationService.GenerateError("El Tipo de Tarifa es obligatorio.")
        If data(_T.TipoTarifa) = enumBdgTipoTarifa.PorKilogrado Or data(_T.TipoTarifa) = enumBdgTipoTarifa.PorVariables Then
            If Length(data(_T.IDVariable)) = 0 Then ApplicationService.GenerateError("La Variable es obligatoria.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarScript(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("TipoTarifa") = enumBdgTipoTarifa.Formula Then
            If Length(data(_T.TextoFormula)) > 0 Then
                Dim StData As New StScriptValidate(data(_T.IDTarifa), data(_T.TextoFormula))
                ProcessServer.ExecuteTask(Of StScriptValidate)(AddressOf ScriptValidate, StData, services)
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StPrecioEntrada
        Public IDTarifa As String
        Public Grado As Double
        Public IDEntrada As Integer
        Public dtEntrada As DataTable
        Public dvVarData As DataView
        Public Srpt As ScriptEngine

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDTarifa As String, _
                       ByVal Grado As Double, _
                       ByVal IDEntrada As Integer, _
                       ByVal dtEntrada As DataTable, _
                       ByVal dvVarData As DataView, _
                       Optional ByVal Srpt As ScriptEngine = Nothing)
            Me.IDTarifa = IDTarifa
            Me.Grado = Grado
            Me.IDEntrada = IDEntrada
            Me.dtEntrada = dtEntrada
            Me.dvVarData = dvVarData
            Me.Srpt = Srpt
        End Sub
    End Class

    <Task()> Public Shared Function PrecioEntrada(ByVal data As StPrecioEntrada, ByVal services As ServiceProvider) As udtPrecioTarifa

        PrecioEntrada = New udtPrecioTarifa

        Dim drTarifa As DataRow
        Dim udtPG As udtPrecioTarifa

        Dim ClsBdgTar As New BdgTarifa
        drTarifa = ClsBdgTar.GetItemRow(data.IDTarifa)

        Select Case drTarifa(_T.TipoTarifa)
            Case enumBdgTipoTarifa.Formula
                If data.Srpt Is Nothing Then data.Srpt = New ScriptEngine
                PrecioEntrada.Precio = data.Srpt.Eval(data.IDTarifa, data.IDEntrada, data.dtEntrada.Rows(0), data.dvVarData)
                PrecioEntrada.GradoTarifa = data.Grado
            Case enumBdgTipoTarifa.Grado
                'se penaliza o premia el grado
                If drTarifa(_T.GradoSup) > 0 Then
                    If data.Grado >= drTarifa(_T.GradoSup) Then
                        data.Grado = data.Grado + drTarifa(_T.MasGrado)
                    End If
                End If
                If drTarifa(_T.GradoInf) > 0 Then
                    If data.Grado <= drTarifa(_T.GradoInf) Then
                        data.Grado = data.Grado - drTarifa(_T.MenosGrado)
                    End If
                End If

                If data.Grado < 0 Then data.Grado = 0
                PrecioEntrada.GradoTarifa = data.Grado

                'se busca el precio para el grado
                Dim StPrecio As New StPrecioGrado(data.IDTarifa, data.Grado)
                udtPG = ProcessServer.ExecuteTask(Of StPrecioGrado, udtPrecioTarifa)(AddressOf PrecioGrado, StPrecio, services)
                If data.Grado < udtPG.GradoTarifa Then
                    PrecioEntrada.Precio = 0
                Else
                    PrecioEntrada.Precio = udtPG.Precio
                End If

                'se ajusta el precio
                If udtPG.GradoTarifa > 0 Then
                    Dim dblIncGrd As Double = drTarifa(_T.IncGrado)
                    Dim dblIncPrc As Double = drTarifa(_T.IncPrecio)

                    If dblIncGrd <> 0 And dblIncPrc <> 0 Then
                        Dim lndLapsos As Integer = Int((data.Grado - udtPG.GradoTarifa) / dblIncGrd)
                        PrecioEntrada.Precio = udtPG.Precio + (lndLapsos * dblIncPrc)
                    End If
                End If
                If PrecioEntrada.Precio < 0 Then PrecioEntrada.Precio = 0
            Case enumBdgTipoTarifa.PorKilogrado, enumBdgTipoTarifa.PorVariables
                'se busca el precio para el valor de la variable
                Dim dataPrecio As New StPrecioGrado(data.IDTarifa, data.Grado)
                udtPG = ProcessServer.ExecuteTask(Of StPrecioGrado, udtPrecioTarifa)(AddressOf PrecioGrado, dataPrecio, services)
                If data.Grado < udtPG.GradoTarifa Then
                    PrecioEntrada.Precio = 0
                    PrecioEntrada.GradoTarifa = 0
                Else
                    If drTarifa(_T.TipoTarifa) = enumBdgTipoTarifa.PorKilogrado Then
                        PrecioEntrada.Precio = udtPG.Precio * data.Grado
                    Else
                        PrecioEntrada.Precio = udtPG.Precio
                    End If

                    PrecioEntrada.GradoTarifa = udtPG.GradoTarifa
                End If

        End Select

    End Function

    <Serializable()> _
    Public Class StPrecioGrado
        Public IDTarifa As String
        Public Grado As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDTarifa As String, ByVal Grado As Double)
            Me.IDTarifa = IDTarifa
            Me.Grado = Grado
        End Sub
    End Class

    <Task()> Public Shared Function PrecioGrado(ByVal data As StPrecioGrado, ByVal services As ServiceProvider) As udtPrecioTarifa

        PrecioGrado = New udtPrecioTarifa

        Dim f As Filter
        f = New Filter
        f.Add(New StringFilterItem(_T.IDTarifa, FilterOperator.Equal, data.IDTarifa))
        f.Add(New NumberFilterItem(_TG.GradoDesde, FilterOperator.LessThanOrEqual, data.Grado))

        Dim fwnTarifaGrd As New BdgTarifaGrado
        Dim dtTarifaGrd As DataTable = fwnTarifaGrd.Filter(f, _TG.GradoDesde & " DESC")

        If dtTarifaGrd.Rows.Count > 0 Then
            PrecioGrado.Precio = dtTarifaGrd.Rows(0)(_TG.Precio)
            PrecioGrado.GradoTarifa = dtTarifaGrd.Rows(0)(_TG.GradoDesde)
        Else
            Dim oF1 As Filter
            oF1 = New Filter
            oF1.Add(_T.IDTarifa, FilterOperator.Equal, data.IDTarifa)
            oF1.Add(_TG.GradoDesde, FilterOperator.GreaterThanOrEqual, data.Grado)

            'strWhere = "IDTarifa = '" & strIDTarifa & "' AND GradoDesde >= " & Str(dblGrado)

            dtTarifaGrd = fwnTarifaGrd.Filter(oF1, _TG.GradoDesde)

            If dtTarifaGrd.Rows.Count > 0 Then
                PrecioGrado.Precio = dtTarifaGrd.Rows(0)(_TG.Precio)
                PrecioGrado.GradoTarifa = dtTarifaGrd.Rows(0)(_TG.GradoDesde)
            End If
        End If
    End Function

    <Task()> Public Shared Sub ValidatePrimaryKey(ByVal IDTarifa As String, ByVal services As ServiceProvider)
        Dim DtAux As DataTable = New BdgTarifa().SelOnPrimaryKey(IDTarifa)
        If Not DtAux Is Nothing AndAlso DtAux.Rows.Count = 0 Then
            ApplicationService.GenerateError("No se encontró la Tarifa: | en la tabla: |.", IDTarifa, cnEntidad)
        End If
    End Sub

    <Task()> Public Shared Sub ValidateDuplicateKey(ByVal IDTarifa As String, ByVal Services As ServiceProvider)
        Dim DtAux As DataTable = New BdgTarifa().SelOnPrimaryKey(IDTarifa)
        If Not DtAux Is Nothing AndAlso DtAux.Rows.Count > 0 Then
            ApplicationService.GenerateError(12516, cnEntidad)
        End If
    End Sub

    <Serializable()> _
    Public Class StScriptValidate
        Public IDTarifa As String
        Public FunctionBody As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDTarifa As String, ByVal FunctionBody As String)
            Me.IDTarifa = IDTarifa
            Me.FunctionBody = FunctionBody
        End Sub
    End Class

    <Task()> Public Shared Sub ScriptValidate(ByVal data As StScriptValidate, ByVal services As ServiceProvider)
        Dim scr As New ScriptEngine
        scr.Validate(data.IDTarifa, data.FunctionBody)
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _T
    Public Const IDTarifa As String = "IDTarifa"
    Public Const DescTarifa As String = "DescTarifa"
    Public Const IncGrado As String = "IncGrado"
    Public Const IncPrecio As String = "IncPrecio"
    Public Const GradoSup As String = "GradoSup"
    Public Const MasGrado As String = "MasGrado"
    Public Const GradoInf As String = "GradoInf"
    Public Const MenosGrado As String = "MenosGrado"
    Public Const TipoTarifa As String = "TipoTarifa"
    Public Const TextoFormula As String = "TextoFormula"
    Public Const IDVariable As String = "IDVariable"
End Class