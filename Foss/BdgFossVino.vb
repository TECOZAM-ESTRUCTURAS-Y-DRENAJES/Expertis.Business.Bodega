Public Class BdgFossVino

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgFossVino"

#End Region

#Region "Constantes"

    Private Const cnAnalisisFoss As String = "BDGFOSVIN"
    Private Const cnAnalisisFossDep As String = "BDGFOSDEP"
    Private Const AnalisisPredeterminado As String = "vNegAnalisisPredeterminadoFoss"
    Private Const VarCodNull As String = "vNegAnalisisFossVinoCodNull"
    Private Const cnFormulaMV As String = "MV"
    Private Const cnFormulaEST As String = "EST"
    Private Const cnFormulaENR As String = "ENR"
    Private Const cnFormulaDE As String = "DE"
    Private Const cnFormulaG As String = "G"
    Private Const cnFormulaAZ As String = "AZ"
    Private Const SepSampleID As String = "/"
    'Variables usadas en Deposito
    Private Const cnDelimiter As String = ";"
    Private Const cnInitVar As Integer = 7
    Private Const cnPosDate As Integer = 35
    Private Const cnPosFinca As Integer = 1
    Private Const cnPosCodFossVino As Integer = 0
    Private Const VarCodNullDep As String = "vNegAnalisisColFossVinoNull"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDFossVino") = AdminData.GetAutoNumeric
    End Sub

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDArticulo", AddressOf CambiosDatos)
        Obrl.Add("IDDeposito", AddressOf CambiosDatos)
        Obrl.Add("Fecha", AddressOf CambiosDatos)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambiosDatos(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Current("Recuperado") = False Then
            data.Current("Validado") = False
            data.Current("Error") = False
            data.Current("DescError") = String.Empty
            data.Current("IdVino") = DBNull.Value
            data.Current("IdVinoAnalisis") = DBNull.Value
        End If
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDFossVino")) = 0 OrElse data("IDFossVino") = 0 Then
                data("IDFossVino") = AdminData.GetAutoNumeric
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

#Region "Leer vino Access"

    <Serializable()> _
    Public Class StProcesarAccessUva
        Public dttData As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal dttData As DataTable)
            Me.dttData = dttData
        End Sub
    End Class

    <Serializable()> _
    Public Class StResultProcesarAccessUva
        Public blnResultOk As Boolean = False
        Public strCodigoFossYaImportado As String = String.Empty

        Public Sub New()
        End Sub
    End Class

    <Task()> Public Shared Function AccessDataFoss(ByVal data As StProcesarAccessUva, ByVal services As ServiceProvider) As StResultProcesarAccessUva
        Dim StResult As New StResultProcesarAccessUva

        Dim oPrm As New Parametro
        Dim e As New Filter
        Dim strParamAnalisis As String = oPrm.ObtenerPredeterminado(cnAnalisisFoss)
        e.Add("IDAnalisis", strParamAnalisis)
        e.Add(New IsNullFilterItem("CodigoFoss", False))
        Dim dttAnalisis As DataTable = New BE.DataEngine().Filter(AnalisisPredeterminado, e, , "Orden")

        If Not IsNothing(dttAnalisis) AndAlso dttAnalisis.Rows.Count > 0 Then
            Dim strSampleId As String = String.Empty
            Dim dttFoss As DataTable = New BdgFossVino().AddNew()
            Dim dttDataAux As DataTable = data.dttData.Copy

            Dim lstCodigosFoss As List(Of Object) = (From c In data.dttData Where Not c.IsNull("Samp") Select c("Samp") Distinct).ToList()
            If Not lstCodigosFoss Is Nothing AndAlso lstCodigosFoss.Count > 0 Then
                For Each drCodigoFoss As Object In lstCodigosFoss
                    Dim strCodigoFoss As String = drCodigoFoss.ToString & String.Empty

                    Dim dtVinoAnalisis As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf MirarCodigoFossEnVariable, strCodigoFoss, services)
                    If Not dtVinoAnalisis Is Nothing AndAlso dtVinoAnalisis.Rows.Count > 0 Then
                        'Ya está importado
                        If Length(StResult.strCodigoFossYaImportado) = 0 Then
                            StResult.strCodigoFossYaImportado = strCodigoFoss
                        Else
                            StResult.strCodigoFossYaImportado &= "," & strCodigoFoss
                        End If
                    Else
                        Dim lstCabeceraFoss As List(Of DataRow) = (From c In data.dttData Where Not c.IsNull("Samp") AndAlso c("Samp") = strCodigoFoss Select c).ToList()
                        If Not lstCabeceraFoss Is Nothing AndAlso lstCabeceraFoss.Count > 0 Then
                            Dim drFoss As DataRow = dttFoss.NewRow
                            Dim strSampleIDArray() As String = Split(lstCabeceraFoss(0)("SampleId"), SepSampleID, , CompareMethod.Text)
                            drFoss("IDFossVino") = AdminData.GetAutoNumeric()
                            If strSampleIDArray.Length = 2 Then
                                drFoss("IDDeposito") = strSampleIDArray(0)
                                drFoss("IDArticulo") = strSampleIDArray(1)
                            End If
                            drFoss("Fecha") = lstCabeceraFoss(0)("DateTime")
                            drFoss("Analisis") = strParamAnalisis
                            drFoss("Recuperado") = False
                            drFoss("Validado") = False
                            drFoss("Error") = False
                            drFoss("DescError") = String.Empty
                            drFoss("CodigoFoss") = lstCabeceraFoss(0)("SampNo")
                            drFoss("Vino") = True

                            Dim i As Integer = 1
                            For Each drAnalisis As DataRow In dttAnalisis.Rows
                                e.Clear()
                                e.Add("Samp", strCodigoFoss)
                                e.Add("SampleId", lstCabeceraFoss(0)("SampleId"))
                                e.Add("ProductName", drAnalisis("Calibracion"))
                                e.Add("ComponentName", drAnalisis("CodigoFoss"))
                                e.Add("IDControl", lstCabeceraFoss(0)("IDControl"))

                                Dim drAux() As DataRow = dttDataAux.Select(e.Compose(New AdoFilterComposer))
                                drFoss("CodVar" & i) = drAnalisis("IDVariable")

                                If drAux.Length > 0 Then '//ha encontrado registros
                                    drFoss("Var" & i) = xRound(drAux(0)("value"), drAnalisis("NDecimales"))
                                Else : drFoss("Var" & i) = System.DBNull.Value
                                End If
                                i = 1 + i
                            Next
                            dttFoss.Rows.Add(drFoss)
                        End If
                    End If
                Next
            End If

            If Not IsNothing(dttFoss) AndAlso dttFoss.Rows.Count > 0 Then
                BusinessHelper.UpdateTable(dttFoss)
                StResult.blnResultOk = True
            End If
        Else
            ApplicationService.GenerateError("El Análisis predeterminado {0} (Parámetro {1}) no existe o no tiene Variables relacionadas con Foss.", strParamAnalisis, cnAnalisisFoss)
        End If

        Return StResult
    End Function

#End Region

#Region "Leer Dep CSV"

    <Serializable()> _
    Public Class StProcesarLineFossVino
        Public Linea As String
        Public DtPlantilla As DataTable
        Public FechaDesde As String
        Public FechaHasta As String
        Public DtData As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal Linea As String, ByVal DtPlantilla As DataTable, ByVal FechaDesde As String, ByVal FechaHasta As String, ByVal DtData As DataTable)
            Me.Linea = Linea
            Me.DtPlantilla = DtPlantilla
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
            Me.DtData = DtData
        End Sub
    End Class

    <Task()> Public Shared Function ProcesarLineFossVino(ByVal data As StProcesarLineFossVino, ByVal services As ServiceProvider) As DataTable
        Dim blnResult As Boolean
        Dim e As New Filter
        Dim blnFechaOK As Boolean
        Dim strDeposito As String
        Dim drLine As DataRow

        If Not data.DtData Is Nothing Then drLine = data.DtData.NewRow

        data.Linea = Replace(data.Linea, """", String.Empty)
        If Len(data.Linea) > 0 Then
            Dim strArrayLine() As String = Split(data.Linea, cnDelimiter)
            If strArrayLine.Length > 0 Then
                If IsDate(strArrayLine(cnPosDate)) Then
                    blnFechaOK = True
                    Dim Dtefecha As New Date(CDate(strArrayLine(cnPosDate)).Year, CDate(strArrayLine(cnPosDate)).Month, CDate(strArrayLine(cnPosDate)).Day)
                    If Len(data.FechaDesde) > 0 Then
                        'Dim dtmFecha As Date = DateAdd(DateInterval.Day, -1, CDate(dtmFecha))
                        If Dtefecha < CDate(data.FechaDesde) Then blnFechaOK = False
                    End If
                    If Len(data.FechaHasta) > 0 And blnFechaOK Then
                        If Dtefecha > CDate(data.FechaHasta) Then blnFechaOK = False
                    End If
                    If blnFechaOK Then
                        If Len(Trim(strArrayLine(cnPosFinca))) > 0 Then
                            If Left(Trim(strArrayLine(cnPosFinca)), 1) = "D" Then
                                strDeposito = Right(Trim(strArrayLine(cnPosFinca)), Len(Trim(strArrayLine(cnPosFinca))) - 1)
                            End If
                        End If
                        If Length(strDeposito) > 0 Then
                            'En el caso de que no tenga el / toda los datos que aparecen detras del la D
                            'Se consideraran un Deposito
                            Dim posArti As Integer = strDeposito.IndexOf("/")
                            If posArti = -1 Then
                                drLine("IDDeposito") = strDeposito
                            Else
                                Dim strDepositoFinal = strDeposito.Remove(posArti, Length(strDeposito) - posArti)
                                Dim strArticulo = strDeposito.Remove(0, posArti + 1)
                                drLine("IDDeposito") = strDepositoFinal
                                drLine("IDArticulo") = strArticulo
                            End If

                            drLine("CodigoFoss") = strArrayLine(cnPosCodFossVino)
                            drLine("Recuperado") = False
                            drLine("Fecha") = Dtefecha
                            Dim objParam As New Parametro
                            drLine("Analisis") = objParam.ObtenerPredeterminado(cnAnalisisFossDep)
                            drLine("Vino") = False
                            Dim strFormat As String

                            Dim i As Integer = 1
                            For Each drVar As DataRow In data.DtPlantilla.Rows
                                If Len(drVar("Orden")) > 0 AndAlso drVar("Orden") > 0 Then
                                    If Len(drVar("ColFossVino") & String.Empty) > 0 Then
                                        drLine("CodVar" & i) = drVar("IDVariable")
                                        If strArrayLine.Length > (drVar("ColFossVino") + cnInitVar) Then
                                            If Length(Trim(strArrayLine(drVar("ColFossVino") + cnInitVar))) = 0 Then
                                                drLine("Var" & i) = System.DBNull.Value
                                            Else
                                                If drVar("NDecimales") = 0 Then
                                                    strFormat = "0"
                                                Else
                                                    strFormat = "0."
                                                End If
                                                If IsNumeric(Trim(strArrayLine(drVar("ColFossVino") + cnInitVar))) Then
                                                    'drLine("Var" & drVar("Orden")) = Format(xRound(Trim(strArrayLine(drVar("ColFossVino") + cnInitVar)), drVar("NDecimales")), strFormat.PadRight(drVar("NDecimales"), "0"))
                                                    drLine("Var" & i) = xRound(Replace(Trim(strArrayLine(drVar("ColFossVino") + cnInitVar)), ".", ","), drVar("NDecimales"))
                                                Else
                                                    drLine("Var" & i) = Trim(strArrayLine(drVar("ColFossVino") + cnInitVar))
                                                End If

                                            End If
                                        Else
                                            drLine("Var" & i) = System.DBNull.Value
                                        End If
                                        i += 1
                                    End If
                                End If
                            Next
                            If Not drLine Is Nothing Then data.DtData.Rows.Add(drLine)
                            blnResult = True
                            strDeposito = String.Empty
                        End If
                    End If
                End If
            End If
        End If
        Return data.DtData
    End Function

#End Region

#Region "Validar Datos, tanto Vino como Dep"

    <Task()> Public Shared Function ValidatedImportDataFoss(ByVal dttData As DataTable, ByVal services As ServiceProvider) As Boolean
        Dim e As New Filter
        If Not dttData Is Nothing AndAlso dttData.Rows.Count > 0 Then
            Dim objVino As New BdgVino
            Dim objAnalisisVino As New BdgVinoAnalisis
            Dim Dte As New BE.DataEngine

            Dim fil As New Filter
            Dim dtVino As DataTable
            fil.Add("Recuperado", FilterOperator.Equal, False)
            fil.Add("Validado", FilterOperator.Equal, False)
            For Each drdata As DataRow In dttData.Select(fil.Compose(New AdoFilterComposer))
                drdata("Validado") = False
                drdata("Error") = False
                drdata("DescError") = String.Empty
                e.Clear()

                'Si trae el IdVino no se tiene que comprobar si existe, porque ya se ha hecho la comprobacion en 
                'Application.FossVino
                If Length(drdata("IdVino").ToString) > 0 Then
                    dtVino = objVino.SelOnPrimaryKey(drdata("IdVino"))
                Else
                    If Length(drdata("IDArticulo")) > 0 AndAlso Length(drdata("IDDeposito")) AndAlso Length(drdata("Fecha")) Then

                        Dim f As New Filter
                        f.Add("IDArticulo", FilterOperator.Equal, drdata("IDArticulo"))
                        f.Add("IDDeposito", FilterOperator.Equal, drdata("IDDeposito"))
                        If Length(drdata("Lote")) > 0 Then
                            f.Add("Lote", FilterOperator.Equal, drdata("Lote"))
                        End If

                        Dim datDepositosFecha As New BdgVino.StDepositosEnFecha(drdata("Fecha"), f, False)
                        Dim VinosEnDeposito As DataTable = ProcessServer.ExecuteTask(Of BdgVino.StDepositosEnFecha, DataTable)(AddressOf BdgVino.DepositosEnFecha, datDepositosFecha, services)

                        dtVino = VinosEnDeposito.Clone
                        For Each dr As DataRow In VinosEnDeposito.Select("", "FechaVino")
                            dtVino.Rows.Add(dr.ItemArray)
                        Next
                    Else
                        dtVino = Nothing
                    End If
                End If

                If Not dtVino Is Nothing AndAlso dtVino.Rows.Count > 0 Then
                    drdata("IDVino") = dtVino.Rows(0)("IDVino")
                    e.Clear()
                    e.Add("IDAnalisis", drdata("Analisis"))
                    e.Add("IDVino", drdata("IDVino"))
                    'e.Add("Fecha", DateValue(drdata("Fecha")))
                    e.Add("Fecha", drdata("Fecha"))
                    Dim dtAnalisisVino As DataTable = objAnalisisVino.Filter(e)
                    If dtAnalisisVino.Rows.Count > 0 Then
                        drdata("IDVinoAnalisis") = dtAnalisisVino.Rows(0)("IDVinoAnalisis")
                        If drdata("Remplazar") = False Then
                            drdata("Error") = True
                            drdata("DescError") = "Analisis Vino existente"
                        Else : drdata("Validado") = True
                        End If
                    Else : drdata("Validado") = True
                    End If
                Else
                    drdata("DescError") = "No existe el Vino"
                    drdata("Error") = True
                End If
            Next
            Dim ClsBdgFossVino As New BdgFossVino
            ClsBdgFossVino.Update(dttData)
        End If
    End Function

#End Region

#Region "Importar Datos, tanto Vino como Dep"

    <Serializable()> _
    Public Class StImportDataFoss
        Public DttData As DataTable
        Public Vino As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal DttData As DataTable, Optional ByVal vino As Boolean = True)
            Me.DttData = DttData
            Me.Vino = vino
        End Sub
    End Class

    <Task()> Public Shared Function ImportDataFoss(ByVal data As StImportDataFoss, ByVal services As ServiceProvider) As Boolean
        AdminData.BeginTx()
        Dim objVinoAn As New BdgVinoAnalisis
        Dim objVinoAnVar As New BdgVinoVariable
        Dim objAnalisis As New BdgAnalisisVariable
        Dim blnReturn As Boolean
        Dim e As New Filter
        Dim eVar As New Filter
        Dim dttOrdenAux As DataTable

        Dim eDelete As New Filter(FilterUnionOperator.Or)

        Dim dttVinoAnalisis As DataTable = objVinoAn.AddNew
        Dim oCont As New Contador
        Dim dtVinoAnCont As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, objVinoAn.GetType.Name, services)
        Dim strVinoAnCont As String
        If Not dtVinoAnCont Is Nothing And dtVinoAnCont.Rows.Count > 0 Then
            strVinoAnCont = dtVinoAnCont.Rows(0)("IdContador")
        End If
        Dim dttVinoAnVar As DataTable = objVinoAnVar.AddNew

        Dim strIDFossIN(-1) As String

        Dim dblFormDE As Double
        Dim dblFormG As Double
        Dim dblFormAZ As Double

        Dim dblFormMV As Double
        Dim dblFormEST As Double
        Dim dblFormENR As Double
        Dim strIDAnalisis As String

        Dim fil As New Filter
        fil.Add("Recuperado", FilterOperator.Equal, False)
        fil.Add("Validado", FilterOperator.Equal, True)
        For Each drData As DataRow In data.DttData.Select(fil.Compose(New AdoFilterComposer))
            Dim delete As Boolean = False
            Dim drVino As DataRow
            If Length(Nz(drData("IDVinoAnalisis").ToString, "")) = 0 Then
                drVino = dttVinoAnalisis.NewRow
                drVino("IDVinoAnalisis") = Guid.NewGuid
                If Len(strVinoAnCont) > 0 Then
                    drVino("NVinoAnalisis") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, strVinoAnCont, services)
                    drVino("IdContador") = strVinoAnCont
                End If
                drVino("IDVino") = drData("IDVino")
                drVino("IDAnalisis") = drData("Analisis")
                'drVino("Fecha") = DateValue(drData("Fecha"))
                drVino("Fecha") = drData("Fecha")
                drVino("CodigoFoss") = drData("CodigoFoss")
            Else
                delete = True
                drVino = objVinoAn.GetItemRow(drData("IDVinoAnalisis"))
            End If

            If Not IsNothing(drVino) Then
                e.Clear()
                strIDAnalisis = drData("Analisis")
                e.Add("IDAnalisis", drData("Analisis"))
                If data.Vino Then
                    e.Add(New IsNullFilterItem("CodigoFoss", False))
                Else
                    e.Add(New IsNullFilterItem("ColFossVino", False))
                End If

                Dim dtAnalisis As DataTable = AdminData.GetData(AnalisisPredeterminado, e, , "Orden")

                e.Clear()
                e.Add("IDAnalisis", drData("Analisis"))
                Dim dtCodNull As DataTable
                If data.Vino Then
                    dtCodNull = AdminData.GetData(VarCodNull, e)
                Else : dtCodNull = AdminData.GetData(VarCodNullDep, e)
                End If


                Dim drFoss As DataRow = New BdgFossVino().GetItemRow(drData("IDFossVino"))
                Dim i As Integer = 1

                dblFormDE = 0 : dblFormG = 0 : dblFormAZ = 0 : dblFormMV = 0 : dblFormEST = 0 : dblFormENR = 0


                For Each drAnalisis As DataRow In dtAnalisis.Rows
                    If drFoss("CodVar" & i) <> cnFormulaMV And _
                        drFoss("CodVar" & i) <> cnFormulaEST And _
                        drFoss("CodVar" & i) <> cnFormulaENR Then

                        Dim drVar As DataRow = dttVinoAnVar.NewRow

                        drVar("IDVinoAnalisis") = drVino("IDVinoAnalisis")
                        drVar("IDVariable") = drFoss("CodVar" & i)
                        If Length(drFoss("Var" & i) & String.Empty) > 0 Then
                            ''drVar("Valor") = xRound(drFoss("Var" & i), 4)
                            ''drVar("ValorNumerico") = xRound(drFoss("Var" & i), 4)

                            drVar("Valor") = xRound(drFoss("Var" & i), drAnalisis("NDecimales"))
                            drVar("ValorNumerico") = xRound(drFoss("Var" & i), drAnalisis("NDecimales"))

                        Else
                            drVar("Valor") = System.DBNull.Value
                            drVar("ValorNumerico") = System.DBNull.Value
                        End If
                        drVar("Orden") = drAnalisis("Orden")

                        Select Case drVar("IDVariable")
                            Case cnFormulaDE
                                dblFormDE = Nz(drFoss("Var" & i), 0)
                            Case cnFormulaG
                                dblFormG = Nz(drFoss("Var" & i), 0)
                            Case cnFormulaAZ
                                dblFormAZ = Nz(drFoss("Var" & i), 0)
                        End Select
                        dttVinoAnVar.Rows.Add(drVar)
                    End If
                    i += 1
                Next

                If Not IsNothing(dtCodNull) AndAlso dtCodNull.Rows.Count Then
                    For Each drCodNull As DataRow In dtCodNull.Rows

                        If drCodNull("IDVariable") <> cnFormulaMV And _
                            drCodNull("IDVariable") <> cnFormulaEST And _
                            drCodNull("IDVariable") <> cnFormulaENR Then

                            Dim drVar As DataRow = dttVinoAnVar.NewRow

                            drVar("IDVinoAnalisis") = drVino("IDVinoAnalisis")
                            drVar("IDVariable") = drCodNull("IDVariable")
                            drVar("Valor") = System.DBNull.Value
                            drVar("ValorNumerico") = System.DBNull.Value
                            drVar("Orden") = drCodNull("Orden")
                            dttVinoAnVar.Rows.Add(drVar)
                        End If
                    Next
                End If
                Dim drVar2 As DataRow = dttVinoAnVar.NewRow

                eVar.Clear()
                eVar.Add("IDAnalisis", strIDAnalisis)
                eVar.Add("IDVAriable", cnFormulaMV)
                dttOrdenAux = AdminData.GetData(AnalisisPredeterminado, eVar)
                If Not dttOrdenAux Is Nothing AndAlso dttOrdenAux.Rows.Count > 0 Then
                    drVar2("IDVinoAnalisis") = drVino("IDVinoAnalisis")
                    drVar2("IDVariable") = cnFormulaMV
                    drVar2("Valor") = xRound(dblFormDE / 1.0018, dttOrdenAux.Rows(0)("NDecimales"))
                    drVar2("ValorNumerico") = xRound(dblFormDE / 1.0018, dttOrdenAux.Rows(0)("NDecimales"))
                    drVar2("Orden") = dttOrdenAux.Rows(0)("Orden")
                    dttVinoAnVar.Rows.Add(drVar2)
                    dblFormMV = xRound(dblFormDE / 1.0018, dttOrdenAux.Rows(0)("NDecimales"))
                End If

                Dim drVar3 As DataRow = dttVinoAnVar.NewRow
                eVar.Clear()
                eVar.Add("IDAnalisis", strIDAnalisis)
                eVar.Add("IDVAriable", cnFormulaEST)
                dttOrdenAux = AdminData.GetData(AnalisisPredeterminado, eVar)
                If Not dttOrdenAux Is Nothing AndAlso dttOrdenAux.Rows.Count > 0 Then
                    drVar3("IDVinoAnalisis") = drVino("IDVinoAnalisis")
                    drVar3("IDVariable") = cnFormulaEST
                    drVar3("Valor") = xRound((2589.8 * dblFormMV) - (0.026 * (dblFormG * dblFormG)) + (3.64 * dblFormG) - 2584.2, dttOrdenAux.Rows(0)("NDecimales"))
                    drVar3("ValorNumerico") = xRound((2589.8 * dblFormMV) - (0.026 * (dblFormG * dblFormG)) + (3.64 * dblFormG) - 2584.2, dttOrdenAux.Rows(0)("NDecimales"))
                    drVar3("Orden") = dttOrdenAux.Rows(0)("Orden")
                    dttVinoAnVar.Rows.Add(drVar3)
                    dblFormEST = xRound((2589.8 * dblFormMV) - (0.026 * (dblFormG * dblFormG)) + (3.64 * dblFormG) - 2584.2, dttOrdenAux.Rows(0)("NDecimales"))
                End If

                Dim drVar4 As DataRow = dttVinoAnVar.NewRow
                eVar.Clear()
                eVar.Add("IDAnalisis", strIDAnalisis)
                eVar.Add("IDVAriable", cnFormulaENR)
                dttOrdenAux = AdminData.GetData(AnalisisPredeterminado, eVar)
                If Not dttOrdenAux Is Nothing AndAlso dttOrdenAux.Rows.Count > 0 Then
                    drVar4("IDVinoAnalisis") = drVino("IDVinoAnalisis")
                    drVar4("IDVariable") = cnFormulaENR
                    drVar4("Valor") = xRound(dblFormEST - dblFormAZ, dttOrdenAux.Rows(0)("NDecimales"))
                    drVar4("ValorNumerico") = xRound(dblFormEST - dblFormAZ, dttOrdenAux.Rows(0)("NDecimales"))
                    drVar4("Orden") = dttOrdenAux.Rows(0)("Orden")
                    dttVinoAnVar.Rows.Add(drVar4)
                End If
                If delete = True Then
                    eDelete.Add("IDVinoAnalisis", drVino("IDVinoAnalisis"))
                Else : dttVinoAnalisis.Rows.Add(drVino)
                End If
                ReDim Preserve strIDFossIN(strIDFossIN.Length)
                strIDFossIN(strIDFossIN.Length - 1) = drFoss("IDFossVino")
            End If
        Next
        If eDelete.Count > 0 Then
            Dim dttDelete As DataTable = objVinoAnVar.Filter(eDelete)
            If Not IsNothing(dttDelete) AndAlso dttDelete.Rows.Count > 0 Then
                objVinoAnVar.Delete(dttDelete)
            End If
        End If
        If Not IsNothing(dttVinoAnalisis) AndAlso dttVinoAnalisis.Rows.Count > 0 Then
            BusinessHelper.UpdateTable(dttVinoAnalisis)
            blnReturn = True
        End If
        If Not IsNothing(dttVinoAnVar) AndAlso dttVinoAnVar.Rows.Count > 0 Then
            BusinessHelper.UpdateTable(dttVinoAnVar)
            blnReturn = True
        End If

        If strIDFossIN.Length > 0 Then
            e.Clear()
            Dim dttFoss As DataTable = New BdgFossVino().Filter(New InListFilterItem("IDFossVino", strIDFossIN, FilterType.String))
            If Not IsNothing(dttFoss) AndAlso dttFoss.Rows.Count > 0 Then
                For Each drFoss As DataRow In dttFoss.Rows
                    drFoss("Recuperado") = True
                Next
                BusinessHelper.UpdateTable(dttFoss)
                blnReturn = True
            End If
        End If
        Return blnReturn
    End Function

#End Region

#Region "Otros..."

    <Serializable()> _
    Public Class StDevolverArticulos
        Public IDDeposito As String
        Public Fecha As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDDeposito As String, ByVal Fecha As Date)
            Me.IDDeposito = IDDeposito
            Me.Fecha = Fecha
        End Sub
    End Class

    <Task()> Public Shared Function DevolverArticulos(ByVal data As StDevolverArticulos, ByVal services As ServiceProvider) As String()
        Dim strCommand As String = "spBdgArticuloFecha " & Quoted(data.IDDeposito) & ", " & data.Fecha.ToString("yyyyMMdd")
        Dim dt As DataTable = AdminData.GetData(strCommand, False)
        Dim rslt(dt.Rows.Count - 1) As String
        For i As Integer = 0 To rslt.Length - 1
            rslt(i) = dt.Rows(i)("IDArticulo")
        Next
        Return rslt
    End Function

    <Serializable()> _
    Public Class StDevolverDepositos
        Public IDArticulo As String
        Public Fecha As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal Fecha As Date)
            Me.IDArticulo = IDArticulo
            Me.Fecha = Fecha
        End Sub
    End Class

    <Task()> Public Shared Function DevolverDepositos(ByVal data As StDevolverDepositos, ByVal services As ServiceProvider) As String()
        Dim strCommand As String = "spBdgDepositoFecha " & Quoted(data.IDArticulo) & ", " & data.Fecha.ToString("yyyyMMdd")
        Dim dt As DataTable = AdminData.GetData(strCommand, False)
        Dim rslt(dt.Rows.Count - 1) As String
        For i As Integer = 0 To rslt.Length - 1
            rslt(i) = dt.Rows(i)("IDDeposito")
        Next
        Return rslt
    End Function

    <Serializable()> _
    Public Class StAccessFilterCompose
        Public FechaDesde As String
        Public FechaHasta As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal FechaDesde As String, ByVal FechaHasta As String)
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
        End Sub
    End Class

    <Task()> Public Shared Function AccessFilterCompose(ByVal data As StAccessFilterCompose, ByVal services As ServiceProvider) As String
        Dim e As New Filter
        If Length(data.FechaDesde) > 0 Then e.Add("DateTime", FilterOperator.GreaterThanOrEqual, CDate(data.FechaDesde), FilterType.DateTime)
        If Length(data.FechaHasta) > 0 Then e.Add("DateTime", FilterOperator.LessThan, DateAdd(DateInterval.Day, 1, CDate(data.FechaHasta)), FilterType.DateTime)
        Return e.Compose(New AdoFilterComposer)
    End Function

    <Task()> Public Shared Function MirarCodigoFossEnVariable(ByVal strCodigoFoss As String, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter("tbBdgVinoAnalisis", New StringFilterItem("CodigoFoss", strCodigoFoss))
    End Function
#End Region

#End Region


#Region " FOSS VINO DESDE PLANTILLA EXCEL "

    <Serializable()> _
   Public Class DataImportarFOSSExcel
        Public PlantillaAnalisis As DataTable
        Public DatosAnalisis As DataTable
        Public EsVino As Boolean

        Public Sub New(ByVal PlantillaAnalisis As DataTable, ByVal DatosAnalisis As DataTable, Optional ByVal EsVino As Boolean = False)
            Me.PlantillaAnalisis = PlantillaAnalisis
            Me.DatosAnalisis = DatosAnalisis
            Me.EsVino = EsVino
        End Sub
    End Class

    <Task()> Public Shared Function ImportarFOSSExcel(ByVal data As DataImportarFOSSExcel, ByVal services As ServiceProvider) As LogProcess
        Dim log As New LogProcess

        Try

            Dim lstFOSS As Dictionary(Of String, DataTable) = ProcessServer.ExecuteTask(Of DataImportarFOSSExcel, Dictionary(Of String, DataTable))(AddressOf BdgFossVino.MergePlantillaFossDatosAnalisis, data, services)
            If lstFOSS Is Nothing Then
                Dim Err As New ClassErrors
                Err.MessageError = AdminData.GetMessageText("Ha ocurrido un error en el proceso de lectura de datos.")
                Err.Elements = ""
                ReDim Preserve log.Errors(log.Errors.Length)
                log.Errors(log.Errors.Length - 1) = Err
            Else
                Dim FossV As New BdgFossVino
                Dim BBDDCurrentID As Guid = AdminData.GetConnectionInfo.IDDataBase
                Dim dtUserDataBase As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Comunes.GetUserDataBases, Nothing, services)
                For Each BBDD As String In lstFOSS.Keys
                    Dim PermisosBBDD As List(Of DataRow) = (From c In dtUserDataBase _
                                                            Where Not c.IsNull("BaseDatos") AndAlso c("BaseDatos") = BBDD).ToList()
                    If Not PermisosBBDD Is Nothing AndAlso PermisosBBDD.Count > 0 Then
                        Dim dt As DataTable = lstFOSS(BBDD)
                        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                            AdminData.SetCurrentConnection(CType(PermisosBBDD(0)("IDBaseDatos"), Guid))
                            FossV.Update(dt)
                        End If
                    End If
                Next

                AdminData.SetCurrentConnection(BBDDCurrentID)

            End If

        Catch ex As Exception
            Dim mensaje As String = AdminData.GetMessageText("Ha ocurrido un error en el proceso de lectura de datos: {0} ")
            mensaje = Engine.ParseFormatString(mensaje, ex.Message)

            Dim Err As New ClassErrors
            Err.MessageError = mensaje
            Err.Elements = ""
            ReDim Preserve log.Errors(log.Errors.Length)
            log.Errors(log.Errors.Length - 1) = Err
        End Try

        Return log

    End Function


    <Serializable()> _
    Public Class DataIdentificador
        Public Empresa As String
        Public IDDeposito As String
        Public IDArticulo As String
        Public Lote As String
    End Class

    <Task()> Public Shared Function GetIdentificador(ByVal Identificador As String, ByVal services As ServiceProvider) As DataIdentificador
        Dim IDFichero() As String = Strings.Split(Identificador, SepSampleID)
        Dim ID As DataIdentificador
        If Not IDFichero Is Nothing AndAlso IDFichero.Length > 0 Then
            '//ID: Empresa/IDDeposito/IDArticulo/Lote
            ID = New DataIdentificador
            ID.Empresa = IDFichero(0)
            ID.IDDeposito = IDFichero(1)
            ID.IDArticulo = IDFichero(2)
            ID.Lote = IDFichero(3)
        End If
        Return ID
    End Function


    <Task()> Public Shared Function MergePlantillaFossDatosAnalisis(ByVal data As DataImportarFOSSExcel, ByVal services As ServiceProvider) As Dictionary(Of String, DataTable)
        Dim FossV As New BdgFossVino
        Dim dt As DataTable
        Dim lstFOSS As New Dictionary(Of String, DataTable)
        If Not data.PlantillaAnalisis Is Nothing AndAlso data.PlantillaAnalisis.Rows.Count > 0 Then
            If Not data.DatosAnalisis Is Nothing AndAlso data.DatosAnalisis.Rows.Count > 0 Then
                Dim EmpresaAnt As String
                Dim RegistrosAnalisis As List(Of DataRow) = (From c In data.DatosAnalisis Order By c("ID")).ToList()
                For Each drDatos As DataRow In RegistrosAnalisis
                    If Length(drDatos("ID")) > 0 Then
                        Dim ID As DataIdentificador = ProcessServer.ExecuteTask(Of String, DataIdentificador)(AddressOf GetIdentificador, drDatos("ID"), services)
                        If Not ID Is Nothing Then

                            Dim Fecha As Date = drDatos("Fecha")
                            If EmpresaAnt & String.Empty <> ID.Empresa & String.Empty Then
                                dt = FossV.AddNew
                                EmpresaAnt = ID.Empresa
                                Dim IDEmpresa As String = AdminData.GetConnectionInfo.DataBase
                                If Length(ID.Empresa) > 0 Then IDEmpresa = ID.Empresa
                                lstFOSS.Add(IDEmpresa, dt)
                            End If
                            Dim drNewLine As DataRow = dt.NewRow
                            drNewLine("IDArticulo") = ID.IDArticulo
                            drNewLine("IDDeposito") = ID.IDDeposito
                            drNewLine("Lote") = ID.Lote

                            drNewLine("CodigoFoss") = drDatos("CodigoFoss")
                            drNewLine("Recuperado") = False
                            drNewLine("Fecha") = Fecha
                            drNewLine("Analisis") = data.PlantillaAnalisis.Rows(0)("IDAnalisis")

                            drNewLine("Vino") = data.EsVino

                            Dim ListaVariables As List(Of DataRow) = (From c In data.PlantillaAnalisis _
                                                                      Order By c("Orden")).ToList()

                            Dim i As Integer = 1
                            For Each drAnalisis As DataRow In ListaVariables
                                Dim IDVariableAnalizar As String = drAnalisis("IDVariable") 'Nz(drAnalisis("CodigoFoss"), drAnalisis("IDVariable"))
                                If i > 60 Then Exit For
                                drNewLine("CodVar" & i) = IDVariableAnalizar
                                drNewLine("Var" & i) = System.DBNull.Value
                                If data.DatosAnalisis.Columns.Contains(IDVariableAnalizar) Then
                                    If Length(Trim(drDatos(IDVariableAnalizar))) > 0 Then
                                        If IsNumeric(Trim(drDatos(IDVariableAnalizar))) Then
                                            drNewLine("Var" & i) = xRound(Trim(drDatos(IDVariableAnalizar)).Replace(".", ","), Nz(drAnalisis("NDecimales"), 0))
                                        Else
                                            drNewLine("Var" & i) = Trim(drDatos(IDVariableAnalizar))
                                        End If
                                    End If
                                End If
                                i += 1
                            Next
                            dt.Rows.Add(drNewLine)


                        End If
                    End If

                Next
            End If
        End If
        Return lstFOSS
    End Function

#End Region

End Class