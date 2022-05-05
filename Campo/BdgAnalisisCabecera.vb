Public Class BdgAnalisisCabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgAnalisisCabecera"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarContador)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_AC.IDAnalisisCabecera)) = 0 Then data(_AC.IDAnalisisCabecera) = Guid.NewGuid
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim Dt As New Contador.DatosDefaultCounterValue(data, _AC.Entidad, _AC.NAnalisisCabecera, _AC.IDContador)
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, Dt, services)
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDAnalisis)
    End Sub

    <Task()> Public Shared Sub ValidarIDAnalisis(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_AC.IDAnalisis)) > 0 Then
            Dim dtAnalisis As DataTable = New BdgAnalisis().SelOnPrimaryKey(data(_AC.IDAnalisis))
            If dtAnalisis Is Nothing OrElse dtAnalisis.Rows.Count = 0 Then
                ApplicationService.GenerateError("No se ha encontrado el análisis")
            End If
        End If
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarNAnalisis)
    End Sub

    <Task()> Public Shared Sub AsignarNAnalisis(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data(_AC.IDContador)) > 0 Then
                data(_AC.NAnalisisCabecera) = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data(_AC.IDContador), services)
            End If
        End If
    End Sub
	
#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DataGenAnalitica
        Public DtData As DataTable
        Public Origen As enumBdgTipoAnalisisCabGen

        Public Sub New()
        End Sub
        Public Sub New(ByVal DtData As DataTable, ByVal Origen As enumBdgTipoAnalisisCabGen)
            Me.DtData = DtData
            Me.Origen = Origen
        End Sub
    End Class

    <Task()> Public Shared Function GenerarAnaliticaFossFinca(ByVal data As DataGenAnalitica, ByVal services As ServiceProvider) As List(Of Guid)
        If Not data.DtData Is Nothing AndAlso data.DtData.Rows.Count > 0 Then
            Dim LstReturn As New List(Of Guid)
            Dim ClsBdgAnCab As New BdgAnalisisCabecera
            Dim ClsBdgVar As New BdgAnalisisLineaValor
            Dim DtBdgAnCab As DataTable = ClsBdgAnCab.AddNew
            Dim DtBDgAnLinVar As DataTable = ClsBdgVar.AddNew

            data.DtData.AcceptChanges()
            Dim LstIDS As New List(Of String)
            Select Case data.Origen
                Case enumBdgTipoAnalisisCabGen.Finca
                    LstIDS = (From Dr As DataRow In data.DtData Select CType(Dr("CFinca"), String)).Distinct.ToList
                Case enumBdgTipoAnalisisCabGen.Observatorio
                    LstIDS = (From Dr As DataRow In data.DtData Select CType(Dr("IDOrigen"), String)).Distinct.ToList
            End Select
            For Each StrID As String In LstIDS
                Dim DrNew As DataRow = DtBdgAnCab.NewRow
                DrNew.ItemArray = ClsBdgAnCab.AddNewForm.Rows(0).ItemArray

                Dim Drlineas As New List(Of DataRow)
                Select Case data.Origen
                    Case enumBdgTipoAnalisisCabGen.Finca
                        Drlineas = (From DrLin As DataRow In data.DtData Select DrLin Where CStr(DrLin("CFinca")) = StrID).ToList
                        DrNew("Valor") = Drlineas(0)("IDOrigen")
                    Case enumBdgTipoAnalisisCabGen.Observatorio
                        Drlineas = (From DrLin As DataRow In data.DtData Select DrLin Where CStr(DrLin("IDOrigen")) = StrID).ToList
                        DrNew("Valor") = Drlineas(0)("IDOrigen")
                End Select
                DrNew("Estado") = enumBdgEstadoAnalisisGen.Solicitado
                DrNew("Fecha") = Drlineas(0)("Fecha")
                DrNew("IDAnalisis") = Drlineas(0)("IDAnalisis")
                DrNew("TipoAnalisis") = data.Origen
                DtBdgAnCab.Rows.Add(DrNew)

                For Each DrLinea As DataRow In Drlineas
                    Dim IntOrden As Integer = 1
                    For Each DcCol As DataColumn In DrLinea.Table.Columns
                        If DcCol.ColumnName.StartsWith("Var/") Then
                            Dim DrNewLin As DataRow = DtBDgAnLinVar.NewRow
                            DrNewLin("IDAnalisisCabecera") = DrNew("IDAnalisisCabecera")
                            Dim StrIDVar As String = DcCol.ColumnName.Substring(DcCol.ColumnName.IndexOf("/") + 1, DcCol.ColumnName.Length - (DcCol.ColumnName.IndexOf("/") + 1))
                            DrNewLin("IDVariable") = StrIDVar
                            If IsNumeric(DrLinea(StrIDVar)) Then
                                DrNewLin("Valor") = DrLinea(StrIDVar)
                                DrNewLin("ValorNumerico") = DrLinea(StrIDVar)
                            Else : DrNewLin("Valor") = DrLinea(StrIDVar)
                            End If
                            DrNewLin("Orden") = IntOrden
                            DtBDgAnLinVar.Rows.Add(DrNewLin)
                            IntOrden += 1
                        End If
                    Next
                Next
                LstReturn.Add(DrNew("IDAnalisisCabecera"))
            Next
            ClsBdgAnCab.Validate(DtBdgAnCab)
            ClsBdgAnCab.Update(DtBdgAnCab)
            ClsBdgVar.Validate(DtBDgAnLinVar)
            ClsBdgVar.Update(DtBDgAnLinVar)
            Return LstReturn
        End If
    End Function

    <Serializable()> _
    Public Class StGetGAnalisis
        Public IDAnalisis As String
        Public Filtro As Filter

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDAnalisis As String, ByVal Filtro As Filter)
            Me.IDAnalisis = IDAnalisis
            Me.Filtro = Filtro
        End Sub
    End Class

    <Task()> Public Shared Function GetAnalisis(ByVal data As StGetGAnalisis, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable = New DataEngine().Filter("frmBdgCIAnalisisFincaObservatorio", data.Filtro)
        Dim htColumas As Hashtable
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
        strSelect += "0"

        'ahora las actualizamos
        Dim bsnBFA As New BdgAnalisisLineaValor
        For Each dr As DataRow In dt.Select
            Dim f As New Filter
            f.Add("IDAnalisisCabecera", dr("IDAnalisisCabecera"))
            f.Add(fVariables)
            Dim dtResult As DataTable = bsnBFA.Filter(f)
            For Each drResult As DataRow In dtResult.Select()
                dr(drResult("IDVariable")) = drResult("Valor")
                dr(drResult("IDVariable") & "_N") = drResult("ValorNumerico")
            Next
        Next
        Return dt
    End Function

#End Region

End Class

<Serializable()> _
Public Class _AC
    Public Const Entidad As String = "BdgAnalisisCabecera"
    Public Const IDAnalisisCabecera As String = "IDAnalisisCabecera"
    Public Const NAnalisisCabecera As String = "NAnalisisCabecera"
    Public Const IDContador As String = "IDContador"
    Public Const IDAnalisis As String = "IDAnalisis"
    Public Const Fecha As String = "Fecha"
End Class