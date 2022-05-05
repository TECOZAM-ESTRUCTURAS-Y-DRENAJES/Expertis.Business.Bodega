Public Class ProcesoBdgOperacionPlan

#Region "Tareas Update BdgOperacionPlan"

    <Task()> Public Shared Function CrearDocumento(ByVal data As UpdatePackage, ByVal services As ServiceProvider) As DocumentoBdgOperacionPlan
        Return New DocumentoBdgOperacionPlan(data)
    End Function

#Region "Tareas Cabecera"


    <Task()> Public Shared Sub AsignarContadorCabecera(ByVal data As DocumentoBdgOperacionPlan, ByVal services As ServiceProvider)
        If data.HeaderRow.RowState = DataRowState.Added Then
            If Length(data.HeaderRow("IDContador")) > 0 Then
                Dim NOperacionPlanProvis As String = data.HeaderRow("NOperacionPlan") & String.Empty

                data.HeaderRow("NOperacionPlan") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data.HeaderRow("IDContador"), services)

                Dim datActualizar As New DataActualizarNOperacionRestoDocumento(data, NOperacionPlanProvis, data.HeaderRow("NOperacionPlan"))
                ProcessServer.ExecuteTask(Of DataActualizarNOperacionRestoDocumento)(AddressOf ActualizarNOperacionRestoDocumento, datActualizar, services)
            End If
        End If
    End Sub


    Public Class DataActualizarNOperacionRestoDocumento
        Public Doc As DocumentoBdgOperacionPlan
        Public NOperacionProvisional As String
        Public NOperacionDefinitiva As String

        Public Sub New(ByVal Doc As DocumentoBdgOperacionPlan, ByVal NOperacionProvisional As String, ByVal NOperacionDefinitiva As String)
            Me.Doc = Doc
            Me.NOperacionProvisional = NOperacionProvisional
            Me.NOperacionDefinitiva = NOperacionDefinitiva
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarNOperacionRestoDocumento(ByVal data As DataActualizarNOperacionRestoDocumento, ByVal services As ServiceProvider)
        '//Actualizamos el NOperacionPlan en el resto de entidades, si se ha cambiado el contador propuesto por un contador fijo
        If data.NOperacionProvisional <> data.NOperacionDefinitiva Then

            Dim datActualizar As New BdgGeneral.DataActualizarPKCabecera(data.Doc.dtOperacionVino, data.NOperacionProvisional, data.NOperacionDefinitiva, "NOperacionPlan")
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            '//

            datActualizar.dtActualizar = data.Doc.dtOperacionMaterial
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionMOD
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionCentro
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionVarios
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            '//

            datActualizar.dtActualizar = data.Doc.dtOperacionVinoMaterial
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionVinoMOD
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionVinoCentro
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionVinoVarios
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

        End If
    End Sub


#End Region

#Region "Tareas Vino Origen / Destino"

    <Task()> Public Shared Sub VinoPlanOrigenDestino(ByVal data As DocumentoBdgOperacionPlan, ByVal services As ServiceProvider)
        For Each DrVinoPlan As DataRow In data.dtOperacionVino.Select
            If DrVinoPlan.RowState = DataRowState.Added Then
                If Length(DrVinoPlan("IDLineaOperacionVinoPlan")) = 0 Then DrVinoPlan("IDLineaOperacionVinoPlan") = Guid.NewGuid
            End If
        Next
    End Sub

#End Region

#Region "Tareas Imputaciones Globales"

    <Task()> Public Shared Sub ImputacionesGlobales(ByVal data As DocumentoBdgOperacionPlan, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionPlan)(AddressOf ProcesoBdgOperacionPlan.TareasMaterialesGlobales, data, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionPlan)(AddressOf ProcesoBdgOperacionPlan.TareasCentrosGlobales, data, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionPlan)(AddressOf ProcesoBdgOperacionPlan.TareasMODGlobales, data, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionPlan)(AddressOf ProcesoBdgOperacionPlan.TareasVariosGlobales, data, services)
    End Sub

#Region "Tareas Materiales Globales"

    <Task()> Public Shared Sub TareasMaterialesGlobales(ByVal data As DocumentoBdgOperacionPlan, ByVal services As ServiceProvider)
        For Each DrMatGlobal As DataRow In data.dtOperacionMaterial.Select
            If DrMatGlobal.RowState = DataRowState.Added Then
                If Length(DrMatGlobal("IDOperacionPlanMaterial")) = 0 Then DrMatGlobal("IDOperacionPlanMaterial") = Guid.NewGuid
            End If
        Next
    End Sub

#End Region

#Region "Tareas Centros Globales"

    <Task()> Public Shared Sub TareasCentrosGlobales(ByVal data As DocumentoBdgOperacionPlan, ByVal services As ServiceProvider)
        For Each DrCentroGlobal As DataRow In data.dtOperacionCentro.Select
            If DrCentroGlobal.RowState = DataRowState.Added Then
                If Length(DrCentroGlobal("IDOperacionPlanCentro")) = 0 Then DrCentroGlobal("IDOperacionPlanCentro") = Guid.NewGuid
            End If
        Next
    End Sub

#End Region

#Region "Tareas MOD Globales"

    <Task()> Public Shared Sub TareasMODGlobales(ByVal data As DocumentoBdgOperacionPlan, ByVal services As ServiceProvider)
        For Each DrMODGlobal As DataRow In data.dtOperacionMOD.Select
            If DrMODGlobal.RowState = DataRowState.Added Then
                If Length(DrMODGlobal("IDOperacionPlanMOD")) = 0 Then DrMODGlobal("IDOperacionPlanMOD") = Guid.NewGuid
            End If
        Next
    End Sub

#End Region

#Region "Tareas Varios Globales"

    <Task()> Public Shared Sub TareasVariosGlobales(ByVal data As DocumentoBdgOperacionPlan, ByVal services As ServiceProvider)
        For Each DrVariosGlobal As DataRow In data.dtOperacionVarios.Select
            If DrVariosGlobal.RowState = DataRowState.Added Then
                If Length(DrVariosGlobal("IDOperacionPlanVarios")) = 0 Then DrVariosGlobal("IDOperacionPlanVarios") = Guid.NewGuid
            End If
        Next
    End Sub

#End Region

#End Region

#Region "Tareas VinoPlan"

    <Task()> Public Shared Sub ValidarOperacionVinoPlan(ByVal data As DocumentoBdgOperacionPlan, ByVal services As ServiceProvider)
        For Each DrVinoPlan As DataRow In data.dtOperacionVino.Select(String.Empty, String.Empty, DataViewRowState.Added Or DataViewRowState.ModifiedCurrent)
            If Length(DrVinoPlan(_OV.Merma)) = 0 Then DrVinoPlan(_OV.Merma) = 0
        Next
    End Sub

#End Region

#End Region

End Class