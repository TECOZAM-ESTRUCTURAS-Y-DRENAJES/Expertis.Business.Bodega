Public Class ProcesoCartillista

#Region "Creación de Documentos"

    <Task()> Public Shared Function CrearDocumento(ByVal data As UpdatePackage, ByVal services As ServiceProvider) As DocumentoCartillista
        Return New DocumentoCartillista(data)
    End Function

#End Region

#Region "Eventos Update"

#Region "Eventos Cabecera Cartillista - BdgCartillista"

    <Task()> Public Shared Sub ActualizarProveedor(ByVal data As DocumentoCartillista, ByVal services As ServiceProvider)
        Dim ClsProv As New BdgProveedor
        Dim dtP As DataTable = ClsProv.SelOnPrimaryKey(data.HeaderRow(_C.IDProveedor))
        If dtP.Rows.Count = 0 Then
            Dim rwP As DataRow = dtP.NewRow
            rwP(_P.IDProveedor) = data.HeaderRow(_C.IDProveedor)
            dtP.Rows.Add(rwP)
            ClsProv.Update(dtP)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarCupos(ByVal data As DocumentoCartillista, ByVal services As ServiceProvider)
        Dim dblHaT, dblHaB, dblMaxT, dblMaxB As Double
        If Not data.HeaderRow.RowState = DataRowState.Deleted Then
            If Length(data.HeaderRow("CupoCartillaPor")) > 0 Then
                Dim IntVendimia As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf BdgVendimia.UltimaVendimia, New Object, services)
                If data.HeaderRow("CupoCartillaPor") = BdgCupoCartillaPor.Municipio Then
                    Dim DrSel() As DataRow = data.dtMunicipios.Select("Vendimia = " & IntVendimia)
                    For Each Dr As DataRow In DrSel
                        dblHaT += Dr(_CV.HaT)
                        dblHaB += Dr(_CV.HaB)
                        dblMaxT += Dr(_CV.MaxT)
                        dblMaxB += Dr(_CV.MaxB)
                    Next
                ElseIf data.HeaderRow("CupoCartillaPor") = BdgCupoCartillaPor.Finca Then
                    Dim DrSelFinca() As DataRow = data.dtFincas.Select("Vendimia = " & IntVendimia)
                    If DrSelFinca.Length > 0 Then
                        For Each DrSel As DataRow In DrSelFinca
                            Dim DtFinca As DataTable = New BdgFinca().Filter(New FilterItem("IDFinca", FilterOperator.Equal, DrSel("IDFinca")))
                            Dim DtVar As DataTable = New BdgVariedad().Filter(New FilterItem("IDVariedad", FilterOperator.Equal, DtFinca.Rows(0)("IDVariedad")))
                            If DtVar.Rows(0)("TipoVariedad") Then
                                dblHaB += DtFinca.Rows(0)("Superficie")
                                dblMaxB += DrSel("Maximo")
                            Else
                                dblHaT += DtFinca.Rows(0)("Superficie")
                                dblMaxT += DrSel("Maximo")
                            End If
                        Next
                    End If
                ElseIf data.HeaderRow("CupoCartillaPor") = BdgCupoCartillaPor.Variedad Then
                    Dim DrSelVar() As DataRow = data.dtVariedades.Select("Vendimia = " & IntVendimia)
                    If DrSelVar.Length > 0 Then
                        For Each DrSel As DataRow In DrSelVar
                            Dim DtVar As DataTable = New BdgVariedad().Filter(New FilterItem("IDVariedad", FilterOperator.Equal, DrSel("IDVariedad")))
                            If DtVar.Rows(0)("TipoVariedad") = BdgTipoVariedad.Blanca Then
                                dblMaxB += DrSel("Maximo")
                            ElseIf DtVar.Rows(0)("TipoVariedad") = BdgTipoVariedad.Tinta Then
                                dblMaxT += DrSel("Maximo")
                            End If
                        Next
                    End If
                End If
                If Not data.HeaderRow("CupoCartillaPor") = BdgCupoCartillaPor.Manual Then
                    Dim DrSelVen() As DataRow = data.dtVendimias.Select("Vendimia = " & IntVendimia)
                    If DrSelVen.Length > 0 Then
                        DrSelVen(0)(_CV.HaB) = dblHaB
                        DrSelVen(0)(_CV.HaT) = dblHaT
                        DrSelVen(0)(_CV.MaxB) = dblMaxB
                        DrSelVen(0)(_CV.MaxT) = dblMaxT
                    Else
                        'Dim DrNew As DataRow = data.dtVendimias.NewRow
                        'DrNew(_CV.IDCartillista) = data.HeaderRow(_C.IDCartillista)
                        'DrNew(_CV.Vendimia) = IntVendimia
                        'DrNew(_CV.HaB) = dblHaB
                        'DrNew(_CV.HaT) = dblHaT
                        'DrNew(_CV.MaxB) = dblMaxB
                        'DrNew(_CV.MaxT) = dblMaxT
                        'data.dtVendimias.Rows.Add(DrNew)
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region "Eventos Fincas - BdgCartillistaFinca"

    <Task()> Public Shared Sub AsignarDatosFincas(ByVal data As DocumentoCartillista, ByVal services As ServiceProvider)
        For Each DrFinca As DataRow In data.dtFincas.Select
            If DrFinca.RowState = DataRowState.Added OrElse DrFinca.RowState = DataRowState.Modified Then
                Dim blnVendimiaPorDefecto As Boolean
                Dim IntVendimia As Integer
                Dim TipoVariedad As BdgTipoVariedad
                Dim dblRdtoxHaT As Double
                Dim dblRdtoxHaB As Double
                Dim dblRdto As Double
                Dim dblSuperficie As Double
                Dim dblIncMaximo As Double
                Dim dblIncHaT As Double
                Dim dblIncHaB As Double
                Dim dblIncMaxT As Double
                Dim dblIncMaxB As Double
                Dim strIDMunicipio As String
                Dim strIDCartillista As String = DrFinca("IDCartillista")

                blnVendimiaPorDefecto = False
                IntVendimia = 0
                dblSuperficie = 0
                dblIncMaximo = 0
                TipoVariedad = -1
                strIDMunicipio = Nothing

                Dim IDFinca As Guid = DrFinca("IDFinca")
                If DrFinca.RowState = DataRowState.Added Then
                    blnVendimiaPorDefecto = True
                    If Length(DrFinca("Vendimia")) > 0 Then blnVendimiaPorDefecto = False
                    If blnVendimiaPorDefecto Then
                        IntVendimia = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf BdgVendimia.UltimaVendimia, New Object, services)
                        DrFinca("Vendimia") = IntVendimia
                    Else : IntVendimia = DrFinca("Vendimia")
                    End If
                ElseIf DrFinca.RowState = DataRowState.Modified Then
                    IntVendimia = DrFinca("Vendimia")
                End If

                Dim DrVendimia As DataRow = New BdgVendimia().GetItemRow(IntVendimia)
                dblRdtoxHaT = DrVendimia("RdtoxHaT")
                dblRdtoxHaB = DrVendimia("RdtoxHaB")

                'calculo de los rendimientos
                Dim StCalc As New BdgCartillistaFinca.StDatosCalculoRendDesdeFinca(IDFinca, IntVendimia)
                Dim udtDatosRdto As New BdgCartillistaFinca.udtBdgDatosCalculoRendimiento
                udtDatosRdto = ProcessServer.ExecuteTask(Of BdgCartillistaFinca.StDatosCalculoRendDesdeFinca, BdgCartillistaFinca.udtBdgDatosCalculoRendimiento)(AddressOf BdgCartillistaFinca.DatosCalculoRendimientoDesdeFinca, StCalc, services)
                If Not udtDatosRdto Is Nothing Then
                    With udtDatosRdto
                        dblSuperficie = .Superficie
                        strIDMunicipio = .MunicipioFinca
                        TipoVariedad = .TipoVariedad
                    End With
                End If
                dblRdto = 0
                If Length(DrFinca("Rdto")) > 0 Then dblRdto = DrFinca("Rdto")
                Select Case TipoVariedad
                    Case BdgTipoVariedad.Tinta
                        If DrFinca.RowState = DataRowState.Added Then
                            dblIncHaT = dblSuperficie
                        ElseIf DrFinca.RowState = DataRowState.Modified Then
                            dblIncHaT = 0
                        End If
                        DrFinca("Maximo") = dblSuperficie * dblRdtoxHaT * dblRdto / 100
                    Case BdgTipoVariedad.Blanca
                        If DrFinca.RowState = DataRowState.Added Then
                            dblIncHaB = dblSuperficie
                        ElseIf DrFinca.RowState = DataRowState.Modified Then
                            dblIncHaB = 0
                        End If
                        DrFinca("Maximo") = dblSuperficie * dblRdtoxHaB * dblRdto / 100
                End Select
                If DrFinca.RowState = DataRowState.Modified Then dblIncMaximo = Nz(DrFinca("Maximo"), 0) - Nz(DrFinca("Maximo", DataRowVersion.Original), 0)
                Select Case TipoVariedad
                    Case BdgTipoVariedad.Tinta
                        dblIncMaxT = dblIncMaximo
                    Case BdgTipoVariedad.Blanca
                        dblIncMaxB = dblIncMaximo
                End Select


                If data.HeaderRow("CupoCartillaPor") = BdgCupoCartillaPor.Finca Then
                    Dim StPrepMun As New BdgCartillista.StPrepararMunicipio(DrFinca("IDCartillista"), IntVendimia, strIDMunicipio, dblIncHaT, dblIncHaB, dblIncMaxT, dblIncMaxB)
                    Dim DtMunicipio As DataTable = ProcessServer.ExecuteTask(Of BdgCartillista.StPrepararMunicipio, DataTable)(AddressOf BdgCartillista.PrepararMunicipio, StPrepMun, services)
                    If Not DtMunicipio Is Nothing Then
                        If DtMunicipio.Rows.Count > 0 Then
                            Dim DrMunicipio As DataRow = DtMunicipio.Rows(0)
                            If DrMunicipio("HaT") <= 0 And DrMunicipio("HaB") <= 0 Then
                                DrMunicipio.Delete()
                            End If
                        End If
                    End If
                    Dim StMunSob As New BdgCartillista.StMunicipiosSobrantes(DrFinca("IDCartillista"), IntVendimia)
                    Dim DtMunicipiosAEliminar As DataTable = ProcessServer.ExecuteTask(Of BdgCartillista.StMunicipiosSobrantes, DataTable)(AddressOf BdgCartillista.MunicipiosSobrantes, StMunSob, services)

                    BusinessHelper.UpdateTable(DtMunicipio)
                    BusinessHelper.UpdateTable(DtMunicipiosAEliminar)
                End If
            End If
        Next
    End Sub

#End Region

#Region "Eventos Municipios - BdgCartillistaMunicipio"

    <Task()> Public Shared Sub AsignarMaximosRdtoMunicipios(ByVal data As DocumentoCartillista, ByVal services As ServiceProvider)
        For Each DrMun As DataRow In data.dtMunicipios.Select
            If DrMun.RowState = DataRowState.Added OrElse DrMun.RowState = DataRowState.Modified Then
                Dim IntVendimia As Integer = 0
                Dim dblRdtoxHaT As Double = 0
                Dim dblRdtoxHaB As Double = 0
                Dim dblHaT As Double = 0
                Dim dblHaB As Double = 0
                Dim dblRdtoT As Double = 0
                Dim dblRdtoB As Double = 0

                Dim dtVendimia As DataTable
                If DrMun.RowState = DataRowState.Added Then
                    If Length(DrMun("Vendimia")) = 0 OrElse DrMun("Vendimia") = 0 Then
                        dtVendimia = New BdgVendimia().Filter("TOP 1 *", , "Vendimia DESC")
                    Else : dtVendimia = New BdgVendimia().SelOnPrimaryKey(DrMun("Vendimia"))
                    End If
                ElseIf DrMun.RowState = DataRowState.Modified Then
                    dtVendimia = New BdgVendimia().SelOnPrimaryKey(DrMun("Vendimia"))
                End If

                If Not dtVendimia Is Nothing AndAlso dtVendimia.Rows.Count > 0 Then
                    IntVendimia = dtVendimia.Rows(0)("Vendimia")
                    dblRdtoxHaT = dtVendimia.Rows(0)("RdtoxHaT")
                    dblRdtoxHaB = dtVendimia.Rows(0)("RdtoxHaB")
                End If

                'calculo de los rendimientos
                dblHaT = 0 : dblHaB = 0
                If Length(DrMun("HaT")) > 0 Then dblHaT = DrMun("HaT")
                If Length(DrMun("HaB")) > 0 Then dblHaB = DrMun("HaB")

                dblRdtoT = 0 : dblRdtoB = 0
                If Length(DrMun("RdtoT")) > 0 Then dblRdtoT = DrMun("RdtoT")
                If Length(DrMun("RdtoB")) > 0 Then dblRdtoB = DrMun("RdtoB")

                DrMun("MaxT") = xRound(dblHaT * dblRdtoxHaT * dblRdtoT / 100)
                DrMun("MaxB") = xRound(dblHaB * dblRdtoxHaB * dblRdtoB / 100)
            End If
        Next
    End Sub

#End Region

#Region "Eventos Variedades - BdgCartillistaVariedad"

    <Task()> Public Shared Sub AsignarVendimiaVariedad(ByVal data As DocumentoCartillista, ByVal services As ServiceProvider)
        For Each DrVar As DataRow In data.dtVariedades.Select
            If DrVar.RowState = DataRowState.Added OrElse DrVar.RowState = DataRowState.Modified Then
                If Length(DrVar("Vendimia")) = 0 Then DrVar("Vendimia") = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf BdgVendimia.UltimaVendimia, New Object, services)
            End If
        Next
    End Sub

#End Region

#End Region


End Class