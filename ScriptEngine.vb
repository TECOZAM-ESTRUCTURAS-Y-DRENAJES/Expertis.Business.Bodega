Public Class ScriptEngine
    Private MyAssembly As Reflection.Assembly

    Public Sub Validate(ByVal strIDTarifa As String, ByVal strFunctionBody As String)

        Dim strClassName As String = GetValidIdentifier(strIDTarifa)
        Dim strCode As String = GetAssemblyCode()
        strCode = strCode & GenerateClassCode(strClassName, strFunctionBody, GenerateVarsCode())
        Dim oAsmb As Reflection.Assembly = Compile(strCode)
        'Dim MyCode As BdgFormula = oAsmb.CreateInstance(strClassName)
        'Return MyCode.Eval

    End Sub

    Private Function GetAssemblyCode() As String
        Dim str As New System.Text.StringBuilder
        str.Append("Option Explicit On")
        str.Append(vbCrLf)
        str.Append("Option Strict Off")
        str.Append(vbCrLf)
        str.Append("Option Compare Text")
        str.Append(vbCrLf)
        str.Append("Imports System")
        str.Append(vbCrLf)
        str.Append("Imports System.Math")
        str.Append(vbCrLf)
        str.Append("Imports Microsoft.VisualBasic")
        str.Append(vbCrLf)
        Return str.ToString
    End Function

    Private Function GenerateClassCode(ByVal strClassName As String, ByVal ClassBody As String, ByVal CommonCode As String) As String
        Dim str As New System.Text.StringBuilder
        str.Append("Public Class " & strClassName)
        str.Append(vbCrLf)
        str.Append("Inherits " & GetType(BdgFormula).FullName)
        str.Append(vbCrLf)

        str.Append(CommonCode)
        str.Append(vbCrLf)

        str.Append("Public Overrides Function Eval() as Double")
        str.Append(vbCrLf)

        str.Append(ClassBody)
        str.Append(vbCrLf)

        str.Append("End Function")
        str.Append(vbCrLf)
        str.Append("End Class")
        str.Append(vbCrLf)
        Return str.ToString
    End Function

    Private Function GenerateVarsCode() As String
        Dim oVar As BdgVariable = New BdgVariable
        Dim dtV As DataTable = oVar.Filter(New NumberFilterItem(_Vr.TipoVariable, FilterOperator.Equal, BdgTipoVariable.Numerica))
        Dim str As New System.Text.StringBuilder
        For Each oRw As DataRow In dtV.Rows
            str.Append("Private Function " & GetValidIdentifier(oRw(_Vr.IDVariable)) & "() as Double")
            str.Append(vbCrLf)
            str.Append("Return GetVar(""" & oRw(_Vr.IDVariable) & """)")
            str.Append(vbCrLf)
            str.Append("End Function")
            str.Append(vbCrLf)
        Next
        Return str.ToString
    End Function

    Private Function GetValidIdentifier(ByVal strName As String) As String
        If Len(strName) Then
            If Char.IsDigit(strName.Chars(0)) Then
                strName = "_" & strName
            End If
            strName = strName.Replace(".", Nothing)
            strName = strName.Replace(" ", Nothing)
            strName = strName.Replace("-", Nothing)
            Return strName
        End If
    End Function

    Private Function Compile(ByVal Code As String) As Reflection.Assembly
        Dim VBCodeProvider As New Microsoft.VisualBasic.VBCodeProvider
        Dim iCompiler As CodeDom.Compiler.ICodeCompiler
        Dim oParams As New CodeDom.Compiler.CompilerParameters(New String() {"System.dll", GetType(BdgFormula).Module.Name})
        Dim oResults As CodeDom.Compiler.CompilerResults

        iCompiler = VBCodeProvider.CreateCompiler
        oParams.GenerateInMemory = True
        oParams.IncludeDebugInformation = False

        oResults = iCompiler.CompileAssemblyFromSource(oParams, Code)
        If oResults.Errors.Count > 0 Then
            Throw New Exception(oResults.Errors(0).ErrorText)
        Else
            Return oResults.CompiledAssembly
        End If
    End Function

    Private Sub Initialize()
        Dim oTrf As BdgTarifa = New BdgTarifa
        Dim dtTrf As DataTable = oTrf.Filter(New NumberFilterItem(_T.TipoTarifa, FilterOperator.Equal, BdgTarifa.enumBdgTipoTarifa.Formula))
        Dim strCommonCode As String = GenerateVarsCode()
        Dim strCode As New System.Text.StringBuilder
        strCode.Append(GetAssemblyCode())
        For Each oRw As DataRow In dtTrf.Rows
            If Not oRw.IsNull(_T.TextoFormula) Then
                strCode.Append(GenerateClassCode(GetValidIdentifier(oRw(_T.IDTarifa)), oRw(_T.TextoFormula), strCommonCode))
            End If
        Next
        MyAssembly = Compile(strCode.ToString)
    End Sub

    Public Function Eval(ByVal IDTarifa As String, _
                        ByVal IDEntrada As Integer, _
                        ByVal Entrada As DataRow, _
                        ByVal VarData As DataView) As Double
        If MyAssembly Is Nothing Then Initialize()
        Dim MiClass As BdgFormula = MyAssembly.CreateInstance(GetValidIdentifier(IDTarifa))
        MiClass.IDEntrada = IDEntrada
        MiClass.Entrada = Entrada
        MiClass.VarData = VarData
        Return MiClass.Eval
    End Function
End Class

Public MustInherit Class BdgFormula
    Public IDEntrada As Integer
    Public MustOverride Function Eval() As Double
    Public VarData As DataView
    Public Entrada As DataRow
    Protected Function GetVar(ByVal IDVar As String) As Double
        If Not VarData Is Nothing Then
            Dim i As Integer = VarData.Find(New Object() {IDEntrada, IDVar})
            If i >= 0 Then
                Return VarData(i)(_EA.ValorNumerico)
            End If
        End If
    End Function

    Protected Function Vendimia() As Integer
        Return Entrada(_E.Vendimia)
    End Function
    Protected Function IDCartillista() As String
        Return Entrada(_E.IDCartillista)
    End Function
    Protected Function NEntrada() As Integer
        Return Entrada(_E.NEntrada)
    End Function
    Protected Function Fecha() As Date
        Return Entrada(_E.Fecha)
    End Function
    Protected Function Hora() As Date
        Return Entrada(_E.Hora)
    End Function
    Protected Function IDMunicipio() As String
        Return Entrada(_E.IDMunicipio)
    End Function
    Protected Function IDVariedad() As String
        Return Entrada(_E.IDVariedad)
    End Function
    Protected Function Bruto() As Double
        Return Entrada(_E.Bruto)
    End Function
    Protected Function BrutoB() As Double
        Return Entrada(_E.BrutoB)
    End Function
    Protected Function Tara() As Double
        Return Entrada(_E.Tara)
    End Function
    Protected Function Neto() As Double
        Return Entrada(_E.Neto)
    End Function
    Protected Function Declarado() As Double
        Return Entrada(_E.Declarado)
    End Function
    Protected Function TipoVariedad() As Integer
        Return Entrada(_E.TipoVariedad)
    End Function

End Class