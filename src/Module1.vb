'Copyright(c) 2016 ClosedXML
'Released under the MIT license
'https://github.com/ClosedXML/ClosedXML/blob/develop/LICENSE

Imports ClosedXML.Excel

Structure TRTResult
    Dim atom, netIntens, background, errorFlag As Integer
    Dim atomPercent, massPercent, sigma As Double
    Dim xLine As String
    Sub New(atom As String,
                  xLine As String,
                  atomPercent As String,
                  massPercent As String,
                  netIntens As String,
                  background As String,
                  sigma As String,
                  errorFlag As String)
        If atom IsNot Nothing Then Me.atom = Integer.Parse(atom)
        If netIntens IsNot Nothing Then Me.netIntens = Integer.Parse(netIntens)
        If background IsNot Nothing Then Me.background = Integer.Parse(background)
        If errorFlag IsNot Nothing Then Me.errorFlag = Integer.Parse(errorFlag)
        If atomPercent IsNot Nothing Then Me.atomPercent = Double.Parse(atomPercent) * 100
        If massPercent IsNot Nothing Then Me.massPercent = Double.Parse(massPercent) * 100
        If sigma IsNot Nothing Then Me.sigma = Double.Parse(sigma)
        If xLine IsNot Nothing Then Me.xLine = xLine
    End Sub
End Structure

Class TRTSpectrum
    Dim atomic_num() As String = {"H", "He", "Li", "Be", "B", "C", "N", "O", "F", "Ne", "Na", "Mg", "Al", "Si", "P", "S", "Cl", "Ar",
        "K", "Ca", "Sc", "Ti", "V", "Cr", "Mn", "Fe", "Co", "Ni", "Cu", "Zn", "Ga", "Ge", "As", "Se", "Br", "Kr",
        "Rb", "Sr", "Y", "Zr", "Nb", "Mo", "Tc", "Ru", "Rh", "Pd", "Ag", "Cd", "In", "Sn", "Sb", "Te", "I", "Xe", "Cs",
        "Ba", "La", "Ce", "Pr", "Nd", "Pm", "Sm", "Eu", "Gd", "Tb", "Dy", "Ho", "Er", "Tm", "Yb", "Lu", "Hf", "Ta", "W", "Re", "Os", "Ir", "Pt", "Au", "Hg", "Tl", "Pb", "Bi", "Po", "At", "Rn", "Fr", "Ra",
        "Ac", "Th", "Pa", "U", "Np", "Pu", "Am", "Cm", "Bk", "Cf", "Es", "Fm", "Md", "No", "Lr", "Rf", "Db", "Sg", "Bh", "Hs", "Mt", "Ds", "Rg", "Cn", "Nh", "Fl", "Mc", "Lv", "Ts", "Og"}
    Public name As String
    Public datetime As DateTime
    Public elements As List(Of TRTResult)
    Public chart As Integer()
    Public Sub New()
        elements = New List(Of TRTResult)
    End Sub
    Public Sub WriteToXlsx(ByRef sheet As IXLWorksheet, ByRef summarySheet As IXLWorksheet)
        'sheet
        sheet.Cell(1, 1).Value = "Results:"
        sheet.Cell(1, 2).Value = name
        sheet.Cell(2, 1).Value = "Date:"
        sheet.Cell(2, 2).Value = datetime.ToString
        sheet.Range(1, 1, 2, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left
        sheet.Cell(4, 1).Value = "Element"
        sheet.Cell(4, 2).Value = "AN"
        sheet.Cell(4, 3).Value = "series"
        sheet.Cell(4, 4).Value = "Net"
        sheet.Cell(4, 5).Value = "[wt.%]"
        sheet.Cell(4, 6).Value = "[norm. wt.%]"
        sheet.Cell(4, 7).Value = "[norm. at.%]"
        sheet.Cell(4, 8).Value = "Error in wt.% (1 Sigma)"
        sheet.Range(4, 1, 4, 8).Style.Fill.BackgroundColor = XLColor.LightGray
        Dim sumMassPercent As Double = 0
        For Each element In elements
            sumMassPercent += element.massPercent
        Next
        Dim a = elements.OrderBy(Function(n) n.atom)
        Dim count As Integer = 0
        For Each element In elements
            sheet.Cell(5 + count, 1).Value = atomic_num(element.atom - 1)
            sheet.Cell(5 + count, 2).Value = element.atom.ToString
            sheet.Cell(5 + count, 3).Value = element.xLine
            sheet.Cell(5 + count, 4).Value = element.netIntens.ToString
            sheet.Cell(5 + count, 5).Value = element.massPercent.ToString
            sheet.Cell(5 + count, 6).Value = (element.massPercent * 100 / sumMassPercent).ToString
            sheet.Cell(5 + count, 7).Value = element.atomPercent.ToString
            sheet.Cell(5 + count, 8).Value = element.sigma.ToString
            count += 1
        Next
        'sum
        sheet.Cell(5 + count, 4).Value = "Sum:"
        sheet.Cell(5 + count, 5).FormulaA1 = "=SUM(E5:E" + (4 + count).ToString + ")"
        sheet.Cell(5 + count, 6).FormulaA1 = "=SUM(F5:F" + (4 + count).ToString + ")"
        sheet.Cell(5 + count, 7).FormulaA1 = "=SUM(G5:G" + (4 + count).ToString + ")"

        'chart
        sheet.Cell(1, 10).Value = "Energy (keV)"
        sheet.Cell(1, 11).Value = "Intensity"
        For i As Integer = 0 To chart.Length - 1
            sheet.Cell(2 + i, 10).Value = 0.005 * i - 0.48
            sheet.Cell(2 + i, 11).Value = chart(i)
        Next

        'style
        sheet.Range(1, 1, 5 + count, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right
        sheet.Range(4, 10, chart.Length + 1, 11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right
        sheet.ColumnsUsed().AdjustToContents()

        'summarySheet
        Dim presentRow As Integer = 4
        While summarySheet.Cell(presentRow, 1).Value <> ""
            presentRow += 1
        End While
        summarySheet.Cell(presentRow, 1).Value = name
        For Each element In elements
            Dim presentColumn As Integer = 2
            Dim tmp As String = summarySheet.Cell(3, presentColumn).Value
            While tmp <> ""
                If Integer.Parse(summarySheet.Cell(3, presentColumn).Value) >= element.atom Then Exit While
                presentColumn += 1
                tmp = summarySheet.Cell(3, presentColumn).Value
            End While
            If tmp = "" Then
                summarySheet.Column(presentColumn).InsertColumnsBefore(1)
                summarySheet.Cell(2, presentColumn).Value = atomic_num(element.atom - 1)
                summarySheet.Cell(3, presentColumn).Value = element.atom.ToString
            ElseIf Integer.Parse(tmp) > element.atom Then
                summarySheet.Column(presentColumn).InsertColumnsBefore(1)
                summarySheet.Cell(2, presentColumn).Value = atomic_num(element.atom - 1)
                summarySheet.Cell(3, presentColumn).Value = element.atom.ToString
            End If
            summarySheet.Cell(presentRow, presentColumn).Value = element.atomPercent.ToString
        Next
        'title
        Dim maxColumn As Integer = 1
        While summarySheet.Cell(3, maxColumn + 1).Value.ToString <> ""
            maxColumn += 1
        End While
        summarySheet.Range(2, 2, 3, maxColumn).Style.Fill.BackgroundColor = XLColor.LightGray
        summarySheet.Cell(1, 2).Value = "Element [norm. at.%]"
        'style
        summarySheet.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right
        summarySheet.ColumnsUsed().AdjustToContents()
    End Sub
End Class

Module Module1
    Sub Main()
        Dim paths As List(Of String) = System.Environment.GetCommandLineArgs.ToList
        paths.RemoveAt(0)
        paths.Sort()
        Console.WriteLine("converting" + paths.Count.ToString + " files to ???.xlsx")

        Dim workbookpath As String = System.AppDomain.CurrentDomain.BaseDirectory + "\converted_" + System.DateTime.Now.ToString("yyMMdd_HHmmss") + ".xlsx"
        Using workbook = New XLWorkbook()
            'compile sheet
            Dim summarySheet As IXLWorksheet = workbook.Worksheets.Add("Summary")
            For Each path As String In paths
                Console.WriteLine(path)
                Dim xDoc As XDocument = XDocument.Load(path)
                Dim data As New TRTSpectrum
                'Name
                Dim eTRTSpectrum As IEnumerable(Of XElement) =
                    From el In xDoc.<TRTSpectrum>.<ClassInstance>
                    Where el.@Type = "TRTSpectrum"
                    Select el
                data.name = eTRTSpectrum.@Name

                'Time
                Dim eTRTSpectrumHeader As IEnumerable(Of XElement) =
                    From el In xDoc.<TRTSpectrum>.<ClassInstance>.<ClassInstance>
                    Where el.@Type = "TRTSpectrumHeader"
                    Select el
                data.datetime = DateTime.ParseExact(eTRTSpectrumHeader.<Date>.Value + " " + eTRTSpectrumHeader.<Time>.Value,
                    "d'.'M'.'yyyy H':'m':'s",
                    System.Globalization.DateTimeFormatInfo.InvariantInfo,
                    System.Globalization.DateTimeStyles.None)

                'Elements
                Dim eTRTResult As IEnumerable(Of XElement) =
                    From el In xDoc.<TRTSpectrum>.<ClassInstance>.<ClassInstance>
                    Where el.@Type = "TRTResult"
                    Select el
                Dim elements = eTRTResult.<Result>
                For Each element In elements
                    Dim elemdata As New TRTResult(element.<Atom>.Value,
                                                  element.<XLine>.Value,
                                                  element.<AtomPercent>.Value,
                                                  element.<MassPercent>.Value,
                                                  element.<NetIntens>.Value,
                                                  element.<Background>.Value,
                                                  element.<Sigma>.Value,
                                                  element.<ErrorFlag>.Value)
                    data.elements.Add(elemdata)
                Next

                'Chart
                Dim eTRTChart As IEnumerable(Of XElement) =
                    From el In xDoc.<TRTSpectrum>.<ClassInstance>.<Channels>
                    Select el
                data.chart = eTRTChart.Value.Split(",").ToList().ConvertAll(Function(str) Int32.Parse(str)).ToArray()

                'writing to sheet
                Dim sheet As IXLWorksheet = workbook.Worksheets.Add(data.name)
                data.WriteToXlsx(sheet, summarySheet)
            Next
            Try
                workbook.SaveAs(workbookpath)
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Using
    End Sub
End Module