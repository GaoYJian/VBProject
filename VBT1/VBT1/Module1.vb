Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.IO.Compression
Imports System.Data.Common
Imports System.Data
Imports Newtonsoft.Json.Linq



Module Module1

    Sub Main()
        'InsertPic()
        'ExportPic()
        'Linq()
        'JsonT()
        'ExcelFormat()
    End Sub

    Sub InsertPic()
        Dim AppXls As Microsoft.Office.Interop.Excel.Application
        Dim AppWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim AppSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim img As System.Drawing.Image

        img = System.Drawing.Image.FromFile("E:\TestDoc\tew.png")
        System.Windows.Forms.Clipboard.SetDataObject(img, True)
        AppXls = New Microsoft.Office.Interop.Excel.Application
        AppXls.Workbooks.Open("E:\TestDoc\T1.xlsx")
        'AppXls.Workbooks.Open("E:\TestDoc\T2.xlsx")
        AppWorkBook = AppXls.Workbooks(1)

        AppSheet = AppWorkBook.Sheets("Char")
        'Dim range As Microsoft.Office.Interop.Excel.Range = AppSheet.Range(AppXls.Cells(1, 1), AppXls.Cells(1, 1))
        AppSheet.Paste(AppSheet.Range("A10", "A10"), img)

        AppWorkBook.Close(True)
        AppXls.Quit()

    End Sub

    Sub ExportPic()
        Dim AppXls As Microsoft.Office.Interop.Excel.Application
        Dim AppWorkBook As Microsoft.Office.Interop.Excel.Workbook
        'Dim AppSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim i As Integer
        Dim b As System.Drawing.Image

        AppXls = New Microsoft.Office.Interop.Excel.Application
        AppXls.Workbooks.Open("E:\TestDoc\T1.xlsx")
        AppWorkBook = AppXls.Workbooks(1)

        For Each sheet As Microsoft.Office.Interop.Excel.Worksheet In AppWorkBook.Sheets
            For i = 1 To sheet.Shapes.Count
                sheet.Shapes.Item(i).Copy()

                If (System.Windows.Forms.Clipboard.ContainsImage()) Then
                    b = System.Windows.Forms.Clipboard.GetImage()
                End If
                b.Save("E:\TestDoc\" + sheet.Name + "_" + i.ToString + ".png")
            Next
        Next
        AppWorkBook.Close(True)
        AppXls.Quit()
    End Sub

    Sub Linq()
        Dim tb As System.Data.DataTable = New System.Data.DataTable()
        tb.Columns.Add("Type1")
        tb.Columns.Add("Type2")
        tb.Columns.Add("Type3")
        tb.Columns.Add("Type4")

        tb.Rows.Add("1", "11", "111", "1111")
        tb.Rows.Add("4", "44", "444", "4444")
        tb.Rows.Add("2", "22", "222", "2222")
        tb.Rows.Add("3", "33", "333", "3333")
        tb.Rows.Add("2", "222", "2222", "22222")

        Dim tb2 As System.Data.DataTable = New System.Data.DataTable()
        tb2.Columns.Add("Key1")
        tb2.Columns.Add("Key2")
        tb2.Columns.Add("Key3")

        tb2.Rows.Add("1", "Beijing", "010")
        tb2.Rows.Add("2", "Guangzhou", "020")
        tb2.Rows.Add("3", "Shenzhen", "021")
        tb2.Rows.Add("5", "Shanghai", "011")


        Dim re As IEnumerable(Of Object)

        ''Return a System.Collections.Generic.IEnumerable(Of Object)() list 
        're = From dr In tb
        '     Where(Convert.ToInt32(dr.Item("Type1").ToString) > 2)
        '     Order By dr.Item("Type4")
        '     Select d4 = dr.Item("Type4"), d3 = dr.Item("Type3")

        ''Return a boolean value aggregate all & any
        'Dim t = Aggregate dr In tb
        '        Into(a = Any(dr.Item("Type1") > 2))

        ''Return a double value average,count & sum
        'Dim t = Aggregate dr In tb
        '        Into(Count(dr.Item("Type1").ToString > "2"))

        ''Distinct 
        're = From dr In tb
        '     Select d1 = dr.Item("Type1")
        '     Distinct

        ''Group by count,sum
        're = From dr1 In tb
        '     Group By t1 = dr1.Item("Type1")
        '     Into(Sum(Convert.ToInt32(dr1.Item("Type2"))))

        're = From dr2 In tb2
        '     Group Join dr1 In tb On
        '     dr2.Item("Key1") Equals dr1.Item("Type1").ToString
        '     Into t = Count()
        '     Select d2 = dr2.Item("Key1"), t

        're = From dr In tb
        '     Group Join dr2 In tb2 On
        '     dr.Item("Type1").ToString Equals dr2.Item("Key1").ToString
        '     Into(tt = Group)
        '     b = tt.FirstOrDefault
        '     Select t1 = dr.Item("Type1"), t2 = dr.Item("Type2"), t3 = dr.Item("Type3"), t4 = dr.Item("Type4"),
        '     y1 = If(b Is Nothing, Nothing, b("Key1")), y2 = If(b Is Nothing, Nothing, b("Key2")), y3 = If(b Is Nothing, Nothing, b("Key3"))

        re = From dr In tb
             Select t1 = dr.Item("Type1").ToString
             Skip While t1 < "5"


        'Console.WriteLine(t)

        For Each tt In re
            Console.WriteLine(tt.ToString)
        Next

        Console.ReadKey()


    End Sub

    Sub JsonT()
        Dim tt As String = "{""@odata.context"":""https://dev01-emea-uipath.roche.com/odata/$metadata#Robots"",""@odata.count"":8,""value"":[{""LicenseKey"":""***ecd72a2"",""MachineName"":""RBAMVPXWPAB0074"",""Name"":""RPW001"",""Username"":""EXBP\\yaojiang"",""Description"":""yaojian cpu"",""Type"":""Development"",""Password"":null,""RobotEnvironments"":""RocheDev"",""Id"":41,""ExecutionSettings"":null},{""LicenseKey"":""***2324edb"",""MachineName"":""RBAMVPXWPAB0063"",""Name"":""RPW003"",""Username"":""EXBP\\HANI2"",""Description"":""Ivy CPU"",""Type"":""Development"",""Password"":null,""RobotEnvironments"":""RocheDev"",""Id"":46,""ExecutionSettings"":null},{""LicenseKey"":""***7a83c86"",""MachineName"":""RBAMVPXWPAB0062"",""Name"":""RPW002"",""Username"":""EXBP\\MAOA2"",""Description"":""Adams CPU"",""Type"":""Development"",""Password"":null,""RobotEnvironments"":""RocheDev"",""Id"":49,""ExecutionSettings"":null},{""LicenseKey"":""***d1bb8b7"",""MachineName"":""RBAMVPXWPAB0069"",""Name"":""RPW004"",""Username"":""EXBP\\SHAOA1"",""Description"":""Athos's CPU"",""Type"":""Development"",""Password"":null,""RobotEnvironments"":""RocheDev"",""Id"":50,""ExecutionSettings"":null},{""LicenseKey"":""***edbe882"",""MachineName"":""SICMV884331"",""Name"":""RPS001"",""Username"":""rmoasia\\klbott01"",""Description"":"""",""Type"":""Unattended"",""Password"":null,""RobotEnvironments"":""RocheDev"",""Id"":52,""ExecutionSettings"":null},{""LicenseKey"":""***a226ea6"",""MachineName"":""RBAMVPXWPAB0070"",""Name"":""RPW005"",""Username"":""EXBP\\CENS2"",""Description"":""SANDY"",""Type"":""Development"",""Password"":null,""RobotEnvironments"":""RocheDev"",""Id"":54,""ExecutionSettings"":null},{""LicenseKey"":""***3fc61aa"",""MachineName"":""RBAMVPXWPAB0064"",""Name"":""RPW006"",""Username"":""EXBP\\ZENGL11"",""Description"":""Lianjian's CPU"",""Type"":""Development"",""Password"":null,""RobotEnvironments"":"""",""Id"":55,""ExecutionSettings"":null},{""LicenseKey"":""***59774a3"",""MachineName"":""RBAMVPXWPAB0067"",""Name"":""RPW007"",""Username"":""EXBP\\liw112"",""Description"":""windy cpu"",""Type"":""Development"",""Password"":null,""RobotEnvironments"":""RocheDev"",""Id"":57,""ExecutionSettings"":null}]}"
        Dim jo As JObject = CType(Newtonsoft.Json.JsonConvert.DeserializeObject(tt), JObject)



        For Each a In jo.Properties()
            Console.WriteLine("key : " + a.Name.ToString + " value : " + a.Value.ToString)
        Next

        Console.ReadKey()
    End Sub

    Sub ExcelFormat()
        Dim AppXls As Microsoft.Office.Interop.Excel.Application
        Dim AppWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim AppSheet As Microsoft.Office.Interop.Excel.Worksheet


        AppXls = New Microsoft.Office.Interop.Excel.Application
        AppXls.Workbooks.Open("E:\TestDoc\T1.xlsx")
        AppWorkBook = AppXls.Workbooks(1)
        AppSheet = AppWorkBook.Sheets("Sheet3")

        'set cell backgroud color
        AppSheet.Range("A1:C7").Interior.Color = System.Drawing.Color.LightBlue
        'set cell font Color
        AppSheet.Range("A2:A3").Font.Color = System.Drawing.Color.LightGreen
        'set cell font Bold
        AppSheet.Range("A2").Font.Bold = True
        'set cell font Italic
        AppSheet.Range("A1:A6").Font.Italic = True
        'set cell font Strikethrough
        AppSheet.Range("A3").Font.Strikethrough = True
        'set cell font size
        AppSheet.Range("A4:A5").Font.Size = 20
        'set cell font style
        AppSheet.Range("A4:A5").Font.Name = "Roman"
        'set cell data format
        'Text
        AppSheet.Range("A4").NumberFormatLocal = "@"
        'Date
        AppSheet.Range("A5").NumberFormatLocal = "yyyy-MM-dd"
        'Number
        AppSheet.Range("A6").NumberFormatLocal = "0.00"

        AppWorkBook.Save()
        AppWorkBook.Close(True)
        AppXls.Quit()
    End Sub

End Module
