#r @".\bin\Debug\ExcelTypeProviderTest19.dll"
#r @"Microsoft.Office.Interop.Excel.dll"
#r @"office.dll"

open Microsoft.Office.Interop

//let xlApp = new Excel.ApplicationClass()
//let fullpath = __SOURCE_DIRECTORY__ + "\BookTest.xls"
//let xlWorkBookInput = xlApp.Workbooks.Open(fullpath)
//let xlWorkSheetInput = xlWorkBookInput.Worksheets.["Sheet1"] :?> Excel.Worksheet
//
//let rows = xlWorkSheetInput.Range(xlWorkSheetInput.Range("A1"), xlWorkSheetInput.Range("A1").End(Excel.XlDirection.xlDown))
//let rowsseq = seq { for row in rows.Rows do
//                     yield row :?> Excel.Range }
//                  |> Seq.skip 1
//
//let data =    
//   seq { for line in rowsseq do 
//            yield ( seq { for cell in line.Columns do
//                           yield (cell :?> Excel.Range).Value2 } 
//                     |> Seq.toList
//                  )
//      }        
//   |> Seq.toList

 
let file = new Samples.FSharpPreviewRelease2011.ExcelProvider.ExcelFile<"BookTest.xls", true>()

let toto = file.Data |> Seq.head

let titi = toto.BID

//
//type T = RegexTyped< @"(?<AreaCode>^\d{3})-(?<PhoneNumber>\d{3}-\d{4}$)">
//let reg = T() 
//let result = T.IsMatch("425-123-2345")
//let r = reg.Match("425-123-2345").AreaCode.Value //r equals "425"
