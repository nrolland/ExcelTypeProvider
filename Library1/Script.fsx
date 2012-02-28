#r @".\bin\Debug\ExcelTypeProviderTest24.dll"
#r @"Microsoft.Office.Interop.Excel.dll"
#r @"office.dll"

open Microsoft.Office.Interop

let filename  = @"C:\Users\e021230\Documents\Visual Studio 11\Projects\exceltypeprovider\Library1\BookTest.xls"
type ExcelFileInternal(filename) =
      let data  = 
         let xlApp = new Excel.ApplicationClass()
         let xlWorkBookInput = xlApp.Workbooks.Open(filename)
         let xlWorkSheetInput = xlWorkBookInput.Worksheets.["Sheet1"] :?> Excel.Worksheet

         // Cache the sequence of all data lines (all lines but the first)
         let firstrow = xlWorkSheetInput.Range(xlWorkSheetInput.Range("A1"), xlWorkSheetInput.Range("A1").End(Excel.XlDirection.xlToRight))
         let rows = xlWorkSheetInput.Range(firstrow, firstrow.End(Excel.XlDirection.xlDown))
         let rows_data = seq { for row  in rows.Rows do 
                                 yield row :?> Excel.Range } |> Seq.skip 1
         let res = 
            seq { for line_data in rows_data do 
                  yield ( seq { for cell in line_data.Columns do
                                 yield (cell  :?> Excel.Range ).Value2} 
                           |> Seq.toArray
                        )
               }
               |> Seq.toArray
         xlWorkBookInput.Close()
         res

      member __.Data = data

if false then
   let file = ExcelFileInternal(filename)
   printf "%A" file.Data
else
   let file = Samples.FSharpPreviewRelease2011.ExcelProvider.ExcelFileInternal(filename)
   printf "%A" file.Data

