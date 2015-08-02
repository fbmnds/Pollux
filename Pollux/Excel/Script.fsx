#r "WindowsBase.dll"
#r "DocumentFormat.OpenXML.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsecCS.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsec.dll"

#r @"C:\Users\Friedrich\projects\Pollux\Pollux\Excel\bin\Debug\Excel.dll"

//#load "Utils.fs"
//#load "Excel.fs"

open Pollux.Excel
open Pollux.Excel.Utils

do
    use workbook = new Workbook (__SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\Cost Summary2.xlsx", false)
    let sheet = Sheet (workbook, "Übersicht", false)
    printfn "%A" sheet.UpperLeft
    printfn "%A" sheet.LowerRight
    sheet.Values
    |> Array2D.iteri (fun i j x -> if x <> CellContent.Empty then printfn "%s %A" (Utils.convertIndex i j) x)
