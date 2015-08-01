#r "WindowsBase.dll"
#r "DocumentFormat.OpenXML.dll"
#r @"..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsecCS.dll"
#r @"..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsec.dll"


#load "Utils.fs"
#load "Excel.fs"

open Pollux.Excel

let sheet = Sheet (@"..\..\..\Desktop\Cost Summary2.xlsx", "Übersicht", false)

sheet.Values()
|> Array2D.iteri (fun i j x -> if x <> CellContent.Empty then printfn "%s %A" (Utils.convertIndex i j) x)
