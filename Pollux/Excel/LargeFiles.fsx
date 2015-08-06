
#r "WindowsBase.dll"
#r "DocumentFormat.OpenXML.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsecCS.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsec.dll"

#r "System.Xml.Linq.dll"


open System.Xml
open System.Xml.Linq
open System.Xml.XPath

open System.IO.Packaging

#time;;
fsi.AddPrinter(fun (x:XmlNode) -> x.OuterXml);;

#load "Log.fs"
#load "Utils.fs"
#load "Excel.fs"

open Pollux.Log
open Pollux.Excel
open Pollux.Excel.Utils

let ``file6000rows.xlsx`` = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\file6000rows.xlsx"

do
    let sheet = Sheet ((new ConsoleLogger()), ``file6000rows.xlsx``, "Random", false)
    Pollux.Log.logInfo "%A" sheet.UpperLeft
    Pollux.Log.logInfo "%A" sheet.LowerRight
    //sheet.Cells() |> Map.iter (fun k v -> printfn "%s:\n %A" k v)
    //sheet.CellFormats |> Map.iter (fun k v -> printfn "%d:\n %A" k v)
    printfn "--------"
    sheet.Values
    |> Array2D.iteri (fun i j x -> 
        if x <> CellContent.Empty then 
            printfn "%s %A" (convertIndex i j) x)