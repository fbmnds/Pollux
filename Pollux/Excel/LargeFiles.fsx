
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

let log = (new Pollux.Log.ConsoleLogger()) :> Pollux.Log.ILogger

let ``file6000rows.xlsx`` = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\file6000rows.xlsx"

do
    let partUri =  sprintf "/xl/worksheets/sheet%s.xml" (getSheetId log ``file6000rows.xlsx`` "Random")
    use xlsx = ZipPackage.Open(``file6000rows.xlsx``, System.IO.FileMode.Open, System.IO.FileAccess.Read)
    let part = 
        xlsx.GetParts()
        |> Seq.filter (fun x -> x.Uri.ToString() = partUri)
        |> Seq.head
    use stream = part.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read)
    use reader = XmlReader.Create(stream)
    log.LogLine Pollux.Log.LogLevel.Info "%s" "start reader..."
    let i = ref 0
    let result = 
        [| 
            while reader.Read() do
                if (reader.MoveToContent() = XmlNodeType.Element && reader.Name = "c") then
                    yield  !i, (reader.ReadOuterXml())
                    i := !i+1 
        |]
    log.LogLine Pollux.Log.LogLevel.Info "finished, %d cells in total, take 5 ..." result.Length
    result |> Seq.ofArray |> Seq.take 5 |> printfn "%A"
    log.LogLine Pollux.Log.LogLevel.Info "%s" "finished, build dict ..."
    result |> dict |> Seq.take 5 |> printfn "%A"
    log.LogLine Pollux.Log.LogLevel.Info "%s" "finished, build map ..."
    result |> Map.ofArray |> Seq.take 5 |> printfn "%A"
    log.LogLine Pollux.Log.LogLevel.Info "%s" "finished"

//    [07:07:53 UTC] Beginning 'getPart2' with xPath //*[name()='sheet' and @name='Random'], partUri /xl/workbook.xml
//    [07:07:53 UTC] 'getPart2' with xPath //*[name()='sheet' and @name='Random'], partUri /xl/workbook.xml finished
//    [07:07:53 UTC] start reader...
//    [07:09:28 UTC] finished, 3324554 cells in total, take 5 ...
//    seq
//      [(0,
//        "<c r="A1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f ca="1">RANDBETWEEN(0,1000)</f><v>437</v></c>");
//       (1,
//        "<c r="C1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>175</v></c>");
//       (2,
//        "<c r="E1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>285</v></c>");
//       (3,
//        "<c r="G1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>397</v></c>");
//       ...]
//    [07:09:28 UTC] finished, build dict ...
//    seq
//      [[0, <c r="A1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f ca="1">RANDBETWEEN(0,1000)</f><v>437</v></c>];
//       [1, <c r="C1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>175</v></c>];
//       [2, <c r="E1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>285</v></c>];
//       [3, <c r="G1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>397</v></c>];
//       ...]
//    [07:09:30 UTC] finished, build map ...
//    seq
//      [[0, <c r="A1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f ca="1">RANDBETWEEN(0,1000)</f><v>437</v></c>];
//       [1, <c r="C1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>175</v></c>];
//       [2, <c r="E1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>285</v></c>];
//       [3, <c r="G1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>397</v></c>];
//       ...]
//    [07:09:59 UTC] finished
//    val it : unit = ()
//    > 


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