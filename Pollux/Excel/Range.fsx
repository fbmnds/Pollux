#r "WindowsBase.dll"
#r "DocumentFormat.OpenXML.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsecCS.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsec.dll"

#r "System.Xml.Linq.dll"



open System.Xml

open System.IO.Packaging

#time;;
fsi.AddPrinter(fun (x:XmlNode) -> x.OuterXml);;

#load "Utils.fs"
#load "Excel.fs"

open Pollux.Excel
open Pollux.Excel.Utils

let ``Cost Summary2.xlsx`` = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\Cost Summary2.xlsx"
let ``file6000rows``       = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\file6000rows.xlsx"


do
    let sheet = Sheet (``Cost Summary2.xlsx``, "Übersicht", false)
    printfn "%A" sheet.UpperLeft
    printfn "%A" sheet.LowerRight
    //sheet.Cells() |> Map.iter (fun k v -> printfn "%s:\n %A" k v)
    //sheet.CellFormats |> Map.iter (fun k v -> printfn "%d:\n %A" k v)
    printfn "--------"
    sheet.Values
    |> Array2D.iteri (fun i j x -> 
        if x <> CellContent.Empty then 
            printfn "%s %A" (convertIndex i j) x)

do
    let partUri = "/xl/styles.xml"
    let xPath = "//*[name()='cellXfs']/*[name()='xf']"
    getPart ``Cost Summary2.xlsx`` xPath partUri
    |> Seq.iter (printfn "%A")

do
    let sheet = Sheet (``Cost Summary2.xlsx``, "CheckSums", false)
    printfn "%A" sheet.UpperLeft
    printfn "%A" sheet.LowerRight
    sheet.Values
    |> Array2D.iteri (fun i j x -> 
        let i',j' = match sheet.UpperLeft  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
        if x <> CellContent.Empty then 
            printfn "%s %A" (convertIndex (i + i') (j + j')) x)

do
    let sheet = Sheet (``Cost Summary2.xlsx``, "CheckSums2", false)
    printfn "%A" sheet.UpperLeft
    printfn "%A" sheet.LowerRight
    sheet.Values
    |> Array2D.iteri (fun i j x -> 
        let i',j' = match sheet.UpperLeft  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
        if x <> CellContent.Empty then 
            printfn "%s %A" (convertIndex (i + i') (j + j')) x)

do
    let sheet = Sheet (``Cost Summary2.xlsx``, "CheckSums", false)
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums2"
           UpperLeft  = match sheet.UpperLeft   with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           LowerRight = match sheet.LowerRight  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           Values = sheet.Values }
    RangeWithCheckSumsRow (range')
    |> fun x -> printfn "%A %A %A" x.CheckSums x.CheckResults x.CheckErrors

do
    let sheet = Sheet (``Cost Summary2.xlsx``, "CheckSums2", false)
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums2"
           UpperLeft  = match sheet.UpperLeft   with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           LowerRight = match sheet.LowerRight  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           Values = sheet.Values }
    RangeWithCheckSumsRow (range')
    |> fun x -> printfn "%A %A %A" x.CheckSums x.CheckResults x.CheckErrors

do
    let sheet = Sheet (``Cost Summary2.xlsx``, "CheckSums", false)
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums"
           UpperLeft  = match sheet.UpperLeft   with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           LowerRight = match sheet.LowerRight  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           Values = sheet.Values }
    RangeWithCheckSumsCol (range')
    |> fun x -> printfn "%A %A %A" x.CheckSums x.CheckResults x.CheckErrors

do
    let sheet = Sheet (``Cost Summary2.xlsx``, "CheckSums2", false)
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums2"
           UpperLeft  = match sheet.UpperLeft   with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           LowerRight = match sheet.LowerRight  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           Values = sheet.Values }
    let conversion (i: int) (j: int) x = 
        match x with
        | StringTableIndex _ | InlineString _ | Empty -> 0M
        | Decimal x -> x
        | Date x -> decimal (toJulianDate x)
    RangeWithCheckSumsCol (range', conversion)
    |> fun x -> x.Eps <- 1M; printfn "%A %A %A" x.CheckSums x.CheckResults x.CheckErrors


    
