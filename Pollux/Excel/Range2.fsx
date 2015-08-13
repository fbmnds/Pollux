
#r "WindowsBase.dll"
#r "DocumentFormat.OpenXML.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsecCS.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsec.dll"

#r "System.Xml.Linq.dll"

open System.Xml

open System.IO.Packaging

#time;;
fsi.AddPrinter(fun (x:XmlNode) -> x.OuterXml);;

#load "Log.fs"
#load "CellParser.fs"
#load "Types.fs"
#load "Utils.fs"
#load "Range.fs"
#load "Excel.fs"

open Pollux.Excel
open Pollux.Excel.Utils
open Pollux.Excel.Range

let log = new Pollux.Log.ConsoleLogger()

let ``Cost Summary2.xlsx``      = __SOURCE_DIRECTORY__ + @"..\UnitTests\data\Cost Summary2.xlsx"
let ``Ranges.xlsx``             = __SOURCE_DIRECTORY__ + @"..\UnitTests\data\Ranges.xlsx"
let ``Base_Data_KA11_v5.xlsx``  = __SOURCE_DIRECTORY__ + @"..\..\..\..\trial 4\fsx\Base_Data_KA11_v5.xlsx"


do
    let sheet = LargeSheet (``Base_Data_KA11_v5.xlsx``, "Base_Data_KA11", false)
    printfn "%A" sheet.Ranges
    printfn "%A" (sheet.RangeValues("Header"))
    printfn "%A" (sheet.RangeValues("data"))


do
    getDefinedNames log (System.Guid.NewGuid()) ``Base_Data_KA11_v5.xlsx``
    |> printfn "%A" 
    let sheet,upperLeft,lowerRight = 
        "Base_Data_KA11!$A$1:$AN$1".Split('!')
        |> fun x -> 
            (x.[1]).Replace("$","").Split(':') 
            |> fun y -> x.[0],CellIndex.Label(y.[0]),CellIndex.Label(y.[1])
    (sheet,upperLeft,lowerRight)
    |> printfn "%A"


do
    let sheet = LargeSheet (``Base_Data_KA11_v5.xlsx``, "Base_Data_KA11", false)
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
    getPart log ``Cost Summary2.xlsx`` xPath partUri id2
    |> Seq.iter (printfn "%A")


do
    let sheet = LargeSheet (``Cost Summary2.xlsx``, "CheckSums", false)
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums"
           DefinedName = None
           UpperLeft  = match sheet.UpperLeft   with | Index(i,j) -> i,j | Label x -> x |> CellIndex.ConvertLabel
           LowerRight = match sheet.LowerRight  with | Index(i,j) -> i,j | Label x -> x |> CellIndex.ConvertLabel
           Values = sheet.Values }
    RangeWithCheckSumsCol (range')
    |> fun x -> printfn "%A %A %A" x.CheckSums x.CheckResults x.CheckErrors


do
    let sheet = LargeSheet (``Cost Summary2.xlsx``, "CheckSums2", false)
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums2"
           DefinedName = None
           UpperLeft  = match sheet.UpperLeft   with | Index(i,j) -> i,j | Label x -> x |> CellIndex.ConvertLabel
           LowerRight = match sheet.LowerRight  with | Index(i,j) -> i,j | Label x -> x |> CellIndex.ConvertLabel
           Values = sheet.Values }
    let conversion (i: int) (j: int) x = 
        match x with
        | StringTableIndex _ | InlineString _ | InlineString2 _ | Pending | Empty -> 0M
        | Decimal x -> x
        | Date x -> decimal (toJulianDate x)
    RangeWithCheckSumsCol (range', conversion)
    |> fun x -> x.Eps <- 1M; printfn "%A %A %A" x.CheckSums x.CheckResults x.CheckErrors

    
do
    CellIndex.ConvertLabel "APO6001"
    |> printfn "%A"
    // val it : int * int = (6000, 482)
