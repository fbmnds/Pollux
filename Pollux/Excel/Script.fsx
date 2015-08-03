#r "WindowsBase.dll"
#r "DocumentFormat.OpenXML.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsecCS.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsec.dll"


#load "Utils.fs"
#load "Excel.fs"

open Pollux.Excel
open Pollux.Excel.Utils

let ``Cost Summary2.xlsx`` = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\Cost Summary2.xlsx"

do
    use workbook = new Workbook (``Cost Summary2.xlsx``, false)
    let sheet = Sheet (workbook, "Übersicht", false)
    printfn "%A" sheet.UpperLeft
    printfn "%A" sheet.LowerRight
    sheet.Values
    |> Array2D.iteri (fun i j x -> 
        if x <> CellContent.Empty then 
            printfn "%s %A" (convertIndex i j) x)

do
    use workbook = new Workbook (``Cost Summary2.xlsx``, false)
    let sheet = Sheet (workbook, "CheckSums", false)
    printfn "%A" sheet.UpperLeft
    printfn "%A" sheet.LowerRight
    sheet.Values
    |> Array2D.iteri (fun i j x -> 
        let i',j' = match sheet.UpperLeft  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
        if x <> CellContent.Empty then 
            printfn "%s %A" (convertIndex (i + i') (j + j')) x)

do
    use workbook = new Workbook (``Cost Summary2.xlsx``, false)
    let sheet = Sheet (workbook, "CheckSums2", false)
    printfn "%A" sheet.UpperLeft
    printfn "%A" sheet.LowerRight
    sheet.Values
    |> Array2D.iteri (fun i j x -> 
        let i',j' = match sheet.UpperLeft  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
        if x <> CellContent.Empty then 
            printfn "%s %A" (convertIndex (i + i') (j + j')) x)

do
    use workbook = new Workbook (``Cost Summary2.xlsx``, false)
    let sheet = Sheet (workbook, "CheckSums", false)
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums2"
           UpperLeft  = match sheet.UpperLeft   with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           LowerRight = match sheet.LowerRight  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           Values = sheet.Values }
    RangeWithCheckSumsRow (range')
    |> fun x -> printfn "%A %A %A" x.CheckSums x.CheckResults x.CheckErrors

do
    use workbook = new Workbook (``Cost Summary2.xlsx``, false)
    let sheet = Sheet (workbook, "CheckSums2", false)
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums2"
           UpperLeft  = match sheet.UpperLeft   with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           LowerRight = match sheet.LowerRight  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           Values = sheet.Values }
    RangeWithCheckSumsRow (range')
    |> fun x -> printfn "%A %A %A" x.CheckSums x.CheckResults x.CheckErrors

do
    use workbook = new Workbook (``Cost Summary2.xlsx``, false)
    let sheet = Sheet (workbook, "CheckSums", false)
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums"
           UpperLeft  = match sheet.UpperLeft   with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           LowerRight = match sheet.LowerRight  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           Values = sheet.Values }
    RangeWithCheckSumsCol (range')
    |> fun x -> printfn "%A %A %A" x.CheckSums x.CheckResults x.CheckErrors

do
    use workbook = new Workbook (``Cost Summary2.xlsx``, false)
    let sheet = Sheet (workbook, "CheckSums2", false)
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums2"
           UpperLeft  = match sheet.UpperLeft   with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           LowerRight = match sheet.LowerRight  with | Index(i,j) -> i,j | Label x -> x |> convertLabel
           Values = sheet.Values }
    let conversion (i: int) (j: int) x = 
        match x with
        | StringTableIndex _ | Empty -> 0M
        | Decimal x -> x
        | Date x -> decimal (toJulianDate x)
    RangeWithCheckSumsCol (range', conversion)
    |> fun x -> x.Eps <- 1M; printfn "%A %A %A" x.CheckSums x.CheckResults x.CheckErrors


    
