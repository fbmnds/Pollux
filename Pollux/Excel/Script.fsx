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

#load "Utils.fs"
#load "Excel.fs"

open Pollux.Excel
open Pollux.Excel.Utils

let ``Cost Summary2.xlsx`` = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\Cost Summary2.xlsx"
let ``file6000rows``       = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\file6000rows.xlsx"

//---------------------------------------------------------------------------------------------------------------------
//
//  Preferred: System.Xml.XPath   
//
//    > getCells ``file6000rows`` "/xl/worksheets/sheet1.xml" |> Seq.take 5 ;;
//    Real: 00:01:43.348, CPU: 00:01:44.578, GC gen0: 243, gen1: 86, gen2: 5
//    val it : seq<string> =
//      seq
//        ["<c r="A1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
//      <f ca="1">RANDBETWEEN(0,1000)</f>
//      <v>437</v>
//    </c>";
//         ...]
//    >
//
//---------------------------------------------------------------------------------------------------------------------


//    Excel Part Uri
//    --------------
//    /docProps/app.xml
//    /docProps/core.xml
//    /xl/calcChain.xml
//    /xl/printerSettings/printerSettings1.bin
//    /xl/sharedStrings.xml
//    /xl/styles.xml
//    /xl/theme/theme1.xml
//    /xl/workbook.xml
//    /xl/worksheets/sheet1.xml
//    /xl/worksheets/sheet2.xml
//    /xl/worksheets/sheet3.xml
//    /xl/worksheets/_rels/sheet1.xml.rels
//    /xl/_rels/workbook.xml.rels
//    /_rels/.rels

let getPart (fileName : string) (xPath : string) (partUri : string) = 
    use xlsx = ZipPackage.Open(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
    let part = 
        xlsx.GetParts()
        |> Seq.filter (fun x -> x.Uri.ToString() = partUri)
        |> Seq.head
    use stream = part.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read)
    let xml = new XPathDocument(stream)
    let navigator = xml.CreateNavigator()
    let manager = new XmlNamespaceManager(navigator.NameTable)
    let expression = XPathExpression.Compile(xPath, manager)
    seq { 
        match expression.ReturnType with
        | XPathResultType.NodeSet -> 
            let nodes = navigator.Select(expression)
            while nodes.MoveNext() do
                yield nodes.Current.OuterXml
        | _ -> failwith <| sprintf "XPath-Expression return type %A not implemented" expression.ReturnType
    }


let getSheetId (fileName : string) (sheetName : string) =
    let partUri = "/xl/workbook.xml"
    let xPath = (sprintf "//*[name()='sheet' and @name='%s']" sheetName)
    getPart fileName xPath partUri
    |> Seq.head
    |> fun x -> 
        let xn s = XName.Get(s)
        let xd = XDocument.Parse(x)
        xd.Root.Attribute(xn "sheetId").Value


let getCells (fileName : string) (sheetName : string) =
    let partUri = sprintf "/xl/worksheets/sheet%s.xml" (getSheetId fileName sheetName)
    let xPath = "//*[name()='c']"
    getPart fileName xPath partUri
    |> Seq.map (fun x -> 
        let xn s = XName.Get(s)
        let xd = XDocument.Parse(x)
        let test name = 
            let x = xd.Root.Descendants() |> Seq.filter (fun x -> x.Name.LocalName = name)
            if x |> Seq.isEmpty then "" else x |> Seq.head |> fun x -> x.Value
        let test' (x: XAttribute) = if (isNull x || isNull x.Value) then "" else x.Value
        let xa s = test' (xd.Root.Attribute(xn s))
        { CellValue          = test "v";
          InlineString       = test "is"
          CellFormula        = test "f";
          ExtensionList      = test "extLst"; 
          CellMetadataIndex  = xa "cm";
          ShowPhonetic       = xa "ph";
          Reference          = (Label(xa "r"));
          StyleIndex         = xa "s";
          CellDataType       = xa "t";
          ValueMetadataIndex = xa "vm" })

do
    getCells ``file6000rows`` "Random" |> Seq.take 3
    |> printfn "%A"
//    seq
//      [{CellValue = "437";
//        InlineString = "";
//        CellFormula = "RANDBETWEEN(0,1000)";
//        ExtensionList = "";
//        CellMetadataIndex = "";
//        ShowPhonetic = "";
//        Reference = Label "A1";
//        StyleIndex = "2";
//        CellDataType = "";
//        ValueMetadataIndex = "";}; {CellValue = "358";
//                                    InlineString = "";
//                                    CellFormula = "RANDBETWEEN(0,1000)";
//                                    ExtensionList = "";
//                                    CellMetadataIndex = "";
//                                    ShowPhonetic = "";
//                                    Reference = Label "B1";
//                                    StyleIndex = "2";
//                                    CellDataType = "";
//                                    ValueMetadataIndex = "";};
//       {CellValue = "175";
//        InlineString = "";
//        CellFormula = "";
//        ExtensionList = "";
//        CellMetadataIndex = "";
//        ShowPhonetic = "";
//        Reference = Label "C1";
//        StyleIndex = "2";
//        CellDataType = "";
//        ValueMetadataIndex = "";}]
//    Real: 00:02:17.895, CPU: 00:02:29.531, GC gen0: 241, gen1: 86, gen2: 3
//    val it : unit = ()
//    > 

do 
    let cell = """<c r="A1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f ca="1">RANDBETWEEN(0,1000)</f><v>437</v><is><t>inline string</t></is></c>"""
    let xn s = XName.Get(s)
    let xd = XDocument.Parse(cell)
    //let test (x: XElement) = if (isNull x || isNull x.Value) then "****" else x.Value
    let test' (x: XAttribute) = if (isNull x || isNull x.Value) then "" else x.Value
    let xa s = test' (xd.Root.Attribute(xn s))
    let test name = 
        let x = xd.Root.Descendants() |> Seq.filter (fun x -> x.Name.LocalName = name)
        if x |> Seq.isEmpty then "" else x |> Seq.head |> fun x -> x.Value
    xd.Root.Descendants() |> Seq.filter (fun x -> x.Name.LocalName = "v") |> Seq.head |> fun x -> x.Value |> printfn "%s"
    xd.Root.Descendants() |> Seq.filter (fun x -> x.Name.LocalName = "is") |> Seq.head |> fun x -> x.Value |> printfn "%s" 
    test "g"
    |> ignore

//---------------------------------------------------------------------------------------------------------------------
//
//  Benchmark System.Xml   
//
//---------------------------------------------------------------------------------------------------------------------


//    > getFirstPart ``file6000rows`` "/xl/worksheets/sheet1.xml" |> ignore;;
//    Real: 00:02:33.794, CPU: 00:02:42.781, GC gen0: 1161, gen1: 408, gen2: 6
//    val it : unit = ()
//    > 

let getFirstPart (fileName : string) (partUri : string) =
    use xlsx = ZipPackage.Open(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
    let part = 
        xlsx.GetParts()
        |> Seq.filter (fun x -> x.Uri.ToString() = partUri)
        |> Seq.head
    use stream = part.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read)
    let xml = new XmlDocument()
    xml.Load(stream) 
    xml





    

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
        | StringTableIndex _ | InlineString _ | Empty -> 0M
        | Decimal x -> x
        | Date x -> decimal (toJulianDate x)
    RangeWithCheckSumsCol (range', conversion)
    |> fun x -> x.Eps <- 1M; printfn "%A %A %A" x.CheckSums x.CheckResults x.CheckErrors


    
