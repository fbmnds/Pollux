

module Pollux.Excel.Utils

#if INTERACTIVE    
open Pollux.Log
open Pollux.Excel
#endif


open FParsec

open System.Xml
open System.Xml.Linq
open System.Xml.XPath

open System.IO.Packaging


let inline isNull x = x = Unchecked.defaultof<_>
let inline isNotNull x = x |> isNull |> not
let inline (|IsNull|) x = isNull x 
    

let inline convertIndex x y = sprintf "%s%d" (CellIndex.ColumnLabel y) (x + 1)
let inline convertIndex2 (x : int*int) = convertIndex (fst x) (snd x)

let convertCellIndex = function
    | Label label -> Index (CellIndex.ConvertLabel label)
    | Index (x,y) -> Label (convertIndex x y)

let convertCellIndex2 = function
    | Label label -> CellIndex.ConvertLabel label
    | Index (x,y) -> x,y

let rec isDateTime (s : string) =
    run (anyOf "ymdhs:") s
    |> function
    | Success _ ->  true
    | _ -> if  s = "" then false else isDateTime (s.Substring 1)

let builtInDateTimeNumberFormatIDs = 
    [| 14u; 15u; 16u; 17u; 18u; 19u;
       20u; 21u; 22u; 27u; 28u; 29u; 
       30u; 31u; 32u; 33u; 34u; 35u; 36u;
       45u; 46u; 47u; 50u;
       51u; 52u; 53u; 54u; 55u; 56u; 57u; 58u |]
    |> Seq.map string
        
let inline fromJulianDate x = 
    // System.DateTime.Parse("30.12.1899").Ticks = 599264352000000000L
    // System.TimeSpan.TicksPerDay = 864000000000L
    System.DateTime(599264352000000000L + (864000000000L * x)) 

let inline toJulianDate (x : System.DateTime) =
    (x.ToBinary() - 599264352000000000L) / 864000000000L


let inline id2 (i: int) (x: 'T) = x

let inline getPart1' (log : Pollux.Log.ILogger) 
                   (fileName : string) (xPath : string) (partUri : string) f = 
    log.LogLine Pollux.Log.LogLevel.Info 
        "Beginning 'getPart1\'' with xPath %s, partUri %s" xPath partUri
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
    let i = ref 0
    let result = 
        [| match expression.ReturnType with
                | XPathResultType.NodeSet -> 
                    let nodes = navigator.Select(expression)
                    while nodes.MoveNext() do
                        yield (f !i nodes.Current.OuterXml)
                        i := !i+1 
                | _ -> failwith <| sprintf "'getPart1\'': unexpected XPath-Expression return type '%A'" expression.ReturnType
        |]
    log.LogLine Pollux.Log.LogLevel.Info 
        "'getPart1\'' with xPath %s, partUri %s finished" xPath partUri
    result

let inline getPart2 (log : Pollux.Log.ILogger) 
                   (fileName : string) (xPath : string) (partUri : string) f = 
    log.LogLine Pollux.Log.LogLevel.Info 
        "Beginning 'getPart2' with xPath %s, partUri %s" xPath partUri
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
    let i = ref 0
    let result = 
        seq { match expression.ReturnType with
                | XPathResultType.NodeSet -> 
                    let nodes = navigator.Select(expression)
                    while nodes.MoveNext() do
                        yield (f !i nodes.Current.OuterXml)
                        i := !i+1 
                | _ -> failwith <| sprintf "'getPart2': unexpected XPath-Expression return type '%A'" expression.ReturnType
        }
    log.LogLine Pollux.Log.LogLevel.Info 
        "'getPart2' with xPath %s, partUri %s finished" xPath partUri
    result

let xn s = System.Xml.Linq.XName.Get(s)
let xd s = System.Xml.Linq.XDocument.Parse(s)


let getSheetId (log : Pollux.Log.ILogger) (fileName : string) (sheetName : string) =
    let partUri = "/xl/workbook.xml"
    let xPath = (sprintf "//*[name()='sheet' and @name='%s']" sheetName)
    getPart2 (log : Pollux.Log.ILogger) fileName xPath partUri id2
    |> Seq.head
    |> fun x -> 
        (xd x).Root.Attribute(xn "sheetId").Value

type CellContentContext =
    { log                  : Pollux.Log.ILogger 
      isCellDateTimeFormat : int -> bool
      rowOffset            : int
      colOffset            : int
      values               : CellContent [,] ref
      inlineString         : Dict<int,string> ref
      cellFormula          : Dict<int,string> ref
      extensionList        : Dict<int,string> ref
      unknownCellFormat    : Dict<int,string> ref }

let setCell i x (log : #Pollux.Log.ILogger)
    (inlineString: Dict<int,string> ref) (cellFormula: Dict<int,string> ref ) (extensionList: Dict<int,string> ref) = 
    let info = Pollux.Log.LogLevel.Info
    let test name = 
        let x' = (xd x).Root.Descendants() |> Seq.filter (fun x'' -> x''.Name.LocalName = name)
        if x' |> Seq.isEmpty then "" else x' |> Seq.head |> fun x'' -> x''.Value
    let test' (x': System.Xml.Linq.XAttribute) = if (isNull x' || isNull x'.Value) then "" else x'.Value
    let xa s = test' ((xd x).Root.Attribute(xn s))
    let test2 x (y: Dict<int,string>)  = 
        let z = test x
        if z = "" then -1 
        else y.Add (i, z); i
    let test3 (x: string) = if (xa x) = "" then -1 else x |> xa |> int
    let cv, cvb =     
        if "" = test "v" then -1M,false
        else
            try (test "v" |> decimal),true
            with | _ -> 
                log.LogLine info "setCell: ignoring invalid cell '%s'" x
                -1M,false
    let rR = xa "r"  |> CellIndex.ConvertLabel |> fst
    let rC = xa "r"  |> CellIndex.ConvertLabel |> snd
    ((rR,rC),
        {   isCellValueValid   = cvb
            CellValue          = cv
            InlineString       = test2 "is" !inlineString
            CellFormula        = test2 "f" !cellFormula
            ExtensionList      = test2 "extLst" !extensionList
            UnknownCellFormat  = -1
            CellMetadataIndex  = test3 "cm"
            ShowPhonetic       = test3 "ph" 
            ReferenceRow       = rR
            ReferenceCol       = rC
            StyleIndex         = test3 "s"  
            CellDataType       = if (xa "t") = "" then ' ' else ((xa "t").ToCharArray()).[0]
            ValueMetadataIndex = test3 "vm" })

let setCell3 (ctx : CellContentContext) index outerXml = 
    try
        let logInfo format = ctx.log.LogLine Pollux.Log.Info format
        let test name = 
            let x' = (xd outerXml).Root.Descendants() |> Seq.filter (fun x'' -> x''.Name.LocalName = name)
            if x' |> Seq.isEmpty then "" else x' |> Seq.head |> fun x'' -> x''.Value
        let test' (x': System.Xml.Linq.XAttribute) = if (isNull x' || isNull x'.Value) then "" else x'.Value
        let xa s = test' ((xd outerXml).Root.Attribute(xn s))
        let test2 x (y: Dict<int,string>)  = 
            let z = test x
            if z = "" then -1 
            else y.Add (index, z); index
        let test3 (x: string) = if (xa x) = "" then -1 else x |> xa |> int
        let cv, cvb =     
            if "" = test "v" then -1M,false
            else
                try (test "v" |> decimal),true
                with | _ -> 
                    logInfo "setCell: ignoring invalid cell '%s'" outerXml
                    -1M,false
        let rR = xa "r"  |> CellIndex.ConvertLabel |> fst
        let rC = xa "r"  |> CellIndex.ConvertLabel |> snd
        {   isCellValueValid   = cvb
            CellValue          = cv
            InlineString       = test2 "is" !(ctx.inlineString)
            CellFormula        = test2 "f" !(ctx.cellFormula)
            ExtensionList      = test2 "extLst" !(ctx.extensionList)
            UnknownCellFormat  = -1
            CellMetadataIndex  = test3 "cm"
            ShowPhonetic       = test3 "ph" 
            ReferenceRow       = rR
            ReferenceCol       = rC
            StyleIndex         = test3 "s"  
            CellDataType       = if (xa "t") = "" then ' ' else ((xa "t").ToCharArray()).[0]
            ValueMetadataIndex = test3 "vm" }
        |> Some
    with _ -> 
            let msg = sprintf "failed in setCell3:\ncell '%s'" outerXml
            ctx.log.LogLine Pollux.Log.Error "%A" msg
            (!ctx.unknownCellFormat).Add(index,outerXml)
            None
    |> fun x -> 
        match x with
        | Some x -> 
            let c =  
                if x.InlineString > -1 then CellContent.InlineString x.InlineString
                else if x.CellDataType = 's' then 
                    CellContent.StringTableIndex (int x.CellValue)
                else if x.isCellValueValid then 
                    if x.StyleIndex > -1 && ctx.isCellDateTimeFormat x.StyleIndex then 
                        CellContent.Date(fromJulianDate (int64 x.CellValue))
                    else CellContent.Decimal(x.CellValue)
                else CellContent.Empty
            x.ReferenceRow,x.ReferenceCol,c
        | None -> -1,-1,CellContent.Empty
    |> fun (rR,rC,x) -> 
        try
            (!ctx.values).[rR-ctx.rowOffset,rC-ctx.colOffset] <- x
        with _ -> 
            let msg = sprintf "failed in setCell3:\ncell '%A'" x 
            ctx.log.LogLine Pollux.Log.Error "%s" msg
            
