﻿

module Pollux.Excel.Utils

#if INTERACTIVE    
open Pollux.Log
open Pollux.Excel
#endif

open Pollux.Excel.Cell.Parser

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

let inline Array2DColSum (x : 'T [,]) col = 
    [ for row in [x.GetLowerBound(0) .. x.GetUpperBound(0)] do yield x.[row,col] ]
    |> Seq.reduce (+)

let inline Array2DRowSum (x : 'T [,]) row = 
    [ for col in [x.GetLowerBound(1) .. x.GetUpperBound(1)] do yield x.[row,col] ]
    |> Seq.reduce (+)

let getDimensions (log : Pollux.Log.ILogger) (fileName : string) sheetName s = 
    try
        parseUnsafe 1 "dimension" (ref s)
        |> Seq.head
        |> fun x -> 
            let len = x.Length 
            (x.Substring(0, len - "\"/>".Length)).Substring("<dimension ref=\"".Length).Split([|':'|])
        |> fun x -> 
            let upperLeft = Index(CellIndex.ConvertLabel x.[0])
            let lowerRight = Index(CellIndex.ConvertLabel x.[1])
            let rowCapacity = (fst (convertCellIndex2 lowerRight)) - (fst (convertCellIndex2 upperLeft)) + 1
            let colCapacity = (snd (convertCellIndex2 lowerRight)) - (snd (convertCellIndex2 upperLeft)) + 1
            upperLeft, lowerRight, rowCapacity, colCapacity
    with _ -> 
        let msg = sprintf "LargeSheet: could not read 'dimension' of sheet '%s' in '%s'" sheetName fileName
        log.LogLine Pollux.Log.LogLevel.Error "%s" msg
        failwith msg


let inline id2 (i: int) (x: 'T) = x

let inline getPart (log : Pollux.Log.ILogger) 
                   (fileName : string) (xPath : string) (partUri : string) f = 
    log.LogLine Pollux.Log.LogLevel.Info 
        "Beginning 'getPart' with xPath %s, partUri %s" xPath partUri
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
                | _ -> failwith <| sprintf "'getPart': unexpected XPath-Expression return type '%A'" expression.ReturnType
        |]
    log.LogLine Pollux.Log.LogLevel.Info 
        "'getPart' with xPath %s, partUri %s finished" xPath partUri
    result

let xn s = System.Xml.Linq.XName.Get(s)
let xd s = System.Xml.Linq.XDocument.Parse(s)
let test x name = 
    let x' = (xd x).Root.Descendants() |> Seq.filter (fun x'' -> x''.Name.LocalName = name)
    if x' |> Seq.isEmpty then "" else x' |> Seq.head |> fun x'' -> x''.Value
let test' (x: System.Xml.Linq.XAttribute) =
    if (isNull x || isNull x.Value) then "" else x.Value
let xa x s = test' ((xd x).Root.Attribute(xn s))


let getSheetId (log : Pollux.Log.ILogger) (fileName : string) (sheetName : string) =
    let partUri = "/xl/workbook.xml"
    let xPath = (sprintf "//*[name()='sheet' and @name='%s']" sheetName)
    getPart (log : Pollux.Log.ILogger) fileName xPath partUri id2
    |> Seq.head
    |> fun x -> 
        x.Replace(" r:id=", " rId=") 
        |> fun x -> (xd x).Root.Attribute(xn @"rId").Value.Substring(3)

let getNumberFormats (log : Pollux.Log.ILogger) (fileName : string) = 
    log.LogLine Pollux.Log.LogLevel.Info
        "%s" "upperLeft, lowerRight, keys finished,  beginning with numberFormats ..."
    let partUri = "/xl/styles.xml"
    let numberFormats = 
        let xPath = "//*[name()='numFmt']"
        getPart log fileName xPath partUri id2
        |> Seq.map (fun x ->             
            { NumberFormatId = xa x "numFmtId"; FormatCode = xa x "formatCode" })
    log.LogLine Pollux.Log.LogLevel.Info 
        "%s" "numberFormats finished,  beginning with cellFormats ..."
    numberFormats

let getCellFormats (log : Pollux.Log.ILogger) (fileName : string) =
    let partUri = "/xl/styles.xml"
    let xPath = "//*[name()='cellXfs']/*[name()='xf']"
    getPart log fileName xPath partUri id2
    |> Seq.mapi (fun i x ->                 
        i,
        { NumFmtId          = xa x "numFmtId";
          BorderId          = xa x "borderId"
          FillId            = xa x "fillId";
          FontId            = xa x "fontId"; 
          ApplyAlignment    = xa x "applyAlignment";
          ApplyBorder       = xa x "applyBorder";
          ApplyFont         = xa x "applyFont";
          XfId              = xa x "xfId";
          ApplyNumberFormat = xa x "applyNumberFormat" })                             
    |> Map.ofSeq

let getSharedStrings (log : Pollux.Log.ILogger) (fileName : string) =
    let partUri = "/xl/sharedStrings.xml"
    let xPath = "//*[name()='sst']/*[name()='si']"   
    getPart log fileName xPath partUri id2
    |> Seq.mapi (fun i x ->                 
        i, test x "t")             // TODO: capture text in runs     
    |> dict

let parseDefinedNames (x: string) = 
    let errMsg = (sprintf "ERROR:unexpected 'definedName' format:VALUE:%s" x),(-1,-1),(-1,-1)
    try
        let name = xa x "name"
        (xd x).Root.Value.Split('!')
        |> fun x -> 
            (x.[1]).Replace("$","").Split(':') 
            |> fun y ->             
                let upperLeft = y.[0] |> CellIndex.ConvertLabel
                if      y.Length = 2 then name, upperLeft, y.[1] |> CellIndex.ConvertLabel
                else if y.Length = 1 then name, upperLeft, upperLeft
                else errMsg
    with _ -> errMsg

let getDefinedNames (log : Pollux.Log.ILogger) sheetGuid (fileName : string) =
    let partUri = "/xl/workbook.xml"
    let xPath = "//*[name()='definedNames']/*[name()='definedName']"
    getPart log fileName xPath partUri id2
    |> Seq.map (fun x ->                 
        let name,upperLeft,lowerRight = parseDefinedNames x
        name,
        { Name       = name
          UpperLeft  = CellIndex.Index(upperLeft)
          LowerRight = CellIndex.Index(lowerRight)
          SheetGuid  = sheetGuid })                             
    |> Map.ofSeq

let GetCellDateTimeFormats numberFormats = 
    numberFormats   
    |> Seq.filter (fun x -> x.FormatCode |> isDateTime)
    |> Seq.map (fun x -> x.NumberFormatId)
    |> Seq.append builtInDateTimeNumberFormatIDs 

let fIsCellDateTimeFormat (cellFormats : Map<int,CellFormat>) cellDateTimeFormats =
    fun x -> 
        if cellFormats.ContainsKey (x) then 
            cellDateTimeFormats
            |> Seq.filter (fun x' -> x' = (cellFormats.[x]).NumFmtId)
            |> Seq.isEmpty
            |> not
        else false

let setCell (ctx : CellContentContext) index outerXml = 
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
        let is = 
            if cvb |> not && (xa "s") = "6" 
            then (!ctx.inlineString).Add (index, (xa "r")) ; index 
            else test2 "is" !(ctx.inlineString)
        {   isCellValueValid   = cvb
            CellValue          = cv
            InlineString       = is 
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
        with _ -> ()
//            let msg = sprintf "failed in setCell3:\ncell '%A'" x 
//            ctx.log.LogLine Pollux.Log.Error "%s" msg
            
