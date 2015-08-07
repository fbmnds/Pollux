﻿namespace Pollux

namespace Pollux.Excel
   
#if INTERACTIVE
    open Pollux.Excel.Utils
    open Pollux.Log
#endif

    type private SpreadsheetDocument'   = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument
    type private WorksheetPart'         = DocumentFormat.OpenXml.Packaging.WorksheetPart
    //type private SharedStringTablePart' = DocumentFormat.OpenXml.Packaging.SharedStringTablePart

    type private Sheet'                 = DocumentFormat.OpenXml.Spreadsheet.Sheet
    type private Row'                   = DocumentFormat.OpenXml.Spreadsheet.Row
    type private Column'                = DocumentFormat.OpenXml.Spreadsheet.Column
    type private Cell'                  = DocumentFormat.OpenXml.Spreadsheet.Cell
    type private CellFormat'            = DocumentFormat.OpenXml.Spreadsheet.CellFormat
    type private CellValues'            = DocumentFormat.OpenXml.Spreadsheet.CellValues
    //type private SharedStringTable'     = DocumentFormat.OpenXml.Spreadsheet.SharedStringTable
    type private SharedStringItem'      = DocumentFormat.OpenXml.Spreadsheet.SharedStringItem
    type private NumberingFormat'       = DocumentFormat.OpenXml.Spreadsheet.NumberingFormat


    type Index = RowIndex*ColIndex
    and RowIndex = int
    and ColIndex = int
    and Label    = string

    [<CustomEquality; CustomComparison>]
    type CellIndex = 
    | Label of Label
    | Index of Index
        member x.ToTuple : int*int = 
            match x with
            | Label x -> convertLabel x
            | Index x -> (fst x), (snd x)

        override x.GetHashCode() = x.GetHashCode()

        override x.Equals(y) =
            match y with
            | :? CellIndex as y -> 
                match y with
                |  Label y as y' -> y'.ToTuple = x.ToTuple
                |  Index y as y' -> y'.ToTuple = x.ToTuple
            | _ -> invalidArg (sprintf "'%A'" y) "is not comparable to CellIndex."

        interface System.IComparable with
           member x.CompareTo y = 
              match y with 
              | :? CellIndex as y -> 
                  match y with
                  | Label y as y' -> 
                       let (a,b) = y'.ToTuple
                       let (a',b') = x.ToTuple
                       if a=a' && b=b' then 0
                       else if a>a' && b>b' then 1
                       else -1
                  | Index y as y' -> 
                       let (a,b) = y'.ToTuple
                       let (a',b') = x.ToTuple
                       if a=a' && b=b' then 0
                       else if a>a' && b>b' then 1
                       else -1              
              | _ -> invalidArg (sprintf "'%A'" y) "is not comparable to CellIndex."
        

    type CellContent =
    | StringTableIndex  of int
    | InlineString      of int
    | Decimal           of decimal
    | Date              of System.DateTime
    | Empty          
    
    [<CustomEquality; NoComparison; CLIMutable>]
    type Cell = 
        { isCellValueValid   : bool
          CellValue          : decimal
          InlineString       : int
          CellFormula        : int
          ExtensionList      : int
          CellMetadataIndex  : int
          ShowPhonetic       : int
          ReferenceRow       : int
          ReferenceCol       : int
          StyleIndex         : int
          CellDataType       : char
          ValueMetadataIndex : int }

        override x.GetHashCode() = (x.ReferenceRow, x.ReferenceCol).GetHashCode()

        override x.Equals(y) =
            match y with
            | :? Cell as y -> x.ReferenceRow = y.ReferenceRow && x.ReferenceCol = y.ReferenceCol
            | _ -> invalidArg (sprintf "'%A'" y) "is not comparable to CellIndex."

    
    type NumberFormat = 
        { NumberFormatId : string
          FormatCode     : string }
    and CellFormat = 
// https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.cellformats(v=office.14).aspx
        { NumFmtId          : string
          BorderId          : string
          FillId            : string
          FontId            : string
          ApplyAlignment    : string
          ApplyBorder       : string
          ApplyFont         : string
          XfId              : string 
          ApplyNumberFormat : string }
          // Alignment as subtype

    type Range =
        { mutable Name : string 
          UpperLeft    : Index
          LowerRight   : Index
          Values       : CellContent [,] }
    and StringRange = 
        { mutable Name : string 
          UpperLeft    : Index
          LowerRight   : Index
          Values       : string [,] }
    and DecimalRange = 
        { mutable Name : string 
          UpperLeft    : Index
          LowerRight   : Index
          Values       : decimal [,] }


    type RangeWithCheckSumsRow (range : DecimalRange) =
        do
            if range.Values.GetUpperBound(0) < 1 then 
                raise (System.ArgumentOutOfRangeException ("CheckSum range row dimension error"))

        let mutable eps = 0.000001M

        let checkSums() : decimal [] =
            [| for col in [0 .. range.Values.GetUpperBound(1)] do
                   yield [| for row in [0 .. range.Values.GetUpperBound(0) - 1] do 
                                yield range.Values.[row,col] |] |> Array.reduce (+) |]
        let checkResults () : bool [] =
            [| for col in [0 .. range.Values.GetUpperBound(1)] do 
                   yield range.Values.[range.Values.GetUpperBound(0),col] |]
            |> Array.zip (checkSums())
            |> Array.map (fun ((x : decimal), y) -> System.Math.Abs (x - y) < eps)
        let checkErrors () : CellIndex [] = 
            checkResults()
            |> Array.mapi (fun j x -> j,x)
            |> Array.filter (fun (_,x) -> not x)
            |> Array.map (fun (j,_) -> 
                Label(convertIndex (fst range.LowerRight) (j + (snd range.UpperLeft))))

        new (range : Range) =
            let defaultConversion = function
            | StringTableIndex _ | Date _ | InlineString _ | Empty  -> 0M
            | Decimal x -> x
            let range' : DecimalRange = 
                {  Name = range.Name
                   UpperLeft = range.UpperLeft
                   LowerRight = range.LowerRight
                   Values = range.Values |> Array2D.map (fun x -> (defaultConversion x))  }
            new RangeWithCheckSumsRow (range')

        new (range : Range, conversion) =
            let range' : DecimalRange = 
                {  Name = range.Name
                   UpperLeft = range.UpperLeft
                   LowerRight = range.LowerRight
                   Values = range.Values |> Array2D.map (fun x -> (conversion x))  }
            new RangeWithCheckSumsRow (range')

        new (range : Range, conversion) =
            let range' : DecimalRange = 
                {  Name = range.Name
                   UpperLeft = range.UpperLeft
                   LowerRight = range.LowerRight
                   Values = range.Values |> Array2D.mapi (fun i j x -> (conversion i j x))  }
            new RangeWithCheckSumsRow (range')

        member x.CheckSums = checkSums()
        member x.CheckResults = checkResults()
        member x.CheckErrors = checkErrors()
        member x.Eps with get() = eps and set(e) =  eps <- e


    type RangeWithCheckSumsCol (range : DecimalRange) =
        do
            if range.Values.GetUpperBound(1) < 1 then 
                raise (System.ArgumentOutOfRangeException ("CheckSum range column dimension error"))

        let mutable eps = 0.000001M 

        let checkSums () : decimal [] =
            [| for row in [0 .. range.Values.GetUpperBound(0)] do
                   yield [| for col in [0 .. range.Values.GetUpperBound(1) - 1] do 
                                yield range.Values.[row,col] |] |> Array.reduce (+) |]
        let checkResults () : bool [] =
            [| for row in [0 .. range.Values.GetUpperBound(0)] do 
                   yield range.Values.[row, range.Values.GetUpperBound(1)] |]
            |> Array.zip (checkSums())
            |> Array.map (fun ((x : decimal), y) -> System.Math.Abs (x - y) < eps)
        let checkErrors () : CellIndex [] = 
            checkResults()
            |> Array.mapi (fun i x -> i,x)
            |> Array.filter (fun (_,x) -> not x)
            |> Array.map (fun (i,_) ->  
                Label(convertIndex (i + (fst range.UpperLeft)) (snd range.LowerRight)))

        new (range : Range) =
            let defaultConversion = function
            | StringTableIndex _ | Date _ | InlineString _ | Empty -> 0M
            | Decimal x -> x
            let range' : DecimalRange = 
                {  Name = range.Name
                   UpperLeft = range.UpperLeft
                   LowerRight = range.LowerRight
                   Values = range.Values |> Array2D.map (fun x -> (defaultConversion x))  }
            new RangeWithCheckSumsCol (range')

        new (range : Range, conversion) =
            let range' : DecimalRange = 
                {  Name = range.Name
                   UpperLeft = range.UpperLeft
                   LowerRight = range.LowerRight
                   Values = range.Values |> Array2D.map (fun x -> (conversion x))  }
            new RangeWithCheckSumsCol (range')

        new (range : Range, conversion) =
            let range' : DecimalRange = 
                {  Name = range.Name
                   UpperLeft = range.UpperLeft
                   LowerRight = range.LowerRight
                   Values = range.Values |> Array2D.mapi (fun i j x -> (conversion i j x))  }
            new RangeWithCheckSumsCol (range')
                    
        member x.CheckSums = checkSums ()
        member x.CheckResults = checkResults ()
        member x.CheckErrors = checkErrors ()
        member x.Eps with get() = eps and set(e) = eps <- e

    
    // for backward compatibility
    type Workbook (fileName : string, editable: bool) =        
        let fileFullName = FileFullName(fileName).Value
        member x.FileFullName = fileFullName



    type Sheet (log : Pollux.Log.ILogger, fileName : string, sheetName: string, editable: bool) =
        let sheetName = sheetName
        //let logInfo  = log.LogLine Pollux.Log.LogLevel.Info 
        //let logError = log.LogLine Pollux.Log.LogLevel.Error
        let inlineString  = ref (System.Collections.Generic.Dictionary<int,string>())
        let cellFormula   = ref (System.Collections.Generic.Dictionary<int,string>())
        let extensionList = ref (System.Collections.Generic.Dictionary<int,string>())

        let fCell i x = 
            let test name = 
                let x' = (xd x).Root.Descendants() |> Seq.filter (fun x'' -> x''.Name.LocalName = name)
                if x' |> Seq.isEmpty then "" else x' |> Seq.head |> fun x'' -> x''.Value
            let test' (x': System.Xml.Linq.XAttribute) = if (isNull x' || isNull x'.Value) then "" else x'.Value
            let xa s = test' ((xd x).Root.Attribute(xn s))
            let test2 x (y: System.Collections.Generic.Dictionary<int,string>)  = 
                let z = test x
                if z = "" then -1 
                else y.Add (i, z); i
            let test3 (x: string) = if (xa x) = "" then -1 else x |> xa |> int
            let cv, cvb =     
                if "" = test "v" then -1M,false
                else
                    try (test "v" |> decimal),true
                    with | _ -> 
                        log.LogLine Pollux.Log.LogLevel.Info "fCell: ignoring invalid cell '%s'" x
                        -1M,false
            let rR = xa "r"  |> convertLabel |> fst
            let rC = xa "r"  |> convertLabel |> snd
            ((rR,rC),
                { isCellValueValid   = cvb
                  CellValue          = cv
                  InlineString       = test2 "is" !inlineString
                  CellFormula        = test2 "f" !cellFormula
                  ExtensionList      = test2 "extLst" !extensionList
                  CellMetadataIndex  = test3 "cm"
                  ShowPhonetic       = test3 "ph" 
                  ReferenceRow       = rR
                  ReferenceCol       = rC
                  StyleIndex         = test3 "s"  
                  CellDataType       = if (xa "t") = "" then ' ' else ((xa "t").ToCharArray()).[0]
                  ValueMetadataIndex = test3 "vm" })
            //|> fun ((row,col),cell) -> cells.[(row,col)] <- cell
        
        let cells =
            let partUri = sprintf "/xl/worksheets/sheet%s.xml" (getSheetId log fileName sheetName)
            let xPath = "//*[name()='c']"
            log.LogLine Pollux.Log.LogLevel.Info 
                "Reading cells from %s, sheet %s in part %s:" fileName sheetName partUri
            getPart1 log fileName xPath partUri fCell

        //let sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable
        //let sharedStringItems = sharedStringTable.Elements<SharedStringItem'>()
        let mutable ranges : Range list = []
        
        let rows = []
        let cols = []

        let upperLeft, lowerRight, keys =
            log.LogLine Pollux.Log.LogLevel.Info 
                "%s" "Beginning with upperLeft, lowerRight, keys ..." 
            let keys = cells |> Array.map (fun x -> fst x)
            let minX,maxX,minY,maxY =
                keys 
                |> Array.fold (fun (minX,maxX,minY,maxY) (x,y) -> 
                    min x minX, max x maxX, min y minY, max y maxY) 
                    (System.Int32.MaxValue,System.Int32.MinValue,System.Int32.MaxValue,System.Int32.MinValue)
            Index(minX, minY), Index(maxX,maxY), keys |> Array.map (fun x -> Index(x))

        let numberFormats, cellFormats = 
            log.LogLine Pollux.Log.LogLevel.Info 
                "%s" "upperLeft, lowerRight, keys finished,  beginning with numberFormats ..."
            let partUri = "/xl/styles.xml"
            let numberFormats = 
                let xPath = "//*[name()='numFmt']"
                getPart2 log fileName xPath partUri id2
                |> Seq.map (fun x -> 
                    let test' (x: System.Xml.Linq.XAttribute) = 
                        if (isNull x || isNull x.Value) then "" else x.Value
                    let xa s = test' ((xd x).Root.Attribute(xn s))
                    { NumberFormatId = xa "numFmtId"; FormatCode = xa "formatCode" })
            log.LogLine Pollux.Log.LogLevel.Info 
                "%s" "numberFormats finished,  beginning with cellFormats ..."
            let cellFormats = 
                let xPath = "//*[name()='cellXfs']/*[name()='xf']"
                getPart2 log fileName xPath partUri id2
                |> Seq.mapi (fun i x ->                 
                    let test' (x: System.Xml.Linq.XAttribute) = 
                        if (isNull x || isNull x.Value) then "" else x.Value
                    let xa s = test' ((xd x).Root.Attribute(xn s))
                    i,
                    { NumFmtId          = xa "numFmtId";
                      BorderId          = xa "borderId"
                      FillId            = xa "fillId";
                      FontId            = xa "fontId"; 
                      ApplyAlignment    = xa "applyAlignment";
                      ApplyBorder       = xa "applyBorder";
                      ApplyFont         = xa "applyFont";
                      XfId              = xa "xfId";
                      ApplyNumberFormat = xa "applyNumberFormat" })                             
                |> Map.ofSeq
            numberFormats, cellFormats

        let mutable cellDateTimeFormats = 
            log.LogLine Pollux.Log.LogLevel.Info 
                "%s" "cellFormats finished,  beginning with cellDateTimeFormats ..."
            numberFormats   
            |> Seq.filter (fun x -> x.FormatCode |> isDateTime)
            |> Seq.map (fun x -> x.NumberFormatId)
            |> Seq.append builtInDateTimeNumberFormatIDs 
        
        let isCellDateTimeFormat x =
            if cellFormats.ContainsKey (x) then 
                cellDateTimeFormats
                |> Seq.filter (fun x' -> x' = (cellFormats.[x]).NumFmtId)
                |> Seq.isEmpty
                |> not
            else false

        let values =
            log.LogLine Pollux.Log.LogLevel.Info 
                "%s" "Building values ..."
            let a,a' = Sheet.ConvertCellIndex2 lowerRight
            let b,b' = Sheet.ConvertCellIndex2 upperLeft 
            let evaluate (x' : (int*int)*Cell) =
                //let i,j,x = (fst (fst x'))+b, (snd (fst x'))+b', (snd x')
                //let index = ((i+b),(j+b'))
                
                let x = snd x'
                if x.InlineString > -1 then CellContent.InlineString x.InlineString
                else if x.CellDataType = 's' then 
                    CellContent.StringTableIndex (int x.CellValue)
                else if x.isCellValueValid then 
                    if x.StyleIndex > -1 && isCellDateTimeFormat x.StyleIndex then 
                        CellContent.Date(fromJulianDate (int64 x.CellValue))
                    else CellContent.Decimal(x.CellValue)
                else CellContent.Empty
            let values = 
                array2D [| for _ in [0 .. (a-b)] do 
                                yield [ for _ in [0 .. (a'-b')] do 
                                            yield CellContent.Empty ] |]
            cells
            |> Array.iter (fun x' -> 
                let i,j = (fst (fst x')), (snd (fst x')) 
                try values.[i-b,j-b'] <- (evaluate x')
                with _ -> log.LogLine Pollux.Log.LogLevel.Info " * failed on cell %A" x')
            values

        new (workbook : Workbook, sheetName: string, editable: bool) = 
            Sheet (((new Pollux.Log.PseudoLogger()) :> Pollux.Log.ILogger), workbook.FileFullName, sheetName , editable)

        new (fileName : string, sheetName: string, editable: bool) = 
            Sheet (((new Pollux.Log.PseudoLogger()) :> Pollux.Log.ILogger), fileName, sheetName , editable)

        member x.Rows = rows
        member x.Cols = cols

        member x.UpperLeft = upperLeft
        member x.LowerRight = lowerRight

        member x.Values : CellContent [,] = values

        member x.Ranges = ranges
        member x.Range (i : Index, j : Index) =
            match  ranges |> List.filter (fun r -> r.UpperLeft = i && r.LowerRight = j) with
            | x :: _ -> x
            | _ -> let name = sprintf "%A:%A" (convertIndex2 i) (convertIndex2 j)
                   let range : Range = { Name = name; UpperLeft = i; LowerRight = j; Values = array2D [||] }
                   ranges <- List.append ranges [ range ]; range               
        member x.Range (name) = 
            match  ranges |> List.filter (fun r -> r.Name = name) with
            | x :: _ -> Some x
            | _ -> None

        member x.Cells () = cells
        member x.Cells (a, b) = ()
        member x.Cells (rangeObj: obj) = ()
        member x.Cells (rangeName: string) = ()

        member x.CellDateTimeFormats 
            with get() = cellDateTimeFormats
            and set(dict) = cellDateTimeFormats <- dict

        static member ConvertCellIndex = function
            | Label label -> Index (convertLabel label)
            | Index (x,y) -> Label (convertIndex x y)

        static member ConvertCellIndex2 = function
            | Label label -> convertLabel label
            | Index (x,y) -> x,y

        member x.CellFormats = cellFormats