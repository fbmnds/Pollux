namespace Pollux

namespace Pollux.Excel
#if INTERACTIVE
    open Pollux.Excel.Utils
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
    | StringTableIndex  of int32
    | InlineString      of string
    | Decimal           of decimal
    | Date              of System.DateTime
    | Empty          


    type Cell = 
        { CellValue          : string 
          InlineString       : string
          CellFormula        : string
          ExtensionList      : string 
          CellMetadataIndex  : string
          ShowPhonetic       : string
          Reference          : CellIndex
          StyleIndex         : string
          CellDataType       : string
          ValueMetadataIndex : string }


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


    type Workbook (fileName : string, editable: bool) =
        
        let fileFullName = FileFullName(fileName).Value
        let sheetDocument = SpreadsheetDocument'.Open(fileFullName, editable)

        let workbookPart = sheetDocument.WorkbookPart

        member x.WorkbookPart = workbookPart
        member x.FileFullName = fileFullName

        interface System.IDisposable with 
            member x.Dispose() = sheetDocument.Dispose()
      

    type Sheet (workbook : Workbook, sheetName: string, editable: bool) =
        let sheetName = sheetName
        
        let workbookPart = workbook.WorkbookPart
        let sheet =
            workbookPart.Workbook.Descendants<Sheet'>()
            |> Seq.filter (fun sheet -> sheet.Name.InnerText = sheetName)
            |> Seq.head
        let sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable
        let sharedStringItems = sharedStringTable.Elements<SharedStringItem'>()
        let mutable ranges : Range list = []
        
        let worksheet = (workbookPart.GetPartById(sheet.Id.Value) :?> WorksheetPart').Worksheet
        let rows = worksheet.Descendants<Row'>() |> Array.ofSeq
        let cols = worksheet.Descendants<Column'>() |> Array.ofSeq
        let upperLeft, lowerRight, cells = 
            let cells : Map<CellIndex, Cell'> ref = ref Map.empty
            rows
            |> Array.map (fun row -> 
                row.Elements<Cell'>()                
                |> Seq.filter (fun cell -> isNotNull cell)
                |> Seq.map (fun cell -> (convertLabel cell.CellReference.Value), cell)
                |> Array.ofSeq)
            |> Array.concat  // unique cell indices
            |> fun cells' ->
                if cells'.Length = 0 then Index(0,0), Index(0,0), !cells 
                else
                    let upperLeft', lowerRight' = ref (fst cells'.[0]), ref (fst cells'.[0])
                    cells'
                    |> Array.iter (fun ((i,j), c) ->
                        upperLeft'  := (min (fst !upperLeft')  i), (min (snd !upperLeft')  j)
                        lowerRight' := (max (fst !lowerRight') i), (max (snd !lowerRight') j)
                        cells := (!cells).Add (Index(i,j), c))
                    Index(!upperLeft'), Index(!lowerRight'), !cells

        let cellFormats = 
            workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Descendants<CellFormat'>() |> Array.ofSeq

        let mutable cellDateTimeFormats = 
            cellFormats   
            |> Array.mapi (fun i x -> 
                workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Descendants<NumberingFormat'>() 
                |> Array.ofSeq
                |> Array.filter (fun x -> x.FormatCode.Value |> isDateTime)
                |> Array.map (fun x -> x.NumberFormatId.Value)
                |> Array.append builtInDateTimeNumberFormatIDs 
                |> Array.map (fun y -> y = x.NumberFormatId.Value) 
                |> Array.fold (fun x' y' -> x' || y') false, i, x)
            |> Array.filter (fun (b, _,_) -> b)
            |> Array.map (fun (_, i, x) -> i, x.NumberFormatId.Value)
            |> Map.ofArray

        let values =
        // https://stackoverflow.com/questions/19034805/how-to-distinguish-inline-numbers-from-ole-automation-date-numbers-in-openxml-sp/19582685
            let values = 
                let a,a' = match lowerRight with | Index(i,j) -> i,j | Label x -> x |> convertLabel 
                let b,b' = match upperLeft  with | Index(i,j) -> i,j | Label x -> x |> convertLabel 
                array2D [| for _ in [0 ..(a-b)] do 
                                yield [ for _ in [0 .. (a'-b')] do yield CellContent.Empty ] |]
            rows
            |> Array.iteri (fun i row -> 
                    row.Elements<Cell'>() 
                    |> Seq.iteri (fun j x -> 
                        if isNull x then 
                            values.[i,j] <- CellContent.Empty
                        else
                            if isNotNull x.DataType then
                                if  x.DataType.Value = CellValues'.SharedString then 
                                    values.[i,j] <- CellContent.StringTableIndex (int32 (x.CellValue.Text))
                                else failwith (sprintf "Data type not covered %A %A" (x.DataType.Value) (x.CellValue.Text))  
                            else 
                                if isNull x.CellValue then
                                    values.[i,j] <- CellContent.Empty
                                else
                                    if isNull x.StyleIndex then
                                        values.[i,j] <- CellContent.Decimal(decimal(x.CellValue.Text))
                                    else
                                        if cellDateTimeFormats.ContainsKey (int x.StyleIndex.Value) then 
                                            values.[i,j] <- CellContent.Date(fromJulianDate (int64 (decimal x.CellValue.Text)))
                                        else 
                                            values.[i,j] <- CellContent.Decimal(decimal(x.CellValue.Text))))                            
            values

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