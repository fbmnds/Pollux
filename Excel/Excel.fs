namespace Pollux

module Excel =

    open Pollux.Excel.Utils

    type private SpreadsheetDocument'   = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument
    type private WorkbookPart'          = DocumentFormat.OpenXml.Packaging.WorkbookPart
    type private WorksheetPart'         = DocumentFormat.OpenXml.Packaging.WorksheetPart
    type private SharedStringTablePart' = DocumentFormat.OpenXml.Packaging.SharedStringTablePart

    type private Sheet'                 = DocumentFormat.OpenXml.Spreadsheet.Sheet
    type private Worksheet'             = DocumentFormat.OpenXml.Spreadsheet.Worksheet
    type private Row'                   = DocumentFormat.OpenXml.Spreadsheet.Row
    type private Column'                = DocumentFormat.OpenXml.Spreadsheet.Column
    type private Cell'                  = DocumentFormat.OpenXml.Spreadsheet.Cell
    type private CellType'              = DocumentFormat.OpenXml.Spreadsheet.CellType
    type private CellFormat'            = DocumentFormat.OpenXml.Spreadsheet.CellFormat
    type private CellValues'            = DocumentFormat.OpenXml.Spreadsheet.CellValues
    type private SharedStringTable'     = DocumentFormat.OpenXml.Spreadsheet.SharedStringTable
    type private SharedStringItem'      = DocumentFormat.OpenXml.Spreadsheet.SharedStringItem
    type private NumberingFormat'       = DocumentFormat.OpenXml.Spreadsheet.NumberingFormat


    type CellIndex = 
    | Label of string
    | Index of Index
    and Index = RowIndex*ColIndex
    and RowIndex = int
    and ColIndex = int
        

    type CellContent =
    | StringTableIndex  of int32
    | Decimal           of decimal
    | Date              of System.DateTime
    | Empty          


    type Range =
        { mutable Name : string 
          UpperLeft    : Index
          LowerRight   : Index
          Values       : CellContent [,] }


    let ConvertCellIndex = function
    | Label label -> Index (convertLabel label)
    | Index (x,y) -> Label (convertIndex x y)


    type Sheet (fileName : string, sheetName: string, editable: bool) =
        let fileFullName = FileFullName(fileName).Value
        let sheetName = sheetName

        let workbookPart = SpreadsheetDocument'.Open(fileFullName, editable).WorkbookPart
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
                        cells := (!cells).Add (Index(i,j), c)
                    )
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

        member x.Rows = rows
        member x.Cols = cols

        member x.UpperLeft = upperLeft
        member x.LowerRight = lowerRight

        member x.Values() : CellContent [,] = 
            // https://stackoverflow.com/questions/19034805/how-to-distinguish-inline-numbers-from-ole-automation-date-numbers-in-openxml-sp/19582685
            let values = 
                let a,a' = match lowerRight with | Index(i,j) -> i,j | Label x -> x |> convertLabel 
                let b,b' = match upperLeft  with | Index(i,j) -> i,j | Label x -> x |> convertLabel 
                array2D [| for i in [0 ..(a-b)] do 
                                yield [ for j in [0 .. (a'-b')] do yield CellContent.Empty ] |]
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
                                            values.[i,j] <- CellContent.Date(fromJulianDate (int64 x.CellValue.Text))
                                        else 
                                            values.[i,j] <- CellContent.Decimal(decimal(x.CellValue.Text))))                            
            values

        member x.Ranges () = ranges
        member x.Range (i : Index, j : Index) =
            match  ranges |> List.filter (fun r -> r.UpperLeft = i && r.LowerRight = j) with
            | x :: _ -> x
            | _ -> let name = sprintf "%A:%A" (convertIndex2 i) (convertIndex2 j)
                   let range = { Name = name; UpperLeft = i; LowerRight = j; Values = array2D [||] }
                   ranges <- List.append ranges [ range ]; range               
        member x.Range (name) = 
            match  ranges |> List.filter (fun r -> r.Name = name) with
            | x :: _ -> Some x
            | _ -> None
        member x.Cells () = ()
        member x.Cells (a, b) = ()
        member x.Cells (rangeObj: obj) = ()
        member x.Cells (rangeName: string) = ()

        member x.CellDateTimeFormats 
            with get() = cellDateTimeFormats
            and set(dict) = cellDateTimeFormats <- dict