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

        let cellFormats = 
            workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Descendants<CellFormat'>() |> Array.ofSeq
        let numberingFormats = 
            workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Descendants<NumberingFormat'>() |> Array.ofSeq
        let dateTimeFormats = 
            numberingFormats 
            |> Array.filter (fun x -> x.FormatCode.Value |> isDateTime)
            |> Array.map (fun x -> x.NumberFormatId.Value)
            |> Array.append builtInDateTimeNumberFormatIDs
        let cellDateTimeFormats = 
            cellFormats   
            |> Array.mapi (fun i x -> 
                dateTimeFormats 
                |> Array.map (fun y -> y = x.NumberFormatId.Value) 
                |> Array.fold (fun x' y' -> x' || y') false, i, x)
            |> Array.filter (fun (b, _,_) -> b)
            |> Array.map (fun (_, i, x) -> i, x.NumberFormatId.Value)
            |> Map.ofArray

        member x.Rows = worksheet.Descendants<Row'>() |> Array.ofSeq
        member x.Cols = worksheet.Descendants<Column'>() |> Array.ofSeq

        member x.UpperLeft  = rows.[0].Elements<Cell'>() |> Seq.head |> fun x -> x.LocalName
        member x.LowerRight = cols.[cols.Length-1].LocalName + rows.[rows.Length-1].LocalName

        member x.Values() : CellContent [,] = 
            // https://stackoverflow.com/questions/19034805/how-to-distinguish-inline-numbers-from-ole-automation-date-numbers-in-openxml-sp/19582685
            let values = array2D [| for i in [0 ..50] do yield [ for j in [0 .. 50] do yield CellContent.Empty ] |]
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
    

 // http://stackoverflow.com/questions/19034805/how-to-distinguish-inline-numbers-from-ole-automation-date-numbers-in-openxml-sp/19582685#19582685
 (*
 public class ExcelHelper
{
    static uint[] builtInDateTimeNumberFormatIDs = new uint[] { 14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58 };
    static Dictionary<uint, NumberingFormat> builtInDateTimeNumberFormats = builtInDateTimeNumberFormatIDs.ToDictionary(id => id, id => new NumberingFormat { NumberFormatId = id });
    static Regex dateTimeFormatRegex = new Regex(@"((?=([^[]*\[[^[\]]*\])*([^[]*[ymdhs]+[^\]]*))|.*\[(h|mm|ss)\].*)", RegexOptions.Compiled);

    public static Dictionary<uint, NumberingFormat> GetDateTimeCellFormats(WorkbookPart workbookPart)
    {
        var dateNumberFormats = workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats
            .Descendants<NumberingFormat>()
            .Where(nf => dateTimeFormatRegex.Match(nf.FormatCode.Value).Success)
            .ToDictionary(nf => nf.NumberFormatId.Value);

        var cellFormats = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats
            .Descendants<CellFormat>();

        var dateCellFormats = new Dictionary<uint, NumberingFormat>();
        uint styleIndex = 0;
        foreach (var cellFormat in cellFormats)
        {
            if (cellFormat.ApplyNumberFormat != null && cellFormat.ApplyNumberFormat.Value)
            {
                if (dateNumberFormats.ContainsKey(cellFormat.NumberFormatId.Value))
                {
                    dateCellFormats.Add(styleIndex, dateNumberFormats[cellFormat.NumberFormatId.Value]);
                }
                else if (builtInDateTimeNumberFormats.ContainsKey(cellFormat.NumberFormatId.Value))
                {
                    dateCellFormats.Add(styleIndex, builtInDateTimeNumberFormats[cellFormat.NumberFormatId.Value]);
                }
            }

            styleIndex++;
        }

        return dateCellFormats;
    }

    // Usage Example
    public static bool IsDateTimeCell(WorkbookPart workbookPart, Cell cell)
    {
        if (cell.StyleIndex == null)
            return false;

        var dateTimeCellFormats = ExcelHelper.GetDateTimeCellFormats(workbookPart);

        return dateTimeCellFormats.ContainsKey(cell.StyleIndex);
    }
}
*)