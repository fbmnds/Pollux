namespace Pollux.Excel

    
    open FParsec

#if INTERACTIVE    
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

    type Dict<'T1,'T2> = System.Collections.Generic.Dictionary<'T1,'T2>
    type FileFullName (fileName) =
        member x.Value = System.IO.FileInfo(fileName).FullName


    [<CustomEquality; CustomComparison>]
    type CellIndex = 
    | Label of string
    | Index of int*int
        static member ColumnLabel columnIndex =
            let rec loop dividend col = 
                if dividend > 0 then
                    let modulo = (dividend - 1) % 26
                    System.Convert.ToChar(65 + modulo).ToString() + col
                    |> loop ((dividend - modulo) / 26) 
                else 
                    col
            loop (columnIndex + 1) ""

        static member ColumnIndex (columnLabel: string) =
            columnLabel.ToUpper().ToCharArray()
            |> Array.map int
            |> Array.fold (fun (value, i, k)  c ->
                let alphabetIndex = c - 64
                if k = 0 then
                    (value + alphabetIndex - 1, i + 1, k - 1)
                else
                    if alphabetIndex = 0 then
                        (value + (26 * k), i + 1, k - 1)
                    else
                        (value + (alphabetIndex * 26 * k), i + 1, k - 1)
                ) (0, 0, (columnLabel.Length - 1))
            |> fun (value,_,_) -> value 

        static member ConvertLabel (label : string) =
            tuple2 (many1Satisfy  isLetter) (many1Satisfy  isDigit)
            |> fun x -> run x (label.ToUpper())
            |> function
            | Success (x, _, _) ->  System.Int32.Parse(snd x) - 1, CellIndex.ColumnIndex (fst x)
            | _ -> failwith (sprintf "Invalid CellIndex '%s'" label)

        member x.ToTuple : int*int = 
            match x with
            | Label x -> CellIndex.ConvertLabel x
            | Index (x,y) -> x, y

        member x.Row = x.ToTuple |> fst

        member x.Col = x.ToTuple |> snd

        override x.ToString() = 
            match x with 
            | Label x -> x
            | Index (x,y) -> sprintf "%s%d" (CellIndex.ColumnLabel y) (x + 1)

        override x.GetHashCode() = x.GetHashCode()

        override x.Equals(y) =
            match y with
            | :? CellIndex as y -> 
                match y with
                |  Label _ as y' -> y'.ToTuple = x.ToTuple
                |  Index (z,z') -> (z,z') = x.ToTuple
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
                  | Index (a,b) -> 
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
    and CellContentContext =
        { log                  : Pollux.Log.ILogger 
          isCellDateTimeFormat : int -> bool
          rowOffset            : int
          colOffset            : int
          values               : CellContent [,] ref
          inlineString         : Dict<int,string> ref
          cellFormula          : Dict<int,string> ref
          extensionList        : Dict<int,string> ref
          unknownCellFormat    : Dict<int,string> ref }
    
    [<CustomEquality; NoComparison; CLIMutable>]
    type Cell = 
        { isCellValueValid   : bool
          CellValue          : decimal
          InlineString       : int
          CellFormula        : int
          ExtensionList      : int
          UnknownCellFormat  : int
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
          DefinedName  : DefinedName option
          UpperLeft    : CellIndex
          LowerRight   : CellIndex
          Values       : CellContent [,] }
    and StringRange = 
        { mutable Name : string 
          DefinedName  : DefinedName option
          UpperLeft    : CellIndex
          LowerRight   : CellIndex
          Values       : string [,] }
    and DecimalRange = 
        { mutable Name : string 
          DefinedName  : DefinedName option
          UpperLeft    : CellIndex
          LowerRight   : CellIndex
          Values       : decimal [,] }
    and DefinedName =
        { Name         : string
          UpperLeft    : CellIndex
          LowerRight   : CellIndex
          SheetGuid    : System.Guid }
    
    // for backward compatibility
    type Workbook (fileName : string, editable: bool) =        
        let fileFullName = FileFullName(fileName).Value
        member x.FileFullName = fileFullName


    type Agent<'T1> = MailboxProcessor<'T1>