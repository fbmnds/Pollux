namespace Pollux.Excel

    open Pollux.Excel.Utils

    type Sheet (log : Pollux.Log.ILogger, fileName : string, sheetName: string, editable: bool) =
        let sheetName = sheetName
        let logInfo format = log.LogLine Pollux.Log.LogLevel.Info format
        let logError format = log.LogLine Pollux.Log.LogLevel.Error format
        let inlineString  = ref (Dict<int,string>())
        let cellFormula   = ref (Dict<int,string>())
        let extensionList = ref (Dict<int,string>())

        let fCell i x = setCell i x log inlineString cellFormula extensionList 
        
        let cells =
            let partUri = sprintf "/xl/worksheets/sheet%s.xml" (getSheetId log fileName sheetName)
            let xPath = "//*[name()='c']"
            logInfo "Reading cells from %s, sheet %s in part %s:" fileName sheetName partUri
            getPart1' log fileName xPath partUri fCell
            |> dict

        //let sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable
        //let sharedStringItems = sharedStringTable.Elements<SharedStringItem'>()
        let mutable ranges : Range list = []
        
        let rows = []
        let cols = []

        let upperLeft, lowerRight, keys =
            logInfo "%s" "Beginning with upperLeft, lowerRight, keys ..." 
            let keys = cells.Keys
            let minX,maxX,minY,maxY =
                keys 
                |> Seq.fold (fun (minX,maxX,minY,maxY) (x,y) -> 
                    min x minX, max x maxX, min y minY, max y maxY) 
                    (System.Int32.MaxValue,System.Int32.MinValue,System.Int32.MaxValue,System.Int32.MinValue)
            Index(minX, minY), Index(maxX,maxY), keys |> Seq.map (fun x -> Index(x))

        let numberFormats, cellFormats = 
            logInfo "%s" "upperLeft, lowerRight, keys finished,  beginning with numberFormats ..."
            let partUri = "/xl/styles.xml"
            let numberFormats = 
                let xPath = "//*[name()='numFmt']"
                getPart2 log fileName xPath partUri id2
                |> Seq.map (fun x -> 
                    let test' (x: System.Xml.Linq.XAttribute) = 
                        if (isNull x || isNull x.Value) then "" else x.Value
                    let xa s = test' ((xd x).Root.Attribute(xn s))
                    { NumberFormatId = xa "numFmtId"; FormatCode = xa "formatCode" })
            logInfo "%s" "numberFormats finished,  beginning with cellFormats ..."
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
            logInfo "%s" "cellFormats finished,  beginning with cellDateTimeFormats ..."
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
            logInfo "%s" "Building values ..."
            let a,a' = Sheet.ConvertCellIndex2 lowerRight
            let b,b' = Sheet.ConvertCellIndex2 upperLeft 
            let evaluate i j =
                let index = i+b, j+b'
                if cells.ContainsKey(index) then
                    let x = cells.[index]
                    if x.InlineString > -1 then CellContent.InlineString x.InlineString
                    else if x.CellDataType = 's' then 
                        CellContent.StringTableIndex (int x.CellValue)
                    else if x.isCellValueValid then 
                        if x.StyleIndex > -1 && isCellDateTimeFormat x.StyleIndex then 
                            CellContent.Date(fromJulianDate (int64 x.CellValue))
                        else CellContent.Decimal(x.CellValue)
                    else CellContent.Empty
                else CellContent.Empty
            array2D [| for i in [0 .. (a-b)] do 
                            yield [ for j in [0 .. (a'-b')] do 
                                       yield (evaluate i j) ] |]

        new (workbook : Workbook, sheetName: string, editable: bool) = 
            Sheet (new Pollux.Log.DefaultLogger(), workbook.FileFullName, sheetName , editable)

        new (fileName : string, sheetName: string, editable: bool) = 
            Sheet (new Pollux.Log.DefaultLogger(), fileName, sheetName , editable)

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
            | Label label -> Index (CellIndex.ConvertLabel label)
            | Index (x,y) -> Label (convertIndex x y)

        static member ConvertCellIndex2 = function
            | Label label -> CellIndex.ConvertLabel label
            | Index (x,y) -> x,y

        member x.CellFormats = cellFormats