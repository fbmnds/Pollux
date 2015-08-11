namespace Pollux.Excel


#if INTERACTIVE
    open Pollux.Log
    open Pollux.Excel
#endif
    open Pollux.Excel.Utils
    open System.IO.Packaging

    type LargeSheet (log : Pollux.Log.ILogger, fileName : string, sheetName: string, editable: bool) =
        let sheetName = sheetName
        let logInfo format  = log.LogLine Pollux.Log.LogLevel.Info format
        let cellContentContext =
            { log = log
              inlineString      = ref (Dict<int,string>())
              cellFormula       = ref (Dict<int,string>())
              extensionList     = ref (Dict<int,string>())
              unknownCellFormat = ref (Dict<int,string>()) }
        
        // fetch sheet as char array
        let sheetAsString =
            let partUri =  sprintf "/xl/worksheets/sheet%s.xml" (getSheetId log fileName sheetName)
            use xlsx = ZipPackage.Open(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
            let part = 
                xlsx.GetParts()
                |> Seq.filter (fun x -> x.Uri.ToString() = partUri)
                |> Seq.head
            use stream = part.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read)        
            use reader = new System.IO.StreamReader(stream,System.Text.Encoding.UTF8)
            reader.ReadToEnd()
            
        // read dimensions, 425 chars
        // -> upperLeft, lowerRight

        // build values array2D
        let capacity1, capacity2 = 10000, 1000
        let initValue = CellContent.Empty
        let values = ref (Array2D.createBased<CellContent> 0 0 capacity1 capacity2 initValue)

        let fCell index outerXml = setCell3 cellContentContext index outerXml
        
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

        let updateValues =
            logInfo "%s" "Building values ..."
            let a,a' = LargeSheet.ConvertCellIndex2 lowerRight
            let b,b' = LargeSheet.ConvertCellIndex2 upperLeft 
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
            for i in [0 .. (a-b)] do 
                for j in [0 .. (a'-b')] do 
                    (!values).[i,j] <- (evaluate i j)

        new (workbook : Workbook, sheetName: string, editable: bool) = 
            LargeSheet (new Pollux.Log.DefaultLogger(), workbook.FileFullName, sheetName , editable)

        new (fileName : string, sheetName: string, editable: bool) = 
            LargeSheet (new Pollux.Log.DefaultLogger(), fileName, sheetName , editable)

        static member ConvertCellIndex = convertCellIndex

        static member ConvertCellIndex2 = convertCellIndex2

        member x.Rows = rows
        member x.Cols = cols

        member x.UpperLeft = upperLeft
        member x.LowerRight = lowerRight

        member x.Values : CellContent [,] = 
            Array2D.initBased<CellContent> 0 0
                ((fst (convertCellIndex2 lowerRight)) - (fst (convertCellIndex2 upperLeft)) + 1) 
                ((snd (convertCellIndex2 lowerRight)) - (snd (convertCellIndex2 upperLeft)) + 1) 
                (fun i j -> (!values).[i,j])

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

        member x.CellFormats = cellFormats
