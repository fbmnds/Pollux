namespace Pollux.Excel


#if INTERACTIVE
    open Pollux.Log
    open Pollux.Excel
#endif
    open Pollux.Excel.Utils
    open Pollux.Excel.Cell.Parser
    open System.IO.Packaging

    type LargeSheet (log : Pollux.Log.ILogger, fileName : string, sheetName: string, editable: bool) =
        let sheetName = sheetName
        let logInfo format  = log.LogLine Pollux.Log.LogLevel.Info format
        let logError format  = log.LogLine Pollux.Log.LogLevel.Error format
        
        let sheetString =
            logInfo "%s" "Reading worksheet ..."
            let partUri =  sprintf "/xl/worksheets/sheet%s.xml" (getSheetId log fileName sheetName)
            use xlsx = ZipPackage.Open(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
            let part = 
                xlsx.GetParts()
                |> Seq.filter (fun x -> x.Uri.ToString() = partUri)
                |> Seq.head
            use stream = part.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read)        
            use reader = new System.IO.StreamReader(stream,System.Text.Encoding.UTF8)
            let s = reader.ReadToEnd()
            logInfo "%s" "... reading worksheet done."
            s

        let upperLeft, lowerRight, rowCapacity, colCapacity = 
            try
                logInfo "%s" "Reading sheet dimension ..."
                let s = ref (sheetString.Substring(0,425))
                parseUnsafe 1 "dimension" s
                |> Seq.head
                |> fun x -> 
                    let len = x.Length 
                    (x.Substring(0, len - "\"/>".Length)).Substring("<dimension ref=\"".Length).Split([|':'|])
                |> fun x -> 
                    let upperLeft = Index(CellIndex.ConvertLabel x.[0])
                    let lowerRight = Index(CellIndex.ConvertLabel x.[1])
                    let rowCapacity = (fst (convertCellIndex2 lowerRight)) - (fst (convertCellIndex2 upperLeft)) + 1
                    let colCapacity = (snd (convertCellIndex2 lowerRight)) - (snd (convertCellIndex2 upperLeft)) + 1
                    logInfo "%s" "... reading sheet dimension done."
                    upperLeft, lowerRight, rowCapacity, colCapacity
            with _ -> 
                let msg = sprintf "LargeSheet: could not read 'dimension' of sheet '%s' in '%s'" sheetName fileName
                logError "%s" msg
                failwith msg

        let values = 
            Array2D.createBased<CellContent> 0 0 rowCapacity colCapacity CellContent.Empty            
        let inlineString      = Dict<int,string>()
        let cellFormula       = Dict<int,string>()
        let extensionList     = Dict<int,string>()
        let unknownCellFormat = Dict<int,string>()

        let numberFormats, cellFormats = 
            logInfo "%s" "Reading 'numberFormats' ..."
            let partUri = "/xl/styles.xml"
            let numberFormats = 
                let xPath = "//*[name()='numFmt']"
                getPart2 log fileName xPath partUri id2
                |> Seq.map (fun x -> 
                    let test' (x: System.Xml.Linq.XAttribute) = 
                        if (isNull x || isNull x.Value) then "" else x.Value
                    let xa s = test' ((xd x).Root.Attribute(xn s))
                    { NumberFormatId = xa "numFmtId"; FormatCode = xa "formatCode" })
            logInfo "%s" "... reading 'numberFormats' done."
            logInfo "%s" "Reading 'cellFormats' ..."
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
            logInfo "%s" "... reading 'cellFormats' done."
            numberFormats, cellFormats

        let mutable cellDateTimeFormats = 
            logInfo "%s" "Building 'cellDateTimeFormats'."
            numberFormats   
            |> Seq.filter (fun x -> x.FormatCode |> isDateTime)
            |> Seq.map (fun x -> x.NumberFormatId)
            |> Seq.append builtInDateTimeNumberFormatIDs 
        
        let isCellDateTimeFormat (cellFormats: Map<int,CellFormat>) =
            (fun x ->
                if cellFormats.ContainsKey (x) then 
                    cellDateTimeFormats
                    |> Seq.filter (fun x' -> x' = (cellFormats.[x]).NumFmtId)
                    |> Seq.isEmpty
                    |> not
                else false)

        let cellContentContext =
            { log                   = log
              isCellDateTimeFormat  = isCellDateTimeFormat cellFormats
              rowOffset             = (fst upperLeft.ToTuple)
              colOffset             = (snd upperLeft.ToTuple)
              values                = ref values
              inlineString          = ref inlineString
              cellFormula           = ref cellFormula
              extensionList         = ref extensionList
              unknownCellFormat     = ref unknownCellFormat }            

        //let fCell index outerXml = setCell3 cellContentContext index outerXml
        
        do 
            logInfo "%s" "Parsing cells ..."
            parseCell 10000000 (ref sheetString)
            |> Seq.iteri (fun index outerXml -> setCell3 cellContentContext index outerXml)
            logInfo "%s" "... parsing cells done."            

        //let sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable
        //let sharedStringItems = sharedStringTable.Elements<SharedStringItem'>()
        let mutable ranges : Range list = []
        
        let rows = []
        let cols = []

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

        member x.Values = values

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

        member x.Cells () = ()
        member x.Cells (a, b) = ()
        member x.Cells (rangeObj: obj) = ()
        member x.Cells (rangeName: string) = ()

        member x.CellDateTimeFormats 
            with get() = cellDateTimeFormats
            and set(dict) = cellDateTimeFormats <- dict

        member x.CellFormats = cellFormats
