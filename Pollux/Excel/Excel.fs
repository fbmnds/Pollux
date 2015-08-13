namespace Pollux.Excel


#if INTERACTIVE
    open Pollux.Log
    open Pollux.Excel
#endif
    open Pollux.Excel.Utils
    open Pollux.Excel.Cell.Parser
    open System.IO.Packaging

    type LargeSheet (log : Pollux.Log.ILogger, fileName : string, sheetName: string, editable: bool) =
        let sheetGuid = System.Guid.NewGuid()
        let logInfo  format  = log.LogLine Pollux.Log.LogLevel.Info  format
        let logError format  = log.LogLine Pollux.Log.LogLevel.Error format
        let dimensionSearchRange = 1000
        let largeSheetCapacityUnit = 10000000

        let sheetString =
            logInfo "%s" "LargeSheet : reading worksheet ..."
            let partUri =  sprintf "/xl/worksheets/sheet%s.xml" (getSheetId log fileName sheetName)
            use xlsx = ZipPackage.Open(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
            let part = 
                xlsx.GetParts()
                |> Seq.filter (fun x -> x.Uri.ToString() = partUri)
                |> Seq.head
            use stream = part.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read)        
            use reader = new System.IO.StreamReader(stream,System.Text.Encoding.UTF8)
            let s = reader.ReadToEnd()
            logInfo "%s" "LargeSheet : reading worksheet done"
            s

        let upperLeft, lowerRight, rowCapacity, colCapacity = 
            getDimensions log fileName sheetName (sheetString.Substring(0,dimensionSearchRange))

        let values = 
            Array2D.createBased<CellContent> 0 0 rowCapacity colCapacity CellContent.Empty            
        let inlineString      = Dict<int,string>()
        let inlineString2     = Dict<int,string>()
        let cellFormula       = Dict<int,string>()
        let extensionList     = Dict<int,string>()
        let unknownCellFormat = Dict<int,string>()

        let numberFormats = getNumberFormats log fileName

        let cellFormats = getCellFormats log fileName

        let mutable cellDateTimeFormats = GetCellDateTimeFormats numberFormats
        
        let isCellDateTimeFormat = fIsCellDateTimeFormat cellFormats cellDateTimeFormats

        let cellContentContext =
            { log                   = log
              isCellDateTimeFormat  = isCellDateTimeFormat
              rowOffset             = (fst upperLeft.ToTuple)
              colOffset             = (snd upperLeft.ToTuple)
              values                = ref values
              inlineString          = ref inlineString
              inlineString2         = ref inlineString2
              cellFormula           = ref cellFormula
              extensionList         = ref extensionList
              unknownCellFormat     = ref unknownCellFormat }            
       
        do 
            let parseAgent = 
                new Agent<CellContentContext*string>(fun x -> 
                    let index' = ref -1
                    let rec loop () =
                        async { let index = System.Threading.Interlocked.Increment(index')
                                let! ctx,outerXml = x.Receive()
                                do setCell3 ctx index outerXml
                                return! loop () }
                    loop () )
            logInfo "%s" "LargeSheet : parsing cells ..."
            parseAgent.Start()
            parseCell largeSheetCapacityUnit (ref sheetString)
            |> Seq.iter (fun outerXml -> parseAgent.Post (cellContentContext,outerXml))
            while parseAgent.CurrentQueueLength > 0 do ()
            logInfo "%s" "LargeSheet : parsing cells done"            

        //let sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable
        //let sharedStringItems = sharedStringTable.Elements<SharedStringItem'>()
        let definedNames = getDefinedNames log sheetGuid fileName

        let ranges = getDefinedNames log sheetGuid fileName
        
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
        member x.RangeValues rangeName = 
            match rangeName |> ranges.TryFind with
            | Some range -> 
                let a,a' = (fst range.LowerRight),(snd range.LowerRight)
                let b,b' = (fst range.UpperLeft), (snd range.UpperLeft)
                printfn "a %d a' %d b %d b' %d upperLeft %A lowerRight %A" a a' b b' upperLeft lowerRight
                //if true then
                if a>=b && a'>=b' && 
                   b>=upperLeft.Row && b'>=upperLeft.Col &&
                   a-upperLeft.Row <= values.GetUpperBound(0) && a'-upperLeft.Col <= values.GetUpperBound(1) then 
                       array2D [| for i in [b .. a] do 
                                      yield [ for j in [b' .. a'] do 
                                                  yield values.[i-upperLeft.Row,j-upperLeft.Col] ] |]
                    |> Some
                else None
            | _ -> None             

        member x.Cells () = ()
        member x.Cells (a, b) = ()
        member x.Cells (rangeObj: obj) = ()
        member x.Cells (rangeName: string) = ()

        member x.CellDateTimeFormats 
            with get() = cellDateTimeFormats
            and set(dict) = cellDateTimeFormats <- dict

        member x.CellFormats = cellFormats
