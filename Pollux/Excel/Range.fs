namespace Pollux.Excel

[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module Range =

    open Pollux.Excel
       
#if INTERACTIVE
    open Pollux.Log
    open Pollux.Excel.Utils
#endif

    type RangeWithCheckSumsRow (range : DecimalRange) =

        do
            if range.Values.GetUpperBound(0) <> (fst range.LowerRight) - (fst range.UpperLeft) ||
               range.Values.GetUpperBound(1) <> (snd range.LowerRight) - (snd range.UpperLeft)
             then 
                invalidArg (sprintf "Range %A" range) "CheckSum range column dimension error"

        let mutable eps = 0.000001M

        let checkSums () : decimal [] =
            [| for col in [0 .. range.Values.GetUpperBound(1)] do
                   yield [| for row in [0 .. range.Values.GetUpperBound(0) - 1] do 
                                yield range.Values.[row,col] |] |> Array.reduce (+) |]
        let checkResults () : bool [] =
            [| for col in [0 .. range.Values.GetUpperBound(1)] do 
                   yield range.Values.[range.Values.GetUpperBound(0),col] |]
            |> Array.zip (checkSums ())
            |> Array.map (fun ((x : decimal), y) -> System.Math.Abs (x - y) < eps)
        let checkErrors () : CellIndex [] = 
            checkResults ()
            |> Array.mapi (fun j x -> j,x)
            |> Array.filter (fun (_,x) -> not x)
            |> Array.map (fun (j,_) -> 
                Label(Utils.convertIndex (fst range.LowerRight) (j + (snd range.UpperLeft))))

        new (range : Range, conversion) =
            let range' : DecimalRange = 
                {  Name = range.Name
                   DefinedName = range.DefinedName
                   UpperLeft = range.UpperLeft
                   LowerRight = range.LowerRight
                   Values = range.Values |> Array2D.mapi (fun i j x -> (conversion i j x)) }
            new RangeWithCheckSumsRow (range')

        new (range : Range, conversion) =
            new RangeWithCheckSumsRow (range, (fun _ _ x -> conversion x))

        new (range : Range) =
            let defaultConversion = function
            | StringTableIndex _ | Date _ | InlineString _ | InlineString2 _ | Pending | Empty -> 0M
            | Decimal x -> x
            new RangeWithCheckSumsRow (range, (fun x -> defaultConversion x))
        
        member x.Range = range
        member x.CheckSums = checkSums ()
        member x.CheckResults = checkResults ()
        member x.CheckErrors = checkErrors ()
        member x.Eps with get() = eps and set(e) =  eps <- e


    type RangeWithCheckSumsCol (range : DecimalRange) =

        do
            if range.Values.GetUpperBound(0) <> (fst range.LowerRight) - (fst range.UpperLeft) ||
               range.Values.GetUpperBound(1) <> (snd range.LowerRight) - (snd range.UpperLeft)
             then 
                invalidArg (sprintf "Range %A" range) "CheckSum range column dimension error"

        let mutable eps = 0.000001M 

        let checkSums () : decimal [] =
            [| for row in [0 .. range.Values.GetUpperBound(0)] do
                   yield [| for col in [0 .. range.Values.GetUpperBound(1) - 1] do 
                                yield range.Values.[row,col] |] |> Array.reduce (+) |]
        let checkResults () : bool [] =
            [| for row in [0 .. range.Values.GetUpperBound(0)] do 
                   yield range.Values.[row, range.Values.GetUpperBound(1)] |]
            |> Array.zip (checkSums ())
            |> Array.map (fun ((x : decimal), y) -> System.Math.Abs (x - y) < eps)
        let checkErrors () : CellIndex [] = 
            checkResults ()
            |> Array.mapi (fun i x -> i,x)
            |> Array.filter (fun (_,x) -> not x)
            |> Array.map (fun (i,_) ->  
                Label(Utils.convertIndex (i + (fst range.UpperLeft)) (snd range.LowerRight)))

        new (range : Range, conversion) =
            let range' : DecimalRange = 
                {  Name = range.Name
                   DefinedName = range.DefinedName
                   UpperLeft = range.UpperLeft
                   LowerRight = range.LowerRight
                   Values = range.Values |> Array2D.mapi (fun i j x -> (conversion i j x)) }
            new RangeWithCheckSumsCol (range')

        new (range : Range, conversion) =
            new RangeWithCheckSumsCol (range, (fun _ _ x -> conversion x))

        new (range : Range) =
            let defaultConversion = function
            | StringTableIndex _ | Date _ | InlineString _ | InlineString2 _ | Pending | Empty -> 0M
            | Decimal x -> x
            new RangeWithCheckSumsCol (range, (fun x -> defaultConversion x))

        member x.Range with get() = range                    
        member x.CheckSums = checkSums ()
        member x.CheckResults = checkResults ()
        member x.CheckErrors = checkErrors ()
        member x.Eps with get() = eps and set(e) = eps <- e
