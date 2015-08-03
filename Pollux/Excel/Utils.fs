

[<AutoOpen>]
module Pollux.Excel.Utils

open FParsec


let inline isNull x = x = Unchecked.defaultof<_>
let inline isNotNull x = x |> isNull |> not

let inline (|IsNull|) x = isNull x 
       
type FileFullName (fileName) =
    member x.Value = System.IO.FileInfo(fileName).FullName 


let inline ColumnLabel columnIndex =
    let rec loop dividend col = 
        if dividend > 0 then
            let modulo = (dividend - 1) % 26
            System.Convert.ToChar(65 + modulo).ToString() + col
            |> loop ((dividend - modulo) / 26) 
        else 
            col
    loop (columnIndex + 1) ""


let inline ColumnIndex (columnLabel: string) =
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


let inline convertLabel (label : string) =
    tuple2 (many1Satisfy  isLetter) (many1Satisfy  isDigit)
    |> fun x -> run x (label.ToUpper())
    |> function
    | Success (x, _, _) ->  System.Int32.Parse(snd x) - 1, ColumnIndex (fst x)
    | _ -> failwith (sprintf "Invalid CellIndex '%s'" label)


let inline convertIndex x y = sprintf "%s%d" (ColumnLabel y) (x + 1)
let inline convertIndex2 (x : int*int) = convertIndex (fst x) (snd x)


let rec isDateTime (s : string) =
    run (anyOf "ymdhs:") s
    |> function
    | Success _ ->  true
    | _ -> if  s = "" then false else isDateTime (s.Substring 1)

let builtInDateTimeNumberFormatIDs = 
    [| 14u; 15u; 16u; 17u; 18u; 19u;
        20u; 21u; 22u; 27u; 28u; 29u; 
        30u; 31u; 32u; 33u; 34u; 35u; 36u;
        45u; 46u; 47u; 50u;
        51u; 52u; 53u; 54u; 55u; 56u; 57u; 58u |]
        
let fromJulianDate x = 
    // System.DateTime.Parse("30.12.1899").Ticks = 599264352000000000L
    // System.TimeSpan.TicksPerDay = 864000000000L
    System.DateTime(599264352000000000L + (864000000000L * x)) 

let toJulianDate (x : System.DateTime) =
    (x.ToBinary() - 599264352000000000L) / 864000000000L