#r "WindowsBase.dll"
#r "DocumentFormat.OpenXML.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsecCS.dll"
#r @"..\..\Pollux\packages\FParsec.1.0.1\lib\net40-client\FParsec.dll"

open FParsec
#r "System.Xml.Linq.dll"
open System.Xml

open System.IO.Packaging

#time;;
fsi.AddPrinter(fun (x:XmlNode) -> x.OuterXml);;

#load "Log.fs"
#load "Types.fs"
#load "Utils.fs"
#load "Range.fs"
#load "Excel.fs"
#load "CellParser.fs"
#load "Excel2.fs"


open Pollux.Log
open Pollux.Excel
open Pollux.Excel.Utils
open Pollux.Excel.Range
open Pollux.Excel.Cell.Parser

let log = (new ConsoleLogger() :> ILogger)


let ``file6000rows.xlsx`` = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\file6000rows.xlsx"

let ``Übersicht`` = 
    __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\Cost Summary2\xl\worksheets\sheet1.xml"
    |> fun x -> System.IO.File.ReadAllText(x)

//let ``Random`` =
//    __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\file6000rows\xl\worksheets\sheet1.xml"
//    |> fun x -> System.IO.File.ReadAllText(x)


let ``Cost Summary2.xlsx``  = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\Cost Summary2.xlsx"
let ``Cost Summary2_1.txt`` = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\Cost Summary2_1.txt"
let ``Cost Summary2_2.txt`` = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\Cost Summary2_2.txt"
let ``Cost Summary2_3.txt`` = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\Cost Summary2_3.txt"

let sheet = LargeSheet (``Cost Summary2.xlsx``, "Übersicht", false)
let sheet2 = LargeSheet (``Cost Summary2.xlsx``, "CheckSums", false)
let sheet3 = LargeSheet (``Cost Summary2.xlsx``, "CheckSums2", false)



do
    let i2',j2' = sheet2.UpperLeft.ToTuple
    [ for i in [0 .. sheet2.Values.GetUpperBound(0)] do
            for j in [0 .. sheet2.Values.GetUpperBound(1)] do
                yield if sheet2.Values.[i,j] <> CellContent.Empty 
                    then sprintf "%s %A\r\n" (convertIndex (i+i2') (j+j2')) sheet2.Values.[i,j];
                    else "" ]
    |> String.concat ""
    |> printfn "%s"

do
    let i2',j2' = sheet2.UpperLeft.ToTuple
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums"
           UpperLeft  = sheet2.UpperLeft.ToTuple
           LowerRight = sheet2.LowerRight.ToTuple
           Values = sheet2.Values }
    let range = (Pollux.Excel.Range.RangeWithCheckSumsRow (range')).Range
    [ for i in [0 .. range.Values.GetUpperBound(0)] do
            for j in [0 .. range.Values.GetUpperBound(1)] do
                yield sprintf "%s %A\r\n" (convertIndex (i+i2') (j+j2')) range.Values.[i,j]; ]
    |> String.concat ""
    |> printfn "%s"




do
    let sheet2 = LargeSheet (``Cost Summary2.xlsx``, "CheckSums", false)
    let i2',j2' = sheet2.UpperLeft.ToTuple
    let range' : Range = 
        {  Name = "Cost Summary2.xlsx : CheckSums"
           UpperLeft  = sheet2.UpperLeft.ToTuple
           LowerRight = sheet2.LowerRight.ToTuple
           Values = sheet2.Values }
    let range = (RangeWithCheckSumsRow (range')).Range
    [ for i in [0 .. sheet2.Values.GetUpperBound(0)] do
            for j in [0 .. sheet2.Values.GetUpperBound(1)] do
                yield sprintf "%s %A %A\r\n" (convertIndex (i+i2') (j+j2')) sheet2.Values.[i,j] range.Values.[i,j]; ]
    |> String.concat ""
    |> printfn "%s"
    [| for col in [0 .. range.Values.GetUpperBound(1)] do
            yield [| for row in [0 .. range.Values.GetUpperBound(0) - 1] do 
                        yield range.Values.[row,col] |] |> Array.reduce (+) |> sprintf "%A" |]
    |> String.concat "\n"
    |> printfn "%s"    
    printfn ""
    [| for row in [0 .. range.Values.GetUpperBound(0)] do
            yield [| for col in [0 .. range.Values.GetUpperBound(1) - 1] do 
                        yield range.Values.[row,col] |] |> Array.reduce (+) |> sprintf "%A" |]
    |> String.concat "\n"
    |> printfn "%s"    






do
    parseCell 1000 (ref ``Übersicht``)
    |> Seq.iter (printfn "%s")

do
    parseCell 10000000 (ref ``Random``)
    |> Seq.take 5
    |> Seq.iter (printfn "%s")

//    <c r="A1" s="2"><f ca="1">RANDBETWEEN(0,1000)</f><v>437</v></c>
//    <c r="B1" s="2"><f t="shared" ref="B1:BM2" ca="1" si="0">RANDBETWEEN(0,1000)</f><v>358</v></c>
//    <c r="C1" s="2"><f t="shared" ca="1" si="0"/>
//    <c r="D1" s="2"><f t="shared" ca="1" si="0"/>
//    <c r="E1" s="2"><f t="shared" ca="1" si="0"/>
//    Real: 00:06:13.092, CPU: 00:06:07.578, GC gen0: 72786, gen1: 8089, gen2: 3
//    val it : unit = ()


do
    let s = 
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1:H33"/><sheetViews><sheetView showGridLines
        """ 
        |> fun x -> x.Substring(0,425)
        |> ref
    s
    |> (Pollux.Excel.Cell.Parser.parse 1 "dimension")
    |> Seq.head
    |> fun x -> 
        let len = x.Length 
        (x.Substring(0, len - "\"/>".Length)).Substring("<dimension ref=\"".Length).Split([|':'|])
    |> fun x -> CellIndex.ConvertLabel x.[0], CellIndex.ConvertLabel x.[1]
    |> printfn "%A"

do
    let inline isNull x = x = Unchecked.defaultof<_>
    let xn s = System.Xml.Linq.XName.Get(s)
    let xd s = System.Xml.Linq.XDocument.Parse(s)
    let outerXml = """<c r="B1" s="2"><f t="shared" ref="B1:BM2" ca="1" si="0">RANDBETWEEN(0,1000)</f><v>358</v></c>"""
    let test name = 
        let x' = (xd outerXml).Root.Descendants() |> Seq.filter (fun x'' -> x''.Name.LocalName = name)
        if x' |> Seq.isEmpty then "" else x' |> Seq.head |> fun x'' -> x''.Value
    let test' (x': System.Xml.Linq.XAttribute) = if (isNull x' || isNull x'.Value) then "" else x'.Value
    let xa s = test' ((xd outerXml).Root.Attribute(xn s))
//    let test2 x (y: Dict<int,string>)  = 
//        let z = test x
//        if z = "" then -1 
//        else y.Add (index, z); index
//    let test3 (x: string) = if (xa x) = "" then -1 else x |> xa |> int
    let cv, cvb =     
        if "" = test "v" then -1M,false
        else
            try (test "v" |> decimal),true
            with | _ -> 
                //logInfo "setCell: ignoring invalid cell '%s'" outerXml
                -1M,false
    let rR = xa "r"  |> CellIndex.ConvertLabel |> fst
    let rC = xa "r"  |> CellIndex.ConvertLabel |> snd
    printfn "cv '%f' cvb '%b' rR '%d' rC '%d'" cv cvb rR rC

do 
    let findTag (tag: string) (s: string) = s.IndexOf(tag) |> fun x -> if x > -1 then s.Substring(x) else ""
    
    let open1 = "<c "
    let open2 = "<c>"
    let close1 = "/>"
    let close2 = "</c>"
    let findNext (s: string) = 
        match s.IndexOf(open1), s.IndexOf(open2) with
        | -1, -1 -> ""
        | -1, _  -> findTag open2 s
        | _,  -1 -> findTag open1 s
        | _ as i,j -> if i<j then findTag open1 s else findTag open2 s

    let tokenizeNext (s: string) =
        let posClose1, posClose2 = s.IndexOf(close1), s.IndexOf(close2)
        ((posClose1 < 0), (posClose2 < 0))
        |> function
        | (true, true)   -> "",""
        | (true, false)  -> s.Substring(0, posClose2 + close2.Length), s.Substring(posClose2 + close2.Length)
        | (false, true)  -> s.Substring(0, posClose1 + close1.Length), s.Substring(posClose1 + close1.Length)
        | _ ->  
            if posClose1 < posClose2 then 
                s.Substring(0, posClose1 + close1.Length), s.Substring(posClose1 + close1.Length) 
            else
                s.Substring(0, posClose2 + close2.Length), s.Substring(posClose2 + close2.Length)

    let rec loop (s: string) (ss: string list) =
        s |> findNext |> tokenizeNext
        |> function 
        | "", _ -> ss |> List.sort
        | _ as x ->  loop (snd x) ((fst x) :: ss)

    loop ``Übersicht`` []
    |> List.iter (printfn "%s")

//    loop ``Random`` []
//    |> List.iteri (fun i x -> if i < 10 then printfn "%s" x)

    findNext ``Übersicht``
    |> tokenizeNext
    |> fun x -> x |> snd |> findNext
    |> printfn "%A"
    

    findNext ``Übersicht``
    |> findNext
    |> printfn "%s"

type Tag = Bold of string
         | Url of string * string
and TagParserMap = System.Collections.Generic.Dictionary<string,Parser<Tag,UserState>>
and UserState = {
        TagParsers: TagParserMap
        }

do 
    let defaultTagParsers = TagParserMap()

    let isTagNameChar1 = fun c -> isLetter c || c = '_'
    let isTagNameChar = fun c -> isTagNameChar1 c || isDigit c
    let expectedTag = expected "tag starting with '['"

    let tag : Parser<Tag, UserState> =
      fun stream ->
        if stream.Skip('[') then
            let name = stream.ReadCharsOrNewlinesWhile(isTagNameChar1, isTagNameChar,
                                                       false)
            if name.Length <> 0 then
                let mutable p = Unchecked.defaultof<_>
                if stream.UserState.TagParsers.TryGetValue(name, &p) then p stream
                else
                    stream.Skip(-name.Length)
                    Reply(ReplyStatus.Error, messageError ("unknown tag name '" + name + "'"))
            else Reply(ReplyStatus.Error, expected "tag name")
        else Reply(ReplyStatus.Error, expectedTag)

    let str s = pstring s
    let ws = spaces
    let text = manySatisfy (function '['|']' -> false | _ -> true)

    defaultTagParsers.Add("b", str "]" >>. text .>> str "[/b]" |>> Bold)

    defaultTagParsers.Add("url",      (str "=" >>. manySatisfy ((<>)']') .>> str "]")
                                 .>>. (text .>> str "[/url]")
                                 |>> Url)

    let parseTagString str =
        runParserOnString tag {TagParsers = TagParserMap(defaultTagParsers)} "" str

    parseTagString "pretext [b]bold text[/b]  between text [b]bold text2[/b] post text"  
    |>  printfn "%A"

do
    let partUri =  sprintf "/xl/worksheets/sheet%s.xml" (getSheetId log ``file6000rows.xlsx`` "Random")
    use xlsx = ZipPackage.Open(``file6000rows.xlsx``, System.IO.FileMode.Open, System.IO.FileAccess.Read)
    let part = 
        xlsx.GetParts()
        |> Seq.filter (fun x -> x.Uri.ToString() = partUri)
        |> Seq.head
    use stream = part.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read)
    use reader = XmlReader.Create(stream)
    log.LogLine Pollux.Log.LogLevel.Info "%s" "start reader..."
    let i = ref 0
    let result = 
        [| 
            while reader.Read() && !i = 0 do
                if (reader.MoveToContent() = XmlNodeType.Element && reader.Name = "c") then
                    let xml = (reader.ReadOuterXml())
                    if xml.StartsWith("<c") then
                        yield !i, xml
                    else
                        yield !i, sprintf "<c>%s</c>" xml
                    i := !i+1 ;      
            while reader.ReadToFollowing("c")                      && !i < 10  do
                if reader.Name = "c" then
                    let xml = (reader.ReadOuterXml())
                    if xml.StartsWith("<c") then
                        yield !i, xml
                    else
                        yield !i, sprintf "<c>%s</c>" xml
                i := !i+1 
        |]

    log.LogLine Pollux.Log.LogLevel.Info "finished, %d cells in total, take 5 ..." result.Length
    result |> Seq.ofArray |> Seq.take 5 |> printfn "%A"
    log.LogLine Pollux.Log.LogLevel.Info "%s" "finished, build dict ..."
    result |> dict |> Seq.take 5 |> printfn "%A"
    log.LogLine Pollux.Log.LogLevel.Info "%s" "finished, build map ..."
    result |> Map.ofArray |> Seq.take 5 |> printfn "%A"
    log.LogLine Pollux.Log.LogLevel.Info "%s" "finished"

//    [07:07:53 UTC] Beginning 'getPart2' with xPath //*[name()='sheet' and @name='Random'], partUri /xl/workbook.xml
//    [07:07:53 UTC] 'getPart2' with xPath //*[name()='sheet' and @name='Random'], partUri /xl/workbook.xml finished
//    [07:07:53 UTC] start reader...
//    [07:09:28 UTC] finished, 3324554 cells in total, take 5 ...
//    seq
//      [(0,
//        "<c r="A1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f ca="1">RANDBETWEEN(0,1000)</f><v>437</v></c>");
//       (1,
//        "<c r="C1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>175</v></c>");
//       (2,
//        "<c r="E1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>285</v></c>");
//       (3,
//        "<c r="G1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>397</v></c>");
//       ...]
//    [07:09:28 UTC] finished, build dict ...
//    seq
//      [[0, <c r="A1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f ca="1">RANDBETWEEN(0,1000)</f><v>437</v></c>];
//       [1, <c r="C1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>175</v></c>];
//       [2, <c r="E1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>285</v></c>];
//       [3, <c r="G1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>397</v></c>];
//       ...]
//    [07:09:30 UTC] finished, build map ...
//    seq
//      [[0, <c r="A1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f ca="1">RANDBETWEEN(0,1000)</f><v>437</v></c>];
//       [1, <c r="C1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>175</v></c>];
//       [2, <c r="E1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>285</v></c>];
//       [3, <c r="G1" s="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f t="shared" ca="1" si="0" /><v>397</v></c>];
//       ...]
//    [07:09:59 UTC] finished
//    val it : unit = ()
//    > 


do
    let sheet = Sheet ((new ConsoleLogger()), ``file6000rows.xlsx``, "Random", false)
    Pollux.Log.logInfo "%A" sheet.UpperLeft
    Pollux.Log.logInfo "%A" sheet.LowerRight
    //sheet.Cells() |> Map.iter (fun k v -> printfn "%s:\n %A" k v)
    //sheet.CellFormats |> Map.iter (fun k v -> printfn "%d:\n %A" k v)
    printfn "--------"
    sheet.Values
    |> Array2D.iteri (fun i j x -> 
        if x <> CellContent.Empty then 
            printfn "%s %A" (convertIndex i j) x)