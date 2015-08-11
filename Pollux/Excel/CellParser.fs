namespace Pollux.Excel

[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module Cell =

    module Parser =

        type State = 
        | Search       of Result
        | Open1        of Result
        | Open2        of Result
        | EOF          of System.Collections.Generic.List<int*int> ref
        and Result =
            { cursor : int
              pos1   : int
              acc    : System.Collections.Generic.List<int*int> ref
              refS   : string ref }


        let work tag =     
            let isEOF (rs: string ref) pos = (pos >= (!rs).Length)
            let testChars (rs: string ref) (cs: string) pos =                 
                if pos + cs.Length |> isEOF rs then false 
                else(!rs).Substring(pos, cs.Length) = cs
            let open1 = sprintf "<%s " tag
            let open2 = sprintf "<%s>" tag
            let close1 = "/>"
            let close2 = sprintf "</%s>" tag
            let isOpen1 (rs: string ref) pos = testChars rs open1 pos
            let isOpen2 (rs: string ref) pos = testChars rs open2 pos
            let isClose1 (rs: string ref) pos = testChars rs close1 pos
            let isClose2 (rs: string ref) pos = testChars rs close2 pos
            (fun state -> 
                state
                |> function
                | Search x -> 
                    if x.cursor |> isEOF x.refS then EOF x.acc
                    else if x.cursor |> isOpen1 x.refS then 
                        Open1 { cursor = x.cursor + open1.Length; pos1 = x.pos1; acc = x.acc; refS = x.refS }
                    else if x.cursor |> isOpen2 x.refS then 
                        Open2 { cursor = x.cursor + open2.Length; pos1 = x.pos1; acc = x.acc; refS = x.refS }
                    else 
                        Search { cursor = x.cursor + 1; pos1 = x.pos1 + 1; acc = x.acc; refS = x.refS }
                | Open1 x -> 
                    if x.cursor |> isEOF x.refS then EOF x.acc
                    else if x.cursor |> isClose1 x.refS then 
                        let cursor' = x.cursor + close1.Length         
                        (!x.acc).Add(x.pos1,cursor')    
                        Search { cursor = cursor' ; pos1 = cursor'; acc = x.acc; refS = x.refS }
                    else if x.cursor |> isClose2 x.refS then 
                        let cursor' = x.cursor + close2.Length
                        (!x.acc).Add(x.pos1,cursor')
                        Search { cursor = cursor' ; pos1 = cursor'; acc = x.acc; refS = x.refS }
                    else
                        Open1 { cursor = x.cursor + 1; pos1 = x.pos1; acc = x.acc; refS = x.refS }
                | Open2 x ->
                    if x.cursor |> isEOF x.refS then EOF x.acc
                    else if x.cursor |> isClose2 x.refS then 
                        let cursor' = x.cursor + close2.Length
                        (!x.acc).Add(x.pos1,cursor')
                        Search { cursor = cursor' ; pos1 = cursor'; acc = x.acc  ; refS = x.refS }
                    else
                        Open2 { cursor = x.cursor + 1; pos1 = x.pos1; acc = x.acc; refS = x.refS }
                | EOF acc -> printfn "EOF"; EOF acc)

        let parseUnsafe (capacity: int) (tag: string) (xml: string ref) =
            let acc = ref (System.Collections.Generic.List<int*int>(capacity))
            let w = (work tag)
            let rec loop state =
                match state with
                | EOF acc -> acc
                | _ -> loop (w state)
            loop (Search { cursor = 0; pos1 = 0; acc = acc; refS = xml })
            |> fun x ->
                !x |> Seq.map (fun (x,y) -> new string([| for i in [x .. y-1] do yield (!xml).[i] |]))

        let workCell =     
            let isEOF (rs: string ref) pos = (pos >= (!rs).Length)
            let testChars (rs: string ref) (cs: string) pos =                 
                if pos + cs.Length |> isEOF rs then false 
                else(!rs).Substring(pos, cs.Length) = cs
            let open1 = "<c "
            let open2 = "<c>"
            let close1 = open1
            let close2 = open2
            let close3 = "</row>"
            let isOpen1 (rs: string ref) pos = testChars rs open1 pos
            let isOpen2 (rs: string ref) pos = testChars rs open2 pos
            let isClose1 (rs: string ref) pos = testChars rs close1 pos
            let isClose2 (rs: string ref) pos = testChars rs close2 pos
            let isClose3 (rs: string ref) pos = testChars rs close3 pos
            (fun state -> 
                state
                |> function
                | Search x -> 
                    if x.cursor |> isEOF x.refS then EOF x.acc
                    else if x.cursor |> isOpen1 x.refS then 
                        Open2 { cursor = x.cursor + open1.Length; pos1 = x.pos1; acc = x.acc; refS = x.refS }
                    else if x.cursor |> isOpen2 x.refS then 
                        Open2 { cursor = x.cursor + open2.Length; pos1 = x.pos1; acc = x.acc; refS = x.refS }
                    else 
                        Search { cursor = x.cursor + 1; pos1 = x.pos1 + 1; acc = x.acc; refS = x.refS }
                | Open1 x -> failwith "never see this"
                | Open2 x -> 
                    if x.cursor |> isEOF x.refS then EOF x.acc
                    else if x.cursor |> isClose1 x.refS || 
                            x.cursor |> isClose2 x.refS || 
                            x.cursor |> isClose3 x.refS then 
                        (!x.acc).Add(x.pos1,x.cursor)
                        Search { cursor = x.cursor; pos1 = x.cursor; acc = x.acc; refS = x.refS }
                    else
                        Open2 { cursor = x.cursor + 1; pos1 = x.pos1; acc = x.acc; refS = x.refS }
                | EOF acc -> EOF acc)

        let parseCell (capacity: int) (xml: string ref) =
            let acc = ref (System.Collections.Generic.List<int*int>(capacity))
            let rec loop state =
                match state with
                | EOF acc -> acc
                | _ -> loop (workCell state)
            loop (Search { cursor = 0; pos1 = 0; acc = acc; refS = xml })
            |> fun x ->
                !x |> Seq.map (fun (x,y) -> new string([| for i in [x .. y-1] do yield (!xml).[i] |]))