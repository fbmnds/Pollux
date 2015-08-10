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
              refS   : char [] ref }


        let work state =     
            let isEOF (rs: char [] ref) pos = (pos >= (!rs).Length)
            let testChars (rs: char [] ref) (cs: char []) pos =
                if pos + cs.Length |> isEOF rs then false 
                else
                    cs |> Array.mapi (fun i x -> x = (!rs).[pos+i]) |> Array.filter id |> fun x -> x.Length = cs.Length
            let isOpen1 (rs: char [] ref) pos = testChars rs [|'<';'c';' '|] pos
            let isOpen2 (rs: char [] ref) pos = testChars rs [|'<';'c';'>'|] pos
            let isClose1 (rs: char [] ref) pos = testChars rs [|'/';'>'|] pos
            let isClose2 (rs: char [] ref) pos = testChars rs [|'<';'/';'c';'>'|] pos
            state
            |> function
            | Search x -> 
                if x.cursor |> isEOF x.refS then EOF x.acc
                else if x.cursor |> isOpen1 x.refS then 
                    Open1 { cursor = x.cursor + "<c ".Length; pos1 = x.pos1; acc = x.acc; refS = x.refS }
                else if x.cursor |> isOpen2 x.refS then 
                    Open2 { cursor = x.cursor + "<c>".Length; pos1 = x.pos1; acc = x.acc; refS = x.refS }
                else 
                    Search { cursor = x.cursor + 1; pos1 = x.pos1 + 1; acc = x.acc; refS = x.refS }
            | Open1 x -> 
                if x.cursor |> isEOF x.refS then EOF x.acc
                else if x.cursor |> isClose1 x.refS then 
                    let cursor' = x.cursor + "/>".Length         
                    (!x.acc).Add(x.pos1,cursor')    
                    Search { cursor = cursor' ; pos1 = cursor'; acc = x.acc; refS = x.refS }
                else if x.cursor |> isClose2 x.refS then 
                    let cursor' = x.cursor + "</c>".Length
                    (!x.acc).Add(x.pos1,cursor')
                    Search { cursor = cursor' ; pos1 = cursor'; acc = x.acc; refS = x.refS }
                else
                    Open1 { cursor = x.cursor + 1; pos1 = x.pos1; acc = x.acc; refS = x.refS }
            | Open2 x ->
                if x.cursor |> isEOF x.refS then EOF x.acc
                else if x.cursor |> isClose2 x.refS then 
                    let cursor' = x.cursor + "</c>".Length
                    (!x.acc).Add(x.pos1,cursor')
                    Search { cursor = cursor' ; pos1 = cursor'; acc = x.acc  ; refS = x.refS }
                else
                    Open2 { cursor = x.cursor + 1; pos1 = x.pos1; acc = x.acc; refS = x.refS }
            | EOF acc -> printfn "EOF"; EOF acc

        let parse (xml: char [] ref) =
            let acc = ref (System.Collections.Generic.List<int*int>(10000000))
            let rec loop state =
                match state with
                | EOF acc -> acc
                | _ -> loop (work state)
            loop (Search { cursor = 0; pos1 = 0; acc = acc; refS = xml })
            |> fun x ->
                !x |> Seq.map (fun (x,y) -> new string([| for i in [x .. y-1] do yield (!xml).[i] |]))

