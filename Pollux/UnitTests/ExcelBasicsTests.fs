namespace Pollux.UnitTests.Excel

module Basics =

    open FsUnit
    open NUnit.Framework


    open Pollux.Excel
    open Pollux.Excel.Utils   
    open Pollux.Excel.Range
 

    [<Test; Category "Pollux.Excel.Utils">]
    let ``Excel : Utils : ColumnLabel``() =
        [0; 7] |> List.map CellIndex.ColumnLabel |> should equal ["A"; "H"]

    [<Test; Category "Pollux.Excel.Utils">]
    let ``Excel : Utils : ColumnIndex``() =
        ["A"; "H"] |> List.map CellIndex.ColumnIndex |> should equal [0; 7]

    [<Test; Category "Pollux.Excel.Utils">]
    let ``Excel : Utils : convertLabel``() =
        ["A1"; "H29"] |> List.map CellIndex.ConvertLabel |> should equal [(0,0); (28,7)]

    [<Test; Category "Pollux.Excel.Utils">]
    let ``Excel : Utils : convertIndex2``() =
        [(0,0); (28,7)] |> List.map convertIndex2 |> should equal ["A1"; "H29"]
        
