namespace Pollux.UnitTests.Excel

module LargeFiles =

    open FsUnit
    open NUnit.Framework


    open Pollux.Excel
    open Pollux.Excel.Utils    

    let ``file6000rows.xlsx``  = __SOURCE_DIRECTORY__ + @"\data\file6000rows.xlsx"
    let ``file6000rows_1.txt``  = __SOURCE_DIRECTORY__ + @"\data\file6000rows_1.txt"

    let sheetRandom = Sheet (``file6000rows.xlsx``, "Random", false)

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeFiles : file6000rows.xlsx : UpperLeft``() =
        sheetRandom.UpperLeft |> should equal (Index(0,0))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeFiles : file6000rows.xlsx : LowerRight``() =
        sheetRandom.LowerRight |> should equal (Index(32,7))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeFiles : file6000rows.xlsx : Values``() =
        let i',j' = sheetRandom.UpperLeft.ToTuple
        [ for i in [0 .. sheetRandom.Values.GetUpperBound(0)] do
              for j in [0 .. sheetRandom.Values.GetUpperBound(1)] do
                  yield if sheetRandom.Values.[i,j] <> CellContent.Empty 
                        then sprintf "%s %A\r\n" (convertIndex (i+i') (j+j')) sheetRandom.Values.[i,j]
                        else "" ]
        |> String.concat ""
        |> should equal (System.IO.File.ReadAllText(``file6000rows_1.txt``))


