namespace Pollux.UnitTests

module ExcelUnitTests =
    
    let ``Cost Summary2.xlsx`` = __SOURCE_DIRECTORY__ + @"\data\Cost Summary2.xlsx"
    let ``Cost Summary2_1.txt`` = __SOURCE_DIRECTORY__ + @"\data\Cost Summary2_1.txt"

    open FsUnit
    open FsCheck
    open NUnit.Framework
    open NUnit.Framework.Constraints
    open Swensen.Unquote

    open Pollux.Excel
    open Pollux.Excel.Utils

    let workbook = new Workbook (``Cost Summary2.xlsx``, false)
    let sheet = Sheet (workbook, "Übersicht", false)
        
    [<Test; Category "Pollux.Excel">]
    let ``Cost Summary2.xlsx: Sheet.UpperLeft``() =
        sheet.UpperLeft |> should equal (Index(0,0))

    [<Test; Category "Pollux.Excel">]
    let ``Cost Summary2.xlsx: Sheet.LowerRight``() =
        sheet.LowerRight |> should equal (Index(32,7))

    [<Test; Category "Pollux.Excel">]
    let ``Cost Summary2.xlsx: Sheet.Values``() =
        [ for i in [0 .. sheet.Values.GetUpperBound(0)] do
              for j in [0 .. sheet.Values.GetUpperBound(1)] do
                  yield if sheet.Values.[i,j] <> CellContent.Empty 
                        then sprintf "%s %A\r\n" (convertIndex i j) sheet.Values.[i,j];
                        else "" ]
        |> String.concat ""
        |> fun x -> printfn "%s" x; x
        |> should equal (System.IO.File.ReadAllText(``Cost Summary2_1.txt``))

    
