[<AutoOpen>]
module Pollux.Log


open System


// #region Logging Interface

type LogLevel =
    | Error
    | Warning
    | Info

type ILogger =
    inherit System.IDisposable
    abstract member Log: LogLevel -> Printf.StringFormat<'a, unit> -> 'a
    abstract member LogLine: LogLevel -> Printf.StringFormat<'a, unit> -> 'a

//#endregion


let now level = 
    (System.DateTime.UtcNow.TimeOfDay.Hours, 
     System.DateTime.UtcNow.TimeOfDay.Minutes,
     System.DateTime.UtcNow.TimeOfDay.Seconds) 
    |> fun (h,m,s) -> sprintf "[%02d:%02d:%02d UTC | '%s' ] " h m s level

let mutable private conbg = Console.BackgroundColor
let mutable private confg = Console.ForegroundColor

let cprintf level bg fg fmt =
  Console.Write(now level)
  Printf.kprintf
    (fun s ->
      let restoreBackgroundColor = conbg
      let restoreForegroundColor = confg
      conbg <- bg
      confg <- fg
      Console.Write(s)
      conbg <- restoreBackgroundColor
      confg <- restoreForegroundColor)
    fmt

let cprintfn level bg fg fmt =
  Console.Write(now level)
  Printf.kprintf
    (fun s ->
      let restoreBackgroundColor = conbg
      let restoreForegroundColor = confg
      conbg <- bg
      confg <- fg
      Console.WriteLine(s)
      conbg <- restoreBackgroundColor
      confg <- restoreForegroundColor)
    fmt

/// ILogger log-to-console implementation
type ConsoleLogger() =
  let log level format =
    match level with
    | Error -> cprintf "ERROR" conbg ConsoleColor.Red format
    | Warning -> cprintf "WARN" conbg ConsoleColor.DarkYellow format
    | Info -> cprintf "INFO" conbg confg format

  let logLine level format =
    match level with
    | Error -> cprintfn "ERROR" conbg ConsoleColor.Red format
    | Warning -> cprintfn "WARN" conbg ConsoleColor.DarkYellow format
    | Info -> cprintfn "INFO" conbg confg format

  interface ILogger with
    member x.Log level format = log level format
    member x.LogLine level format = logLine level format
    member x.Dispose() = ()


/// ILogger ignore-log implementation
type DefaultLogger() =
  interface ILogger with
    member x.Log level format = Printf.kprintf (fun s -> ()) format
    member x.LogLine level format = Printf.kprintf (fun s -> ()) format
    member x.Dispose() = ()

let consoleLogger     = (new ConsoleLogger() :> ILogger)
let logError format   = consoleLogger.LogLine LogLevel.Error format
let logWarning format = consoleLogger.LogLine LogLevel.Warning format
let logInfo format    = consoleLogger.Log LogLevel.Info format


