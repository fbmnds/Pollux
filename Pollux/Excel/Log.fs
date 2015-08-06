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


let now () = 
    (System.DateTime.UtcNow.TimeOfDay.Hours, 
     System.DateTime.UtcNow.TimeOfDay.Minutes,
     System.DateTime.UtcNow.TimeOfDay.Seconds) 
    |> fun (h,m,s) -> sprintf "[%02d:%02d:%02d UTC] " h m s


let cprintf bg fg fmt =
  Console.Write(now())
  Printf.kprintf
    (fun s ->
      let restoreBackgroundColor = Console.BackgroundColor
      let restoreForegroundColor = Console.ForegroundColor
      Console.BackgroundColor <- bg
      Console.ForegroundColor <- fg
      Console.Write(s)
      Console.BackgroundColor <- restoreBackgroundColor
      Console.ForegroundColor <- restoreForegroundColor)
    fmt

let cprintfn bg fg fmt =
  Console.Write(now())
  Printf.kprintf
    (fun s ->
      let restoreBackgroundColor = Console.BackgroundColor
      let restoreForegroundColor = Console.ForegroundColor
      Console.BackgroundColor <- bg
      Console.ForegroundColor <- fg
      Console.WriteLine(s)
      Console.BackgroundColor <- restoreBackgroundColor
      Console.ForegroundColor <- restoreForegroundColor)
    fmt

/// ILogger log-to-console implementation
type ConsoleLogger() =
  let log level format =
    match level with
    | Error -> cprintf Console.BackgroundColor ConsoleColor.DarkRed format
    | Warning -> cprintf Console.BackgroundColor ConsoleColor.DarkYellow format
    | Info -> cprintf Console.BackgroundColor Console.ForegroundColor format

  let logLine level format =
    match level with
    | Error -> cprintfn Console.BackgroundColor ConsoleColor.DarkRed format
    | Warning -> cprintfn Console.BackgroundColor ConsoleColor.DarkYellow format
    | Info -> cprintfn Console.BackgroundColor Console.ForegroundColor format

  interface ILogger with
    member x.Log level format = log level format
    member x.LogLine level format = logLine level format
    member x.Dispose() = ()


/// ILogger do-not-log implementation
type PseudoLogger() =
  interface ILogger with
    member x.Log level format = Printf.kprintf (fun s -> ()) format
    member x.LogLine level format = Printf.kprintf (fun s -> ()) format
    member x.Dispose() = ()

let consoleLogger = (new ConsoleLogger() :> ILogger)
let logError format = consoleLogger.LogLine LogLevel.Error format
let logWarning format = consoleLogger.LogLine LogLevel.Warning format
let logInfo format = consoleLogger.Log LogLevel.Info format


