
#r @"../../packages/FAKE/tools/FakeLib.dll"
open Fake
open System
open System.IO
open System.Text.RegularExpressions

//// <summary>
//// Searches a text file for a regex.  Relies on default file encoding.
//// </summary>
//// <param name="filename">the full path to the source text file, no guards for fileexceptions (yet!)</param>
//// <param name="pattern">the regex to search for.  presumed to be a valid regex, no checking done in this function</param>
//// <returns>true in the event the pattern is matched, false otherwise</returns>
let public searchText filename pattern = 
    
    let (|ExistsAtLeastOnce|_|) (pattern: string) (input : string) = 
        let result = Regex.Match(input, pattern)
        if result.Success then
            Some "match"
        else
            None

    let mutable result = false
    let lines = File.ReadAllLines(filename)
    for line in lines do
        match line with 
        | ExistsAtLeastOnce pattern "match" -> result <- true
        | _ -> ()

    result

// Generate
let scriptArgsSearchFile = Environment.GetCommandLineArgs().[2]   // it is passed as the 3rd argument in the command that launches fsi
let scriptArgsSearchPattern = Environment.GetCommandLineArgs().[3]
Fake.TraceHelper.trace scriptArgsSearchFile
Fake.TraceHelper.trace scriptArgsSearchPattern
searchText scriptArgsSearchFile scriptArgsSearchPattern