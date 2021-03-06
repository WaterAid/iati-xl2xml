// --------------------------------------------------------------------------------------
// FAKE build script
// --------------------------------------------------------------------------------------

#r @"packages/FAKE/tools/FakeLib.dll"
#r @"lib/VbaMd.dll"

open VbaMd.Main
open Fake
open Fake.Git
open Fake.ReleaseNotesHelper
open Fake.UserInputHelper
open System
open System.IO

// The name of the project
// (used by attributes in AssemblyInfo, name of a NuGet package and directory in 'src')
let project = "iati-xl2xml"

// Short summary of the project
// (used as description in AssemblyInfo and as a short summary for NuGet package)
let summary = "Aid project data reporting."

// Longer description of the project
// (used as a description for NuGet package; line breaks are automatically cleaned up)
let description = "Aid project data reporting to the IATI standard."

// File system information 
let mutable solutionFile  = "xl2xml.xlsm"
let mutable solutionPath = @"C:\mark\excel\iati-xl2xml\src\"

// Git configuration (used for publishing documentation in gh-pages branch)
// The profile where the project is posted
let gitOwner = "WaterAid" 
let gitHome = "https://github.com/" + gitOwner

// The name of the project on GitHub
let gitName = "iati-xl2xml"

// The url for the raw files hosted
let gitRaw = environVarOrDefault "gitRaw" "https://raw.github.com/wateraid"

// Read additional information from the release notes document
let release = LoadReleaseNotes "RELEASE_NOTES.md"

// --------------------------------------------------------------------------------------
// Clean build results
Target "CleanSource" (fun _ ->
    !! (solutionPath + "*.*") -- (solutionPath + "*.xlsm")
        |> Seq.iter DeleteFile
)

Target "CleanDocs" (fun _ ->
    CleanDirs ["docs/output"]
)

// --------------------------------------------------------------------------------------
// Validate the application against the OpenXml schema
let getValidationErrors (filename:string) = 
    traceFAKE "starting OpenXml validation of source document"
    let outcome = executeFSIWithScriptArgsAndReturnMessages "docs/tools/openxmlprojecthelper.fsx" [|filename|]
    match outcome with
    | (true, _) -> trace "passed OpenXml validation"
    | (false, _) -> traceError "failed OpenXml validation with these errors" 
    ()
    traceFAKE "ending OpenXml validation of source document"

Target "ValidateSourceDocument" (fun _ ->
    let path = getUserInput "Enter the absolute path to the spreadsheet source file or leave blank to accept the default: "
    match System.String.IsNullOrEmpty(path), System.String.IsNullOrWhiteSpace(path) with
    | (true, false) -> solutionFile <- path
    | (false, true) -> solutionFile <- path
    | (false, false) -> solutionFile <- path
    | (_,_) -> () 

    match Fake.FileSystemHelper.fileExists(solutionFile) with
    | false -> traceError "file does not exist" 
    | true -> getValidationErrors solutionFile
)

//---------------------------------------------------------------------------------------
// Extract the source files from the application
let getSource (filename: string) (docDirectory: string) = 
    let outcome = executeFSIWithScriptArgsAndReturnMessages "docs/tools/vbaprojecthelper.fsx" [|filename;docDirectory|]
    match outcome with
    | (true, _) -> trace "Source files successfully extracted."
    | (false, _) -> snd(outcome)|> Seq.iter (fun f -> traceError f.Message)
    ()

Target "ExtractSource" (fun _ ->
   let path = getUserInput "Enter the absolute path to the solution directory or leave blank to accept the default: "
   match String.IsNullOrWhiteSpace(path), String.IsNullOrEmpty(path) with
   | (false, false) -> 
       match Fake.FileSystemHelper.isValidPath(path), Fake.FileSystemHelper.isDirectory(path) with
       | (true, true) -> solutionPath <-path
       | _ -> ()
    | (_,_) -> ()
   Fake.TraceHelper.trace solutionPath
   // presume the default project structure is preserved and calculate the full path to the api doc directory
   let apiDocPath = Fake.EnvironmentHelper.combinePaths solutionPath @"..\docs\content\api"
   getSource (solutionPath + solutionFile) apiDocPath
)

//---------------------------------------------------------------------------------------
// Extract the markdown from the source files
let getPublicComments (filename: string) (parent: string) = 
    parseFile filename parent
    ()

Target "ReferenceDocumentation" (fun _ ->
    !! (solutionPath + @"*.*") -- (solutionPath + "*.xlsm")
        |> Seq.iter(fun f -> getPublicComments f (solutionPath + solutionFile))
        
    !! (solutionPath + @"*.md")    
        |> Seq.iter(fun f -> Fake.FileHelper.MoveFile @"docs/contents/api" f)
)

//---------------------------------------------------------------------------------------
// Print the pattern search report from the source code
let getPatternFromSource (filename: string) (pattern: string) = 
    let outcome = executeFSIWithScriptArgsAndReturnMessages "docs/tools/textfilegrepper.fsx" [|filename;pattern|]
    match outcome with
    | (true, _) -> traceImportant ("Pattern found in file: " + filename) ; true
    | (false, _) -> trace("Pattern not found in file: " + filename) ; false
    
Target "SearchSource" (fun _ ->
   let path = getUserInput "Enter the absolute path to the source directory or leave blank to accept the default: "
   match String.IsNullOrWhiteSpace(path), String.IsNullOrEmpty(path) with
   | (false, false) -> 
       match Fake.FileSystemHelper.isValidPath(path), Fake.FileSystemHelper.isDirectory(path) with
       | (true, true) -> solutionPath <-path
       | _ -> ()
    | (_,_) -> ()
   Fake.TraceHelper.trace solutionPath

   let pattern = getUserInput "Enter the pattern to search for: "
   match String.IsNullOrWhiteSpace(pattern), String.IsNullOrEmpty(pattern) with
   | (false, false) -> 
        use reportStream = new System.IO.StreamWriter(solutionPath + "report.txt",true)
        !! (solutionPath + @"*.*") -- (solutionPath + "*.xlsm")
            |> Seq.where(fun f -> getPatternFromSource f pattern)
            |> Seq.iter(fun f -> reportStream.WriteLine f)
        reportStream.Flush()
        reportStream.Close()
        Fake.TraceHelper.traceEndTarget "report generated"
   | (_,_) -> Fake.TraceHelper.traceError "You must enter a non-empty pattern to search for."
    
)


// --------------------------------------------------------------------------------------
// Generate the documentation

let generateHelp' fail debug =
    let args =
        if debug then ["--define:HELP"]
        else ["--define:RELEASE"; "--define:HELP"]
    if executeFSIWithArgs "docs/tools" "generate.fsx" args [] then
        traceImportant "Help generated"
    else
        if fail then
            failwith "generating help documentation failed"
        else
            traceImportant "generating help documentation failed"

let generateHelp fail =
    generateHelp' fail false

Target "GenerateHelp" (fun _ ->
    DeleteFile "docs/content/release-notes.md"
    CopyFile "docs/content/" "RELEASE_NOTES.md"
    Rename "docs/content/release-notes.md" "docs/content/RELEASE_NOTES.md"

    DeleteFile "docs/content/license.md"
    CopyFile "docs/content/" "LICENSE.txt"
    Rename "docs/content/license.md" "docs/content/LICENSE.txt"

    generateHelp true
)

Target "GenerateHelpDebug" (fun _ ->
    DeleteFile "docs/content/release-notes.md"
    CopyFile "docs/content/" "RELEASE_NOTES.md"
    Rename "docs/content/release-notes.md" "docs/content/RELEASE_NOTES.md"

    DeleteFile "docs/content/license.md"
    CopyFile "docs/content/" "LICENSE.txt"
    Rename "docs/content/license.md" "docs/content/LICENSE.txt"

    generateHelp' true true
)

let createIndexFsx lang =
    let content = """(*** hide ***)
// This block of code is omitted in the generated HTML documentation. Use 
// it to define helpers that you do not want to show in the documentation.
#I "../../../bin"

(**
WaterAid IATI XL2XML Project ({0})
=========================
*)
"""
    let targetDir = "docs/content" @@ lang
    let targetFile = targetDir @@ "index.fsx"
    ensureDirectory targetDir
    System.IO.File.WriteAllText(targetFile, System.String.Format(content, lang))


// --------------------------------------------------------------------------------------
// Release Scripts

Target "ReleaseDocs" (fun _ ->
    let tempDocsDir = "temp/gh-pages"
    CleanDir tempDocsDir
    Repository.cloneSingleBranch "" (gitHome + "/" + gitName + ".git") "gh-pages" tempDocsDir

    CopyRecursive "docs/output" tempDocsDir true |> tracefn "%A"
    StageAll tempDocsDir
    Git.Commit.Commit tempDocsDir (sprintf "Update generated documentation for version %s" release.NugetVersion)
    Branches.push tempDocsDir
)

#load "./paket-files/fsharp/FAKE/modules/Octokit/Octokit.fsx"
open Octokit

Target "Release" (fun _ ->
    let user =
        match getBuildParam "github-user" with
        | s when not (String.IsNullOrWhiteSpace s) -> s
        | _ -> getUserInput "Username: "
    let pw =
        match getBuildParam "github-pw" with
        | s when not (String.IsNullOrWhiteSpace s) -> s
        | _ -> getUserPassword "Password: "
    let remote =
        Git.CommandHelper.getGitResult "" "remote -v"
        |> Seq.filter (fun (s: string) -> s.EndsWith("(push)"))
        |> Seq.tryFind (fun (s: string) -> s.Contains(gitOwner + "/" + gitName))
        |> function None -> gitHome + "/" + gitName | Some (s: string) -> s.Split().[0]

    StageAll ""
    Git.Commit.Commit "" (sprintf "Bump version to %s" release.NugetVersion)
    Branches.pushBranch "" remote (Information.getBranchName "")

    Branches.tag "" release.NugetVersion
    Branches.pushTag "" remote release.NugetVersion
    
    // release on github
    createClient user pw
    |> createDraft gitOwner gitName release.NugetVersion (release.SemVer.PreRelease <> None) release.Notes 
    // TODO: |> uploadFile "PATH_TO_FILE"
    |> releaseDraft
    |> Async.RunSynchronously
)


// --------------------------------------------------------------------------------------
// Run all targets by default. Invoke 'build <Target>' to override

Target "All" DoNothing
    
"ReleaseDocs"
  ==> "Release"

RunTargetOrDefault "All"
