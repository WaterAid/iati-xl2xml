#I "../../packages/FAKE/tools/"
#I @"C:/Program Files (x86)/Reference Assemblies/Microsoft/Framework/.NETFramework/v4.5.2/"
#r "FakeLib.dll"
#r @"..\..\packages\DocumentFormat.OpenXml\lib\DocumentFormat.OpenXml.dll"
#r "WindowsBase.dll"

open Fake
open System
open System.IO
open System.Xml
open FSharp.Data
open System.Linq
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Validation

#I "../../packages/FAKE/tools/"
#r "FakeLib.dll"
#r @"..\..\packages\DocumentFormat.OpenXml\lib\DocumentFormat.OpenXml.dll"
#r @"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5.2\WindowsBase.dll"
#r @"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5.2\System.IO.Compression.FileSystem.dll"
#r @"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5.2\System.IO.Compression.dll"
#r @"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5.2\System.Data.Linq.dll"


open Fake
open System
open System.IO
open System.Xml
open FSharp.Data
open System.Linq
open System.Data.Linq
open System.IO.Compression
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Validation


//// <summary>
//// Replaces the malformed uri with (currently) static fixed ones
//// </summary>
//// <param name="fs">the filestream containing the xml</param>
//// <param name="invalidUriHandler">the lambda function that does the uri replacement</param>
//// <returns>nothing yet, unit</results>
let updateBrokenUris (invalidUriHandler:Func<string, Uri>) (entry:ZipArchiveEntry) (relNs:string) = 
    use entryStream = entry.Open()
    try
        let entryXDoc = XDocument.Load(entryStream)
        match isNull(entryXDoc.Root), entryXDoc.Root.Namespace  with
        | (false,relNs) -> 
            let urisToCheck = 
                entryXDoc.Descendants(relNs + "Relationship")
                |> Seq.where(fun r -> !isNull(r.Attribute("TargetMode"))
                |> Seq.where(fun r -> r.Attribute("TargetMode") = "External") 
            replaceBrokenUri urisToCheck invalidUriHandler, entryXDoc           
        | (_,_) -> ()
    with
    | :? XmlException as xex -> ()
            
let internal replaceBrokenUri (urisToCheck:seq<'a>) (invalidUriHandler:Func<string, Uri>) = 
    let generator a = 
            try
                let uri = new Uri(u.Target)
                Some u
            with
            | :? UriFormatException -> 
                u.Target.Value = invalidUriHandler u.Target
                Some u
            | _ -> None

    results
// if the root is not null and the namespace is the relns namespace
// get the list of uris to check by taking the doc and getting the Relationship descendants
// which have the attribute TargetMode not null and cast to a string the attribute TargetMode equal to External
// for each of these then get the cast to string attribute Target and check it's not null
// then try and new up a Uri object with it and catch the UriFormatException
// when you do then replace the uri with the invalidUriHandler lamda and 
// and remember that you have done this
// if you have done this replacement then for each uri
// delete the existing one and write the new one in its place

                    
//                    
//                }
//                if (replaceEntry)
//                {
//                    var fullName = entry.FullName;
//                    entry.Delete();
//                    var newEntry = za.CreateEntry(fullName);
//                    using (StreamWriter writer = new StreamWriter(newEntry.Open()))
//                    using (XmlWriter xmlWriter = XmlWriter.Create(writer))
//                    {
//                        entryXDoc.WriteTo(xmlWriter);
//                    }

//// <summary>
//// Replaces the malformed uri with (currently) static fixed ones
//// </summary>
//// <param name="fs">the filestream containing the xml</param>
//// <param name="invalidUriHandler">the lambda function that does the uri replacement</param>
//// <returns>nothing yet, unit</results>
let internal fixInvalidUri (fs:System.IO.FileStream)  (invalidUriHandler:Func<string, Uri>) =  
    // open the stream, locate all the urls, test their instantiation as uri, if they throw then replace them

    let relNs = "http://schemas.openxmlformats.org/package/2006/relationships"
    use za = new ZipArchive(fs, ZipArchiveMode.Update)
    let entries = za.Entries.ToList<ZipArchiveEntry>() |> List.where(fun e -> !e.Name.EndsWith(".rels"))

    // open the entries and identify invalid uri's then replace
    entries |> List.iter(fun e -> updateBrokenUris invalidUriHandler e)


//// <summary>
//// Compensates for the Malformed URI error detailed here: https://github.com/OfficeDev/Open-XML-SDK/issues/38
//// </summary>
//// <param name="filename">the full path to the source spreadsheet file</param>
//// <param name="fixedfilename">the full path to the file that gets fixed
//// <returns>the full path of the fixed file</results>
let public fixUriOpenXml filename fixedFilename = 
    File.Copy(filename, fixedFilename)
    try
        use document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(filename, false)
        document.Close()
        filename
    with
    | :? OpenXmlPackageException as pkEx -> 
        using (new System.IO.FileStream(fixedFilename, FileMode.OpenOrCreate, FileAccess.ReadWrite)) fixInvalidUri (new Uri("http://broken-link/"))
        fixedFilename
    | _ as othEx -> throw othEx : fixedFilename


//// <summary>
//// Validates the document against the OpenXml schema
//// </summary>
//// <param name="filename">the full path to the source spreadsheet file</param>
//// <param name="fixedfilename">the full path to the file that gets fixed</param>
let validateOpenXml filename fixedfilename = 
    
    let mutable resultFile = fixedfilename
    File.Copy(filename, resultFile)
    try
        use document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(filename, false)
        document.Close()
    with
    | :? OpenXmlPackageException as pkEx -> 
        using (new System.IO.FileStream(resultFile, FileMode.OpenOrCreate, FileAccess.ReadWrite)) fixInvalidUri (new Uri("http://broken-link/"))

    match Fake.FileSystemHelper.fileExists(fixedfilename) with
    | false -> resultFile <- filename
    | true -> resultFile <- fixedfilename

    //// finally do the additional validation checks on this file
    use processDoc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(resultFile, false)
    let wbPart = processDoc.WorkbookPart
    let validator = new OpenXmlValidator()
    let validationErrors = validator.Validate(processDoc)
    printfn "%A" validationErrors.Count
    //for err in validationErrors do printfn "%s" (err.ToString())
    resultFile
        
// Generate
let scriptArgsSourceFile = Environment.GetCommandLineArgs().[2]   // it is passed as the 3rd argument in the command that launches fsi
validateOpenXml scriptArgsSourceFile