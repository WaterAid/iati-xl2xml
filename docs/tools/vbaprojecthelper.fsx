#I "../../packages/FAKE/tools/"
#I "C:/Program Files (x86)/Microsoft Visual Studio 12.0/Visual Studio Tools for Office/PIA/Office15"
#r "FakeLib.dll"
#r "Microsoft.Office.Interop.Excel"
#r "System.Runtime.InteropServices"
#r "Microsoft.Vbe.Interop"

open Fake
open Fake.FileHelper
open Fake.TraceHelper
open System
open System.IO
open Microsoft.Office.Interop.Excel
open System.Runtime.InteropServices
open Microsoft.Vbe.Interop

//// <summary>
//// Persists the module to a file by the same name
//// </summary>
//// <param name="vbComponent">the code module</param>
//// <param name="filepath">the path to save the module at</param>
//// <param name="extension">the file extension of the source file.  One of: .cls or .bas</param> 
let private writeModule (vbComponent : VBComponent) (filepath : string) (extension : string) = 
    let codeModule = vbComponent.CodeModule
    let filename = filepath + @"\" + codeModule.Name + extension
    try
        try
            use outFile = File.CreateText(filename)
            traceImportant filename
            for i in 0 .. codeModule.CountOfLines do
                outFile.WriteLine(codeModule.Lines(i+1,1))
        with
        | ex-> raise ex
    finally
        if(not(obj.ReferenceEquals(codeModule, null)) && Marshal.IsComObject(codeModule)) then printfn "Remaining RCW references to codeModule: %i " (Marshal.ReleaseComObject(codeModule))
           
//// <summary>
//// Opens the workbook and gets access to the code modules.
//// </summary>
//// <param name="filename">the full path to the source spreadsheet file</param>
let public getVbaSource filename = 
    let xlApp = new Microsoft.Office.Interop.Excel.ApplicationClass(Visible = true)
    let workbooks = xlApp.Workbooks
    let filepath = Directory.GetParent(filename).FullName
    let mutable workbook = null
    let mutable codemodule = null
    let mutable vbProject = null
    let mutable moduleColl = null

    try
        try
            workbook <- workbooks.Open(filename, Microsoft.Office.Interop.Excel.XlUpdateLinks.xlUpdateLinksNever, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Microsoft.Office.Interop.Excel.XlCorruptLoad.xlNormalLoad)
            vbProject <- workbook.VBProject
            moduleColl <- vbProject.VBComponents
            for i in 1 .. moduleColl.Count do
                match moduleColl.[i].Type with
                | vbext_ComponentType.vbext_ct_ClassModule -> writeModule moduleColl.[i] filepath ".cls"
                | vbext_ComponentType.vbext_ct_StdModule -> writeModule moduleColl.[i] filepath ".bas" 
                | _ -> ()
            
            workbook.Close(false)
        with
            | :? COMException as ex -> printfn "%s threw com exception: %s" filename ex.Message
            | :? ArgumentNullException as ex -> printfn "%s threw argnull exception: %s" filename ex.Message
            | ex -> printfn "%s threw an exception: %s" filename ex.Message
    finally
        if(not(obj.ReferenceEquals(vbProject, null)) && Marshal.IsComObject(vbProject)) then printfn "Remaining RCW references to vbProject: %i " (Marshal.ReleaseComObject(vbProject))
        if(not(obj.ReferenceEquals(moduleColl, null)) && Marshal.IsComObject(moduleColl)) then printfn "Remaining RCW references to module collection: %i " (Marshal.ReleaseComObject(moduleColl))
        if(not(obj.ReferenceEquals(workbook, null)) && Marshal.IsComObject(workbook)) then printfn "Remaining RCW references to workbook: %i " (Marshal.ReleaseComObject(workbook))
        if(not(obj.ReferenceEquals(workbooks, null)) && Marshal.IsComObject(workbooks)) then printfn "Remaining RCW references to workbooks: %i " (Marshal.ReleaseComObject(workbooks))
        xlApp.Quit()
        if(not(obj.ReferenceEquals(xlApp, null)) && Marshal.IsComObject(xlApp)) then printfn "Remaining RCW references to xlApp: %i " (Marshal.ReleaseComObject(xlApp))

// Generate
let scriptArgsSourceFile = Environment.GetCommandLineArgs().[2]   // it is passed as the 3rd argument in the command that launches fsi
getVbaSource scriptArgsSourceFile