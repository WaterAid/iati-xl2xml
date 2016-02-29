## Contributing to the project 

We welcome contributions to the project.  We're not quite set up so you can just fork the project, pick an issue and send us a pull request so in the first instance please contact [Mike Smith](https://github.com/drmrsmith).

Once that's done or if you just want to poke around then you'll need to follow these steps:

- Clone the repo
- Install [FAKE](http://fsharp.github.io/FAKE/index.html)
- Run `build.cmd` in the project root directory. *

We are presuming you are building on Windows for the time being.

###Project structure

```
project
|
|---.fake
|---.git
|---.paket
|---docs
     |
	 |---content		* put markdown documentation in here *
	      |
		  |---api		* code documentation will be put in here *
		  style.css
		  tips.js
	 |---files			* put static documentation in here *
	 |---images
	 |---output		* target directory for all public documentation, gets built in to here *
	 |---specs		* contains requirements
	 |---tools
	      |
		  |---templates	* contains *.cshtml templates for generating html from documentation *
		  generate.fsx		* script to control documentation generation *
|---packages
|---paket-files
|---src					* contains the source for the application *
	 |
	 *.[bas|cls|frm]
	 xl2xml.xlsm	* the application *
|---test
|   .gitignore
|   build.cmd
|   build.fsx
|   LICENSE.txt
|   packages.config
|   paket.dependencies
|   paket.lock
|   README.md
|   RELEASE_NOTES.md

```

Once you have the code the typical developer workflow that is being followed on this project is as follows:

1.	Open the application in a version of Excel greater than or equal to 2007.
2.  Make and test the appropriate changes.
3.  Update the developer documentation in the code as you are making the code changes.
4.	Compile and save the workbook.
5.	In a command line at the project root run the `Clean` target which will clean out the old documentation and source directories.
6.	Run the `ExtractSource` target which will open the workbook and export the source code files in to the `src` directory.  *Note: this step currently depends on the Interop Assemblies being located at: `C:/Program Files (x86)/Microsoft Visual Studio 12.0/Visual Studio Tools for Office/PIA/Office15`.
7.	Run the `ReferenceDocumentation` target which will parse the source files for their markdown and copy the markdown to the `\docs\content\api` directory.
8.	Update the version in the RELEASE_NOTES.md file and save.
9.	Run the `GenerateDocumentation` target which will create the html documentation from the markdown files.
10.	Run the `Release` target when you are ready to go which will create a pull request on GitHub and notify the repo owners.  If you are making a pull request please keep the version number the same as it is but add `_yourgithubname` to the revision.  For example: `v1.0.1` becomes `v1.0.1_markstownsend` (in my case).