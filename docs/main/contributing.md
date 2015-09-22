## Contributing to the project 

We welcome contributions to the project.  We're not quite set up so you can just fork the project, pick an issue and send us a pull request so in the first instance please contact [Mike Smith](@mrdrsmith).

Once that's done or if you just want to poke around then you'll need to follow these steps:

- Clone the repo
- Install the NuGet command line.  We use [chocolatey](https://chocolatey.org/) to do this but your options are all [here](http://docs.nuget.org/consume/Command-Line-Reference).
- Open a PowerShell window, and navigate to the local project directory.  Restore the packages using the following command ![nuget restore](./images/restorePackages.jpg)
- For the time being we use this to generate the documentation.

### Now you can do other things.
In order to run the documentation you must do this:
![generate docs](./images/generateDocsMain.jpg).
[Here](https://tpetricek.github.io/FSharp.Formatting/commandline.html) is the link to using the commandline tool.

Or in code:
	[lang=powershell]
	PS C:\mark\excel\iati-xl2xml\packages\FSharp.Formatting.CommandTool.2.10.3\tools> .\fsformatting.exe literate
	 --processDirectory --inputDirectory C:\mark\excel\iati-xl2xml\docs\main 
	 --outputDirectory C:\mark\excel\iati-xl2xml --templateFile C:\mark\excel\iati-xl2xml\docs\tools\template-main.html 
	 --replacements "project-name" "IATI XL2XML" "page-title" "iati-xl2xml" "github-link" "https://github.com/WaterAid/iati-xl2xml"

	 