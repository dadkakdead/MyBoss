### What is ntOrgchart? ###
**ntOrgchart** is a tool to automatically create nice org charts in PowerPoint on PC and Mac. Technically, it's a VBA solution — [PowerPoint file] with macros and a custom tab.

---

### How do these org charts look like? ###

The slide below was automatically created by **ntOrgchart** using [sample data].

![ntOrgchart-example_slide](https://github.com/devrazdev/ntOrgchart/blob/master/misc/screenshot.jpg?raw=true)

Also, [product demo] is available on Youtube. 

### Why should I use it? ###
There is a [growing] interest in creating org charts, and the most [common] approach is creating them in PowerPoint manually, but, starting from ~50 people it becomes too time-consuming. Once you start [searching] for automation software, you stumble over [Microsoft Visio]. Indeed, Visio let's you visualize org structures of any size by uploading existing data and automatically arranging shapes, but it has 2 problems:

1. Visual appeal. Initial arrangement of shapes looks clumsy, there is still much manual tuning to do
2. Flexibility. Visio org charts can't be natively pasted/edited in PowerPoint, so you get locked on Visio

Org chart automation has already become a feature of many cloud applications ([Google Sheets], [OrgChartNow], [Lucidchart] to name a few), but neither of them successfully addressed the 1st problem and very few addressed the 2nd. **ntOrgchart** creates beautiful org charts out of the box, and it's native for Microsoft Office suite. Headshot.

### What do I need to try it? ###
You should just have Microsoft PowerPoint 16.* and Microsoft Excel 16.* installed on your computer. Last time **ntOrgchart** was manually tested with:
- PC: Microsoft PowerPoint 2016 MSO (16.0.9126.2259) 32-bit + Microsoft Excel 2016 MSO (16.0.9126.2259) 32-bit
- Mac: Microsoft PowerPoint for Mac Version 16.15 (180709) + Microsoft Excel for Mac Version 16.15 (180709)

### How to run it? ###
Since **ntOrgchart** is just a PowerPoint file, there is no installation required. However:

1. Since it's written in VBA, PowerPoint may ask you to enable running VBA code, when you open the file (depending on your Trust Center settings). If "Security warning" comes up, just click "Enable content".
2. **MAC ONLY:** Due to the file access restrictions  ([reading Excel files from PowerPoint require AppleScript]), you will need to put *ntOrgchart-browse_files_on_mac.scpt* to ~/Library/Application Scripts/com.microsoft.Powerpoint/, before running the **ntOrgchart**. Do not change the name of script since it's hardcoded!

To perform a test run:
1. Open *ntorgchart.pptm*
2. Go to "Org chart" tab
3. Click "Open"
4. Select *sample input.xlsx*
5. Select the spreadsheet you like
6. Click "Upload"
7. Once loaded, click "Create org chart" -> "Single slide org structure"
8. Wait until it finishes

### It's not working! Can you help? ###
1. Restart your computer and check the issue
2. If issue persists, shoot me an e-mail at devrazdev@gmail.com - I will try to help as soon as possible
3. If issue is gone, it's a lucky day

---

## Developers corner ##

### Are there any hidden dependencies? ###
The only third-party thing I have included was Tim Hall's custom implementation of Dictionary class ([available on github]) to support Mac. Thanks, Tim.

### Why did you choose to make a VBA solution? ###
There are 3 ways to build custom solutions for Office suite:
1. VBA solution
2. Visual Studio Tools for Office (VSTO) Add-in
3. JavaScript API Add-in

Their comparison is presented [here] and [there]. Basically, Office for Mac doesn't support VSTO Add-ins and JavaScript API for PowerPoint is yet (July'18) too limited.

### What is the right data format? ###
Take a look at *ntOrgchart-sample_input.xlsx*. Basic rules:
1. One org structure — one spreadsheet.
2. One employee — one line.
3. Minimum required data for every employee (not CEO) - 4 fields:
    - Employee Surname
    - Employee Name
    - Reports To Surname (manager's surname)
    - Reports To Name (manager's surname)
4. Minimum required data for CEO - 2 fields:
    - Employee Surname
    - Employee Name
5. CEO must have these fields empty (because nobody is managing CEO):
    - Reports To Surname
    - Reports To Name
    - Dotted Line Manager Surname
    - Dotted Line Manager

See "data_minimum" tab in *ntOrgchart-sample_input.xlsx* for the example of minimum data input.

### How to customize it on my own? ###
The sad news is that VBA code is stored inside the *ntorgchart.pptm* file in binary format. So, if you want to customize the **ntOrgchart**, you will have to use Office Visual Basic Editor.

1. To customize logic (VBA):
    - PC: type Alt+F11 or go to "Developer" tab -> Visual Basic;
    - Mac: go to Tools -> Macro -> Visual Basic Editor.
2. To customize UI (Ribbon XML):
[Reference guide on UI], [Mac Ribbon examples], [Win Ribbon examples]
    - PC: Suggest using utility [OfficeCustomUIEditorSetup] 
    - Mac: suggest you find a PC and check the point below. [However], if you change the extension of *ntorgchart.pptm* from PPTM to ZIP and look inside the archive, you will find the Ribbon XML, which you can then edit...

## Farewell ##
I would be happy to hear any feedback/news from about how you use **ntReport** in real life. Please fee free to contact me by e-mail at devrazdev@gmail.com. Thank you.

[PowerPoint file]: <https://github.com/devrazdev/ntOrgchart/raw/master/ntOrgchart.pptm>
[sample data]: <https://github.com/devrazdev/ntOrgchart/raw/master/misc/sample%20input.xlsx>
[product demo]: <https://www.youtube.com/watch?v=Do3c5ff7b1c>

[growing]: <https://trends.google.com/trends/explore?q=create%20org%20chart&date=all>
[common]: <https://www.youtube.com/results?search_query=create+org+chart>
[searching]: <https://support.office.com/en-us/article/create-an-org-chart-in-office-9419815f-0d7f-4d8b-8220-822036b1fe2b>

[Microsoft Visio]: <https://products.office.com/en-us/visio/flowchart-software>
[Google Sheets]: <https://www.bettercloud.com/monitor/the-academy/create-an-org-structure-chart-in-google-sheets/>
[OrgChartNow]: <https://www.orgchartpro.com/products/orgchart-now-2/>
[Lucidchart]: <https://www.lucidchart.com/pages/how-to-make-an-org-chart>

[reading Excel files from PowerPoint require AppleScript]: <https://developer.microsoft.com/en-us/office/blogs/VBA-improvements-in-Office-2016/>

[OfficeCustomUIEditorSetup]: http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2006/05/26/customuieditor.aspx
[available on github]: <https://github.com/VBA-tools/VBA-Dictionary>

[think-cell]: <https://www.think-cell.com/en/>

[here]: <https://docs.microsoft.com/en-us/visualstudio/vsto/vba-and-office-solutions-in-visual-studio-compared>
[there]: <https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins#StartBuildingApps_TypesofApps>

[Reference guide on UI]: <https://msdn.microsoft.com/en-us/library/dd926139(v=office.12).aspx>
[Mac Ribbon examples]: <https://www.rondebruin.nl/mac/macfiles/MacRibbonExamples.dmg>
[Win Ribbon examples]: <https://www.rondebruin.nl/win/winfiles/RibbonExampleFiles.zip>

[However]: <https://support.office.com/en-us/article/extract-files-or-objects-from-a-powerpoint-file-85511e6f-9e76-41ad-8424-eab8a5bbc517>
