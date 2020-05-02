### What is MyBoss? ###
**MyBoss** is a PowerPoint tool which lets you generate professional-looking org charts from employee data in Excel. Works on PC and Mac.

### How does it work? ###
**MyBoss** takes org structure data from the Excel file and automatically builds org chart from it. For example, the slide below was built based on data from the [sample Excel file].

[sample Excel file]: <https://github.com/devrazdev/MyBoss/raw/master/MyBoss-Sample%20data.xlsx>
[Excel template]: <https://github.com/devrazdev/MyBoss/raw/master/MyBoss-Template.xlsx>

![MyBoss-Demo](https://github.com/devrazdev/MyBoss/raw/master/misc/MyBoss-Demo.gif)

### How do i use it? ###
Open **MyBoss**, click the *Org chart* tab and import the pre-filled [Excel template].
- PC users: [download MyBoss here](https://github.com/devrazdev/MyBoss/raw/master/MyBoss.pptm).
- Mac users: scroll down to **How to run MyBoss on Mac?** section.

### How to fill the Excel template?
Learn from example by playing with [sample Excel file]. Despite column names are self-explanatory, there are some useful tips below in **How to fill the Excel template?** section.

### How to give feedback?
- Option 1: Fill the [2-minute survey](https://forms.gle/9EE1sbwSakhsVVNf7) about **MyBoss**
- Option 2: Shoot me an email at nikitobot@gmail.com

📣 If **MyBoss** saved you time, give it a tribute by spreading a word among friends and colleagues. One day it may save time to them too.

---

### Why should I try MyBoss if I already use X? ###
Creating org chart (aka team slide) is [good and old] routine. 
You have [likely tried] to create it manually in PowerPoint, but, starting from ~10 people, it should have become too time-consuming, hence annoying.
If you [searched] for automation software, you should have discovered [Microsoft Visio]. Still Visio charts have room for improvement:

1. Design. The charts look clumsy, they usually require manual adjustments
2. Format. Visio org charts can't be natively pasted/edited in PowerPoint

You could try going for cloud applications like [Google Sheets], [OrgChartNow] or [Lucidchart], but there is nothing more pedantic and native to PowerPoint than **MyBoss**.

[good and old]: <https://trends.google.com/trends/explore?q=create%20org%20chart&date=all>
[likely tried]: <https://www.youtube.com/results?search_query=create+org+chart>
[searched]: <https://support.office.com/en-us/article/create-an-org-chart-in-office-9419815f-0d7f-4d8b-8220-822036b1fe2b>

[Microsoft Visio]: <https://products.office.com/en-us/visio/flowchart-software>
[Google Sheets]: <https://www.bettercloud.com/monitor/the-academy/create-an-org-structure-chart-in-google-sheets/>
[OrgChartNow]: <https://www.orgchartpro.com/products/orgchart-now-2/>
[Lucidchart]: <https://www.lucidchart.com/pages/how-to-make-an-org-chart>

### What are the main features of MyBoss? ###
- Minimum **4** clicks to get the org chart
- Automatic design of slides
- Automatic validation of data in your Excel template
- Automatic calculation of headcount statistics
- Automatic slides' cross-linking

And cool UX features!
- Customizable design
- Editable org structure

For power users: extended [product demo](<https://www.youtube.com/watch?v=Do3c5ff7b1c>) is available on Youtube. 

### How to fill Excel template? ###
Follow these basic rules when filling the Excel template:

1. One tab - one org structure.
2. One line - one employee.
3. "Must have" fields for CEO (boss, head of org tree):
    - Employee Surname
    - Employee Name
4. "Must have" fields for every other employee:
    - Employee Surname
    - Employee Name
    - Reports To Surname (manager's surname)
    - Reports To Name (manager's name)

Learning from [sample Excel file] is highly recommended. 

### What are system requirements? ###
You only need Microsoft PowerPoint 16.* and Microsoft Excel 16.*

**MyBoss** was manually tested under:
- PC: Microsoft PowerPoint for Office 365 MSO (16.0.12624.20422) 32-bit
- Mac: Microsoft PowerPoint for Mac Version 16.36 (version 20041300)

### How to run MyBoss on Mac? ###
You need to [download the tool](https://github.com/devrazdev/MyBoss/raw/master/MyBoss.pptm) AND put [this AppleScript file] to this folder:
```bash
~/Library/Application Scripts/com.microsoft.Powerpoint/
```
This overhead is Mac-specific and comes from file access restrictions: [reading Excel files from PowerPoint require AppleScript].

*NB: Do not change the name of this file since it's hardcoded in MyBoss!*

To perform a test run:
1. [Open MyBoss](https://github.com/devrazdev/MyBoss/raw/master/MyBoss.pptm)
2. Go to "Org chart" tab
3. Click "Select Excel file"
4. Select [sample Excel file]
5. Select the spreadsheet you want
6. Click "Import"
7. Once imported, click "Create org chart" -> "Single slide org structure"
8. Wait until **MyBoss** finishes running

[this AppleScript file]: <https://github.com/devrazdev/MyBoss/raw/master/misc/MyBoss-browse_files_on_mac.scpt>
[reading Excel files from PowerPoint require AppleScript]: <https://developer.microsoft.com/en-us/office/blogs/VBA-improvements-in-Office-2016/>

### F.A.Q ###
> "Org chart" tab is missing in PowerPoint.

1. Make sure you opened **MyBoss.pptm** file you should have downloaded from this page.
2. Make sure your PowerPoint security settings allow to run macros and allow access to
document model.
3. Make sure you can see *Org chart* tab name in ribbon settings in PowerPoint and it's checked to be visible.
4. If you followed pp. 1-3 and it still does not come up — drop me a line at nikitobot@gmail.com, i will try to help asap.

> When I open the file, security warning comes up.

It usually happens because PowerPoint restricts running VBA code without user consent, while **MyBoss** is a 100% VBA solution. So, if that happens, just click "Enable content".

> When I open **MyBoss**, I see empty "Org chart" tab

To resolve the issue:
1. *Go to PowerPoint -> Preferences -> Security*
2. *Click "Enable all macros"*
3. *Click "Trust access to the VBA project object model"*
4. *Allow actions to run programs without notification*
5. *Reopen the file*
---

## Developer's corner ##
### What is the technology behind MyBoss? ###
Technically, **MyBoss** is a VBA solution — PowerPoint file with macros and a custom tab. 

### Are there any hidden dependencies? ###
The only third-party module used is Tim Hall's custom implementation of Dictionary class  to support Mac ([available on github]). Thanks, Tim.

[available on github]: <https://github.com/VBA-tools/VBA-Dictionary>

### Why did you choose to make a VBA solution? ###
There are 3 ways to build custom solutions for Office suite:
1. VBA solution
2. Visual Studio Tools for Office (VSTO) Add-in
3. JavaScript API Add-in

Their comparison is presented [here] and [there]. Basically, Office for Mac doesn't support VSTO Add-ins and JavaScript API for PowerPoint is yet too limited (July'18).

[here]: <https://docs.microsoft.com/en-us/visualstudio/vsto/vba-and-office-solutions-in-visual-studio-compared>
[there]: <https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins#StartBuildingApps_TypesofApps>

### How to customize MyBoss on my own? ###
Foreword: it's highly recommended to check [Ron De Bruin's website](https://www.rondebruin.nl/) with tutorials and examples about building VBA solutions and customizing ribbons. That's an amazing source of VBA wisdom.

1. To customize core, you will have to use Office Visual Basic Editor, since VBA code is stored inside the *MyBoss.pptm* file in binary format.
    - PC: type Alt+F11 or go to "Developer" tab -> Visual Basic;
    - Mac: go to Tools -> Macro -> Visual Basic Editor.
2. To customize UI tab (Ribbon XML):
[Reference guide on UI], [Mac Ribbon examples], [Win Ribbon examples]
    - PC: Suggest using utility [OfficeCustomUIEditorSetup] (requires [.NET 3.0](https://www.microsoft.com/en-us/p/surface-laptop-3/8VFGGH1R94TM))
    - Mac: Suggest you find a PC. [However], if you change the extension of *MyBoss.pptm* from PPTM to ZIP and look inside the archive, you will find the Ribbon XML, which then you can edit (folder "customUI")
3. To customize the design of organizational charts:
    - PC & Mac: edit the Templates on slide 1, preserving their names and their grouping. [Learn how to check the names of shapes using Selection pane]

[Reference guide on UI]: <https://msdn.microsoft.com/en-us/library/dd926139(v=office.12).aspx>
[Mac Ribbon examples]: <https://www.rondebruin.nl/mac/macfiles/MacRibbonExamples.dmg>
[Win Ribbon examples]: <https://www.rondebruin.nl/win/winfiles/RibbonExampleFiles.zip>
[OfficeCustomUIEditorSetup]: http://www.rondebruin.nl/win/winfiles/OfficeCustomUIEditorSetup.zip
[However]: <https://support.office.com/en-us/article/extract-files-or-objects-from-a-powerpoint-file-85511e6f-9e76-41ad-8424-eab8a5bbc517>
[Learn how to check the names of shapes using Selection pane]:<https://support.office.com/en-us/article/manage-objects-with-the-selection-pane-a6b2fd3e-d769-46c1-9b9c-b94e04a72550>

## Thanks to contributors ❤️
- **Anna Glushkova** — for keeping an eagle eye on technical part and being a devoted user
- **[Alexey Makurin](https://github.com/amakurin)** — for professionalism and persistence in debugging the rebellious VBA scripts
