# Excel VBA to create a comprehensible Chart based on your Data
The project provides an Excel VBA macro to create a chart with data extracted from Excel worksheets.
![Excel Chart](https://github.com/Sky4Lion/vba_create_chart/blob/master/demo/Demo.png)

## Presentation - Functionality
Within the demo folder you can download an Excel demo document that can help understanding the functionality and the ability of the macro.
The Excel document needs to match the following structure:
![Excel Chart](https://github.com/Sky4Lion/vba_create_chart/blob/master/demo/Structure.png)
####Element (required):
The element is non optional and needs to be declared for every row in the worksheet. The element could be any data like e.g. the name of an employee (or a task, a software module, …).
####Section (optional):
If your elements are not unique itself, you can use the optional section to build a unique key. This key is used to identify the correct element. If the key is not unique, the macro will always link to the first matching key. In my example, the section could be the bureau or department of the employee (or a project).
(Hint: So far, the macro assumes there is a section. If you don’t use a section, the created chart will add empty lines instead -> not looking nicely.)
####Dependency (optional):
Your element might depend on other elements (or sub elements). The chart will show these dependencies and also searches for other elements that depend on the considered on. If there is some direct interference between different elements, the chart will show this as arrows between the elements. In my example, the employee could depend on its team or some equipment (or depending tasks, software sub modules / functions, …).

## Installation
By just inserting the 3 vba code files into one (or more) macro modules of your Excel document, you can use the macro.

## First Steps
If you first use the macro on your own data you need to adjust some settings. Therefore please read through the CONFIG file and adjust the parameters to meet your data. Especially the “Processed Columns” section need to be adjusted.
