# VBA Create Chart Macro Code
This folder contains the 3 macro modules needed to provide the functionality: 
* Auto_Open 
* CONFIG
* main

## Auto_Open
The file provides functions that are automatically performed by Excel when opening the file:
* Apply the CONFIG settings to the macro
* Create 2 useful buttons (one to deactivate the double click event handler that makes it easier to edit cells)
* Start the double click event handler to execute the main macro at click action

## CONFIG
The file contains easy to configure variables to adjust the macro. You can activate options by deleting the ‘ at the beginning of a line or by adjusting their value within the specification. 
Comments are marked by ‘#. Please do not activate comments as this will result in a VBA error. The comments also contain the allowed variable type as a hint.

## main
This is the main macro code. The macro will first read all necessary data of the Excel sheet and processes this data further. Processed are the corresponding columns for the elements, the section, the dependency and the link (marked in bold in the first row of the picture):
![Excel Worksheet Demo]( https://github.com/Sky4Lion/vba_create_chart/blob/master/doc/pictures/Table_full.png)

#### Element
Most important is the element column as it is part of the unique key. There shouldn’t be any empty element cell in your worksheet as its row may only be processed fractionally. An element itself doesn’t have to be unique. In case an element is not unique, the section should be used to form a unique key. Else the macro will only regard the first matching key. 

#### Section
The section column is optional and should be used to differ equal elements. As an example the section could be seen as a group’s name whereas the element is the name of a single member. 
Hint: The current version of the macro assumes there is a section so it is wisely to do not have empty cells in this column.

#### (unique) Key
The keys are created by combining the optional section, the delimiter and the element. The key should be unique, because in case it is not, the macro will only consider the first matching key, ignoring every other match. Right after reading the data, the macro creates these keys and uses them for every comparison and assignment.

#### Delimiter
The delimiter is a selectable string or char (see CONFIG) that is used to differ the section and the following element. It is not allowed to use this string in either the section or the element content. The default value for the delimiter is “$” (dollar-sign).

#### Dependency
If there are any dependencies or interference between the elements or outward stuff (not included in this worksheet), this can be noted in this column. The chart will show these dependencies and also searches for other elements that depend on the considered on. If there is some direct interference between different elements, the chart will show this as arrows between the elements. 
The column can be empty or contain any strings or empty lines (even the delimiter). The macro splits this cell content at every new line (ALT+Enter) and processes these parts further. If the parts (content of one line in this cell) fulfil certain criteria, they are regarded by the macro: 
* If the part contains the delimiter, the part is regarded as a key and is compared to the other keys. If it matches another key, this dependency is linked to the corresponding element of the matched key. (This is the only way to refer to an element that has a different section than the considered element.)
* If the part doesn’t contain the delimiter, a key is built by combining this part with the section of its row. The result is then compared to the other keys and in case of a match, it is linked to the corresponding element of the matched key.
* If no match was possible, the part is regarded to be not a key and is marked red in the chart. Red marked dependencies are not linked and cannot be followed in contrast to the blue ones. 

The dependencies can be split in 2 groups:
* Backward dependency 
* Forward dependency

#### Backward dependency
Backward dependencies of the current element are all dependencies that are listed in the dependency cell of the current row of the considered element. In other words: These Backward dependencies are somehow needed by the considered element or are steps or actions to be performed in ADVANCE to the considered element. (In a flow chart these are all elements with an arrow TO the considered element. So you have to follow the arrow BACKWARD to get them.)

#### Forward dependency
A dependency is a called a forward dependency, if there is another element that contains the considered element as its dependency. In other words: The Forward element can only be performed AFTER performing the considered element. Meaning: The considered element is listed in the dependency cell of the Forward element. (In a flow chart these are all elements with an arrow FROM the considered element. So you have to follow the arrow FORWARD to get them.)

#### The chart will display these dependencies as follows:
![Chart Explaining](https://github.com/Sky4Lion/vba_create_chart/blob/master/doc/pictures/Chart_explain.png)
Picture Legend:
* the element calling the chart first is colored in green (so you can easily track the origin).
* every successfully linked Back- or Forward element is colored in blue (by clicking on them you can rebuild the chart using them as considered element)
* every element that couldn’t be linked is colored red (because no valid key was found -> no rebuild possible)
* if an dependency appears as Forward AND Backward, the Backward element is replaced by an arrow from the corresponding Forward to the considered central element (like blue or red arrow).
* if there is a short loop between the dependencies like: “considered element -> Forward dependency -> Backward element -> considered element”, an arrow links the Forward to the Backward element (like green arrow).

## Result
Executing the macro in row 5 of the picture above results in an chart like this:
![Demo Chart]( https://github.com/Sky4Lion/vba_create_chart/blob/master/doc/pictures/Demo.png)

### Hint
The macro always tries to fit the whole chart into the window size, so maximize the Excel window or using the Excel zoom may help if there is only little space.
