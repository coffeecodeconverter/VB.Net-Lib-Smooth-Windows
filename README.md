# VB.Net-Smooth-Windows
Smoothly Animate WinForms - Position, Resize, Fade, Module Code - works with any project 


Introducing Smooth Windows for VB.Net<br>

# Smooth Form Movement and Resizing Module

A module designed to simplify the process of moving and resizing a form with smooth animations, including fading transparency. It allows resizing from a specific corner of the window for a more visually appealing effect, based on the screen quadrant.  
Compatible with .NET Framework 4.0 and above.

## Usage

### Exposed Functions

#### `Smooth.CreateDynamicForm()`
Displays a dynamic form for testing and exploring the module out of the box. This function can be removed when importing the module into your own projects.

#### `Smooth.MoveWindow(ByRef Form, newX, newY, optional newWidth, optional newHeight, optional opacityStart, optional opacityEnd, optional callbackFunction)`
The primary function responsible for moving and/or resizing a window, adjusting its opacity, and animating the process.

### Other Available Functions

#### `Smooth.GetTaskbarPosition()`
Returns the position of the taskbar as `"Top"`, `"Bottom"`, `"Left"`, or `"Right"`.

#### `Smooth.GetTaskbarThickness()`
Returns an integer representing the thickness of the taskbar, either in height or width depending on orientation.

#### `Smooth.SetWindowTransparency(ByRef form, opacityInt1to100)`
Sets the transparency level of a form, where opacity is an integer value between 1 and 100.

#### `Smooth.MoveWindowToScreen(ByRef form, screenIndex)`
Moves a form to a specified monitor index with coordinates set to `x=0`, `y=0`.

#### `Smooth.GetScreenInfo(Optional ByRef form)`
Returns detailed information about the monitor as a `Dictionary(Of String, Dictionary(Of String, Object))`.  
Optional: Can accept a form to return screen data for that specific form.

#### `Smooth.FormatScreenInfo(dictionary)`
Converts the dictionary returned by `GetScreenInfo()` into a neatly formatted string.

#### `Smooth.GetMonitorForForm(dictionary, ByRef form)`
Determines the monitor index where a given form is located.

---

## Example Code Using the Dictionary to Retrieve Screen Information:

```vb
Dim tempDict As Dictionary(Of String, Dictionary(Of String, Object)) = GetScreenInfo(Me)
Dim tempString As String = FormatScreenInfo(tempDict)
MessageBox.Show(tempString)
