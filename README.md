# Generate RoadMap for Projects

This project was created to help a Project Management Team of an Orgainization, where the team has all the resources like the tasks, start dates, end dates, predecessor, successor as well as the chronology which is also regraded as vertical levels/position. The goal was to generate a roadmap of the project planning on a Google Slides with the aid of the available formatted resources in Google Sheets. Following that, the script was expected to be avialable by all the team members with the expectation that some memebers might not be able to deal with the raw codes on Visual Studios or other IDE, thus I had to append it as a Google Sheet Extension only for the members of the team.

## How it works

All the team members loging in with the organisation's Google domain are automatically verified to access the "RoadMap Generator" extension in their Google Sheets. Once they download it as add-on, and run the extension, the user is required to give some inputs like the Pre-made Template/Base-Layer Name, Source Sheet Name, Scales, and Cartisian Start Position of the RoadMap. Following the inputs, the AppScript Code reads the source files, calculates the rectangular object of Google Slide for the tasks with the aid of Scales and Start Position. Lastly, with the aid of Google Drive, Sheets, and Slides API it AppScript generates the RoadMap presentation in the team member's Google Drive/Slide in a ready to Present Slide that comes well labelled and have the minor details.

## Steps to setup the environment for Add-on

1. The source sheet should adhere to the standard structure of a spreadsheet, with the following guidelines:
  a. Row 1 serves as the header, and all subsequent rows contain data.
  b. Column headers include Activity ID, Task Name*, Start Date*, Finish Date*, Successor, Predecessor, and <Anything Else>. These headers     begin from column 1 in sequential order. Column headers marked with an asterisk are crucial for the program to generate taskbars on the      roadmap. They must remain in their designated positions (Task Name*, Start Date*, Finish Date*, and Vertical Position/Level* in columns      B, C, D, and E respectively). Other columns can be rearranged as needed.
2. Each cell/row should be filled with the user's chosen color for the task bar. The program reads the background color of individual row cells when reading the roadmap data and uses it to color the generated taskbars.
3. Dates should follow one of these formats: either as a date object, like “July 1, 2027 8:00 AM” or “26-Sep-22”, or as a string, like “16-Jul-21 A”, where the initial part of the string contains the date to be used.
4. As the template of the Google Slide used for generating the roadmap slide may vary, there is no pre-generated template. Hence, ensure that you create an appropriate template before running the program.
5. The template should be designed using the “insert table” feature. Add columns and rows according to the specific requirements of the roadmap. Make sure that the table is not placed in the theme/layout layer of the slide, as this information is not duplicated when the slide is copied.
6. Each Google presentation should comprise only one slide. The designated template slide for drawing should be the sole slide in the template.
7. The height of the grid on which the roadmap task bars will be drawn should be a minimum of 12.25 cm.
8. Verify that the language setting for your slide is English (Canada) to ensure that the measurements used are in the metric system. If they appear in imperial format, confirm that the language is set to Canada and not the US.

Link: https://github.com/Maheer1207/Project-GenerateProjectRoadMapOnSlideFromSheets/blob/f4d3c12774f29406d5e043f099ef5147072ee3f7/Steps%20to%20setup%20the%20environment%20for%20Add-on.pdf

## Steps to use the Add-on

1. Install the 'RoadMap Automation Extension' as an add-on in your Google Sheets.
2. Once this is completed, you'll notice a new feature added under 'Extensions' called 'Draw Task Bars and Milestones.'
3. To utilize this feature, begin by entering the name of the source sheet in your Google Spreadsheet.
4. Following that, the system will prompt you to specify the name of the Google Slide template you intend to use for your roadmap.
5. Subsequently, you will be required to input the length of the taskbar for a month. You can determine this length by utilizing the 'Insert Shape' function and adjusting its format. Add a rectangular shape to the grid and calculate the monthly length of the taskbar using the 'Size & Rotation' options.
6. Moving on, input the X-Coordinate from which you want the roadmap's taskbar to commence, employing the 'Position' feature.
7. Lastly, provide the Y-Coordinates indicating the starting point of the roadmap's taskbar, also using the 'Position' feature.

That should generate a Google slide with the Roadmap.

Link: https://github.com/Maheer1207/Project-GenerateProjectRoadMapOnSlideFromSheets/blob/f4d3c12774f29406d5e043f099ef5147072ee3f7/Steps%20to%20use%20the%20Add-on.pdf
