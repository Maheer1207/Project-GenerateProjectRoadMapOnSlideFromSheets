# Generate RoadMap for Projects

[Video Demonstration](https://youtu.be/Qo8xihB8vJQ)

## Project Overview

This project was developed to assist the Project Management Team of an organization in generating project roadmaps using Google Slides. The team possesses all necessary resources, including tasks, start and end dates, predecessors, successors, and task chronology (referred to as vertical levels or positions). The objective was to create a Google Slides roadmap based on these formatted resources available in Google Sheets. Furthermore, the script was designed to be accessible to all team members, recognizing that not all members may be proficient in using raw code in Visual Studio or other IDEs. Consequently, it was integrated as a Google Sheets add-on, exclusively available to the team.

## How It Works

Team members logging in with the organization's Google domain are automatically granted access to the "RoadMap Generator" extension within Google Sheets. After downloading and installing the add-on, users are prompted to provide inputs, including the pre-made template name, source sheet name, scales, and Cartesian start position for the roadmap. The AppScript then reads the source files and calculates the rectangular objects on Google Slides for the tasks, utilizing the provided scales and start position. Finally, leveraging the Google Drive, Sheets, and Slides APIs, the AppScript generates the roadmap presentation in the user's Google Slides, producing a ready-to-present slide that is well-labeled and includes all relevant details.

## Tools and Technologies

1. **Google Drive API**: Used to access the existing "template" slides that form the foundation for the roadmap and to create copies of the slide presentation.
2. **Google Sheets API**: Utilized to read the source data sheets, extract necessary data, and format it into another sheet.
3. **Google Slides API**: Employed to create the roadmap on a copy of the "template" slide stored in Google Drive.
4. **Google OAuth**: Integrated to authenticate users during the installation of the add-on. Given its focus on organizational use, external access is restricted. OAuth also handles permissions for reading and writing in the user's Google Drive.
5. **Google Marketplace**: Used to publish and promote the script as an add-on available in Google Sheets for the organization's employees.

## Environment Setup for the Add-On

To set up the environment, follow these steps:

1. **Template Slide**: Ensure that the template slide for the roadmap's base layer exists (a sample template slide is provided in the repository).
2. **Source Sheet Structure**: The source sheet should follow a standard structure:
   - **Row 1**: Serves as the header, with all subsequent rows containing data.
   - **Column Headers**: Include Activity ID, Task Name*, Start Date*, Finish Date*, Successor, Predecessor, and Vertical Position/Level*. These headers should start from column 1 in sequential order. Columns marked with an asterisk (*) are crucial for generating taskbars on the roadmap and must remain in their designated positions (Task Name*, Start Date*, Finish Date*, and Vertical Position/Level* in columns B, C, D, and E, respectively). Other columns can be rearranged as needed.
3. **Task Bar Colors**: Each cell/row should be filled with the chosen color for the taskbar. The program reads the background color of individual row cells when processing the roadmap data and applies it to the generated taskbars.
4. **Date Formats**: Dates should follow one of these formats: a date object (e.g., “July 1, 2027, 8:00 AM” or “26-Sep-22”) or a string (e.g., “16-Jul-21 A”), where the initial part of the string contains the date to be used.
5. **Slide Template Design**: The template Google Slide used for generating the roadmap may vary. Therefore, ensure an appropriate template is created before running the program. The template should be designed using the “Insert Table” feature, with columns and rows added according to the roadmap's specific requirements. Ensure that the table is not placed in the theme/layout layer of the slide, as this information is not duplicated when the slide is copied.
6. **Single Slide Requirement**: Each Google presentation should contain only one slide. The template slide for drawing should be the sole slide in the presentation.
7. **Grid Height**: The height of the grid on which the roadmap taskbars will be drawn should be a minimum of 12.25 cm.
8. **Language Settings**: Verify that the language setting for your slide is English (Canada) to ensure that the measurements used are in the metric system. If measurements appear in the imperial format, confirm that the language is set to Canada, not the United States.

For further setup details, refer to the [Steps to setup the environment for Add-on](https://github.com/Maheer1207/Project-GenerateProjectRoadMapOnSlideFromSheets/blob/f4d3c12774f29406d5e043f099ef5147072ee3f7/Steps%20to%20setup%20the%20environment%20for%20Add-on.pdf).

## Steps to Use the Add-On

1. **Install the Add-On**: Install the 'RoadMap Automation Extension' as an add-on in Google Sheets.
2. **Access the Feature**: Once installed, a new feature called 'Draw Task Bars and Milestones' will appear under the 'Extensions' menu.
3. **Enter Source Sheet Name**: Start by entering the name of the source sheet in your Google Spreadsheet.
4. **Select Google Slide Template**: Specify the name of the Google Slide template you intend to use for your roadmap.
5. **Input Taskbar Length**: Enter the length of the taskbar for a month. You can determine this by using the 'Insert Shape' function and adjusting its format. Add a rectangular shape to the grid and calculate the monthly length of the taskbar using the 'Size & Rotation' options.
6. **Set X-Coordinate**: Input the X-coordinate from which you want the roadmap's taskbar to commence, using the 'Position' feature.
7. **Set Y-Coordinate**: Provide the Y-coordinate indicating the starting point of the roadmap's taskbar, also using the 'Position' feature.

Following these steps will generate a Google Slide with the completed roadmap.

For detailed usage instructions, refer to the [Steps to use the Add-on](https://github.com/Maheer1207/Project-GenerateProjectRoadMapOnSlideFromSheets/blob/f4d3c12774f29406d5e043f099ef5147072ee3f7/Steps%20to%20use%20the%20Add-on.pdf).
