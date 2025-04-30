# DF_and-_Accuracy_Report_v3
This code is created to automate the creation of accuracy reports and daily reports.


DF and Accuracy Report is the application based on python programming with PyQt v5 GUI. This application functions to determine whether or not a UAV is precise in carrying out the dot spraying mission. To run this application, UAV (Unmanned Aerial Vehicle) log data and target tree point vectors are needed.

Following are the steps for using the application:

The user selects the directory containing the UAV log data via the "Folder .log" button.
Users select the directory containing flying route data in shapefile format via the "Folder .waypoints" button.
The user selects the directory to save the validation results via the "Folder Output" button.
The user selects the UTM zone that corresponds to the data collection location.
Users select the level of accuracy required for data validation.
After selecting all directories and UTM zones, users can press the "Process" button to start the validation process.
Later the results of the process will be grouped into 3 files, the .xlsx file which contains the report all flight missions that have been entered (usually called Daily Flight), the .pdf file which contains  Accuracy report overview of tree points per flight, and the .csv file which contains raw data from application processing.

By using this application, users can easily validate point spraying on oil palm plants by integrating data from UAVs and geospatial data. This can help farmers or plantation managers increase efficiency and accuracy in spraying pesticides or nutrients on oil palm plants.

