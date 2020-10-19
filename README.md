## GOAL

Read csv file (using pandas) and write in a new csv file using PyAutoGUI

## Requirements
1. pyAutoGUI
2. pandas

## Pre - requisites

1. Download disaster_names.csv, botfile.csv and place it in same path
2. Open botfile.csv

## Implementation steps

  1. open a csv file (pandas)                                                
  2. read data from csv (pandas)                                              
  3. use pyauto gui write the data into another csv file                      
       a. minimize all windows                                                 
       b. locate minimized excel file by identifying the respective icon       
       c. open excel file and maximize the window                              
       d. move cursor to initial location                                      
       e. write data in x-axis from pandas frame                               
       f. write data in y-axis from pandas frame                               
  4. use pyauto gui to save csv file                                          
       a. Use Hotkeys to save file 

## Known Issues

1. PyAutoGUI unable to locate excel-icon.PNG

        To localize the icon PyAutoGUI uses open pythons Pillow library. The accuracy of excel icon screen shot detection is 
        quite low. In case of no detection by PIL, make a fresh screen shot of excel file icon in the "windows task bar" and 
        replace it with the excel-icon.PNG file in folder 
