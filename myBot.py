#################################################################################
# author : Hamza Rafique                                                        #
# First test using pyautogui :                                                  #
#  Read csv file and write in a new csv file using bot                          #
# Steps :                                                                       #
#   1. open a csv file (pandas)                                                 #
#   2. read data from csv (pandas)                                              #
#   3. use pyauto gui write the data into another csv file                      #
#       a. minimize all windows                                                 #
#       b. locate minimized excel file by identifying the respective icon       #
#       c. open excel file and maximize the window                              #
#       d. move cursor to initial location                                      #
#       e. write data in x-axis from pandas frame                               #
#       f. write data in y-axis from pandas frame                               #
#   4. use pyauto gui to save csv file                                          #
#       a. Use Hotkeys to save file                                             #
#################################################################################

import pyautogui
from os import path
import pandas as pd


DEBUGGING = False

def test_size():
    width, height = pyautogui.size()

    return width, height


def minimize_windows():
    
    width, height = pyautogui.size()
    pyautogui.click(x=width-2, y=height-2, button = 'left')

def find_icon(icon_path):
    print(icon_path)
    try:
        x,y = pyautogui.locateCenterOnScreen(str(icon_path))
        return x,y
    except TypeError:
        print("Cannot find Excel file . Please Open the Excel File for the Bot.")

def click_and_write(loc,data):

    a,b = loc
    pyautogui.doubleClick(a, b, button = 'left')
    pyautogui.typewrite(str(data))
    pyautogui.hotkey('tab')

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

    return os.path.join(base_path, relative_path)

def main(args):

    # Welcome Note
    ans = pyautogui.confirm('Ready to Launch Our Bot ?')

    if ans == 'OK' :

        # read data file
        data_file_tmp = path.join(args.input_dir, args.file_name)
        # converting relative path to absolute path for PyIstaller
        data_file = resource_path(data_file_tmp)

        if DEBUGGING :
            print (data_file)
        
        df = pd.read_csv(data_file,index_col=0)

        # Window Operations
        #####################
        
        # Minimize all windows
        minimize_windows()
        
        # Locate excel file
        icon_path_tmp = args.input_dir + 'excel-icon.PNG'
        # converting relative path to absolute path for PyIstaller
        icon_path = resource_path(icon_path_tmp)
        x,y = find_icon(icon_path)

        
        # Open excel window
        pyautogui.click(x, y, button = 'left')
        # Maximize window
        pyautogui.hotkey('alt', 'space')
        pyautogui.typewrite('x')

        # Data Write Operations
        #####################    

        # Move cursor to reference point
        master_loc = (150,380)

        # Write data in x-axis
        for i in range(len(df.columns)):
            steps = i * 80
            init_loc = (master_loc[0]+steps , master_loc[1])
            
            click_and_write(init_loc, df.columns[i])

        # Write data in y-axis
        for m in range(len(df.columns)):

            if DEBUGGING :
                print("mmmmmmmmmmmmmmmmmm = ", m)
            k = 0    
            for j in range(1,10):
                
                steps = j * 25
                init_loc = (master_loc[0], master_loc[1]+steps)

                if DEBUGGING :
                    print(m,k, init_loc, df[df.columns[m]][k])

                click_and_write(init_loc, df[df.columns[m]][k])
                k+=1
                
            master_loc = (master_loc[0] + 80, master_loc[1])

        # File Save Operations
        #####################
        
        ans = pyautogui.confirm('Process Completed. Save File?')
        if ans == 'OK' :
                
            pyautogui.hotkey('alt')
            pyautogui.typewrite('f')
            pyautogui.typewrite('s')

        else :
            return

    else:
        pyautogui.alert('Process Terminated')



if __name__ == "__main__":
    
    import argparse
    parser = argparse.ArgumentParser(
    description=__doc__)

    parser.add_argument('--input-dir', default='', help='path to data csv file')
    parser.add_argument('--file-name', default='./', help='name of data csv file')

    args = parser.parse_args()

    main(args)

#python myBot.py --file-name disaster_names.csv
