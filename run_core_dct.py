
from Py_Selenium_MainGUI_except_Desg import *
from Designation_Contribution_Joblevel import *

'''  Created to run both the files at the same time.
This will help user to use functions from both the files at the same time  '''

 #pyi-makespec run_core_dct.py --onefile --noconsole --add-binary "driver\chromedriver.exe;driver\" --add-data "Img\Dbox4.png;Img\"  --name Core_DCT_thread
 #pyinstaller --clean Core_DCT_thread.spec

 #pyi-makespec Py_Selenium_Copy.py --onefile --noconsole --add-binary "driver\chromedriver.exe;driver\" --add-data "Img\Dbox4.png;Img\"  --name Core_DCT_demo_thread
 #pyinstaller --clean Core_DCT_demo_thread.spec

if __name__ == "__main__":

    # Please dont remove the below given line
    multiprocessing.freeze_support()

    p1 = multiprocessing.Process(target=Main_root_window)
    p2 = multiprocessing.Process(target=Open_designation_window)

    p1.start()
    p2.start()

    p1.join()
    p2.join()
