##############################################################################
#
# Photoshop Playstore Screenshot and Feature Graphic Translation Exporter
# Author: urcm, 2021
#
# Usage:
# python change_text_with_ps_store_ss.py c:/folder/test.psd c:/folder/translation.txt
##############################################################################

# Import required modules
import win32com.client
# To install this package with conda run:
# conda install -c anaconda pywin32
# if still have problem :
# in pywin32 folder "Lib/site-packages/pywin32_system32",
# which including 3 dll libs, copy them to the "/Lib/site-packages/win32" directory,
# which including the win32apython pi.pyd or win32api.pyc.

import sys
import os
import re


def main():
    if len(sys.argv) < 3:
        print("\nUsage error")
        print("\nex. change_text_with_ps_store_ss.py test.psd translate.txt\n\n")
        sys.exit()

    directory = sys.argv[1]
    file_translation = sys.argv[2]

    print("\nPhotoshop File: '{}'\n".format(directory))
    print("\nTranslated Text File: '{}'\n".format(file_translation))
    
    # Definition for just escaping warning "local variable 'ps_app' might be referenced before assignment"
    ps_app = None
    # sys.exit()
    try:
        ps_app = win32com.client.Dispatch("Photoshop.Application")
    except Exception as e:
        print("\nDispatching Photoshop is not working...\n", repr(e))
        print("\nDo you have installed Photoshop?\n")

    # path = "D:\calismalar\lina\OK2\test\translate.txt"
    path = sys.argv[2]
    
    
    

if __name__ == "__main__":
    main()
