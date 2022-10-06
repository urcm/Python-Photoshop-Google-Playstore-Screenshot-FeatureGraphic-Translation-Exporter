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



if __name__ == "__main__":
    main()
