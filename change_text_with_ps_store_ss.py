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
    
    
    # path = "C:\folder\translate.txt"
    path = sys.argv[2]

    s = "------------------------------------------"
    # file_name = r"C:\folder\translate.psd"
    ps_source = sys.argv[1]

    file_location = os.path.dirname(ps_source)
    ps_app.Open(ps_source)
    doc = ps_app.Application.ActiveDocument
    
    translation_dict = {}
    translation_arr, dic_counter = [], []
    
    # Read text file for translated text
    with open(path, 'r', encoding='utf-8') as fh:
        for b, line in enumerate(fh):
            if re.search(s, line):
                # print(b, 1)
                dic_counter.append(b)
                # pass
            else:
                try:
                    # print(b , "b")
                    if (b - 1) in dic_counter:
                        # print(b-1)
                        translation_dict = translation_dict.copy()
                    else:
                        (key, val) = line.split(":")
                        translation_dict[key.strip()] = val.strip()
                        # print(key)
                        if translation_dict not in translation_arr:
                            translation_arr.append(translation_dict)
                except Exception as e:
                    print(repr(e), "Do nothing")
                    pass
                
    # Read translation dictionary
    for c in translation_arr:
        # print(c)
        # Change the text content
        print(c["language"], " language starts for layer to update...")
        file_layer_container = 0
        for i in doc.LayerSets:
            for lyr in i.ArtLayers:
                if lyr.Name == "text" + str(file_layer_container):
                    lyr.TextItem.contents = c["ss" + str(file_layer_container)]
                    # lyr.TextItem.contents = c["ss '{}'".format(str(file_layer_container))]
                    print(lyr.Name, file_layer_container, "Text has updated.", lyr.TextItem.contents)
                elif lyr.Name == 'ss_fea':
                    lyr.TextItem.contents = c["ss_fea"]
                    print(lyr.Name, file_layer_container, "Text has updated.", lyr.TextItem.contents)

            file_layer_container += 1                
    

if __name__ == "__main__":
    main()
