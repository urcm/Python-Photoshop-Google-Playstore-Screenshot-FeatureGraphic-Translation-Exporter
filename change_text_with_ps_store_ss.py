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
            
        """
        for i in doc.LayerSets:    
            for a in i.ArtLayers:
                # print(a.Name)
                # print(dict_texts["ss1"])

                if a.name == 'text1':
                    print(a.TextItem.contents)
                    a.TextItem.contents = c["ss1"]
                elif a.name == 'text2':
                    print(a.TextItem.contents)
                    a.TextItem.contents = c["ss2"]
                elif a.name == 'text3':
                    print(a.TextItem.contents)
                    a.TextItem.contents = c["ss3"]
                elif a.name == 'text4':
                    print(a.TextItem.contents)
                    a.TextItem.contents = c["ss4"]
                elif a.name == 'text5':
                    print(a.TextItem.contents)
                    a.TextItem.contents = c["ss5"]
                elif a.name == 'text6':
                    print(a.TextItem.contents)
                    a.TextItem.contents = c["ss6"]
                elif a.name == 'text7':
                    print(a.TextItem.contents)
                    a.TextItem.contents = c["ss7"]
                elif a.name == 'text8':
                    print(a.TextItem.contents)
                    a.TextItem.contents = c["ss8"]
                elif a.name == 'ss_fea':
                    print(a.TextItem.contents)
                    a.TextItem.contents = c["ss_fea"] """ 
            

        # Save for Web dispatch
        options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
        # Define JPEG save options
        options.Format = 6  # JPEG
        options.Quality = 100  # Img Quality Val 0-100

        # Define PNG save options
        # options.Format = 13  # PNG
        # options.PNG8 = False  # Get PNG-24 bit
        
        
        export_location = os.path.join(file_location, c['language'])

        # Check if there is a folder with language name, not create a folder
        if not os.path.exists(export_location):
            os.makedirs(export_location)

        # Save history state to get initial file
        saved_history_state = doc.activeHistoryState
        
        file_del_counter = 0
        for d, i in enumerate(doc.layers):
            fname = os.path.join(export_location, i.name + ".jpg")
            print('Exporting', fname)
            # Tried to set layers invisible but it has not work for Save for Web
            # so all artboards except saving artboard deleting...
            for s in reversed(range(len(doc.layers))):            
                if s is file_del_counter:
                    pass    
                else:
                    print(s, "Deleting artboard Save for Web")
                    if doc.layers[s].Name == 'fea':
                        doc.layers["fea"].Delete()
                    else:
                        doc.layers["ss-" + str(s)].Delete()
                     # time.sleep(1)

if __name__ == "__main__":
    main()
