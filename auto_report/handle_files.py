from pathlib import Path
import os
import re
from PyQt5.QtCore import QObject, pyqtSignal
from .constant_value import *
from .image_to_string import convert_img_duration_to_string, convert_img_freq_to_string, set_tesseract_cmd
# import ptvsd

'''
    though each char in input_string and return index of first digit
    @params:
        input_string: string need to scan
    @return: 
        first index of digit char
        False: if string haven't digit char
'''
def has_number(input_string):
    for index, char in enumerate(input_string):
        if char.isdigit():
            return index
    return False

"""
    Check a attr_name has attr of chime
    @params: attr_name: a string contain chime attribute
    @return:
        key for dict chime
        NOT_OK if string hasn't attr
"""
def filter_chime_attr(attr_name):
    attr = attr_name.lower()
    attr_type = ""
    if "freq" in attr:
        attr_type = FREQ_IMG_COORD
    elif "attack" in attr:
        attr_type = ATTACK_IMG_COORD
    elif "peak" in attr:
        attr_type = PEAK_IMG_COORD
    elif "decaywithin" in attr:
        attr_type = DECAY_WITHIN_IMG_COORD
    elif "fall" in attr or "decayafter" in attr:
        attr_type = FALL_IMG_COORD
    elif "zero" in attr:
        attr_type = ZERO_GAIN_IMG_COORD
    elif "total" in attr:
        attr_type = TOTAL_DURATION_IMG_COORD
    elif "amplitude" in attr or "amp" in attr:
        attr_type = AMPLITUDE_IMG_COORD
    elif "waveform" in attr or "wareform":
        attr_type = WAVEFORM_IMG_COORD
    else:
        attr_type = NOT_OK
    return attr_type

'''
    Scan folder and get all image type of '.png'
    @params:
        folder_path: folder shall to scan
    @return:
        list of image path
'''
def get_all_images(folder_path):
    with os.scandir(folder_path) as entries:
        list_chimes = []
        # loop in Image folder
        for entry in entries:
            if os.path.isfile(entry.path) and entry.name.lower().endswith('.png'):
                list_chimes.append(entry)
    return list_chimes

class FileHandler(QObject):
    # emit signal when handled a file
    handledFileSignal = pyqtSignal(bool)
    # emit progressed of handle files
    progressed = pyqtSignal(int)
    # emit when tesseract path not exist
    tesseractNotExist = pyqtSignal()
    # emit when finished handled all files
    finished = pyqtSignal(dict)

    '''
    init of class FileHandler
    @params:
        - files: all files *.png
        - is_fill_actual_value: 
            + true -> try to convert img to string using tesseract OCR
            + false -> do nothing
        - tesseract_path: path to folder install tesseract app
        - report_type:
    '''
    def __init__(self, files, is_fill_actual_value, tesseract_path, report_type):
        super().__init__()
        self._files = files
        self._is_fill_actual_value = is_fill_actual_value
        self._tesseract_path = tesseract_path
        self._report_type = report_type
    
    """ set tesseract_cmd by your tesseract app path """
    def set_tesseract_path(self):
        if self._is_fill_actual_value:
            if self._tesseract_path != '.' and self._tesseract_path != '':
                set_tesseract_cmd(self._tesseract_path)
            else:
                pass

    ''' 
    hand single image path
    steps:
        1. split files name to id, Chime name, attr, index.
        2. 

    @params:
        - img_entry: image path with wrapper Path from pathlib
        - dict_chimes: an empty dict to get data struct of chimes after analyze image path
    @return:
        - True/False
    '''
    def _handle_chime_image(self, img_entry, dict_chimes):
        # split files
        split_file_name = re.split('-|_|\.', img_entry.name)
        # remove extension
        if 'png' in split_file_name:
            split_file_name.remove('png')
        elif 'PNG' in split_file_name:
            split_file_name.remove('PNG')
        else:
            return False
        # return False if path not contain enough info   
        if len(split_file_name) <= 2:
            return False
        ''' handle image path '''
        # cursor will through each part of split name
        current_cursor = 0
        # get chime id
        chime_id = split_file_name[current_cursor]
        # get chime name
        current_cursor += 1
        chime_name = split_file_name[current_cursor]
        if chime_name[:5] == CHIME_TEXT:
            if chime_name[5:].isdigit():
                current_cursor += 1
            else:
                current_cursor += 1
                chime_name = ''.join(chime_name + split_file_name[current_cursor])
                current_cursor += 1
        else:
            return False
        # check and get zone name
        if ((chime_name == CHIME16) or
            (chime_name == CHIME17)) and \
            (ZONE in split_file_name[current_cursor]):
            zone_name = split_file_name[current_cursor]
            if zone_name[4:].isdigit():
                current_cursor += 1
            else:
                current_cursor += 1
                zone_name = ''.join(zone_name + split_file_name[current_cursor])
                current_cursor += 1
            chime_name = '{}_{}'.format(chime_name, zone_name)
        # handle index
        if has_number(split_file_name[current_cursor]):
            index = has_number(split_file_name[current_cursor])
            chime_index = split_file_name[current_cursor][index:]
            chime_attr = split_file_name[current_cursor][:index]
            # print(chime_index)
        # shoud remove BELLOW CONDITION if analyze ratio
        elif current_cursor == (len(split_file_name) - 1):
            chime_index = '1'
            chime_attr = split_file_name[current_cursor]
        elif split_file_name[current_cursor + 1].isdigit():
            chime_index = split_file_name[current_cursor + 1]
            chime_attr = split_file_name[current_cursor]
        else:
            return False
        # check chime in dict
        if chime_name not in dict_chimes.keys():
            dict_chimes.update({chime_name: {}})
        else:
            pass
        # check index in dict
        if chime_index not in dict_chimes[chime_name].keys():
            dict_chimes[chime_name].update({chime_index: {}})
        else:
            pass

        # filter chime attribute
        attr = filter_chime_attr(chime_attr)
        if attr == NOT_OK:
            return False

        # create sub_dict attribute
        dict_chimes[chime_name][chime_index].update({attr: {}})
        # add attribute path to dict
        dict_chimes[chime_name][chime_index][attr][PATH] = img_entry.path
        # add attribute basename to dict
        dict_chimes[chime_name][chime_index][attr][BASENAME] = img_entry.name
        #check fill or not
        if self._is_fill_actual_value and (self._report_type == OSC_TYPE or self._report_type == WAVEFORM_TYPE):
        # get duration and freq
            if attr == FREQ_IMG_COORD:
                freq = 0
                try:
                    freq = convert_img_freq_to_string(img_entry.path)
                except:
                    self.tesseractNotExist.emit()
                dict_chimes[chime_name][chime_index][attr][FREQ] = freq
            elif attr == AMPLITUDE_IMG_COORD:
                dict_chimes[chime_name][chime_index][attr][DURATION] = 0
            else:
                duration = 0
                try:
                    duration = convert_img_duration_to_string(img_entry.path)
                    #        if attr == TOTAL_DURATION_IMG_COORD:
                    #           duration = float(duration/1000)
                    #        else:
                    #            pass
                except:
                    self.tesseractNotExist.emit()
                dict_chimes[chime_name][chime_index][attr][DURATION] = duration
        else:
            pass
        return True

    ''' 
    This function loop in a list of file, handle file name and return result of handle
    @params: None
    @return: None
    '''
    def handle_chime_images_v2(self):
        # set tesseract path for image processing
        self.set_tesseract_path()
        # ptvsd.debug_this_thread()
        dict_chimes = {}
        # loop list of files
        for index, entry in enumerate(self._files):
            return_value = self._handle_chime_image(entry, dict_chimes)
            self.handledFileSignal.emit(return_value)
            self.progressed.emit(index)
        self.finished.emit(dict_chimes)
        return

