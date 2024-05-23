# -*- coding: utf-8 -*-
"""
Created on Thu May 23 11:08:47 2024

@author: Misha
"""

import pywinauto as pwa
import xlwings as xw

import win32gui
import win32ui
import win32con

import numpy as np
import pandas as pd
import cv2

from mltu.inferenceModel import OnnxInferenceModel
from mltu.utils.text_utils import ctc_decoder
from mltu.configs import BaseModelConfigs
import typing

import os

class ImageToWordModel(OnnxInferenceModel):
    def __init__(self, char_list: typing.Union[str, list], *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.char_list = char_list

    def predict(self, image: np.ndarray):
        image = cv2.resize(image, self.input_shape[:2][::-1])

        image_pred = np.expand_dims(image, axis=0).astype(np.float32)

        preds = self.model.run(None, {self.input_name: image_pred})[0]

        text = ctc_decoder(preds, self.char_list)[0]

        return text

def get_model():
    path = str(os.path.abspath(os.getcwd()))
    configs = BaseModelConfigs.load(path + "/configs.yaml")
    model = ImageToWordModel(model_path=configs.model_path, char_list=configs.vocab)
    return model

def find_browser():
    browsers = []
    def winEnumHandler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            n = win32gui.GetWindowText(hwnd)
            #print(n)
            if n =='Сервисы - Федеральная служба судебных приставов — Профиль 1: Microsoft​ Edge':
                browsers.append(hwnd)
    win32gui.EnumWindows(winEnumHandler, None)
    return browsers


model = get_model()
hwnd=find_browser()[0]
appUIA=pwa.Application(backend='uia').connect(handle=hwnd)
mainUIA=appUIA.window(handle=hwnd)

import numpy as np
import cv2
from time import sleep
def solve_captcha():
    send_btn = mainUIA.descendants(control_type='Button', title='Отправить')
    while len(send_btn) != 0:
        img = mainUIA.descendants(control_type='Image')[-1].capture_as_image()
        cv_image = np.array(img)
        padding = 0
        while True:
            if list(cv_image[cv_image.shape[0]//2, padding]) == [255,255,255] or list(cv_image[cv_image.shape[0]//2, padding]) == [140,140,140]:
                padding+=1
            else:
                break
        cv_image_fixed = cv_image[padding:cv_image.shape[0]-padding, padding:cv_image.shape[1]-padding, :]
        cv_image_fixed_scaled = cv2.resize(cv_image_fixed, (200, 60))
        ans = model.predict(cv_image_fixed_scaled)
        mainUIA.descendants(control_type='Edit')[-2].set_edit_text(ans)
        send_btn[0].invoke()
        sleep(1)
        send_btn = mainUIA.descendants(control_type='Button', title='Отправить')
        
def get_dt():
    elems = mainUIA.descendants()
    for i in range(len(elems)):
        if elems[i].window_text() == 'Дата, причина окончания или прекращения ИП (статья, часть, пункт основания)':
            return elems[i+9].window_text()
    return ''

def wait_elem(control_type, title, visible=True):
    elems = mainUIA.descendants(control_type=control_type, title=title)
    c = 0
    while (len(elems) == 0) == visible:
        sleep(2)
        elems = mainUIA.descendants(control_type=control_type, title=title)
        c+=1
        if c > 10:
            raise Exception(f'Не найден элемент {control_type} - {title}')
    

book=xw.Book('data.xlsx')
sht=book.sheets['Sheet1']
df = sht.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value

i = 0
while i != df.shape[0]:
    wait_elem(control_type=None, title='Дата, причина окончания или прекращения ИП (статья, часть, пункт основания)', visible=False)
    mainUIA.descendants(title='Поиск по номеру ИП')[0].invoke()
    sleep(1)
    mainUIA.descendants(control_type='Edit')[1].set_edit_text(df.loc[i, 'Номер ИП'])
    find_button = mainUIA.descendants(control_type='Button', title='Найти')[0]
    find_button.invoke()
    wait_elem(control_type='Button', title='Отправить')
    solve_captcha()
    again = False
    while True:
        if len(mainUIA.descendants(title='По вашему запросу ничего не найдено')) != 0:
            dt = 'Не найдено'
            break
        elif len(mainUIA.descendants(title='Ваш запрос обрабатывается')):
            sleep(20)
            again = True
            break
        elif len(mainUIA.descendants(title='Дата, причина окончания или прекращения ИП (статья, часть, пункт основания)')):
            dt = get_dt()
            break
        sleep(2)
    df.loc[i, 'Дата и статья прекращения'] = dt
    sht.range('A1').options(index=False).value = df
    mainUIA.descendants(title='Обновить')[0].invoke()
    i+=1