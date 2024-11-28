# cython: language_level=3
# Boa:Frame:Frame1
# 老版本代码

import os
import time
import threading
import datetime
import json
import traceback
import io

import wx
from wx.adv import TaskBarIcon as wxTaskBarIcon
from wx.adv import EVT_TASKBAR_LEFT_DCLICK

import pyWinhook
import win32con, win32print
import win32api, win32gui
import pyperclip
from playsound import playsound
from playsound import PlaysoundException
import re

import i18n
import config


VERSION = '3.3'


wx.NO_3D = 0
HOT_KEYS = ['F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11', 'F12']

conf = config.getconfig()
i18n.load_path.append('i18n')
i18n.set('locale', conf['language'])
ID_MAP = {'en':0, 'zh-tw':1}
RID_MAP = {0:'en', 1:'zh-tw'}

def GetMondrianStream():
    data = b'\x89\x50\x4e\x47\x0d\x0a\x1a\x0a\x00\x00\x00\x0d\x49\x48\x44\x52\x00\x00\x00\x20\x00\x00\x00\x20\x08\x06\x00\x00\x00\x73\x7a\x7a\xf4\x00\x00\x00\x01\x73\x52\x47\x42\x00\xae\xce\x1c\xe9\x00\x00\x05\xf6\x49\x44\x41\x54\x58\x47\xc5\x97\x7b\x4c\x53\x57\x1c\xc7\xbf\xe7\xb6\x22\x02\xca\xcb\x82\x62\xa5\x55\x06\x94\xb7\x93\xc5\xa9\x80\x33\x8a\x9b\x31\xd9\x66\x8c\x9a\xc8\x43\x45\x97\x81\x82\x60\x7c\xcd\xb7\x9d\x73\xd1\x6c\x9a\xa9\xd3\xc9\x4c\x9c\xaf\x2d\x53\xa3\x1b\xce\x25\xa8\x63\x51\x54\x14\x71\x13\x85\x22\x14\x4b\x29\x22\x20\x02\x05\x29\x0c\x4b\xb9\xf7\x2c\xf7\x32\x18\x95\x56\x0a\x9b\xf1\xf7\x4f\x6f\xef\x3d\xe7\xfc\x3e\xe7\xfb\x7b\x9c\x7b\x09\x5e\xb3\x91\xd7\xec\x1f\xaf\x14\x20\x2c\xd4\x3b\x9f\x35\xb5\x85\x81\x12\xa8\x4a\x9e\x32\x96\x36\xdb\x0b\x20\x44\xe1\x46\x41\x45\xb4\x50\x5d\x67\x71\x82\x2d\x8a\x6d\x58\x9b\x12\x71\xe1\xc2\xa9\xeb\xca\xe4\x0d\xc4\x6d\x88\x33\x52\x77\x6f\xe4\x01\x2c\x6e\xd6\xec\x66\x88\xc2\x95\x8b\x9b\x16\x45\xae\xaa\xb4\xa8\x7c\x5a\x07\x10\xe6\x2f\x55\x71\x8d\xa3\x2d\x4e\xbb\xc6\xc8\xe5\x72\x17\xc9\x30\xa3\x7e\xe1\xfb\x31\x44\x31\x26\x00\x45\xb7\xf3\x70\xe4\xda\x39\x14\xaa\x1b\xfa\x06\x08\x56\x0c\xa7\xfb\xd6\xec\x04\x01\x83\xbc\xfc\x5f\x70\xa5\x50\x8b\xea\x27\x35\x88\x88\x9c\x7a\x36\xfd\xbb\xb3\xf3\xfa\x02\x51\x2a\x95\x4c\xc6\xe9\x83\x1d\x61\xf2\x40\xf2\xc1\xb4\x59\xa8\xd6\x55\x22\x68\x5c\x28\x56\xee\xd9\x84\x82\x92\xba\xfe\x00\xf0\xae\x28\x38\xb0\xb8\x71\x27\x13\x57\xff\x2c\x82\xde\xd0\x8c\xc3\xc7\x7f\x1c\x31\x71\x62\x74\x6d\x98\x9f\x07\x0b\x50\x06\x20\x7c\x78\xcd\xac\xf3\x3f\x41\xa4\x62\x3c\x7c\x25\xde\x18\x37\x21\xbc\x1f\x00\x01\x23\x5b\x77\xaf\xdd\xea\x60\xc7\x89\x71\x2f\xef\x0f\x50\x42\x11\x38\x71\x02\x54\xb9\xd9\x78\x50\x55\x8a\x42\x8d\x0e\x1d\x94\x43\xcc\xf4\x08\x98\x18\x31\x18\x4a\xc0\x11\x8a\x45\x6b\xbe\x00\x43\xa8\x00\xc2\x76\xb0\xa8\x28\xd3\xe1\xd8\xf1\x74\xe4\xdd\xbd\x85\x5d\xab\x37\x63\xfd\x97\x9f\xa2\x40\x6d\x43\x0e\xf0\x0b\x04\x2b\xdc\xe9\xfe\x35\xbb\xc0\xaf\x57\x90\x9f\x0f\xdf\xd0\x60\x3c\x6f\x6e\x41\x99\xf6\x0e\xa8\xc8\x08\x0a\x11\xc4\x0e\x4e\x58\x90\xa2\xe4\x45\xe2\x37\xdb\xcb\x4c\x1d\xed\xa8\x7a\x5c\x83\x67\x0d\x0d\xd8\xb1\x63\x03\x9a\x5b\x9b\x51\xd8\x5f\x80\xa6\x06\x3d\x5c\x5c\x5d\xa1\xd3\xe9\xe0\xec\xec\x0c\x4d\x79\x3e\x3c\x7d\x7d\x20\x16\x89\x10\x3d\x3b\x0e\xd7\x6e\xe6\x61\xec\x1b\xbe\x28\xd7\xa8\x11\x11\x11\x61\x06\xa1\xd1\xa8\x71\xe6\xa7\x0b\xb8\xfa\xfb\x15\x7c\xb4\x38\x06\x47\x8e\x1e\xb0\x1d\x60\x79\xe2\xd2\x99\xfa\x2a\x4d\xe6\xdc\x29\xb3\xd0\xa4\x6f\x84\x6c\x8c\x0c\xa5\xc5\x45\x90\x85\x29\xf0\xe6\x94\x19\x20\x62\x31\x36\x6e\xda\x82\x6b\xd9\x39\x50\x04\x29\xb0\x38\x7e\x01\x22\x23\xa3\xc0\x50\x8a\x2b\x19\xe7\x21\x93\x8f\x41\x73\x6b\x0b\xec\xdc\xdd\x50\xdf\x50\x87\xa7\xd5\xf5\x00\x8c\xb8\x9e\x73\xe3\xe4\xe9\x8c\x8b\x0b\x5f\x94\xcb\x62\x66\x06\x06\x48\xe8\x3e\x3e\x0c\x46\x23\x0a\x55\xf9\x68\x47\x23\xa6\x7e\x18\x03\xa9\x22\x14\x04\x04\x7a\xbd\x1e\x42\xf6\x11\x0e\x6e\x6e\x6e\x42\x1c\x2a\x34\x1a\x98\x0c\xad\xb8\x97\x9d\x83\xf0\xe9\x53\xe1\xe1\xed\x0d\x47\x67\x07\xb0\x2c\x50\xa1\xd1\xc2\x60\x68\xc4\xfa\x6d\x6b\x55\xf7\x55\x95\x21\x3d\x21\x2c\x02\x04\x29\x24\x34\xe1\x9d\x39\x20\xc4\x04\x96\x69\x12\x92\x11\x84\xc1\xfc\xe5\x9b\x31\xd8\x71\x58\xf7\x7c\xad\xfa\x61\xe7\x35\x01\x52\x56\x2c\x11\x7e\x09\xc7\x82\x23\x22\x24\x2c\x4a\xc4\x5b\xe1\x6f\x63\xac\x9f\x2f\x40\x29\xca\x4b\xcb\x90\x94\x16\x83\xc2\x12\xf3\x7e\x60\xb5\x15\x8f\xf3\x97\xd0\x05\xd1\x93\xbb\x9d\x15\x3f\xae\x45\x7c\xc2\x3a\x04\x86\x85\x82\x32\x04\x5a\xb5\x06\x0c\x47\x90\xb4\x32\x16\x84\x12\x78\xd9\x39\x99\xa9\xdb\xc8\x1a\xd1\x4a\x8d\xf8\xe6\xab\x93\xf0\x09\xf0\x45\xb9\x5a\x03\xca\x71\x48\x49\x5b\xc2\xdd\x53\x3f\x11\x75\x0d\xb6\x0a\x10\xea\xe7\x41\x63\x67\x4c\x12\x94\xae\x7b\xd6\x86\x07\xba\x2a\xac\x5a\xf5\x19\x1c\x1c\x1c\x3b\xd5\xa7\x04\x89\x69\x71\x90\x89\x9d\xc0\x81\x0a\xa1\xe9\x32\xa1\x20\x09\xc0\x70\x40\x95\xa9\x19\x07\xbe\x3e\x01\x1f\x85\x1f\xca\x4b\xb4\x48\x5a\x11\x87\x82\xd2\x7f\x4b\xd2\x3a\x80\xbf\x3b\x95\x4b\x3c\x30\x3e\xc8\x17\xb7\x8b\xb5\x82\xc3\x55\xab\xb7\x62\xc8\x90\xce\x10\x2c\x4b\x8d\xc7\x68\xb1\x03\x28\x79\xf9\x91\xc1\x50\xe0\x91\xc9\x80\x83\xfb\x4f\x40\x0c\x11\x12\x53\x6d\x05\x50\xb8\x73\x52\x89\x17\x01\x38\xc1\xe1\xdc\x79\x4b\xa1\xf0\x0f\x12\x4a\x9f\xb7\xa4\x54\x7e\xf7\x43\xc1\xf5\x71\x9e\xf2\xe3\x2b\x4d\x06\xa4\xef\x3f\x09\x10\x82\xe5\x29\xf1\xb6\x29\xc0\x3b\x09\xf5\x1b\x4e\xa5\x9e\x9e\x56\x01\xe4\xe2\xa1\x60\xfb\x00\xe0\x01\x1f\xb7\x0f\x10\x20\xcc\xcf\x93\x95\x8e\x90\x30\x94\x12\x3a\x6f\xfe\x22\xe2\xeb\x1f\x2c\x84\xa2\x4b\x01\xa9\xdd\x50\xa1\x63\x5a\x33\x7e\x28\x1f\x02\x5d\x87\x01\xdf\xee\x1b\x80\x02\x4a\xa5\x52\x9c\x71\xea\x50\xfb\xf4\xa8\x77\xc9\x9c\xb9\xf1\xc2\x01\xd5\x55\x77\x49\xa9\xb1\x98\xe4\xe2\x8d\xaa\x96\xc6\x5e\x07\x52\x4f\xa0\x51\x8e\x2e\xc8\x6d\xaa\xec\x0e\xc1\xb2\xfe\x84\x20\x2b\xeb\xbc\xe7\xae\x6d\xdb\x6b\xb6\x6f\xfd\x9c\xf0\x79\xde\xd3\x56\xaf\x4f\x86\x8f\xbd\x13\x9c\x38\x11\x6a\x5a\x9a\x84\x47\x1d\xff\xc8\xc1\xef\x5a\x04\x02\x77\xa7\x61\xc2\xbd\x52\x63\x0b\x76\xef\x3c\x08\x06\x44\xc8\x1d\x9b\xaa\x40\xc8\x81\x80\x91\xdc\xa1\xbd\xc7\xac\x46\x79\x6f\xfa\x17\xc0\xa3\x5a\xb0\x60\xe1\x3a\xd8\x01\x0d\x6d\x2d\x60\x38\x0a\x7b\xd1\x20\x40\x2c\x82\xc9\x64\x82\x68\xf4\x08\xac\x48\x5e\x07\x80\x41\x91\xba\x80\x66\x65\x5f\x3e\x7b\xe6\xdc\x6f\xf3\xfb\xec\x03\x0b\x63\x66\x6f\x99\x19\xfd\xde\xc7\xb2\x51\x63\xa5\x56\x83\x4c\x38\xdc\xbd\x73\x0b\x9c\x88\xe0\xe7\xef\x8f\x09\xb5\xcf\xf7\x84\x41\x20\x08\x9f\x3c\x15\x52\x99\x0c\xe1\xe3\x27\x0b\xd9\xcf\x5b\x62\x5a\x2c\x54\x25\xf5\x66\x1b\xb2\xde\x07\x14\x1e\xdc\xe1\x03\x3f\x70\x2c\xcb\x75\x77\x2d\x8b\x20\x84\xf2\x9d\x16\xcd\x86\x7a\x50\xfe\x82\xcf\x14\xc2\xc0\xd1\xde\x09\x76\x83\xec\x51\x5d\x53\x01\x2f\x2f\x6f\x7c\xb2\x39\x85\xee\xdc\x73\x54\x1e\x15\x15\xf5\xa8\xe7\x3a\x2f\x69\x44\x1e\xf4\x62\x66\x0e\x5b\xfe\x50\xfb\x72\x00\x0b\x54\x6d\xcf\x5b\x50\xd7\xf0\x44\x78\x69\x92\x8e\x1a\x0d\x86\x11\x09\xb1\x7f\xf1\x1c\xe0\xa7\x5a\x04\x08\x09\xf2\xe8\xc8\xfc\xf5\xa6\x88\xf7\xac\x2d\x2d\xb3\x5e\x67\x56\x9e\xf0\x4a\x10\x3e\x13\x29\x03\x0a\x16\xc9\x69\x09\xb8\x5f\x52\xcb\xb7\xcc\x5e\x45\x6b\x11\x20\xd8\x5f\x42\x2f\x5f\xba\x8d\xf2\xae\xd3\xae\xdf\x08\xe6\x13\x96\xa5\xc5\xa3\xc0\x96\xd7\xf2\xae\x69\x3c\xc0\xa5\x8b\xb9\xa8\x28\xd5\xf4\x46\x1e\x00\xcc\x00\x00\x3c\x69\xd6\xa5\x5c\x94\xa9\x35\x03\x70\x67\x3e\x85\x10\x8a\xa4\x54\x5e\x81\xda\xbe\x5f\xcb\x7b\x2a\xf0\x9f\x3d\xbf\xb0\x80\x4a\x6d\xc3\x77\xc1\xff\xed\xd4\x96\xf5\x5e\xe9\xc7\xa9\x2d\x00\x7f\x03\x99\xb1\x75\x3f\x1a\x2e\x31\x2f\x00\x00\x00\x00\x49\x45\x4e\x44\xae\x42\x60\x82'
    stream = io.BytesIO(data)
    return stream


def GetMondrianBitmap():
    stream = GetMondrianStream()
    image = wx.ImageFromStream(stream)
    return wx.BitmapFromImage(image)


def GetMondrianIcon():
    icon = wx.EmptyIcon()
    icon.CopyFromBitmap(GetMondrianBitmap())
    return icon


def create(parent):
    return Frame1(parent)


def current_ts():
    return int(time.time() * 1000)
    

[wxID_FRAME1, wxID_FRAME1BTRECORD, wxID_FRAME1BTRUN, wxID_FRAME1BTPAUSE, wxID_FRAME1BUTTON1,
 wxID_FRAME1CHOICE_SCRIPT, wxID_FRAME1CHOICE_START, wxID_FRAME1CHOICE_STOP, wxID_FRAME1CHOICE_RECORD,
 wxID_FRAME1PANEL1, wxID_FRAME1STATICTEXT1, wxID_FRAME1STATICTEXT2,
 wxID_FRAME1STATICTEXT3, wxID_FRAME1STATICTEXT4, wxID_FRAME1STIMES,
 wxID_FRAME1TEXTCTRL1, wxID_FRAME1TEXTCTRL2, wxID_FRAME1TNUMRD,
 wxID_FRAME1TSTOP, wxID_FRAME1STATICTEXT5, wxID_FRAME1TEXTCTRL3,
 wxID_LANGUAGECHOICE
] = [wx.NewId() for _init_ctrls in range(22)]


# SW = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
# SH = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
hDC = win32gui.GetDC(0)
SW = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
SH = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)


class Frame1(wx.Frame):
    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_FRAME1, name='', parent=prnt,
              pos=wx.Point(SW / 2 - 183, SH / 2 - 115.5), size=wx.Size(366, 271),
              style=wx.STAY_ON_TOP | wx.DEFAULT_FRAME_STYLE,
              title='KeymouseGo v%s' % VERSION)
        self.SetClientSize(self.FromDIP(wx.Size(361, 270)))

        self.panel1 = wx.Panel(id=wxID_FRAME1PANEL1, name='panel1', parent=self,
              pos=self.FromDIP(wx.Point(0, 0)), size=self.FromDIP(wx.Size(350, 245)),
              style=wx.NO_3D | wx.CAPTION)

        self.btrecord = wx.Button(id=wxID_FRAME1BTRECORD, label=i18n.t('lang.record'),
              name='btrecord', parent=self.panel1, pos=self.FromDIP(wx.Point(213, 12)),
              size=self.FromDIP(wx.Size(56, 32)), style=0)
        self.btrecord.Bind(wx.EVT_BUTTON, self.OnBtrecordButton,
              id=wxID_FRAME1BTRECORD)

        self.btrun = wx.Button(id=wxID_FRAME1BTRUN, label=i18n.t('lang.launch'),
              name='btrun', parent=self.panel1, pos=self.FromDIP(wx.Point(285, 12)),
              size=self.FromDIP(wx.Size(56, 32)), style=0)
        self.btrun.Bind(wx.EVT_BUTTON, self.OnBtrunButton, id=wxID_FRAME1BTRUN)

        # 暂停/继续 功能不适合用按钮的形式来做，所以暂时隐去
        # self.btpause = wx.Button(id=wxID_FRAME1BTPAUSE, label='暂停',
        #       name='btpause', parent=self.panel1, pos=self.FromDIP(wx.Point(274, 141)),
        #       size=self.FromDIP(wx.Size(56, 32)), style=0)
        # self.btpause.Bind(wx.EVT_BUTTON, self.OnBtpauseButton, id=wxID_FRAME1BTPAUSE)

        # 暂停录制
        self.btpauserecord = wx.Button(id=wxID_FRAME1BTPAUSE, label=i18n.t('lang.pauserecord'),
               name='btpauserecording', parent=self.panel1, pos=self.FromDIP(wx.Point(213, 135)),
               size=self.FromDIP(wx.Size(86, 32)), style=0)
        self.btpauserecord.Bind(wx.EVT_BUTTON, self.OnPauseRecordButton, id=wxID_FRAME1BTPAUSE)
        self.btpauserecord.Enable(False)

        self.tnumrd = wx.StaticText(id=wxID_FRAME1TNUMRD, label='ready..',
              name='tnumrd', parent=self.panel1, pos=self.FromDIP(wx.Point(17, 245)),
              size=self.FromDIP(wx.Size(100, 36)), style=0)

        self.button1 = wx.Button(id=wxID_FRAME1BUTTON1, label='test',
              name='button1', parent=self.panel1, pos=self.FromDIP(wx.Point(128, 296)),
              size=self.FromDIP(wx.Size(75, 24)), style=0)
        self.button1.Bind(wx.EVT_BUTTON, self.OnButton1Button,
              id=wxID_FRAME1BUTTON1)

        self.tstop = wx.StaticText(id=wxID_FRAME1TSTOP,
              label='If you want to stop it, Press F12', name='tstop',
              parent=self.panel1, pos=self.FromDIP(wx.Point(25, 332)), size=self.FromDIP(wx.Size(183, 18)),
              style=0)
        self.tstop.Show(False)

        self.stimes = wx.SpinCtrl(id=wxID_FRAME1STIMES, initial=0, max=1000,
              min=0, name='stimes', parent=self.panel1, pos=self.FromDIP(wx.Point(217, 101)),
              size=self.FromDIP(wx.Size(45, 18)), style=wx.SP_ARROW_KEYS)
        self.stimes.SetValue(int(conf['looptimes']))

        self.label_run_times = wx.StaticText(id=wxID_FRAME1STATICTEXT2,
              label=i18n.t('lang.tip'),
              name='label_run_times', parent=self.panel1, pos=self.FromDIP(wx.Point(214, 57)),
              size=self.FromDIP(wx.Size(146, 36)), style=0)

        self.textCtrl1 = wx.TextCtrl(id=wxID_FRAME1TEXTCTRL1, name='textCtrl1',
              parent=self.panel1, pos=self.FromDIP(wx.Point(24, 296)), size=self.FromDIP(wx.Size(40, 22)),
              style=0, value='119')

        self.textCtrl2 = wx.TextCtrl(id=wxID_FRAME1TEXTCTRL2, name='textCtrl2',
              parent=self.panel1, pos=self.FromDIP(wx.Point(80, 296)), size=self.FromDIP(wx.Size(36, 22)),
              style=0, value='123')

        self.label_script = wx.StaticText(id=wxID_FRAME1STATICTEXT3,
              label=i18n.t('lang.script'), name='label_script', parent=self.panel1,
              pos=self.FromDIP(wx.Point(17, 20)), size=self.FromDIP(wx.Size(40, 32)), style=0)

        self.choice_script = wx.Choice(choices=[], id=wxID_FRAME1CHOICE_SCRIPT,
              name='choice_script', parent=self.panel1, pos=self.FromDIP(wx.Point(100, 15)),
              size=self.FromDIP(wx.Size(108, 25)), style=0)

        self.label_start_key = wx.StaticText(id=wxID_FRAME1STATICTEXT1,
              label=i18n.t('lang.launchhotkey'), name='label_start_key',
              parent=self.panel1, pos=self.FromDIP(wx.Point(16, 61)), size=self.FromDIP(wx.Size(80, 36)),
              style=0)

        self.label_stop_key = wx.StaticText(id=wxID_FRAME1STATICTEXT4,
              label=i18n.t('lang.terminatehotkey'), name='label_stop_key',
              parent=self.panel1, pos=self.FromDIP(wx.Point(16, 102)), size=self.FromDIP(wx.Size(76, 32)),
              style=0)

        self.label_record_key = wx.StaticText(id=wxID_FRAME1STATICTEXT1,
              label=i18n.t('lang.recordhotkey'), name='label_record_key',
              parent=self.panel1, pos=self.FromDIP(wx.Point(16, 141)),
              size=self.FromDIP(wx.Size(80, 36)),
              style=0)

        self.choice_start = wx.Choice(choices=[], id=wxID_FRAME1CHOICE_START,
              name='choice_start', parent=self.panel1, pos=self.FromDIP(wx.Point(100, 58)),
              size=self.FromDIP(wx.Size(108, 25)), style=0)
        self.choice_start.SetLabel('')
        self.choice_start.SetLabelText('')
        self.choice_start.Bind(wx.EVT_CHOICE, self.OnChoice_startChoice,
              id=wxID_FRAME1CHOICE_START)

        self.choice_stop = wx.Choice(choices=[], id=wxID_FRAME1CHOICE_STOP,
              name='choice_stop', parent=self.panel1, pos=self.FromDIP(wx.Point(100, 98)),
              size=self.FromDIP(wx.Size(108, 25)), style=0)
        self.choice_stop.Bind(wx.EVT_CHOICE, self.OnChoice_stopChoice,
              id=wxID_FRAME1CHOICE_STOP)

        self.choice_record = wx.Choice(choices=[], id=wxID_FRAME1CHOICE_RECORD,
              name='choice_record', parent=self.panel1, pos=self.FromDIP(wx.Point(100, 138)),
              size=self.FromDIP(wx.Size(108, 25)), style=0)
        self.choice_record.Bind(wx.EVT_CHOICE, self.OnChoice_recordChoice,
              id=wxID_FRAME1CHOICE_RECORD)

        self.label_mouse_interval = wx.StaticText(
              label=i18n.t('lang.precision'), name='label_mouse_interval',
              parent=self.panel1, pos=self.FromDIP(wx.Point(16, 181)), size=self.FromDIP(wx.Size(56, 32)),
              style=0)

        self.mouse_move_interval_ms = wx.SpinCtrl(initial=int(conf['precision']), max=999999,
              min=0, name='mouse_move_interval_ms', parent=self.panel1, pos=self.FromDIP(wx.Point(100, 181)),
              size=self.FromDIP(wx.Size(68, 18)), style=wx.SP_ARROW_KEYS)

        self.label_mouse_interval_tips = wx.StaticText(
              label=i18n.t('lang.tip2'), name='label_mouse_interval_tips',
              parent=self.panel1, pos=self.FromDIP(wx.Point(171, 180)), size=self.FromDIP(wx.Size(170, 60)),
              style=0)

        self.label_execute_speed = wx.StaticText(
              label=i18n.t('lang.speed'), name='label_execute_speed',
              parent=self.panel1, pos=self.FromDIP(wx.Point(16, 216)), size=self.FromDIP(wx.Size(70, 32)),
              style=0)

        self.execute_speed = wx.SpinCtrl(initial=int(conf['executespeed']), max=500,
              min=20, name='execute_speed', parent=self.panel1,
              pos=self.FromDIP(wx.Point(100, 216)),
              size=self.FromDIP(wx.Size(68, 18)), style=wx.SP_ARROW_KEYS)

        self.label_execute_speed_tips = wx.StaticText(
            label=i18n.t('lang.range'), name='label_execute_speed_tips',
            parent=self.panel1, pos=self.FromDIP(wx.Point(171, 216)), size=self.FromDIP(wx.Size(120, 20)),
            style=0)

        self.label_language = wx.StaticText(
            label=i18n.t('lang.language'), style=wx.ALIGN_RIGHT,
            parent=self.panel1, pos=self.FromDIP(wx.Point(120, 245)), size=self.FromDIP(wx.Size(150, 50)),
            )

        self.choice_language = wx.Choice(choices=['简体中文', 'English'], id=wxID_LANGUAGECHOICE,
            parent=self.panel1, pos=self.FromDIP(wx.Point(281, 240)),
            size=self.FromDIP(wx.Size(70, 25)), style=0)
        self.choice_language.SetSelection(ID_MAP[conf['language']])

        # ===== if use SetProcessDpiAwareness, comment below =====
        # self.label_scale = wx.StaticText(id=wxID_FRAME1STATICTEXT5,
        #       label='屏幕缩放', name='staticText5',
        #       parent=self.panel1, pos=self.FromDIP(wx.Point(16, 141)), size=self.FromDIP(wx.Size(56, 32)),
        #       style=0)
        # self.text_scale = wx.TextCtrl(id=wxID_FRAME1TEXTCTRL3, name='textCtrl3',
        #       parent=self.panel1, pos=self.FromDIP(wx.Point(79, 138)), size=self.FromDIP(wx.Size(108, 22)),
        #       style=0, value='100%')
        # =========================================================

    def __init__(self, parent):

        self._init_ctrls(parent)

        self.SetIcon(GetMondrianIcon())
        self.taskBarIcon = TaskBarIcon(self)
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.Bind(wx.EVT_ICONIZE, self.OnIconfiy)

        if not os.path.exists('../scripts'):
            os.mkdir('../scripts')
        self.scripts = os.listdir('../scripts')[::-1]

        self.scripts = list(filter(lambda s: s.endswith('.txt'), self.scripts))
        self.choice_script.SetItems(self.scripts)
        if self.scripts:
            self.choice_script.SetSelection(0)

        self.choice_start.SetItems(HOT_KEYS)
        self.choice_start.SetSelection(int(conf['starthotkeyindex']))

        self.choice_stop.SetItems(HOT_KEYS)
        self.choice_stop.SetSelection(int(conf['stophotkeyindex']))

        self.choice_record.SetItems(HOT_KEYS)
        self.choice_record.SetSelection(int(conf['recordhotkeyindex']))

        self.running = False
        self.recording = False
        self.record = []
        self.ttt = current_ts()

        # for pause-resume feature
        self.paused = False
        self.pause_event = threading.Event()

        # Pause-Resume Record
        self.pauserecord = False

        self.actioncount = 0

        # For better thread control
        self.runthread = None
        self.isbrokenorfinish = True

        self.hm = pyWinhook.HookManager()

        def on_mouse_event(event):

            # print('MessageName:',event.MessageName)  #事件名称
            # print('Message:',event.Message)          #windows消息常量 
            # print('Time:',event.Time)                #事件发生的时间戳        
            # print('Window:',event.Window)            #窗口句柄         
            # print('WindowName:',event.WindowName)    #窗口标题
            # print('Position:',event.Position)        #事件发生时相对于整个屏幕的坐标
            # print('Wheel:',event.Wheel)              #鼠标滚轮
            # print('Injected:',event.Injected)        #判断这个事件是否由程序方式生成，而不是正常的人为触发。
            # print('---')

            if not self.recording or self.running or self.pauserecord:
                return True

            message = event.MessageName
            if message == 'mouse wheel':
                message += ' up' if event.Wheel == 1 else ' down'
            elif message in config.swapmousemap and config.swapmousebuttons:
                message = config.swapmousemap[message]
            all_messages = ('mouse left down', 'mouse left up', 'mouse right down', 'mouse right up', 'mouse move',
                            'mouse middle down', 'mouse middle up', 'mouse wheel up', 'mouse wheel down')
            if message not in all_messages:
                return True

            pos = win32api.GetCursorPos()

            delay = current_ts() - self.ttt

            # 录制鼠标轨迹的精度，数值越小越精准，但同时可能产生大量的冗余
            mouse_move_interval_ms = self.mouse_move_interval_ms.Value or 999999

            if message == 'mouse move' and delay < mouse_move_interval_ms:
                return True

            self.ttt = current_ts()
            if not self.record:
                delay = 0

            x, y = pos
            tx = x / SW
            ty = y / SH
            tpos = (tx, ty)

            print(delay, message, tpos)

            # self.record.append([delay, 'EM', message, tpos])
            self.record.append([delay, 'EM', message, ['{0}%'.format(tx), '{0}%'.format(ty)]])
            self.actioncount = self.actioncount + 1
            text = '%d actions recorded' % self.actioncount

            self.tnumrd.SetLabel(text)
            return True

        def on_keyboard_event(event):

            # print('MessageName:',event.MessageName)          #同上，共同属性不再赘述
            # print('Message:',event.Message)
            # print('Time:',event.Time)
            # print('Window:',event.Window)
            # print('WindowName:',event.WindowName)
            # print('Ascii:', event.Ascii, chr(event.Ascii))   #按键的ASCII码
            # print('Key:', event.Key)                         #按键的名称
            # print('KeyID:', event.KeyID)                     #按键的虚拟键值
            # print('ScanCode:', event.ScanCode)               #按键扫描码
            # print('Extended:', event.Extended)               #判断是否为增强键盘的扩展键
            # print('Injected:', event.Injected)
            # print('Alt', event.Alt)                          #是某同时按下Alt
            # print('Transition', event.Transition)            #判断转换状态
            # print('---')

            message = event.MessageName
            message = message.replace(' sys ', ' ')

            if message == 'key up':
                # listen for start/stop script
                key_name = event.Key.lower()
                # start_name = 'f6'  # as default
                # stop_name = 'f9'  # as default

                start_index = self.choice_start.GetSelection()
                stop_index = self.choice_stop.GetSelection()
                record_index = self.choice_record.GetSelection()

                # Predict potential conflict
                if start_index == stop_index:
                    stop_index = (stop_index + 1) % len(HOT_KEYS)
                    self.choice_stop.SetSelection(stop_index)
                if start_index == record_index:
                    record_index = (record_index + 1) % len(HOT_KEYS)
                    if record_index == stop_index:
                        record_index = (record_index + 1) % len(HOT_KEYS)
                    self.choice_record.SetSelection(record_index)
                start_name = HOT_KEYS[start_index].lower()
                stop_name = HOT_KEYS[stop_index].lower()
                record_name = HOT_KEYS[record_index].lower()

                if key_name == start_name and not self.running and not self.recording:
                    print('script start')
                    # t = RunScriptClass(self, self.pause_event)
                    # t.start()
                    self.runthread = RunScriptClass(self, self.pause_event)
                    self.runthread.start()
                    self.isbrokenorfinish = False
                    print(key_name, 'host start')
                elif key_name == start_name and self.running and not self.recording:
                    if self.paused:
                        print('script resume')
                        self.paused = False
                        self.pause_event.set()
                        print(key_name, 'host resume')
                    else:
                        print('script pause')
                        self.paused = True
                        self.pause_event.clear()
                        print(key_name, 'host pause')
                elif key_name == stop_name and self.running and not self.recording:
                    print('script stop')
                    self.tnumrd.SetLabel('broken')
                    self.isbrokenorfinish = True
                    if self.paused:
                        self.paused = False
                        self.pause_event.set()
                    print(key_name, 'host stop')
                elif key_name == stop_name and self.recording:
                    self.recordMethod()
                    print(key_name, 'host stop record')
                elif key_name == record_name:
                    if not self.recording:
                        self.recordMethod()
                        print(key_name, 'host start record')
                    else:
                        self.pauseRecordMethod()
                        print(key_name, 'host pause record')

            if not self.recording or self.running or self.pauserecord:
                return True

            all_messages = ('key down', 'key up')
            if message not in all_messages:
                return True

            # 不录制热键
            hot_keys = [HOT_KEYS[self.choice_start.GetSelection()],
                        HOT_KEYS[self.choice_stop.GetSelection()],
                        HOT_KEYS[self.choice_record.GetSelection()]]
            if event.Key in hot_keys:
                return True

            key_info = (event.KeyID, event.Key, event.Extended)

            delay = current_ts() - self.ttt
            self.ttt = current_ts()
            if not self.record:
                delay = 0

            print(delay, message, key_info)

            self.record.append([delay, 'EK', message, key_info])
            self.actioncount = self.actioncount + 1
            text = '%d actions recorded' % self.actioncount
            self.tnumrd.SetLabel(text)
            return True

        self.hm.MouseAll = on_mouse_event
        self.hm.KeyAll = on_keyboard_event
        self.hm.HookMouse()
        self.hm.HookKeyboard()

    def get_script_path(self):
        i = self.choice_script.GetSelection()
        if i < 0:
            return ''
        script = self.scripts[i]
        path = os.path.join(os.getcwd(), '../scripts', script)
        print(path)
        return path

    def new_script_path(self):
        now = datetime.datetime.now()
        script = '%s.txt' % now.strftime('%m%d_%H%M')
        if script in self.scripts:
            script = '%s.txt' % now.strftime('%m%d_%H%M%S')
        self.scripts.insert(0, script)
        self.choice_script.SetItems(self.scripts)
        self.choice_script.SetSelection(0)
        return self.get_script_path()

    def OnHide(self, event):
        self.Hide()
        event.Skip()

    def OnIconfiy(self, event):
        self.Hide()
        event.Skip()

    def OnClose(self, event):
        config.saveconfig({
            'starthotkeyindex':self.choice_start.GetSelection(),
            'stophotkeyindex':self.choice_stop.GetSelection(),
            'recordhotkeyindex': self.choice_record.GetSelection(),
            'looptimes':self.stimes.GetValue(),
            'precision':self.mouse_move_interval_ms.GetValue(),
            'executespeed':self.execute_speed.GetValue(),
            'language':RID_MAP[self.choice_language.GetSelection()]
            })
        self.taskBarIcon.Destroy()
        self.Destroy()
        event.Skip()

    def OnButton1Button(self, event):
        event.Skip()

    def pauseRecordMethod(self):
        if self.pauserecord:
            print('record resume')
            self.pauserecord = False
            self.btpauserecord.SetLabel(i18n.t('lang.pauserecord'))
        else:
            print('record pause')
            self.pauserecord = True
            self.btpauserecord.SetLabel(i18n.t('lang.continuerecord'))
            self.tnumrd.SetLabel('record paused')

    def OnPauseRecordButton(self, event):
        self.pauseRecordMethod()
        event.Skip()

    def recordMethod(self):
        if self.recording:
            print('record stop')
            self.recording = False
            self.record = self.record[:-2]
            output = json.dumps(self.record, indent=1)
            output = output.replace('\r\n', '\n').replace('\r', '\n')
            output = output.replace('\n   ', '').replace('\n  ', '')
            output = output.replace('\n ]', ']')
            open(self.new_script_path(), 'w').write(output)
            self.btrecord.SetLabel(i18n.t('lang.record'))
            self.tnumrd.SetLabel('finished')
            self.record = []
            self.btpauserecord.Enable(False)
            self.btrun.Enable(True)
            self.actioncount = 0
            self.pauserecord = False
            self.btpauserecord.SetLabel(i18n.t('lang.pauserecord'))
        else:
            print('record start')
            self.recording = True
            self.ttt = current_ts()
            status = self.tnumrd.GetLabel()
            if 'running' in status or 'recorded' in status:
                return
            self.btrecord.SetLabel(i18n.t('lang.finish'))  # 结束
            self.tnumrd.SetLabel('0 actions recorded')
            self.choice_script.SetSelection(-1)
            self.record = []
            self.btpauserecord.Enable(True)
            self.btrun.Enable(False)

    def OnBtrecordButton(self, event):
        self.recordMethod()
        event.Skip()

    def OnBtrunButton(self, event):
        print('script start by btn')
        # t = RunScriptClass(self, self.pause_event)
        # t.start()
        self.runthread = RunScriptClass(self, self.pause_event)
        self.runthread.start()
        self.isbrokenorfinish = False
        event.Skip()

    def OnBtpauseButton(self, event):
        print('script pause button pressed')
        if self.paused:
            print('script is resumed')
            self.pause_event.set()
            self.paused = False
            self.btpause.SetLabel(i18n.t('lang.pause'))
        else:
            print('script is paused')
            self.pause_event.clear()
            self.paused = True
            self.btpause.SetLabel(i18n.t('lang.continue'))
        event.Skip()

    def OnChoice_startChoice(self, event):
        event.Skip()

    def OnChoice_stopChoice(self, event):
        event.Skip()

    def OnChoice_recordChoice(self, event):
        event.Skip()


class RunScriptClass(threading.Thread):

    def __init__(self, frame: Frame1, event: threading.Event):
        self.frame = frame
        self.event = event
        self.event.set()
        super(RunScriptClass, self).__init__()

    def run(self):

        status = self.frame.tnumrd.GetLabel()
        if self.frame.running or self.frame.recording:
            return

        if 'running' in status or 'recorded' in status:
            return

        script_path = self.frame.get_script_path()
        if not script_path:
            self.frame.tnumrd.SetLabel('script not found, please self.record first!')
            return

        self.frame.running = True

        self.frame.btrun.Enable(False)
        self.frame.btrecord.Enable(False)

        try:
            self.run_times = self.frame.stimes.Value
            self.running_text = '%s running..' % script_path.split('/')[-1].split('\\')[-1]
            self.frame.tnumrd.SetLabel(self.running_text)
            self.frame.tstop.Shown = True
            self.run_speed = self.frame.execute_speed.Value

            self.j = 0
            while self.j < self.run_times or self.run_times == 0:
                self.j += 1
                current_status = self.frame.tnumrd.GetLabel()
                if  current_status in ['broken', 'finished']:
                    self.frame.running = False
                    break
                RunScriptClass.run_script_once(script_path, self.j, thd=self)

            self.frame.tnumrd.SetLabel('finished')
            self.frame.tstop.Shown = False
            self.frame.running = False
            PlayPromptTone.play_end_sound()
            print('script run finish!')

        except Exception as e:
            print('run error', e)
            traceback.print_exc()
            self.frame.tnumrd.SetLabel('failed')
            self.frame.tstop.Shown = False
            self.frame.running = False
        finally:
            self.frame.btrun.Enable(True)
            self.frame.btrecord.Enable(True)

    @classmethod
    def run_script_once(cls, script_path, step, thd=None):

        content = ''

        lines = []

        try:
            lines = open(script_path, 'r', encoding='utf8').readlines()
        except Exception as e:
            print(e)
            try:
                lines = open(script_path, 'r', encoding='gbk').readlines()
            except Exception as e:
                print(e)

        for line in lines:
            # 去注释
            if '//' in line:
                index = line.find('//')
                line = line[:index]
            # 去空字符
            line = line.strip()
            content += line

        # 去最后一个元素的逗号（如有）
        content = content.replace('],\n]', ']\n]').replace('],]', ']]')

        print(content)
        s = json.loads(content)
        steps = len(s)

        for i in range(steps):

            print(s[i])

            delay = s[i][0] / (thd.run_speed/100)
            event_type = s[i][1].upper()
            message = s[i][2].lower()
            action = s[i][3]

            if 1 == step and 0 == i:
                play = PlayPromptTone(1, delay)
                play.start()

            time.sleep(delay / 1000.0)

            if thd:
                # current_status = thd.frame.tnumrd.GetLabel()
                # if current_status in ['broken', 'finished']:
                #     break
                if thd.frame.isbrokenorfinish:
                    break
                thd.event.wait()
                text = '%s  [%d/%d %d/%d] %d%%' % (thd.running_text, i+1, steps, thd.j, thd.run_times, thd.run_speed)
                thd.frame.tnumrd.SetLabel(text)

            if event_type == 'EM':
                x, y = action
                # 兼容旧版的绝对坐标
                if not isinstance(x, int) and not isinstance(y, int):
                    x = float(re.match('([0-1].[0-9]+)%', x).group(1))
                    y = float(re.match('([0-1].[0-9]+)%', y).group(1))

                if action == [-1, -1]:
                    # 约定 [-1, -1] 表示鼠标保持原位置不动
                    pass
                else:
                    # 挪动鼠标 普通做法
                    # ctypes.windll.user32.SetCursorPos(x, y)
                    # or
                    # win32api.SetCursorPos([x, y])

                    # 更好的兼容 win10 屏幕缩放问题
                    if isinstance(x, int) and isinstance(y, int):
                        nx = int(x * 65535 / SW)
                        ny = int(y * 65535 / SH)
                    else:
                        nx = int(x * 65535)
                        ny = int(y * 65535)
                    win32api.mouse_event(win32con.MOUSEEVENTF_ABSOLUTE|win32con.MOUSEEVENTF_MOVE, nx, ny, 0, 0)

                if message == 'mouse left down':
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                elif message == 'mouse left up':
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                elif message == 'mouse right down':
                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                elif message == 'mouse right up':
                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                elif message == 'mouse middle down':
                    win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0)
                elif message == 'mouse middle up':
                    win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
                elif message == 'mouse wheel up':
                    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, win32con.WHEEL_DELTA, 0)
                elif message == 'mouse wheel down':
                    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, -win32con.WHEEL_DELTA, 0)
                elif message == 'mouse move':
                    pass
                else:
                    print('unknow mouse event:', message)

            elif event_type == 'EK':
                key_code, key_name, extended = action

                # shift ctrl alt
                # if key_code >= 160 and key_code <= 165:
                #     key_code = int(key_code/2) - 64

                # 不执行热键
                hot_keys = [HOT_KEYS[thd.frame.choice_start.GetSelection()], HOT_KEYS[thd.frame.choice_stop.GetSelection()]]
                if key_name in hot_keys:
                    continue

                base = 0
                if extended:
                    base = win32con.KEYEVENTF_EXTENDEDKEY

                if message == 'key down':
                    win32api.keybd_event(key_code, 0, base, 0)
                elif message == 'key up':
                    win32api.keybd_event(key_code, 0, base | win32con.KEYEVENTF_KEYUP, 0)
                else:
                    print('unknow keyboard event:', message)

            elif event_type == 'EX':

                if message == 'input':
                    text = action
                    pyperclip.copy(text)
                    # Ctrl+V
                    win32api.keybd_event(162, 0, 0, 0)  # ctrl
                    win32api.keybd_event(86, 0, 0, 0)  # v
                    win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(162, 0, win32con.KEYEVENTF_KEYUP, 0)
                else:
                    print('unknow extra event:', message)


class TaskBarIcon(wxTaskBarIcon):
    ID_About = wx.NewId()
    ID_Closeshow = wx.NewId()

    def __init__(self, frame):
        wxTaskBarIcon.__init__(self)
        self.frame = frame
        self.SetIcon(GetMondrianIcon())
        self.Bind(EVT_TASKBAR_LEFT_DCLICK, self.OnTaskBarLeftDClick)
        self.Bind(wx.EVT_MENU, self.OnAbout, id=self.ID_About)
        self.Bind(wx.EVT_MENU, self.OnCloseshow, id=self.ID_Closeshow)

    def OnTaskBarLeftDClick(self, event):
        if self.frame.IsIconized():
            self.frame.Iconize(False)
        if not self.frame.IsShown():
            self.frame.Show(True)
        self.frame.Raise()

    def OnAbout(self, event):
        wx.MessageBox('https://github.com/rushionn/KeymouseGo', 'KeymouseGo v%s' % VERSION)
        event.Skip()

    def OnCloseshow(self, event):
        self.frame.Close(True)
        event.Skip()

    def CreatePopupMenu(self):
        menu = wx.Menu()
        menu.Append(self.ID_About, 'About')
        menu.Append(self.ID_Closeshow, 'Exit')
        return menu


class PlayPromptTone(threading.Thread):

    def __init__(self, op, delay):
        self._delay = delay
        self._op = op
        super().__init__()

    def run(self):
        if 1 == self._op:
            if self._delay >= 1000:
                time.sleep((self._delay - 500.0) / 1000.0)
            self._play_start_sound()

    def _play_start_sound(self):
        try:
            path = os.path.join(os.getcwd(), '../sounds', 'start.mp3')
            playsound(path)
        except PlaysoundException as e:
            print(e)

    @classmethod
    def play_end_sound(cls):
        try:
            path = os.path.join(os.getcwd(), '../sounds', 'end.mp3')
            playsound(path)
        except PlaysoundException as e:
            print(e)
