#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlwt
import wx
from wx.lib.pubsub import Publisher as pub
import os
import math

def pluginProperties():
    props = {}
    props['name'] = 'Export DVHs to Excel'
    props['description'] = 'Export DVHs to Excel'
    props['author'] = 'Sebastien Jodogne'
    props['version'] = 0.2
    props['plugin_type'] = 'menu'
    props['plugin_version'] = 1
    props['min_dicom'] = ['rtss', 'rtplan', 'rtdose', 'images']
    props['recommended_dicom'] = ['rtss', 'rtplan', 'rtdose', 'images']
    return props

class plugin:
    def __init__(self, parent):
        self.parent = parent
        pub.subscribe(self.OnUpdatePatient, 'patient.updated.parsed_data')

    def OnUpdatePatient(self, msg):
        if (msg.data.has_key('dvhs') and
            msg.data.has_key('structures')):
            self.structures = msg.data['structures']
            self.dvhs = msg.data['dvhs']

            self.maxBins = 0
            for i in self.dvhs:
                bins = len(self.dvhs[i]['data'])
                self.maxBins = max(self.maxBins, bins)

    def pluginMenu(self, evt):
        fileDialog = wx.FileDialog(self.parent, 
                                   style = wx.SAVE | wx.OVERWRITE_PROMPT,
                                   wildcard = 'Excel file (*.xls)|*.xls',
                                   message = 'Choose the target file')

        if fileDialog.ShowModal() == wx.ID_OK:
            path = fileDialog.GetPath()

            w = xlwt.Workbook()
            wr = w.add_sheet('Relative volumes (%)')
            wr.write(0, 0, 'Dose (cGy)')
            wa = w.add_sheet('Absolute volumes (cc)')
            wa.write(0, 0, 'Dose (cGy)')

            for y in range(self.maxBins):
                wr.write(y + 1, 0, y)
                wa.write(y + 1, 0, y)

            x = 1
            for i in self.structures:
                name = self.structures[i]['name']
                structureId = self.structures[i]['id']

                if not self.dvhs.has_key(structureId):
                    continue

                dvh = self.dvhs[structureId]
                if len(dvh['data']) == 0:
                    continue

                volume = dvh['data'][0]
                
                for y in range(self.maxBins):
                    if y < len(dvh['data']):
                        absoluteValue = dvh['data'][y]
                    else:
                        absoluteValue = 0

                    if math.isnan(absoluteValue) or math.isinf(absoluteValue):
                        absoluteValue = 0

                    relativeValue = absoluteValue / volume * 100.0

                    if math.isnan(relativeValue) or math.isinf(relativeValue):
                        relativeValue = 0

                    wr.write(y + 1, x, relativeValue)
                    wa.write(y + 1, x, absoluteValue)

                wr.write(0, x, name)
                wa.write(0, x, name)
                x += 1
            
            w.save(path)

        fileDialog.Destroy()
