# -*- coding: utf-8 -*-
"""
Created on Thu May 17 14:35:51 2018

@author: zachary.shaver
"""
import win32com.client as w32
import time

outlook = w32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

item = outlook.CreateItem ( 1 )
recip = item.Recipients.Add ( 'Hennigan, Kara' )
recip.Resolve ()
Folder = namespace.GetSharedDefaultFolder ( recip, 9 )

items = Folder.Items




    
    
    