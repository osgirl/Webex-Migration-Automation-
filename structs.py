# -*- coding: utf-8 -*-
"""
Created on Thu May 17 11:38:29 2018
learning about structs 
@author: zachary.shaver
"""

import struct

packed_data = struct.pack('iif', 6, 19, 21.123)

print(struct.calcsize('i'))
print(struct.calcsize('iif'))

print(struct.unpack('iif', packed_data))