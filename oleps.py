## -*- coding: UTF-8 -*-
## oleps.py
##
## Copyright (c) 2018 analyzeDFIR
## 
## Permission is hereby granted, free of charge, to any person obtaining a copy
## of this software and associated documentation files (the "Software"), to deal
## in the Software without restriction, including without limitation the rights
## to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
## copies of the Software, and to permit persons to whom the Software is
## furnished to do so, subject to the following conditions:
## 
## The above copyright notice and this permission notice shall be included in all
## copies or substantial portions of the Software.
## 
## THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
## IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
## FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
## AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
## LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
## OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
## SOFTWARE.

import math
from datetime import datetime, timedelta

try:
    from lib.parsers import ByteParser
    from lib.parsers.utils import StructureProperty, WindowsTime
    from structures import oleps as olepsstructs
except ImportError:
    from .lib.parsers import ByteParser, contexted
    from .lib.parsers.utils import StructureProperty, WindowsTime
    from .structures import oleps as olepsstructs

class OLETypedPropertyValue(ByteParser):
    '''
    Class for TypedPropertyValue structures from a Windows 
    Object Embedding and Linking (OLE) Property Set
    '''
    header = StructureProperty(0 ,'header')
    value = StructureProperty(1, 'value', deps=['header'])

    def _parse_clsid(self):
        '''
        Args:
            N/A
        Returns:
            Container<String, Any>
            NTFSGUID (see structures.shared_structures.windows.guid.NTFSGUID)
        Preconditions:
            N/A
        '''
        return olepsstructs.NTFSGUID.parse_stream(self.stream)
    def _parse_cf(self):
        '''
        Args:
            N/A
        Returns:
            Integer 
            OLEPropertyIdentifier (see structures.OLEPropertyIdentifier)
        Preconditions:
            N/A
        '''
        return olepsstructs.OLEPropertyIdentifier.parse_stream(self.stream)
    def _parse_vt_blob_object(self):
        '''
        @OLETypedPropertyValue._parse_vt_blob
        '''
        return self._parse_vt_blob()
    def _parse_vt_blob(self):
        '''
        Args:
            N/A
        Returns:
            Container<String, Any>
            OLEBlob (see structures.OLEBlob)
        Preconditions:
            N/A
        '''
        return olepsstructs.OLEBlob.parse_stream(self.stream)
    def _parse_vt_filetime(self):
        '''
        Args:
            N/A
        Returns:
            Container<String, Any>
            NTFSFILETIME (see structures.shared_structures.windows.misc.NTFSFILETIME)
        Preconditions:
            N/A
        '''
        return WindowsTime.parse_filetime(
            olepsstructs.NTFSFILETIME.parse_stream(self.stream)
        )
    def _parse_vt_lpwstr(self):
        '''
        Args:
            N/A
        Returns:
            Container<String, Any>
            OLEUnicodeString
        Preconditions:
            N/A
        '''
        return olepsstructs.OLEUnicodeString.parse_stream(self.stream)
    def _parse_vt_lpstr(self):
        '''
        @OLETypedPropertyValue._parse_vt_bstr
        '''
        return self._parse_vt_bstr()
    def _parse_vt_uint(self):
        '''
        @OLETypedPropertyValue._parse_vt_ui2
        '''
        return self._parse_vt_ui2()
    def _parse_vt_int(self):
        '''
        @OLETypedPropertyValue._parse_vt_i2
        '''
        return self._parse_vt_i2()
    def _parse_vt_ui8(self):
        '''
        Args:
            N/A
        Returns:
            Integer
        Preconditions:
            N/A
        '''
        return olepsstructs.Int64ul.parse_stream(self.stream)
    def _parse_vt_i8(self):
        '''
        Args:
            N/A
        Returns:
            Integer
        Preconditions:
            N/A
        '''
        return olepsstructs.Int64sl.parse_stream(self.stream)
    def _parse_vt_ui4(self):
        '''
        Args:
            N/A
        Returns:
            Integer
        Preconditions:
            N/A
        '''
        return olepsstructs.Int32ul.parse_stream(self.stream)
    def _parse_vt_ui2(self):
        '''
        Args:
            N/A
        Returns:
            Integer
        Preconditions:
            N/A
        '''
        return olepsstructs.Int16ul.parse_stream(self.stream)
    def _parse_vt_ui1(self):
        '''
        Args:
            N/A
        Returns:
            Integer
        Preconditions:
            N/A
        '''
        return olepsstructs.Int8ul.parse_stream(self.stream)
    def _parse_vt_i1(self):
        '''
        Args:
            N/A
        Returns:
            Integer
        Preconditions:
            N/A
        '''
        return olepsstructs.Int8sl.parse_stream(self.stream)
    def _parse_vt_decimal(self):
        '''
        Args:
            N/A
        Returns:
            Container<String, Any>
            OLEDecimal (structures.OLEDecimal)
        Preconditions:
            N/A
        '''
        return olepsstructs.OLEDecimal.parse_stream(self.stream)
    def _parse_vt_bool(self):
        '''
        Args:
            N/A
        Returns:
            Boolean
            OLEVariantBool (see structures.OLEVariantBool)
        Preconditions:
            N/A
        '''
        return olepsstructs.OLEVariantBool.parse_stream(self.stream)
    def _parse_vt_error(self):
        '''
        Args:
            N/A
        Returns:
            Container<String, Any>
            OLEHResult
            (see https://msdn.microsoft.com/en-us/library/cc231196.aspx
            and structures.OLEHResult)
        Preconditions:
            N/A
        '''
        return olepsstructs.OLEHResult.parse_stream(self.stream)
    def _parse_vt_bstr(self):
        '''
        Args:
            N/A
        Returns:
            Container<String, Any>
            CodePageString (see structures.OLECodePageString)
        Preconditions:
            N/A
        '''
        return olepsstructs.OLECodePageString.parse_stream(self.stream)
    def _parse_vt_date(self):
        '''
        Args:
            N/A
        Returns:
            DateTime
        Preconditions:
            N/A
        '''
        raw_date = olepsstructs.OLEDate.parse_stream(self.stream)
        day_part, time_part = math.modf(raw_date)
        if day_delta < 0 or time_delta < 0:
            raise Exception('Found negative date or time part in date %d'%raw_date)
        return datetime(1899, 12, 30) + \
            timedelta(day_part, ( time_part * ( 60*60*24 ) ))
    def _parse_vt_cy(self):
        '''
        Args:
            N/A
        Returns:
            Integer
        Preconditions:
            N/A
        '''
        return olepsstructs.OLECurrency.parse_stream(self.stream)
    def _parse_vt_r8(self):
        '''
        Args:
            N/A
        Returns:
            Integer
        Preconditions:
            N/A
        '''
        return olepsstructs.Float64l.parse_stream(self.stream)
    def _parse_vt_r4(self):
        '''
        Args:
            N/A
        Returns:
            Integer
        Preconditions:
            N/A
        '''
        return olepsstructs.Float32l.parse_stream(self.stream)
    def _parse_vt_i4(self):
        '''
        Args:
            N/A
        Returns:
            Integer
        Preconditions:
            N/A
        '''
        return olepsstructs.Int32sl.parse_stream(self.stream)
    def _parse_vt_i2(self):
        '''
        Args:
            N/A
        Returns:
            Integer
        Preconditions:
            N/A
        '''
        return olepsstructs.Int16sl.parse_stream(self.stream)
    def _parse_vt_null(self):
        '''
        @OLETypedPropertyValue._parse_vt_empty
        '''
        return self._parse_vt_empty()
    def _parse_vt_empty(self):
        '''
        Args:
            N/A
        Returns:
            None
        Preconditions:
            N/A
        '''
        return None
    def _parse_value(self):
        '''
        Args:
            N/A
        Returns:
            Any
            Value associated with this TypedPropertyValue
        Preconditions:
            N/A
        '''
        parser = '_parse_%s'%self.header.Type.lower()
        if not (hasattr(self, parser) and callable(getattr(self, parser))):
            return None
        return getattr(self, parser)()
    def _parse_header(self):
        '''
        Args:
            N/A
        Returns:
            Container<String, Any>
            TypedPropertyValue header (see structures.OLETypedPropertyValueHeader)
        Preconditions:
            N/A
        '''
        return olepsstructs.OLETypedPropertyValueHeader.parse_stream(self.stream)
