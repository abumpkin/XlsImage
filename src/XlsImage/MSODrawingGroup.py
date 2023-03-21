"""
创建: 2022年9月15日
"""
from struct import unpack


class OfficeArtRecordHeader:
    def __init__(self, data):
        tmv = memoryview(data)
        tVal = unpack("<H", tmv[0:2])[0]
        self.recVer = tVal & 0xf
        self.recInstance = tVal >> 4
        tVal = unpack("<H", tmv[2:4])[0]
        self.recType = tVal
        self.recLen = unpack("<I", tmv[4:8])[0]

    def GetTotalSize(self):
        return self.recLen + 8


class OfficeArtFDGG:
    def __init__(self, data):
        tmv = memoryview(data)
        self.spidMax = unpack("<I", tmv[0:4])[0]
        self.cidcl = unpack("<I", tmv[4:8])[0]
        self.cspSaved = unpack("<I", tmv[8:12])[0]
        self.cdgSaved = unpack("<I", tmv[12:16])[0]


class OfficeArtIDCL:
    def __init__(self, data):
        tmv = memoryview(data)
        self.dgid = unpack("<I", tmv[0:4])[0]
        self.cspidCur = unpack("<I", tmv[4:8])[0]


class OfficeArtFDGGBlock:
    def __init__(self, data):
        tmv = memoryview(data)
        self.rh = OfficeArtRecordHeader(tmv[0:8])
        self.head = OfficeArtFDGG(tmv[8:24])
        self.Rgidcl = [OfficeArtIDCL(tmv[24 + 8 * i: 24 + 8 * i + 8])
                       for i in range(self.head.cidcl - 1)]

class OfficeArtBlip:
    OfficeArtBlipEMFType = 0xF01A # as defined in section 2.2.24.
    OfficeArtBlipWMFType = 0xF01B # as defined in section 2.2.25.
    OfficeArtBlipPICTType = 0xF01C # as defined in section 2.2.26.
    OfficeArtBlipJPEGType = 0xF01D # as defined in section 2.2.27.
    OfficeArtBlipPNGType = 0xF01E # as defined in section 2.2.28.
    OfficeArtBlipDIBType = 0xF01F # as defined in section 2.2.29.
    OfficeArtBlipTIFFType = 0xF029 # as defined in section 2.2.30.
    OfficeArtBlipJPEGType2 = 0xF02A # 0xF02A is treated as 0xF01D.

    class MSOBLIPTYPE:
        msoblipERROR = 0x00 # Error reading the file.
        msoblipUNKNOWN = 0x01 # Unknown BLIP type.
        msoblipEMF = 0x02 # EMF.
        msoblipWMF = 0x03 # WMF.
        msoblipPICT = 0x04 # Macintosh PICT.
        msoblipJPEG = 0x05 # JPEG.
        msoblipPNG = 0x06 # PNG.
        msoblipDIB = 0x07 # DIB
        msoblipTIFF = 0x11 # TIFF
        msoblipCMYKJPEG = 0x12 # JPEG in the YCCK or CMYK color space.

    class OfficeArtBlipJPEG:
        suffix = "jpeg"
        def __init__(self, rh:OfficeArtRecordHeader, data):
            tmv = memoryview(data)
            self.rgbUid1 = tmv[0:16]
            tpos = 16
            if rh.recInstance == 0x46b or \
                rh.recInstance == 0x6e3:
                self.rgbUid2 = tmv[tpos:tpos + 16]
                tpos += 16
            self.tag = unpack("B", tmv[tpos:tpos + 1])
            tpos += 1
            self.BLIPFileDataSize = rh.recLen - tpos
            self.BLIPFileData = tmv[tpos:tpos + self.BLIPFileDataSize]

    class OfficeArtBlipPNG:
        suffix = "png"
        def __init__(self, rh:OfficeArtRecordHeader, data):
            tmv = memoryview(data)
            self.rgbUid1 = tmv[0:16]
            tpos = 16
            if rh.recInstance == 0x6e1:
                self.rgbUid2 = tmv[tpos:tpos + 16]
                tpos += 16
            self.tag = unpack("B", tmv[tpos:tpos + 1])
            tpos += 1
            self.BLIPFileDataSize = rh.recLen - tpos
            self.BLIPFileData = tmv[tpos:tpos + self.BLIPFileDataSize]

    def __init__(self, data):
        tmv = memoryview(data)
        self.rh = OfficeArtRecordHeader(tmv[0:8])
        if self.rh.recType == self.OfficeArtBlipJPEGType or \
            self.rh.recType == self.OfficeArtBlipJPEGType2:
            self._blip = self.OfficeArtBlipJPEG(self.rh,
                                                tmv[8:self.rh.GetTotalSize()])
        elif self.rh.recType == self.OfficeArtBlipPNGType:
            self._blip = self.OfficeArtBlipPNG(self.rh,
                                               tmv[8:self.rh.GetTotalSize()])


class OfficeArtFBSE:
    def __init__(self, data):
        tmv = memoryview(data)
        self.rh = OfficeArtRecordHeader(tmv[0:8])
        self.btWin32 = unpack("<B", tmv[8:9])[0]
        self.btMacOS = unpack("<B", tmv[9:10])[0]
        self.rgbUid = tmv[10:26]
        self.tag = tmv[26:28]
        self.size = unpack("<I", tmv[28:32])[0]
        self.cRef = unpack("<I", tmv[32:36])[0]
        self.foDelay = tmv[36:40]
        self.cbName = unpack("<B", tmv[41:42])[0]
        self.nameData = None
        tpos = 44
        if self.cbName > 0:
            self.nameData = unpack("{}s".format(self.cbName),
                                   tmv[tpos:tpos + self.cbName])[0]
            self.nameData:str = self.nameData.decode("utf-16-le")
            self.nameData.rstrip(b"\x00".decode())
            tpos += self.cbName
        self.embedded = OfficeArtBlip(tmv[tpos:tpos + self.rh.recLen])
        self._blip = self.embedded._blip

def OfficeArtBStoreContainerFileBlock(data):
    tmv = memoryview(data)
    theader = OfficeArtRecordHeader(tmv[0:8])
    if theader.recType == 0xF007:
        return OfficeArtFBSE(tmv[0:theader.GetTotalSize()])
    return OfficeArtBlip(tmv[0:theader.GetTotalSize()])

class OfficeArtBStoreContainer:
    def __init__(self, data):
        tmv = memoryview(data)
        self.rh = OfficeArtRecordHeader(tmv[0:8])
        self.rgfb = []
        tpos = 8
        for i in range(self.rh.recInstance):
            theader = OfficeArtRecordHeader(tmv[tpos:tpos + 8])
            tstore = OfficeArtBStoreContainerFileBlock(
                tmv[tpos:tpos + theader.GetTotalSize()])
            tpos += theader.GetTotalSize()
            self.rgfb.append(tstore)


###########################################################################

class OfficeArtDggContainer:
    def __init__(self, data: bytes):
        self.dataMv = memoryview(data)
        tmv = self.dataMv
        tpos = 0
        self.rh = OfficeArtRecordHeader(tmv[tpos:tpos + 8])
        tpos += 8
        theader = OfficeArtRecordHeader(tmv[tpos:tpos + 8])
        self.drawingGroup = OfficeArtFDGGBlock(
            tmv[tpos:tpos + theader.GetTotalSize()])
        tpos += theader.GetTotalSize()
        theader = OfficeArtRecordHeader(tmv[tpos:tpos + 8])
        self.blipStore = OfficeArtBStoreContainer(
            tmv[tpos:tpos + theader.GetTotalSize()])
        # self.drawingPrimaryOptions
        # self.drawingTertiaryOptions
        # self.colorMRU
        # self.splitColors

    def __del__(self):
        if type(self.dataMv) == memoryview:
            self.dataMv.release()
