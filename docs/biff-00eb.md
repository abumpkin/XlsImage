<div align="right"> 2022年9月15日 </div>

---

# MSODRAWINGGROUP: Microsoft Office Drawing Group (EBh)

* MSODRAWINGGROUP
  * rgMSODrawiGr: `OfficeArtDggContainer` 类型

## OfficeArtDggContainer 类型

* OfficeArtDggContainer
  * rh (8 bytes): `OfficeArtRecordHeader` 类型
    * rh.recVer: A value that MUST be 0xF.
    * rh.recInstance: A value that MUST be 0x000.
    * rh.recType: A value that MUST be 0xF000.
    * rh.recLen (4 bytes): 接下来的长度
  * drawingGroup: `OfficeArtFDGGBlock` 类型
  * blipStore: `OfficeArtBStoreContainer` 类型
  * drawingPrimaryOptions: `OfficeArtFOPT` 类型
  * drawingTertiaryOptions: `OfficeArtTertiaryFOPT` 类型
  * colorMRU: `OfficeArtColorMRUContainer` 类型
  * splitColors: `OfficeArtSplitMenuColorContainer` 类型

## OfficeArtFDGGBlock 类型

* OfficeArtFDGGBlock
  * rh (8 bytes): `OfficeArtRecordHeader` 类型
    * rh.recVer: A value that MUST be 0x0.
    * rh.recInstance: A value that MUST be 0x000.
    * rh.recType: A value that MUST be 0xF006.
    * rh.recLen (4 bytes): A value that MUST be 0x00000010 + ((head.cidcl - 1) * 0x00000008)
  * head (16 bytes): `OfficeArtFDGG` 类型
    * spidMax (4 bytes): specifying the current maximum shape identifier that is used in any drawing. This value MUST be less than 0x03FFD7FF.
    * cidcl (4 bytes): An unsigned integer that specifies the number of OfficeArtIDCL records
    * cspSaved (4 bytes): An unsigned integer specifying the total number of shapes that have been saved in all of the drawings.
    * cdgSaved (4 bytes): An unsigned integer specifying the total number of drawings that have been saved in the file.
  * Rgidcl: An array of `OfficeArtIDCL` elements, The number of elements in the array is specified by (head.cidcl – 1).
    * dgid (4 bytes): specifying the drawing identifier that owns this identifier cluster.
    * cspidCur (4 bytes):  An unsigned integer that, if less than 0x00000400, specifies the largest shape identifier that is currently assigned in this cluster, or that otherwise specifies that no shapes can be added to the drawing.

## OfficeArtBStoreContainer 类型

* OfficeArtBStoreContainer
  * rh (8 bytes): `OfficeArtRecordHeader` 类型
    * rh.recVer (4 bits): A value that MUST be 0x*F.
    * rh.recInstance (12 bits): An unsigned integer that specifies the number
of contained `OfficeArtBStoreContainerFileBlock` records. 0x___*
    * rh.recType: A value that MUST be 0xF001.
    * rh.recLen (4 bytes): 接下来几个 `OfficeArtBStoreContainerFileBlock` 的长度(即 **rgfb** 的长度).
  * rgfb: ` OfficeArtBStoreContainerFileBlock ` 类型数组

## OfficeArtBStoreContainerFileBlock 类型

* OfficeArtBStoreContainerFileBlock
  * `OfficeArtFBSE` 类型 | `OfficeArtBlip` 类型

先把它当作 `OfficeArtRecordHeader` 类型读取其 **recType**:

* recType = 0xF007, 则为 `OfficeArtFBSE` 类型
* recType = 0xF018-0xF117, 则为 `OfficeArtBlip` 类型

* `OfficeArtFBSE` 类型
  * rh (8 bytes): `OfficeArtRecordHeader` 类型
    * rh.recVer (4 bits): A value that MUST be 0x*2.
    * rh.recInstance (12 bits): An `MSOBLIPTYPE` enumeration value. 0x___*
    * rh.recType: A value that MUST be 0xF007.
    * rh.recLen (4 bytes): 无符号整型, 接下来的长度. 如果 BLIP 内嵌, 则此值为 (sizeof(**nameData**) + 36), 否则为 (sizeof(**nameData**) + **size** + 36)
  * btWin32 (1 bytes): `MSOBLIPTYPE` 枚举类型, 如果 **btMacOS** 值被 windows 支持, 则应与 **btMacOS** 一样. 如果与 **btMacOS** 不同, 则必须存在与 **rh.recInstance** 匹配的BLIP, 另一个可能存在.
  * btMacOS (1 bytes): ...
  * rgbUid (16 bytes): specifies the unique identifier of the pixel data in the BLIP.
  * tag (2 bytes): This value MUST be 0xFF for external files.
  * size (4 bytes):  An unsigned integer that specifies the size, in bytes, of the BLIP in the stream.
  * cRef(4 bytes): 无符号整型, 指定这个 BLIP 被引用的计数
  * foDelay (4 bytes): : An MSOFO structure, as defined in section 2.1.4, that specifies the file offset into the associated OfficeArtBStoreDelay record, as defined in section 2.2.21, (delay stream). A value of 0xFFFFFFFF specifies that the file is not in the delay stream, and in this case, cRef MUST be 0x00000000.
  * unused1 (1 bytes)
  * cbName (1 bytes): 无符号整型, 指定 **nameData** 长度, 包括结尾 null, 必须是偶数. 如果为 0 则没有写入 **nameData** (从 **nameData** 开始进入 `OfficeArtBlip` 类型)
  * unused2 (1 bytes)
  * unused3 (1 bytes)
  * [nameData: Unicode null-terminated 字符串]
  * embeddedBlip: `OfficeArtBlip` 类型

* `OfficeArtBlip` 类型, 先当作 `OfficeArtRecordHeader` 类型读取其 **recType**, 判断其值获得对应的类型:
  * 0xF01A `OfficeArtBlipEMF`, as defined in section 2.2.24.
  * 0xF01B `OfficeArtBlipWMF`, as defined in section 2.2.25.
  * 0xF01C `OfficeArtBlipPICT`, as defined in section 2.2.26.
  * 0xF01D `OfficeArtBlipJPEG`, as defined in section 2.2.27.
  * 0xF01E `OfficeArtBlipPNG`, as defined in section 2.2.28.
  * 0xF01F `OfficeArtBlipDIB`, as defined in section 2.2.29.
  * 0xF029 `OfficeArtBlipTIFF`, as defined in section 2.2.30.
  * 0xF02A `OfficeArtBlipJPEG`, 0xF02A is treated as 0xF01D.

* `OfficeArtBlipJPEG` 类型
  * rh (8 bytes): `OfficeArtRecordHeader` 类型
    * rh.recVer: A value that MUST be 0x*0.
    * rh.recInstance: 如下:

    | Value of recInstance | 意义 | uid 数量 |
    | --- | --- | --- |
    | 0x46A | JPEG in RGB color space  | 1 |
    | 0x46B | JPEG in RGB color space  | 2 |
    | 0x6E2 | JPEG in CMYK color space | 1 |
    | 0x6E3 | JPEG in CMYK color space | 2 |

    * rh.recType: A value that MUST be 0xF01D.
    * rh.recLen (4 bytes): 无符号整型, 接下来的长度. 如果 **recInstance** 是 0x46a 或 0x6e2, 则值为 (sizeof(**BLIPFileData**) + 17), 如果是 0x46b 或 0x6e3, 则值为 (sizeof(**BLIPFileData**) + 33)
  * rgbUid1 (16 bytes)
  * [rgbUid2 (16 bytes)]
  * tag (1 byte): An unsigned integer that specifies an application-defined internal resource tag. This value MUST be 0xFF for external files.
  * BLIPFileData: 数据

* `OfficeArtBlipPNG` 类型
  * rh (8 bytes): `OfficeArtRecordHeader` 类型
    * rh.recVer: A value that MUST be 0x*0.
    * rh.recInstance: A value that MUST be 0x6e_*.
    * rh.recType: A value that MUST be 0xF01E.
    * rh.recLen (4 bytes): 无符号整型, 接下来的长度. 如果 **recInstance** 是 0x6e0, 则值为 (sizeof(**BLIPFileData**) + 17), 如果是 0x6e1, 则值为 (sizeof(**BLIPFileData**) + 33)
  * rgbUid1 (16 bytes)
  * [rgbUid2 (16 bytes)]
  * tag (1 byte): An unsigned integer that specifies an application-defined internal resource tag. This value MUST be 0xFF for external files.
  * BLIPFileData: 数据

# 枚举类型

## MSOBLIPTYPE

* msoblipERROR 0x00 Error reading the file.
* msoblipUNKNOWN 0x01 Unknown BLIP type.
* msoblipEMF 0x02 EMF.
* msoblipWMF 0x03 WMF.
* msoblipPICT 0x04 Macintosh PICT.
* msoblipJPEG 0x05 JPEG.
* msoblipPNG 0x06 PNG.
* msoblipDIB 0x07 DIB
* msoblipTIFF 0x11 TIFF
* msoblipCMYKJPEG 0x12 JPEG in the YCCK or CMYK color space.
