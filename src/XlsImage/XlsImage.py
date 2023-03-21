from .MSODrawingGroup import OfficeArtDggContainer
import os, struct

def XlsGetImages(path:str) -> list:
    """
    获得 .xls 文件的图片数据.

    Args:
        path (str): 输入文件目录

    Return:
        list[tuple]: 返回以 tuple(图片后缀, bytes) 组成的列表
    """
    ret = []
    if not os.path.exists(path):
        return ret
    with open(path, "rb") as f:
        f.seek(0, os.SEEK_END)
        size = f.tell()
        f.seek(0, os.SEEK_SET)
        xlsBin = f.read(size)
    # 文件格式标识
    identifier = xlsBin[:8]
    if b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" != identifier:
        # 文件格式标识不正确
        return ret
    # 段大小
    secSize = 2 ** struct.unpack("<H", xlsBin[30: 32])[0]
    # 短段大小
    shortSecSize = 2 ** struct.unpack("<H", xlsBin[32:34])[0]
    # SAT 用的段数量
    secAllocNumUsed = struct.unpack("<I", xlsBin[44:48])[0]
    # 目录段ID
    directorySecID = struct.unpack("<i", xlsBin[48:52])[0]
    # 标准流最小大小
    minimumSizeOfStandardStream = struct.unpack("<I", xlsBin[56: 60])[0]
    # SSAT 段ID
    shortAllocTableSecID = struct.unpack("<i", xlsBin[60: 64])[0]
    # SSAT 占用的段数量
    shortSecNumUsed = struct.unpack("<I", xlsBin[64: 68])[0]
    # MSAT 段ID
    masterAllocTableSecID = struct.unpack("<i", xlsBin[68:72])[0]
    # MSAT 占用的段数量
    masterSecNumUsed = struct.unpack("<I", xlsBin[72: 76])[0]

    def get_sec_off(sec_id):
        """获得段偏移"""
        return 512 + sec_id * secSize
    def get_short_sec_off(sec_id):
        """获得短段偏移"""
        return short_id * shortSecSize
    def to_alloc_table(sec_id):
        """转成分配表"""
        t = xlsBin[get_sec_off(sec_id):get_sec_off(sec_id) + secSize]
        return [struct.unpack("<i", t[i : i + 4])[0]
                for i in range(0, len(t), 4)]
    def get_data(sec_id, table):
        """从分配表获取数据"""
        chain = [sec_id]
        i = sec_id
        ret = bytearray()
        while (i := table[i]) >= 0:
            chain.append(i)
        for i in chain:
            off = get_sec_off(i)
            ret.extend(xlsBin[off:off + secSize])
        return ret
    def get_sec_data(sec_id):
        """获得段数据"""
        off = get_sec_off(sec_id)
        return xlsBin[off:off + secSize]
    # MSAT
    masterAllocTable = bytearray()
    masterAllocTable.extend(xlsBin[76: 76 + 436])
    masterAllocTable = [struct.unpack("<i", masterAllocTable[i: i + 4])[0]
                        for i in range(0, len(masterAllocTable), 4)]
    # SAT, 接下来拼接 SAT表
    secAllocTable = []
    for i in masterAllocTable:
        # MSAT 存 SAT, 且是顺序存
        if i >= 0:
            secAllocTable.extend(to_alloc_table(i))
            continue
        break
    # 完善 MSAT
    while True:
        if masterAllocTableSecID >= 0:
            msatNext = to_alloc_table(masterAllocTableSecID)
            masterAllocTableSecID = msatNext.pop(len(msatNext) - 1)
            # 完善 SAT
            for i in msatNext:
                if i >= 0:
                    secAllocTable.extend(to_alloc_table(i))
                    continue
                break
            masterAllocTable.extend(msatNext)
            del msatNext
            continue
        break
    # 获得目录
    directory = get_data(directorySecID, secAllocTable)
    # 获得目录项
    directoryEntries = [struct.unpack("<64sHBBiii16siQQiIi",
                                      directory[i: i + 128])
                        for i in range(0, len(directory), 128)]
    del directory
    directoryFields = ["name", "name_size", "type", "node_color",
                       "left_child_node_id", "right_child_node_id",
                       "root_node_entry_id", "unique_identifier",
                       "user_flag", "timestamp_creation",
                       "timestamp_modification", "short_sec_or_sec_id",
                       "total_stream_size", "not_used"]
    directoryEntries = [dict(zip(directoryFields, i))
                        for i in directoryEntries]
    # 寻找 Workbook
    for i in directoryEntries:
        i["name"]:str = i["name"].decode(encoding="utf-16-le") \
            .rstrip(b'\x00'.decode())
        if i["name"] == "Workbook":
            # 获取用户流
            stream = get_data(i["short_sec_or_sec_id"], secAllocTable)
    # 不再需要 xlsBin
    del xlsBin
    # 获得内存视图
    streammv = memoryview(stream)
    # 所有 BIFF 对象
    BIFFs = []
    curPos = 0
    # 获得流的所有 BIFF 对象
    while True:
        if curPos + 4 < streammv.nbytes:
            biffSize = struct.unpack("<H", streammv[curPos + 2:curPos + 4])[0]
            biffSize = biffSize + 4
            BIFFs.append((curPos, biffSize, streammv[curPos:curPos + 2],
                          streammv[curPos + 4:curPos + biffSize]))
            curPos += biffSize
            continue
        break
    # 获得 Microsoft Office 绘图组 (MSODrawingGroup)
    curPos = 0
    drawingGroup = []
    tBIFF = bytearray()
    for i in BIFFs:
        if curPos == 0:
            # 判断绘图组标识和继续标识
            if i[2] == b'\xeb\x00' or i[2] == b'\x3c\x00':
                curPos = i[0] + i[1]
                tBIFF.extend(i[3])
                continue
        else:
            if curPos != i[0]:
                drawingGroup.append(tBIFF)
                tBIFF.clear()
                curPos = 0
                continue
            tBIFF.extend(i[3])
            curPos += i[1]
    if len(tBIFF) > 0:
        drawingGroup.append(tBIFF)
    del tBIFF
    del curPos
    # 判断是否存在绘图组
    if len(drawingGroup) == 0:
        return ret
    # 获得美术容器
    try:
        artContianer = OfficeArtDggContainer(drawingGroup[0])
        for blip in artContianer.blipStore.rgfb:
            ret.append((blip._blip.suffix, blip._blip.BLIPFileData))
    except:
        pass
    return ret
