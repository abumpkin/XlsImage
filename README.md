# 简介

XlsImage 是一个从 xls 表格文件里读取图片的 Python 库.

## 安装

```sh
pip install xlsimage
```

## 使用

在代码中使用:

```py
import XlsImage as xls

imgs = xls.XlsGetImages("example.xls")
print("total images: ", len(imgs))
for k, v in enumerate(imgs):
    with open("{}.{}".format(k, v[0]), "wb") as f:
        f.write(v[1])
```

在命令行中使用:

```sh
python -m XlsImage example.xls ./output
```
