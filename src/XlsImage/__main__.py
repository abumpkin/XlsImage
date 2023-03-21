from XlsImage import *

if __name__ == "__main__":
    import sys, os
    allFuncs = dir()
    allObjs = vars()
    argv = sys.argv
    Usage = f'''usage: XlsImage <xls file> [output dir]
'''
    if len(argv) > 1:
        argv = argv[1:]
        output = ""
        xlsFile = argv.pop(0)
        if argv:
            output = os.path.join(output, argv.pop(0))
        try:
            os.makedirs(output)
        except:
            pass
        if not os.path.exists(xlsFile):
            print("file doesn't exsits!")
            exit(-1)
        imgs = XlsGetImages(xlsFile)
        print("total images: ", len(imgs))
        for k, v in enumerate(imgs):
            filename = "{}.{}".format(k, v[0])
            filename = os.path.join(output, filename)
            print("write to: ", filename)
            with open(filename, "wb") as f:
                f.write(v[1])
        print("done!")
    else:
        print(Usage)