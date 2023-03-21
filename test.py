import src.XlsImage as xls

def xls_test():
    imgs = xls.XlsGetImages("./untrack/test.xls")
    print("total images: ", len(imgs))
    for k, v in enumerate(imgs):
        with open("./untrack/{}.{}".format(k, v[0]), "wb") as f:
            f.write(v[1])



if __name__ == "__main__":
    import sys
    allFuncs = dir()
    allObjs = vars()
    argv = sys.argv
    if len(argv) > 1:
        funcName = argv[1]
        argv = argv[1:]
        if funcName == "-a":
            for i in allFuncs:
                if callable(allObjs[i]):
                    print("----- {} -----".format(i))
                    allObjs[i]()
        else:
            for i in argv:
                if i in allFuncs:
                    print("----- {} -----".format(i))
                    allObjs[i]()
                else:
                    print("没有函数: " + funcName)
    else:
        print("使用方法:\npython mtest.py {测试函数1} {测试函数2} ...\n"
              "python mtest.py -a 执行全测试")