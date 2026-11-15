# NX图层切换 - 极简测试版
# 只做一件事：设置工作层

import NXOpen

TARGET_LAYER = 200   # ← 改成你要切到的层号

def main():
    theSession = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work

    if workPart is None:
        return

    # 切换工作层
    workPart.Layers.WorkLayer = TARGET_LAYER

  

if __name__ == '__main__':
    main()
