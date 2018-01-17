import xlwt
import xlrd 
import os
#在哪里搜索多个表格  
filelocation="E:\\医生"  
#当前文件夹下搜索的文件名后缀  
filedestination="E:\\"  
#合并后的表格命名为file  
wfile="test"  
# 表头与sheet名关联关系
biao = {}
# 当前新sheet编号
sheetIndex = 0
#遍历目标目录下所有文件 
writeFile=xlwt.Workbook()

for filename in os.listdir(filelocation):
    if filename.split(".")[-1] != "xls" and filename.split(".")[-1] != "xlsx":
        continue
    print("开始处理%s..."%(filename,))  
    # 遍历文件所有sheet
    for sh in  xlrd.open_workbook(filelocation + "\\" + filename).sheets():
        nrows=sh.nrows
        ncols=sh.ncols
        # row为0时是表头，根据表头放到不同sheet里
        strbiao = ""
        biaotou = []
        for k in range(0,ncols):
            strbiao += sh.cell(0,k).value.strip()
            biaotou.append(sh.cell(0,k).value.strip())
            pass
        print("获取当前表表头...")
        if strbiao == '':
            continue
            pass
        # 表头集合中无该表
        if strbiao not in biao.keys():
            sheet = writeFile.add_sheet("sheet"+str(sheetIndex))
            print("不存在相同表头的sheet，新建sheet%s存放当前sheet数据"%(str(sheetIndex),))
            # 创建表头
            for i in range(0,len(biaotou)):  
                sheet.write(0,i,biaotou[i])
                pass
            biao[strbiao] = {"sheetNam":"sheet"+str(sheetIndex),"biaotou":biaotou,"rowIndex":1}
            sheetIndex += 1
            pass

        # 根据当前表头获取要操作的sheet
        sheet = writeFile.get_sheet(biao[strbiao]["sheetNam"])
        rowIndex = biao[strbiao]["rowIndex"]
        # 从第二行也就是数据行开始遍历当前sheet
        for j in range(1,nrows):  
            for k in range(0,ncols):
                sheet.write(rowIndex,k,sh.cell(j,k).value)
                pass
            rowIndex += 1
            pass
        # 行数更新回dict
        biao[strbiao]["rowIndex"] = rowIndex
        print("当前sheet处理完毕")
        pass
    print("%s处理完毕..."%(filename,))
    writeFile.save(filedestination + wfile + ".xls")
    pass

print("操作完成")  
  