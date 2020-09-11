# Excel 模块
- xlrd:read excel
- xlwt:write excel
- xlsxwriter:写入文件，比xlwt更新
- xlutils:辅助工具
## 1、xlrd
```
# 得到整个文件
wb = xlrd.open_workbook("文件路径")
# 得到文件中的第0号sheet
ws = wb.sheet_by_index(0)
# 显示行数列数
ws.ncols
ws.nrows
# 获取索引为2的一整行/列的数据，返回值为列表
# row_v = ws.row_values(row, start, step)
row_v = ws.row_values(2)
```

## 2、xlwt
- xlwt不能改写原文件，只是创建一个新的文件然后覆盖原来的文件，步骤如下：
    - 1.打开文件，即创建了新的文件
    - 2.为新的文件添加一个sheet
    - 3.在新的sheet上添加数据
    - xlsxwriter是基于xlwt的更新版本的库，功能更加细化
```
# 创建一个指定名字的文件，没有就创建，有就覆盖掉
wb = xlwt.Workbook("文件名字")
# 为文件添加一个指定名字的sheet
ws = wb.add_sheet("sheet的名字")
# 写内容
ws.write(row, col, msg)

# 最后必须保存
wb.save('文件名')
# 以上就完成了最基本的数据写入功能
```
- xlwt还可以设置写的风格
    ```
    wb = xlwt.Workbook("文件名字")
    ws = wb.add_sheet("sheet的名字")
  
    # -----设置单元格对齐方式-----
    style = xlwt.XFStyle()
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平方向对齐方式
    al.vert = 0x01  # 设置竖直方向对齐方式
        # VERT_TOP    = 0x00    竖直向上对齐  
        # VERT_CENTER = 0x01    竖直方向居中对齐
        # VERT_BOTTOM = 0x02    竖直向下对齐
        # HORZ_LEFT   = 0x01    水平靠左对齐
        # HORZ_CENTER = 0x02    水平居中对齐
        # HORZ_RIGHT  = 0x03    水平靠右对齐
    style.alignment = al
    # --------------------------
  
    ws.write(row, col, msg, style)
    
    # 最后必须保存
    wb.save('文件名')
    ```

- xlwriter基于xlwt,并且使用起来更加的方便
- 如果需要进行对表格进行内容的添加，需要使用到xlutils
    - 实现方法如下
    ```
  from xlutils.copy import copy
  
  read_data = xlrd.open_workbook('源文件名')
  work_book = copy(read_data)
  work_sheet = work_book.get_sheet(0)
  # 写入数据
  work_sheet.write(要写入的内容)
  
  work_book.save(文件名)
  ```