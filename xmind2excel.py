
from typing import List, Any
import xlwt
from xmindparser import xmind_to_dict
import xlrd



def resolve_path(dict_, lists, title):
    """
    通过递归取出每个主分支下的所有小分支并将其作为一个列表
    :param dict_:
    :param lists:
    :param title:
    :return:
    """
    # 去除title首尾空格
    title = title.strip()
    # 若title为空，则直接取value
    if len(title) == 0:
        concat_title = dict_["title"].strip()
    else:
        concat_title = title + "\t" + dict_["title"].strip()
    if not dict_.__contains__("topics"):
        lists.append(concat_title)
    else:
        for d in dict_["topics"]:
            resolve_path(d, lists, concat_title)


def merge_cells(input_file,row,col,style):

    style_top=style
    style_top.alignment.horz=xlwt.Alignment.HORZ_LEFT  # 垂直顶部对齐
    # 使用xlrd打开已有的Excel文件
    workbook_rd = xlrd.open_workbook(input_file)
    sheet_rd = workbook_rd.sheet_by_index(0)

    # 创建一个新的工作簿并添加工作表
    workbook_wt = xlwt.Workbook()
    sheet_wt = workbook_wt.add_sheet('Sheet1')

    sheet_wt.col(0).width = 256 * 10
    sheet_wt.col(1).width = 256 * 20
    sheet_wt.col(2).width = 256 * 30
    sheet_wt.col(3).width = 256 * 40
    sheet_wt.col(4).width = 256 * 60

    # 获取行数和列数
    row_count = row
    col_count = col

    # 使用双指针遍历每一列并合并相同的单元格
    for col_index in range(col_count):
        start_row = 0  # 起始指针
        end_row = 0  # 结束指针

        while start_row < row_count:
            # 找到相同内容的单元格范围
            while end_row + 1 < row_count and sheet_rd.cell_value(end_row, col_index) == sheet_rd.cell_value(
                    end_row + 1, col_index):
                end_row += 1

            # 如果start_row和end_row不相同，说明有需要合并的单元格
            if start_row != end_row:#预期结果不合并单元格
                # 合并单元格
                sheet_wt.write_merge(start_row, end_row, col_index, col_index,
                                     sheet_rd.cell_value(start_row, col_index),style)
            else:
                # 如果没有合并，直接写当前单元格
                if col_index==3 or col_index==4:


                    #操作步骤和预期结果左对齐
                    print(sheet_rd.cell_value(start_row, col_index))
                    sheet_wt.write(start_row, col_index, sheet_rd.cell_value(start_row, col_index), style_top)
                else:
                    sheet_wt.write(start_row, col_index, sheet_rd.cell_value(start_row, col_index), style)

            # 移动指针到下一个未处理的单元格
            end_row += 1
            start_row = end_row

    # 保存新的Excel文件
    workbook_wt.save(input_file)

#将最后一层的多个叶子结点合并
def mergr_list(list):
    if len(list)<2:
        return
    slow=0
    fast=1
    serial=1
    drop_index=[]
    ans=[]
    while fast<len(list):
        while fast<len(list) and list[slow][-2]==list[fast][-2] :
            if serial==1:
                list[slow][-1]=str(serial)+'.'+list[slow][-1]+'\n'
                serial+=1
            list[slow][-1]=list[slow][-1]+str(serial)+'.'+list[fast][-1]+'\n'
            drop_index.append(fast)
            fast+=1
            serial+=1
        slow=fast
        fast+=1
        serial=1
    for i in range(0,len(list)):
        if i in drop_index:
            continue
        else:
            ans.append(list[i])
    return ans

def process_operation_steps(arr):
    if not arr:
        return ""
    serial=1
    ans=""
    for elem in arr:
        ans=ans+str(serial) + '.' + elem+'\n'
        serial += 1
    return ans

def xmind_to_excel(list_, excel_path):
    f = xlwt.Workbook()
    # 生成单sheet的Excel文件，sheet名自取
    sheet = f.add_sheet("XX模块", cell_overwrite_ok=True)
    style = xlwt.XFStyle()
    # 设置单元格的对齐方式
    alignment = xlwt.Alignment()
    alignment.wrap = 1  # 开启换行
    # alignment.vert = xlwt.Alignment.VERT_TOP  # 垂直顶部对齐
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平居中
    alignment.vert = xlwt.Alignment.VERT_CENTER  # 垂直居中
    style.alignment = alignment

    # 第一行固定的表头标题
    row_header = ["模块", "功能", "子功能","操作步骤","预期结果"]
    for i in range(0, len(row_header)):
        sheet.write(0, i, row_header[i])

    # 设置列宽度（单位是1/256字符宽度）
    sheet.col(0).width = 256 * 10
    sheet.col(1).width = 256 * 20
    sheet.col(2).width = 256 * 30
    sheet.col(3).width = 256 * 40
    sheet.col(4).width = 256 * 60


    # 增量索引,表的行数
    index = 1

    for h in range(0, len(list_)):
        lists: List[Any] = []
        resolve_path(list_[h], lists, "")
        # print(lists)
        for i in range(0, len(lists)):
            str=lists[i]
            list=str.split("\t")
            lists[i]=list
        print(lists)

        lists=mergr_list(lists)
        for j in range(0, len(lists)):

            print(lists[j])
            n=len(lists[j])

            #最小有两个（模块-预期结果）
            module=lists[j][0] #模块
            expect_result=lists[j][-1] #预期结果
            sheet.write(index, 0, module,style)
            sheet.write(index, 4, expect_result,style)


            if n==3:
                function = lists[j][1]  # 功能
                sheet.write(index, 1, function, style)
            elif n==4:
                function = lists[j][1]  # 功能
                sub_function = lists[j][2]  # 子功能
                sheet.write(index, 1, function, style)
                sheet.write(index, 2, sub_function, style)
            elif n>=5:

                function = lists[j][1]  # 功能
                sub_function = lists[j][2]  # 子功能
                operation_steps = lists[j][3:-1]  # 操作步骤
                steps = process_operation_steps(operation_steps)
                sheet.write(index, 1, function, style)
                sheet.write(index, 2, sub_function, style)
                sheet.write(index, 3, steps, style)

            index+=1
            f.save(excel_path)

            # for n in range(0, len(lists[j])):
            #     # 生成第一列的序号
            #     sheet.write(j + index + 1, 0, j + index + 1)
            #     sheet.write(j + index + 1, n + 1, lists[j][n])
            #     # 自定义内容，比如：测试点/用例标题、预期结果、实际结果、操作步骤、优先级……
            #     # 这里为了更加灵活，除序号、模块、功能点的标题固定，其余以【自定义+序号】命名，如：自定义1，需生成Excel表格后手动修改
            #     if n >= 2:
            #         sheet.write(0, n + 1, "自定义" + str(n - 1))
            # # 遍历完lists并给增量索引赋值，跳出for j循环，开始for h循环
            # if j == len(lists) - 1:
            #     index += len(lists)
    f.save(excel_path)
    merge_cells(excel_path,index,5,style)



def run(xmind_path):
    # 将XMind转化成字典
    xmind_dict = xmind_to_dict(xmind_path)
    # print("将XMind中所有内容提取出来并转换成列表：", xmind_dict)
    # Excel文件与XMind文件保存在同一目录下
    excel_name = xmind_path.split('\\')[-1].split(".")[0] + '.xls'
    excel_path = "\\".join(xmind_path.split('\\')[:-1]) + "\\" + excel_name
    print(excel_path)
    # print("通过切片得到所有分支的内容：", xmind_dict[0]['topic']['topics'])
    xmind_to_excel(xmind_dict[0]['topic']['topics'], excel_path)


if __name__ == '__main__':
    xmind_path_ = r"C:\Users\Chen\Desktop\【通用】排行榜.xmind"
    run(xmind_path_)

