from typing import List, Any
import xlwt
from xmindparser import xmind_to_dict
import xlrd

# 增量索引,表的行数
index = 2
big_title=""
def resolve_path(dict_,f,file_path,sheet,style,level=0):
    """
    遍历字典结构，当某个层级的所有子节点都不含有 topics 时，打印这些子节点的 title；
    否则继续递归到下一个子节点。
    :param dict_: 当前分支的字典
    :param lists: 存储拼接标题的列表
    :param title: 上一级的标题
    :param level: 当前处理的层级
    :return:
    """
    global index

    # 判断当前节点是否有子节点
    if "topics" not in dict_:
        return
    else:
        # 检查子节点是否都不含有 topics
        all_leaf_nodes = all("topics" not in sub_dict for sub_dict in dict_["topics"])

        if all_leaf_nodes:
            # 打印当前层级下的所有子节点的标题
            numbered_topics=[f"{i + 1}. {sub_dict['title']}" for i, sub_dict in enumerate(dict_['topics'])]
            topics_str='\n'.join(numbered_topics)
            sheet.write(index, 0, level,style)
            sheet.write(index, 1, dict_['title'],style)
            sheet.write(index, 4, topics_str,style)
            index += 1
            f.save(file_path)
            print(f"Level {level}: {dict_['title']} -> 子节点: {topics_str}")
            return
        else:
            # 打印当前层级并继续递归处理
            print(f"Level {level}: {dict_['title']}(继续递归下一级)")
            sheet.write(index, 0, level,style)
            sheet.write(index, 1, dict_['title'],style)
            index+=1
            print(index)
            f.save(file_path)
            for sub_dict in dict_["topics"]:
                resolve_path(sub_dict,f,file_path,sheet,style,level + 1)




def xmind_to_excel(list_, excel_path):
    f = xlwt.Workbook()
    # 生成单sheet的Excel文件，sheet名自取
    sheet = f.add_sheet("模块", cell_overwrite_ok=True)
    style = xlwt.XFStyle()
    # 设置单元格的对齐方式
    alignment = xlwt.Alignment()
    alignment.wrap = 1  # 开启换行
    # alignment.vert = xlwt.Alignment.VERT_TOP  # 垂直顶部对齐
    alignment.horz=xlwt.Alignment.HORZ_LEFT
    # alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平居中
    alignment.vert = xlwt.Alignment.VERT_CENTER  # 垂直居中
    style.alignment = alignment

    # 第一行固定的表头标题
    # 创建加粗样式
    bold_style = xlwt.XFStyle()
    font = xlwt.Font()
    font.bold = True  # 设置字体为加粗
    bold_style.font = font

    row_header = ["级别", "功能点", "测试点", "操作步骤", "预期结果","测试结果","负责人","备注"]
    for i in range(0, len(row_header)):
        sheet.write(0, i, row_header[i],bold_style)

    # 设置列宽度（单位是1/256字符宽度）
    sheet.col(0).width = 256 * 10
    sheet.col(1).width = 256 * 50
    sheet.col(2).width = 256 * 50
    sheet.col(3).width = 256 * 60
    sheet.col(4).width = 256 * 80
    sheet.col(5).width = 256 * 15
    sheet.col(6).width = 256 * 15
    sheet.col(7).width = 256 * 15
    sheet.col(8).width = 256 * 15

    sheet.write(1, 0, 1, style)
    sheet.write(1, 1, list_['title'], style)
    f.save(excel_path)

    for h in range(0, len(list_['topics'])):
        lists: List[Any] = []
        resolve_path(list_['topics'][h],f,excel_path,sheet,style,2)
        # print(lists)

def run(xmind_path):
    # 将XMind转化成字典
    xmind_dict = xmind_to_dict(xmind_path)
    # print("将XMind中所有内容提取出来并转换成列表：", xmind_dict)
    # Excel文件与XMind文件保存在同一目录下
    excel_name = xmind_path.split('\\')[-1].split(".")[0] + '.xls'
    excel_path = "\\".join(xmind_path.split('\\')[:-1]) + "\\" + excel_name
    print(excel_path)
    # print("通过切片得到所有分支的内容：", xmind_dict[0]['topic']['topics'])
    xmind_to_excel(xmind_dict[0]['topic'], excel_path)


if __name__ == '__main__':
    xmind_path_ = r"C:\Users\v_ahongchen\Desktop\【通用】排行榜.xmind"
    run(xmind_path_)

