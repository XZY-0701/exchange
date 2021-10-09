from xmindparser import xmind_to_dict
import xlwt
import patterns as patterns

def traversal_xmind(root, rootstring, listcontainer):

    if isinstance(root, dict):
        if 'title' in root.keys() and 'topics' in root.keys():
            traversal_xmind(root['topics'], str(rootstring), listcontainer)
        if 'title' in root.keys() and 'topics' not in root.keys():
            traversal_xmind(root['title'], str(rootstring), listcontainer)
    elif isinstance(root, list):
        for sonroot in root:
            traversal_xmind(sonroot, str(rootstring) + "&" + sonroot['title'], listcontainer)
            if 'makers' in sonroot and 'callout' in sonroot:
                traversal_xmind(sonroot, str(rootstring) + "&" + sonroot['title'] + "&" + str(sonroot['makers'][0]) +
                                "&" + str(sonroot['callout'][0]), listcontainer)
            elif 'callout' in sonroot and 'makers' not in sonroot:
                traversal_xmind(sonroot, str(rootstring) + "&" + sonroot['title'] + "&" + str(sonroot['callout'][0]),
                                listcontainer)
            elif 'makers' in sonroot and 'callout' not in sonroot:
                traversal_xmind(sonroot, str(rootstring) + "&" + sonroot['title'] + "&" + str(sonroot['makers'][0]),
                                listcontainer)

    elif isinstance(root, str):
        listcontainer.append(str(rootstring))

def get_case(root):
    rootstring = root['title']
    listcontainer = []
    traversal_xmind(root, rootstring, listcontainer)
    return listcontainer

def maker_judgement(makers):
    maker = "用例等级不合法"
    if 'priority-1' in makers:
        maker = "高"
    elif 'priority-2' in makers:
        maker = "中"
    elif 'priority-3' in makers:
        maker = "低"
    return maker

def write_sheet(b, filename, name, demand_id, callout, step, result, case_type, case_state, maker, creator):
    font = xlwt.Font()
    font.height = 20 * 11  # 设置字体大小
    # 设置单元格对齐方式
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x01
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01
    # 设置自动换行
    alignment.wrap = 1
    # 设置背景颜色
    pattern = xlwt.Pattern()
    # 设置背景颜色的模式
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # 背景颜色
    pattern.pattern_fore_colour = 26
    borders = xlwt.Borders()
    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    # 初始化
    style1 = xlwt.XFStyle()
    style1.font = font
    style1.alignment = alignment
    style1.pattern = pattern
    style1.borders = borders

    worksheet.write(b, 0, filename, style1)
    worksheet.write(b, 1, name, style1)
    worksheet.write(b, 2, demand_id, style1)
    worksheet.write(b, 3, callout, style1)
    worksheet.write(b, 4, step, style1)
    worksheet.write(b, 5, result, style1)
    worksheet.write(b, 6, case_type, style1)
    worksheet.write(b, 7, case_state, style1)
    worksheet.write(b, 8, maker, style1)
    worksheet.write(b, 9, creator, style1)


def deal_with_list(list):

    b = 1
    for i in list:
        j = i.split("&")

        if 'priority-1' in j or 'priority-2' in j or 'priority-3' in j:
            print(j)
            # print(j[1])
            # print(j[1].split("\n")[0].split("：")[1])
            x = 0
            if 'priority-1' in j:
                x = j.index('priority-1')
            elif 'priority-2' in j:
                x = j.index('priority-2')
            elif 'priority-3' in j:
                x = j.index('priority-3')

            if j[-1] == j[x]:
                result = ""
                step = ""
                callout = ""
            else:
                result = j[-1]
                step = j[-2]
                if j[-2] == j[x + 1]:
                    callout = ""
                else:
                    callout = j[x + 1:x + 2]
            maker = maker_judgement(j[x])
            filename = j[2:x - 1]
            catalogue = ""
            for f in filename:
                catalogue += f.replace("*", "-")
            name = ""
            demand_id = j[1].split("\n")[0].split("：")[1]
            creator = j[1].split("\n")[1].split("：")[1]
            case_type = j[1].split("\n")[2].split("：")[1]
            case_state = j[1].split("\n")[3].split("：")[1]
            for a in j[x - 1]:
                name += a
            write_sheet(b, j[0] + catalogue, name.split("#")[0], demand_id, callout, step, result, case_type, case_state, maker, creator)
            b += 1

def get_path():
    path = input("请输入文件地址：")
    if "/" in path:
        path0 = path.replace('/', "//")
    else:
        path0 = path.replace("\\", "/")
    return path0

if __name__ == '__main__':
    try:
        path = get_path()
        root = xmind_to_dict(path)[0]['topic']     # xmind文件路径
    except IOError:
        print("找不到文件，请重新输入")
    else:
        row0 = ["用例目录", "用例名称", "需求ID", "前置条件", "用例步骤", "预期结果", "用例类型", "用例状态", "用例等级", "创建人"]
        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet(str(root['title']), cell_overwrite_ok=True)

        # 设置首行字体格式
        font = xlwt.Font()
        font.name = '方正书宋_GBK'
        font.bold = True
        font.height = 20 * 16  # 设置字体大小

        # 设置背景颜色
        pattern = xlwt.Pattern()
        # 设置背景颜色的模式
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # 背景颜色
        pattern.pattern_fore_colour = 47
        # pattern.pattern_fore_colour = 3

        borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1

        #  初始化样式
        style0 = xlwt.XFStyle()
        style0.font = font
        style0.pattern = pattern
        style0.borders = borders

        for i in range(len(row0)):
            worksheet.write(0, i, row0[i], style0)
            worksheet.col(i).width = 25 * 256

        case = get_case(root)
        deal_with_list(case)
        workbook.save(root['title'] + '.xls')
        print("用例转换完成！")