import json
import pandas as pd
import os
import math

def read_fileindex(fileindex_path):
    fileindex_map = {}
    with open(fileindex_path, 'r') as f:
        lines = f.readlines()
        for line in lines:
            parts = line.strip().split("->")
            if len(parts) == 3:
                image_filename = parts[1]
                pile_number = parts[0].split(".")[0]
                fileindex_map[pile_number] = image_filename
    return fileindex_map

def process_label_file(label_file, fileindex_map):
    with open(label_file, 'r') as f:
        label_data = json.load(f)

    image_filename = os.path.basename(label_data['imagePath'])
    pile_number = None

    for key, value in fileindex_map.items():
        if value == image_filename:
            pile_number = key
            break

    if pile_number is None:
        return None

    report_entries = []

    for shape in label_data['shapes']:
        label = get_name(shape['label'])
        points = shape['points']
        area = count_area(shape['label'], points)
        length = abs(points[0][0] - points[1][0])
        width = abs(points[0][1] - points[1][1])

        report_entries.append({
            '桩号': pile_number,
            '病害名称': label,
            '长度(m)': length,
            '宽度(m)': width,
            '面积(㎡)': area
        })

    return report_entries

def count_area(class_id,points):
    #current classes: ["kuaizhuangliefeng", "hengxiangliefeng", "tiaozhuangxiubu", "kuaizhuangxiubu", "zongxiangliefeng", "junlie", "kengcao"]
    x1, y1 = points[0]
    x2, y2 = points[1]
    width = abs(x2 - x1)
    height = abs(y2 - y1)

    if class_id == "kuaizhuangliefeng":
        return width * height
    elif class_id == "hengxiangliefeng":
        return math.sqrt(width**2 + height**2) * 0.2
    elif class_id == "tiaozhuangxiubu": 
        return math.sqrt(width**2 + height**2) * 0.2
    elif class_id == "kuaizhuangxiubu": 
        return width * height * 0.2
    elif class_id == "zongxiangliefeng": 
        return math.sqrt(width**2 + height**2) * 0.2
    elif class_id == "junlie":
        return width * height
    elif class_id == "kengcao": 
        return width * height
    
def get_name(class_id):
    #current classes: ["kuaizhuangliefeng", "hengxiangliefeng", "tiaozhuangxiubu", "kuaizhuangxiubu", "zongxiangliefeng", "junlie", "kengcao"]
    if class_id == "kuaizhuangliefeng":
        return "块状裂缝"
    elif class_id == "hengxiangliefeng":
        return "横向裂缝"
    elif class_id == "tiaozhuangxiubu": 
        return "条状修补"
    elif class_id == "kuaizhuangxiubu": 
        return "块状修补"
    elif class_id == "zongxiangliefeng": 
        return "纵向裂缝"
    elif class_id == "junlie":
        return "龟裂"
    elif class_id == "kengcao": 
        return "坑槽"
    
def generate_report(label_dir, fileindex_path, output_path):
    # 读取fileindex.txt文件
    fileindex_map = read_fileindex(fileindex_path)

    # 遍历label文件夹，生成报表数据
    report_data = []
    for label_file in os.listdir(label_dir):
        if label_file.endswith('.json'):
            label_file_path = os.path.join(label_dir, label_file)
            report_entries = process_label_file(label_file_path, fileindex_map)
            if report_entries:
                report_data.extend(report_entries)

    # 创建DataFrame
    df = pd.DataFrame(report_data)

    # 添加序号列
    df.insert(0, '序号', range(1, len(df) + 1))

    # 生成报表
    df.to_excel(output_path, index=False)