import json
import pandas as pd
import os
import math
import openpyxl

def read_fileindex(fileindex_path):
    fileindex_map = {}
    with open(fileindex_path, 'r') as f:
        lines = f.readlines()
        for line in lines:
            parts = line.strip().split("->")
            if len(parts) == 3:
                image_filename = parts[1]
                pile_number = parts[0].split(".")[0]
                pile_number = pile_number.replace("+", "")[:-3]
                fileindex_map[pile_number] = image_filename
    return fileindex_map

def process_label_file(label_file, fileindex_map, imageHeight, imageWidth):
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
    length_per_pixer = imageHeight / label_data['imageHeight'] 
    width_per_pixer = imageWidth / label_data['imageWidth'] 

    for shape in label_data['shapes']:
        label = get_name(shape['label'])
        points = shape['points']
        length = abs(points[0][0] - points[1][0]) * length_per_pixer
        width = abs(points[0][1] - points[1][1]) * width_per_pixer
        area = count_area(shape['label'], length, width)
        report_entries.append({
            '桩号': pile_number,
            '病害名称': label,
            '长度(m)': length,
            '宽度(m)': width,
            '面积(㎡)': area
        })

    return report_entries

def count_area(class_id, height ,width):
    #current classes: ["kuaizhuangliefeng", "hengxiangliefeng", "tiaozhuangxiubu", "kuaizhuangxiubu", "zongxiangliefeng", "junlie", "kengcao"]
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
    
def generate_report(label_dir, fileindex_path, output_path, imageHeight, imageWidth):
    fileindex_map = read_fileindex(fileindex_path)
    df_10m_report = pd.DataFrame(columns=['起点', '终点', '龟裂(㎡)', '块状裂缝(㎡)', '纵向裂缝(㎡)', '横向裂缝(㎡)', '坑槽(㎡)', '块状修补(㎡)', '条状修补(㎡)'])
    current_start_pile = None
    current_end_pile = None
    current_area = {
        '龟裂(㎡)': 0,
        '块状裂缝(㎡)': 0,
        '纵向裂缝(㎡)': 0,
        '横向裂缝(㎡)': 0,
        '坑槽(㎡)': 0,
        '块状修补(㎡)': 0,
        '条状修补(㎡)': 0
    }
    lines = list(fileindex_map.keys())
    report_data = []
    for i, pile_number in enumerate(lines):
        image_filename = fileindex_map[pile_number]
        label_file_path = os.path.join(label_dir, image_filename.replace(".jpg", ".json"))
        if os.path.exists(label_file_path):
            report_entries = process_label_file(label_file_path, fileindex_map, imageHeight, imageWidth)
            if report_entries:
                report_data.extend(report_entries)
                for entry in report_entries:
                    label = entry['病害名称'] + '(㎡)'
                    area = entry['面积(㎡)']
                    if label in current_area:
                        current_area[label] += area
        if i % 5 == 0:
            current_start_pile = str(int(pile_number) - 2)
        if i % 5 == 4 or i == len(lines) - 1:
            current_end_pile = pile_number
            df_10m_report = pd.concat([df_10m_report, pd.DataFrame([{
                '起点': current_start_pile,
                '终点': current_end_pile,
                **current_area
            }])], ignore_index=True)
            # 重置起点、终点和病害统计数据
            current_start_pile = None
            current_end_pile = None
            current_area = {
                '龟裂(㎡)': 0,
                '块状裂缝(㎡)': 0,
                '纵向裂缝(㎡)': 0,
                '横向裂缝(㎡)': 0,
                '坑槽(㎡)': 0,
                '块状修补(㎡)': 0,
                '条状修补(㎡)': 0
            }
    df = pd.DataFrame(report_data)
    df.insert(0, '序号', range(1, len(df) + 1))
    writer = pd.ExcelWriter(output_path, engine='openpyxl')

    # 生成报表
    df.to_excel(writer, index=False, sheet_name='病害明细')
    title_sheet = writer.book['病害明细']
    title_sheet.insert_rows(0, amount=1) 
    title_sheet['A1'] = "病害明细"
    title_sheet.merge_cells('A1:F1')

    title_cell = title_sheet['A1']
    title_cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")


    df_10m_report.to_excel(writer, index=False, sheet_name='路面损坏十米统计')
    title_sheet = writer.book['路面损坏十米统计']
    title_sheet.insert_rows(0, amount=1) 
    title_sheet['A1'] = "路面损坏十米统计"
    title_sheet.merge_cells('A1:I1')
    title_cell = title_sheet['A1']
    title_cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")


    writer.save()


