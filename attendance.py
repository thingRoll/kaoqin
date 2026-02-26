import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.cell.cell import MergedCell
import warnings
import sys
import re
import os
import tkinter as tk
from tkinter import filedialog
import time

# 强制输出编码
sys.stdout.reconfigure(encoding='utf-8')
warnings.filterwarnings('ignore')

# ================= 0. 核心工具：获取真实路径 (新增) =================
def get_application_path():
    """
    获取程序运行的真实目录。
    兼容：Python脚本运行模式 和 打包后的EXE运行模式
    """
    if getattr(sys, 'frozen', False):
        # 如果是打包后的 exe，使用 exe 所在路径
        return os.path.dirname(sys.executable)
    else:
        # 如果是 python 脚本，使用脚本所在路径
        return os.path.dirname(os.path.abspath(__file__))

# 获取基础路径
BASE_DIR = get_application_path()

# ================= 1. 配置区域 =================
# 使用绝对路径，确保百分百找到文件
TEMPLATE_FILE = os.path.join(BASE_DIR, '模板-考勤.xlsx')
OUTPUT_FILE_PREFIX = os.path.join(BASE_DIR, '结果-本月考勤_')

# 统计分类库
LOC_PROVINCE_IN = [
    '济南', '威海', '济宁', '曲阜', '兖州', '龙口', '烟台', '青岛', '淄博', 'emc', '公司', '本部', 
    '会展', '大安机场', '文化中心', 
    '济', '白', '曲', '郓', '枣', '新', '梁', '博', '聊'
]
LOC_PROVINCE_OUT = [
    '北京', '门源', '邵寨', '方城', '上海', '深圳', '河南', '甘肃', '南京', 
    '京', '蒙', '贵'
]
SITE_DAYS_DEPT_KEYWORDS = ['运维', '工程技术']

PROJECT_MAPPING = {
    '黄河国际会展中心': '会展',
    '济宁大安机场': '大安机场',
    '济宁文化中心': '文化中心',
    '美年大健康': '南京',
    '邵寨': '邵寨'
}
CITY_ABBREVIATIONS = {
    '梁宝寺': '梁', '郓': '郓', '郓城': '郓', '白庄': '白',
    '曲阜': '曲', '尼山': '曲', '北京': '京', '博兴': '博',
    '聊城': '聊', '内蒙': '蒙', '枣庄': '枣', '新驿': '新', '贵州': '贵'
}

def pause_and_exit(code=0):
    print("\n" + "="*30)
    input("👉 程序执行完毕，请按【回车键】关闭窗口...")
    sys.exit(code)

# ================= 2. 文件选择模块 (修复版) =================
print("******************************************")
print("      全自动考勤计算系统 V16.0      ")
print("******************************************\n")
print(f"当前工作目录: {BASE_DIR}")

# 检查模板
if not os.path.exists(TEMPLATE_FILE):
    print(f"!!! 错误：未找到模板文件！")
    print(f"!!! 请确保 '模板-考勤.xlsx' 位于文件夹:\n    {BASE_DIR}")
    pause_and_exit(1)

print(">>> [1/6] 正在唤起文件选择窗口...")

# --- 修复弹窗不显示的问题 ---
try:
    root = tk.Tk()
    root.withdraw() # 隐藏主窗口
    root.attributes('-topmost', True) # 关键：强制置顶，防止被控制台遮挡
    
    SOURCE_FILE = filedialog.askopenfilename(
        parent=root,
        title='请选择本月的钉钉考勤导出报表 (Excel)',
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    
    root.destroy() # 选完后立即销毁窗口资源
except Exception as e:
    print(f"!!! 弹窗启动失败: {e}")
    print("请尝试以管理员身份运行。")
    pause_and_exit(1)
# ----------------------------

if not SOURCE_FILE:
    print("!!! 未选择文件，操作已取消。")
    pause_and_exit()

print(f"    已选择: {os.path.basename(SOURCE_FILE)}")

# ================= 3. 智能日期提取 =================
print(">>> [2/6] 分析考勤周期...")
try:
    df_meta = pd.read_excel(SOURCE_FILE, sheet_name='月度汇总', header=None, nrows=1)
    meta_text = str(df_meta.iloc[0, 0]) 
    dates_found = re.findall(r'(\d{4}-\d{2}-\d{2})', meta_text)
    if len(dates_found) >= 2:
        DATE_RANGE_START = dates_found[0]
        DATE_RANGE_END = dates_found[1]
        print(f"    周期: {DATE_RANGE_START} 至 {DATE_RANGE_END}")
        month_str = DATE_RANGE_START.split('-')[1]
    else:
        print("    警告：无法提取日期，使用默认配置。")
        DATE_RANGE_START = '2025-11-26'
        DATE_RANGE_END = '2025-12-25'
        month_str = "XX"
    date_list = pd.date_range(start=DATE_RANGE_START, end=DATE_RANGE_END)
except Exception as e:
    print(f"!!! 日期提取失败: {e}")
    pause_and_exit(1)

# 生成输出文件名 (使用绝对路径)
OUTPUT_FILE = f"{OUTPUT_FILE_PREFIX}{month_str}月.xlsx"

# ================= 4. 数据读取 =================
print(">>> [3/6] 读取数据中...")
try:
    df_stats = pd.read_excel(SOURCE_FILE, sheet_name='月度汇总', header=2)
    if '姓名' not in df_stats.columns: df_stats.rename(columns={df_stats.columns[0]: '姓名'}, inplace=True)
    df_stats['match_name'] = df_stats['姓名'].astype(str).str.replace(' ', '').str.strip()

    df_daily_source = pd.read_excel(SOURCE_FILE, sheet_name='月度汇总', header=3)
    df_daily_source.rename(columns={df_daily_source.columns[0]: '姓名'}, inplace=True)
    df_daily_source['match_name'] = df_daily_source['姓名'].astype(str).str.replace(' ', '').str.strip()

    df_records = pd.read_excel(SOURCE_FILE, sheet_name='原始记录', header=2)
    if '姓名' not in df_records.columns: df_records.rename(columns={df_records.columns[0]: '姓名'}, inplace=True)
    df_records['match_name'] = df_records['姓名'].astype(str).str.replace(' ', '').str.strip()
    df_records['date_clean'] = df_records['考勤日期'].astype(str).apply(lambda x: str(x).split(' ')[0])
except Exception as e:
    print(f"!!! 数据读取失败: {e}\n请确保选择了正确的钉钉导出文件。")
    pause_and_exit(1)

# ================= 5. 模板清理与备份 =================
print(">>> [4/6] 备份存班并清空旧数据...")
def safe_write(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    if isinstance(cell, MergedCell): return
    cell.value = value

try:
    # 使用绝对路径打开模板
    wb = openpyxl.load_workbook(TEMPLATE_FILE)
    ws = wb['当月考勤']
    
    stat_col_map = {}
    date_start_col = 0
    for r in [2, 3]:
        for col in range(1, 50):
            val = str(ws.cell(row=r, column=col).value).strip()
            if val in ['出勤日', '省内', '省外', '加班', '病假', '请假', '调休', '迟到', '旷工', '工地天数', '存班']:
                stat_col_map[val] = col
            if val == '存班': date_start_col = col + 1
    if date_start_col == 0: date_start_col = 14 
    
    old_banked_data = {}
    name_col = 2
    
    for row in range(4, ws.max_row + 1):
        name_cell = ws.cell(row=row, column=name_col).value
        if not name_cell: continue
        name = str(name_cell).replace(' ', '').strip()
        
        if '存班' in stat_col_map:
            val = ws.cell(row=row, column=stat_col_map['存班']).value
            s_val = str(val).strip()
            nums = re.findall(r"[-+]?\d*\.\d+|\d+", s_val)
            old_banked_data[name] = float(nums[0]) if nums else 0.0
            
        for col_idx in stat_col_map.values(): safe_write(ws, row, col_idx, None)
        for col_idx in range(date_start_col, ws.max_column + 1): safe_write(ws, row, col_idx, None)

except Exception as e:
    print(f"!!! 模板清理失败: {e}\n请检查模板文件是否被占用。")
    pause_and_exit(1)

# ================= 6. 重绘表头 =================
print(">>> [5/6] 绘制新表头...")
try:
    week_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
    current_col = date_start_col
    date_col_map = {} 
    date_strs = [d.strftime('%Y-%m-%d') for d in date_list]
    
    for i, dt in enumerate(date_list):
        cell = ws.cell(row=3, column=current_col)
        if not isinstance(cell, MergedCell):
            cell.value = dt.day
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        cell = ws.cell(row=4, column=current_col)
        if not isinstance(cell, MergedCell):
            cell.value = week_map[dt.weekday()]
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        date_col_map[current_col] = date_strs[i]
        current_col += 1
except Exception as e:
    print(f"!!! 表头重绘失败: {e}")
    pause_and_exit(1)

# ================= 7. 逻辑工具 =================
def get_day_type(date_str):
    try:
        dt = pd.to_datetime(date_str)
        if dt.weekday() >= 5: return 'weekend'
        return 'workday'
    except: return 'workday'

def parse_number(val):
    if pd.isna(val) or val == '': return 0.0
    s = str(val).strip()
    nums = re.findall(r"[-+]?\d*\.\d+|\d+", s)
    if nums: return float(nums[0])
    return 0.0

def extract_city(address):
    if not isinstance(address, str): return None
    match = re.search(r'([\u4e00-\u9fa5]{2,})市', address)
    if match: return match.group(1) 
    for city in ['北京', '上海', '天津', '重庆']:
        if city in address: return city
    return None

def analyze_attendance(name, full_date_str, daily_row, records_df, department):
    day_str = str(int(full_date_str.split('-')[-1]))
    status = "正常"
    val = None
    if day_str in daily_row: val = daily_row[day_str]
    elif int(day_str) in daily_row: val = daily_row[int(day_str)]
    if val is not None: status = str(val)

    short_date = full_date_str[5:]
    daily_recs = records_df[
        (records_df['match_name'] == name) & 
        (records_df['date_clean'].str.endswith(short_date))
    ]
    
    found_locs = []
    has_field_work = False 
    
    for _, rec in daily_recs.iterrows():
        addr = str(rec['打卡地址']) + str(rec['打卡备注'])
        res = str(rec['打卡结果'])
        if '外勤' in res or '外勤' in status: has_field_work = True
        
        curr = None
        for keyword, symbol in PROJECT_MAPPING.items():
            if keyword in addr:
                curr = symbol
                break
        
        if not curr:
            for keyword, symbol in CITY_ABBREVIATIONS.items():
                if keyword in addr:
                    curr = symbol
                    break
        
        if not curr:
            if '威海' in addr: curr = '威海'
            elif '门源' in addr: curr = '门源'
            elif '龙口' in addr: curr = '龙口'
            elif '方城' in addr: curr = '方城'
            elif '兖州' in addr: curr = '兖州'
            elif '济宁' in addr: curr = '济宁' 
            
            if not curr and has_field_work:
                city = extract_city(rec['打卡地址'])
                if city: curr = city

        if curr and curr not in found_locs:
            found_locs.append(curr)

    base_text = '√'
    loc_type = 'company'
    
    if found_locs:
        base_text = '/'.join(found_locs)
        if any(l in LOC_PROVINCE_OUT for l in found_locs): loc_type = 'province_out'
        elif any(l in LOC_PROVINCE_IN for l in found_locs): loc_type = 'province_in'
        elif has_field_work: loc_type = 'province_in'
    elif '邵寨' in department:
        base_text = '邵寨'
        loc_type = 'province_out'
    else:
        if daily_recs.empty and get_day_type(full_date_str) == 'weekend':
             base_text = '○'
             loc_type = 'rest'

    if '节假日' in status: return '※', 'rest'
    if '休息' in status: return '○', 'rest'
    
    if '请假' in status:
        if '0.5' in status or '半天' in status:
             time_match = re.search(r'(\d{2}:\d{2})', status)
             if time_match and int(time_match.group(1).split(':')[0]) < 12:
                 return f"假/{base_text}", loc_type
             else:
                 return f"{base_text}/假", loc_type
        return '假', 'leave'
    
    if '事假' in status: return '假', 'leave'
    if '病假' in status: return '病假', 'leave'
    if '年假' in status: return '年', 'leave'

    if '调休' in status:
        if '0.5' in status or '半天' in status:
             time_match = re.search(r'(\d{2}:\d{2})', status)
             if time_match and int(time_match.group(1).split(':')[0]) < 12:
                 return f"调/{base_text}", loc_type
             else:
                 return f"{base_text}/调", loc_type
        return '调休', 'comp_leave'

    if '旷工' in status and '旷工迟到' not in status: return '×', 'absent'
    if '迟到' in status: return '迟', loc_type

    is_weekend = get_day_type(full_date_str) == 'weekend'
    if is_weekend and base_text == '√' and loc_type != 'rest':
        return '+', 'company_ot'
    
    if daily_recs.empty and is_weekend and '正常' not in status:
        return '○', 'rest'
        
    return base_text, loc_type

# ================= 8. 数据填充 =================
print(f">>> [6/6] 开始计算与填充...")
processed_cnt = 0

for row in range(4, ws.max_row + 1):
    try:
        name_cell = ws.cell(row=row, column=name_col).value
        if not name_cell: continue
        name = str(name_cell).replace(' ', '').strip()
        
        stats_row = df_stats[df_stats['match_name'] == name]
        daily_row = df_daily_source[df_daily_source['match_name'] == name]
        
        if stats_row.empty or daily_row.empty: continue
        processed_cnt += 1
        
        stats_data = stats_row.iloc[0]
        daily_data = daily_row.iloc[0]
        dept = str(stats_data.get('部门', ''))
        
        old_banked = old_banked_data.get(name, 0.0)

        dt_comp_leave = parse_number(daily_data.get('调休(天)', 0))
        dt_personal_leave = parse_number(daily_data.get('事假(天)', 0))
        dt_sick_leave = parse_number(daily_data.get('病假(天)', 0))
        
        dt_ot = 0.0
        for k in ['工作日加班', '休息日加班', '节假日加班']:
            dt_ot += parse_number(daily_data.get(k, 0))

        balance = old_banked + dt_ot - dt_comp_leave
        new_banked = 0.0
        write_comp_leave = 0.0
        write_personal_leave = 0.0
        
        if balance < 0:
            deficit = abs(balance)
            new_banked = 0
            avail = max(0, old_banked + dt_ot)
            write_comp_leave = min(dt_comp_leave, avail)
            write_personal_leave = dt_personal_leave + deficit
        else:
            new_banked = balance
            write_comp_leave = dt_comp_leave
            write_personal_leave = dt_personal_leave

        local_prov_in = 0
        local_prov_out = 0
        
        for col, d_str in date_col_map.items():
            txt, l_type = analyze_attendance(name, d_str, daily_data, df_records, dept)
            safe_write(ws, row, col, txt)
            if l_type == 'province_in': local_prov_in += 1
            if l_type == 'province_out': local_prov_out += 1
        
        if '存班' in stat_col_map: safe_write(ws, row, stat_col_map['存班'], new_banked if new_banked != 0 else None)
        if '调休' in stat_col_map: safe_write(ws, row, stat_col_map['调休'], write_comp_leave if write_comp_leave > 0 else None)
        if '请假' in stat_col_map: safe_write(ws, row, stat_col_map['请假'], write_personal_leave if write_personal_leave > 0 else None)
        if '病假' in stat_col_map and dt_sick_leave > 0: safe_write(ws, row, stat_col_map['病假'], dt_sick_leave)
        if '加班' in stat_col_map and dt_ot > 0: safe_write(ws, row, stat_col_map['加班'], dt_ot)
        if '迟到' in stat_col_map:
            late = parse_number(stats_data.get('迟到次数', 0)) + parse_number(stats_data.get('旷工迟到次数', 0))
            if late > 0: safe_write(ws, row, stat_col_map['迟到'], late)
        if '旷工' in stat_col_map:
            absent = parse_number(stats_data.get('旷工天数', 0))
            if absent > 0: safe_write(ws, row, stat_col_map['旷工'], absent)
        if '出勤日' in stat_col_map: safe_write(ws, row, stat_col_map['出勤日'], stats_data.get('出勤天数', 0))
        if '省内' in stat_col_map: safe_write(ws, row, stat_col_map['省内'], local_prov_in)
        if '省外' in stat_col_map: safe_write(ws, row, stat_col_map['省外'], local_prov_out)
        if '工地天数' in stat_col_map:
            if any(k in dept for k in SITE_DAYS_DEPT_KEYWORDS):
                safe_write(ws, row, stat_col_map['工地天数'], local_prov_in + local_prov_out)
            else:
                safe_write(ws, row, stat_col_map['工地天数'], None)

    except Exception as row_error:
        print(f"!!! 出错 [行{row} {name}]: {row_error}")
        continue

# ================= 9. 保存 =================
print(">>> 正在保存文件...")
try:
    wb.save(OUTPUT_FILE)
    print(f"\n{'='*40}")
    print(f"✅ 处理成功！\n✅ 生成文件: {os.path.basename(OUTPUT_FILE)}\n✅ 处理人数: {processed_cnt}")
    print(f"{'='*40}")
except Exception as e:
    print(f"!!! 保存失败: {e}\n请确保文件未被占用！")

pause_and_exit()