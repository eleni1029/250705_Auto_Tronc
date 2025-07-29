import os
import glob
import pandas as pd
from datetime import datetime
import re

from sub_todolist_result import process_sheet_data
from sub_todolist_resource import extract_resources_from_result, get_resource_statistics
import json

def get_analyzed_files():
    """獲取 to_be_executed 目錄中所有 analyzed 檔案，按時間排序"""
    pattern = os.path.join('to_be_executed', 'course_structures_analyzed_*.xlsx')
    files = glob.glob(pattern)
    
    if not files:
        print("❌ 在 to_be_executed 目錄中找不到 analyzed 檔案")
        return []
    
    # 按修改時間排序（最新的在前）
    files.sort(key=os.path.getmtime, reverse=True)
    return files

def extract_timestamp(filename):
    """從檔案名中提取時間戳"""
    match = re.search(r'analyzed_(\d{8}_\d{6})', filename)
    if match:
        return match.group(1)
    return "unknown"

def select_file(files):
    """讓用戶選擇檔案 - GUI終端友善版本"""
    print("\n📁 找到以下 analyzed 檔案（按時間排序，最新的在前）：")
    for i, file in enumerate(files, 1):
        timestamp = extract_timestamp(file)
        print(f"{i}. {os.path.basename(file)} (時間戳: {timestamp})")
    
    print("\n📝 選擇說明：")
    print(f"- 輸入 1-{len(files)} 之間的數字選擇特定檔案")
    print("- 輸入 '0' 使用最新的檔案 (預設)")
    
    while True:
        try:
            print(f"\n請選擇檔案 (1-{len(files)}) 或輸入'0'使用預設:")
            print(">>> ", end="", flush=True)
            choice = input().strip()
            if not choice:
                print("⚠️ 請輸入有效值，或輸入 '0' 使用預設值")
                continue
            if choice == '0':
                print(f"\n✅ 已選擇預設檔案: {os.path.basename(files[0])}")
                return files[0]  # 預設選擇最新的
            
            choice_num = int(choice)
            if 1 <= choice_num <= len(files):
                print(f"\n✅ 已選擇檔案: {os.path.basename(files[choice_num - 1])}")
                return files[choice_num - 1]
            else:
                print(f"❌ 請輸入 1-{len(files)} 之間的數字")
        except ValueError:
            print("❌ 請輸入有效的數字")

def get_sheet_names(file_path):
    """獲取 Excel 檔案中的所有 sheet 名稱"""
    try:
        excel_file = pd.ExcelFile(file_path)
        return excel_file.sheet_names
    except Exception as e:
        print(f"❌ 讀取 Excel 檔案時出錯: {e}")
        return []

def analyze_duplicate_names(all_result_data):
    """分析重複名稱統計"""
    duplicate_stats = {
        'courses': {'total': 0, 'duplicates': [], 'unique_originals': 0},
        'chapters': {'total': 0, 'duplicates': [], 'unique_originals': 0},
        'units': {'total': 0, 'duplicates': [], 'unique_originals': 0}
    }
    
    # 分析課程重複
    course_originals = {}
    for item in all_result_data:
        if item['類型'] == '課程':
            course_name = item['名稱']
            # 檢查是否為編號版本
            if '_' in course_name and course_name.split('_')[-1].isdigit():
                original_name = '_'.join(course_name.split('_')[:-1])
            else:
                original_name = course_name
            
            if original_name not in course_originals:
                course_originals[original_name] = []
            course_originals[original_name].append(course_name)
    
    duplicate_stats['courses']['unique_originals'] = len(course_originals)
    for original, versions in course_originals.items():
        if len(versions) > 1:
            duplicate_stats['courses']['duplicates'].append({
                'original': original,
                'versions': versions,
                'count': len(versions)
            })
    duplicate_stats['courses']['total'] = sum(len(v) for v in course_originals.values())
    
    # 分析章節重複（按課程分組）
    chapter_by_course = {}
    for item in all_result_data:
        if item['類型'] == '章節':
            course_name = item['所屬課程']
            chapter_name = item['名稱']
            
            if course_name not in chapter_by_course:
                chapter_by_course[course_name] = {}
            
            # 檢查是否為編號版本
            if '_' in chapter_name and chapter_name.split('_')[-1].isdigit():
                original_name = '_'.join(chapter_name.split('_')[:-1])
            else:
                original_name = chapter_name
            
            if original_name not in chapter_by_course[course_name]:
                chapter_by_course[course_name][original_name] = []
            chapter_by_course[course_name][original_name].append(chapter_name)
    
    total_chapters = 0
    for course, chapters in chapter_by_course.items():
        for original, versions in chapters.items():
            total_chapters += len(versions)
            if len(versions) > 1:
                duplicate_stats['chapters']['duplicates'].append({
                    'course': course,
                    'original': original,
                    'versions': versions,
                    'count': len(versions)
                })
    duplicate_stats['chapters']['total'] = total_chapters
    duplicate_stats['chapters']['unique_originals'] = sum(len(chapters) for chapters in chapter_by_course.values())
    
    # 分析單元重複（按課程和章節分組）
    unit_by_course_chapter = {}
    for item in all_result_data:
        if item['類型'] == '單元':
            course_name = item['所屬課程']
            chapter_name = item['所屬章節']
            unit_name = item['名稱']
            
            key = (course_name, chapter_name)
            if key not in unit_by_course_chapter:
                unit_by_course_chapter[key] = {}
            
            # 檢查是否為編號版本
            if '_' in unit_name and unit_name.split('_')[-1].isdigit():
                original_name = '_'.join(unit_name.split('_')[:-1])
            else:
                original_name = unit_name
            
            if original_name not in unit_by_course_chapter[key]:
                unit_by_course_chapter[key][original_name] = []
            unit_by_course_chapter[key][original_name].append(unit_name)
    
    total_units = 0
    for (course, chapter), units in unit_by_course_chapter.items():
        for original, versions in units.items():
            total_units += len(versions)
            if len(versions) > 1:
                duplicate_stats['units']['duplicates'].append({
                    'course': course,
                    'chapter': chapter,
                    'original': original,
                    'versions': versions,
                    'count': len(versions)
                })
    duplicate_stats['units']['total'] = total_units
    duplicate_stats['units']['unique_originals'] = sum(len(units) for units in unit_by_course_chapter.values())
    
    return duplicate_stats

def select_sheet(sheet_names):
    """讓用戶選擇 sheet - GUI終端友善版本"""
    print(f"\n📋 檔案包含以下 {len(sheet_names)} 個 sheet:")
    for i, sheet in enumerate(sheet_names, 1):
        print(f"{i}. {sheet}")
    
    print("\n📝 選擇說明：")
    print(f"- 輸入 1-{len(sheet_names)} 之間的數字選擇單個 sheet")
    print("- 輸入多個數字並用逗號分隔選擇多個 sheet")
    print("  例如: 1,3,5 (選擇第1、3、5個sheet)")
    print("  例如: 2,4,6,8 (選擇第2、4、6、8個sheet)")
    print("- 輸入 'all' 選擇所有 sheet")
    
    while True:
        try:
            # 分離提示訊息和輸入，確保提示訊息能顯示
            print(f"\n請選擇 sheet (1-{len(sheet_names)})、多個數字用逗號分隔、或輸入'all':")
            print(">>> ", end="", flush=True)
            choice = input().strip().lower()
            
            if choice == 'all':
                print(f"\n✅ 已選擇所有 sheet: {len(sheet_names)} 個")
                return sheet_names
            
            # 檢查是否包含逗號（多選模式）
            if ',' in choice:
                # 解析多個數字
                numbers = []
                selected_sheets = []
                error_occurred = False
                
                for num_str in choice.split(','):
                    num_str = num_str.strip()
                    if not num_str:
                        continue
                    
                    try:
                        num = int(num_str)
                        if 1 <= num <= len(sheet_names):
                            if num not in numbers:  # 避免重複選擇
                                numbers.append(num)
                                selected_sheets.append(sheet_names[num - 1])
                        else:
                            print(f"❌ 數字 {num} 超出範圍 (1-{len(sheet_names)})")
                            error_occurred = True
                            break
                    except ValueError:
                        print(f"❌ '{num_str}' 不是有效的數字")
                        error_occurred = True
                        break
                
                if error_occurred:
                    continue
                
                if selected_sheets:
                    print(f"\n✅ 已選擇 {len(selected_sheets)} 個 sheet:")
                    for i, sheet in enumerate(selected_sheets, 1):
                        print(f"  {i}. {sheet}")
                    return selected_sheets
                else:
                    print("❌ 沒有選擇任何有效的 sheet")
                    continue
            
            else:
                # 單選模式
                choice_num = int(choice)
                if 1 <= choice_num <= len(sheet_names):
                    selected_sheet = sheet_names[choice_num - 1]
                    print(f"\n✅ 已選擇 sheet: {selected_sheet}")
                    return [selected_sheet]
                else:
                    print(f"❌ 請輸入 1-{len(sheet_names)} 之間的數字，或輸入 'all'")
        except ValueError:
            print("❌ 請輸入有效的數字、用逗號分隔的多個數字、或 'all'")

def analyze_course_statistics(all_result_data, all_resource_data, timestamp):
    """
    分析各課程統計數據並生成詳細報告
    
    Args:
        all_result_data: 所有結果數據
        all_resource_data: 所有資源數據
        timestamp: 時間戳
    """
    # HTML 擴展名配置（與 sub_mp4filepath_identifier.py 保持一致）
    HTML_EXTENSIONS = ['.html', '.htm']
    
    # 分析重複名稱統計
    duplicate_stats = analyze_duplicate_names(all_result_data)
    
    # 按課程分組統計
    course_stats = {}
    course_order = []
    
    # 建立課程與來源Sheet的映射
    course_to_sheet = {}
    
    # 統計各課程的基本數據
    for item in all_result_data:
        course_name = item['所屬課程']
        if not course_name or course_name == 'nan':
            continue
            
        # 記錄課程與來源Sheet的對應關係
        if course_name not in course_to_sheet and item.get('來源Sheet'):
            course_to_sheet[course_name] = item['來源Sheet']
            
        if course_name not in course_stats:
            course_stats[course_name] = {
                'chapters': set(),
                'units': set(), 
                'activities': [],
                'resources': set(),
                'html_files': 0,
                'video_activities': 0  # 修正：直接統計影音教材活動數量
            }
            course_order.append(course_name)
        
        stats = course_stats[course_name]
        
        if item['類型'] == '章節':
            stats['chapters'].add(item['名稱'])
        elif item['類型'] == '單元':
            stats['units'].add(item['名稱'])
        elif item['類型'] == '學習活動':
            stats['activities'].append(item)
            
            # 統計資源
            if item['檔案路徑']:
                stats['resources'].add(item['檔案路徑'])
            
            # 修正1：統計HTML線上連結數量（檢查網址路徑）
            if item['網址路徑']:
                # 檢查網址路徑中是否包含HTML檔案
                web_path = str(item['網址路徑']).lower()
                if any(ext.lower() in web_path for ext in HTML_EXTENSIONS):
                    stats['html_files'] += 1
            
            # 修正2：直接統計影音教材活動數量，不依賴檔案副檔名
            if item['學習活動類型'] in ['影音教材_影片', '影音教材_音訊']:
                stats['video_activities'] += 1
    
    # 生成日誌檔案
    log_filename = f"log/todolist_course_analysis_{timestamp}.log"
    os.makedirs("log", exist_ok=True)
    
    # 從 all_resource_data 中提取重複引用的資源統計
    # 建立檔案路徑到引用次數的映射
    file_path_references = {}
    for resource in all_resource_data:
        file_path = resource['檔案路徑']
        # 從 all_result_data 中計算實際引用次數
        reference_count = sum(1 for item in all_result_data 
                            if item['類型'] == '學習活動' and item['檔案路徑'] == file_path)
        file_path_references[file_path] = reference_count
    
    # 統計全局重複引用資源
    multiple_reference_resources = [
        {'檔案路徑': path, '引用次數': count} 
        for path, count in file_path_references.items() if count > 1
    ]
    total_saved_uploads = sum(r['引用次數'] - 1 for r in multiple_reference_resources)
    
    # 生成統計報告
    report_data = {
        'generation_time': datetime.now().isoformat(),
        'timestamp': timestamp,
        'total_courses': len(course_stats),
        'overall_summary': {
            'total_chapters': sum(len(stats['chapters']) for stats in course_stats.values()),
            'total_units': sum(len(stats['units']) for stats in course_stats.values()),
            'total_activities': sum(len(stats['activities']) for stats in course_stats.values()),
            'total_unique_resources': len(set().union(*[stats['resources'] for stats in course_stats.values()])),
            'total_html_files': sum(stats['html_files'] for stats in course_stats.values()),
            'total_video_activities': sum(stats['video_activities'] for stats in course_stats.values()),  # 修正
            'multiple_reference_resources': len(multiple_reference_resources),
            'saved_uploads': total_saved_uploads
        },
        'course_details': []
    }
    
    # 逐課程統計
    with open(log_filename, 'w', encoding='utf-8') as log_file:
        log_file.write("="*80 + "\n")
        log_file.write("TronClass 課程結構提取統計報告\n")
        log_file.write("="*80 + "\n")
        log_file.write(f"生成時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        log_file.write(f"時間戳: {timestamp}\n")
        log_file.write(f"總課程數: {len(course_stats)}\n\n")
        
        # 總體匯總
        log_file.write("總體匯總:\n")
        log_file.write("-"*40 + "\n")
        log_file.write(f"總章節數: {report_data['overall_summary']['total_chapters']}\n")
        log_file.write(f"總單元數: {report_data['overall_summary']['total_units']}\n") 
        log_file.write(f"總學習活動數: {report_data['overall_summary']['total_activities']}\n")
        log_file.write(f"總唯一資源數: {report_data['overall_summary']['total_unique_resources']}\n")
        log_file.write(f"總HTML線上連結數: {report_data['overall_summary']['total_html_files']}\n")
        log_file.write(f"總影音教材活動數: {report_data['overall_summary']['total_video_activities']}\n")  # 修正
        log_file.write(f"被多個學習活動引用的資源數: {report_data['overall_summary']['multiple_reference_resources']}\n")
        log_file.write(f"節省重複上傳的資源數: {report_data['overall_summary']['saved_uploads']}\n\n")
        
        # 重複名稱統計
        log_file.write("重複名稱處理統計:\n")
        log_file.write("-"*40 + "\n")
        log_file.write(f"課程重複: {len(duplicate_stats['courses']['duplicates'])} 組重複，共 {duplicate_stats['courses']['total']} 個課程\n")
        log_file.write(f"章節重複: {len(duplicate_stats['chapters']['duplicates'])} 組重複，共 {duplicate_stats['chapters']['total']} 個章節\n")
        log_file.write(f"單元重複: {len(duplicate_stats['units']['duplicates'])} 組重複，共 {duplicate_stats['units']['total']} 個單元\n")
        
        if duplicate_stats['courses']['duplicates']:
            log_file.write("\n重複課程詳情:\n")
            for dup in duplicate_stats['courses']['duplicates']:
                log_file.write(f"  - {dup['original']}: {', '.join(dup['versions'])}\n")
        
        if duplicate_stats['chapters']['duplicates']:
            log_file.write("\n重複章節詳情:\n")
            for dup in duplicate_stats['chapters']['duplicates']:
                log_file.write(f"  - {dup['course']} > {dup['original']}: {', '.join(dup['versions'])}\n")
        
        if duplicate_stats['units']['duplicates']:
            log_file.write("\n重複單元詳情:\n")
            for dup in duplicate_stats['units']['duplicates']:
                log_file.write(f"  - {dup['course']} > {dup['chapter']} > {dup['original']}: {', '.join(dup['versions'])}\n")
        
        log_file.write("\n")
        
        # 各課程詳細統計
        log_file.write("各課程詳細統計:\n")
        log_file.write("="*80 + "\n")
        
        # 顯示重複名稱統計
        if (duplicate_stats['courses']['duplicates'] or 
            duplicate_stats['chapters']['duplicates'] or 
            duplicate_stats['units']['duplicates']):
            print(f"\n🔄 重複名稱處理統計:")
            print(f"  課程重複: {len(duplicate_stats['courses']['duplicates'])} 組")
            print(f"  章節重複: {len(duplicate_stats['chapters']['duplicates'])} 組") 
            print(f"  單元重複: {len(duplicate_stats['units']['duplicates'])} 組")
        
        print(f"\n📊 各課程統計詳情:")
        print("="*60)
        
        for i, course_name in enumerate(course_order, 1):
            stats = course_stats[course_name]
            
            # 獲取課程的來源Sheet名稱
            source_sheet = course_to_sheet.get(course_name, '')
            display_name = f"{course_name}（{source_sheet}）" if source_sheet else course_name
            
            # 計算該課程的資源重複引用情況
            course_file_paths = stats['resources']
            course_multiple_reference = []
            course_saved_uploads = 0
            
            for file_path in course_file_paths:
                reference_count = file_path_references.get(file_path, 1)
                if reference_count > 1:
                    course_multiple_reference.append({
                        '檔案路徑': file_path,
                        '檔案名稱': os.path.basename(file_path),
                        '引用次數': reference_count
                    })
                    course_saved_uploads += reference_count - 1
            
            course_detail = {
                'order': i,
                'name': course_name,
                'chapters': len(stats['chapters']),
                'units': len(stats['units']),
                'activities': len(stats['activities']),
                'resources': len(stats['resources']),
                'html_files': stats['html_files'],
                'video_activities': stats['video_activities'],  # 修正
                'multiple_reference_resources': len(course_multiple_reference),
                'saved_uploads': course_saved_uploads
            }
            
            report_data['course_details'].append(course_detail)
            
            # 寫入日誌檔案
            log_file.write(f"{i}. {display_name}\n")
            log_file.write(f"   課程匯總: 章節數={len(stats['chapters'])}, 單元數={len(stats['units'])}, 學習活動數={len(stats['activities'])}, 資源數={len(stats['resources'])}\n")
            log_file.write(f"   細節匯總: HTML線上連結數={stats['html_files']}, 影音教材活動數={stats['video_activities']}, 被多個學習活動引用的資源數={len(course_multiple_reference)}, 節省重複上傳的資源數={course_saved_uploads}\n")  # 修正
            
            if course_multiple_reference:
                log_file.write(f"   重複引用資源詳情:\n")
                for res in course_multiple_reference:
                    log_file.write(f"     - {res['檔案名稱']}: 被{res['引用次數']}個學習活動引用\n")
            
            log_file.write("\n")
            
            # 控制台輸出
            print(f"{i:2d}. {display_name}")
            print(f"    📊 課程匯總: 章節={len(stats['chapters'])}, 單元={len(stats['units'])}, 活動={len(stats['activities'])}, 資源={len(stats['resources'])}")
            print(f"    🔍 細節匯總: HTML連結={stats['html_files']}, 影音教材={stats['video_activities']}, 重複引用資源={len(course_multiple_reference)}, 節省上傳={course_saved_uploads}")  # 修正
            if course_multiple_reference:
                # 顯示前3個重複引用的資源
                resource_names = []
                for r in course_multiple_reference[:3]:
                    resource_names.append(f"{r['檔案名稱']}({r['引用次數']})")
                resources_text = ', '.join(resource_names)
                
                if len(course_multiple_reference) > 3:
                    resources_text += f"... 等{len(course_multiple_reference)}個"
                
                print(f"    🔄 重複引用: {resources_text}")
            print()
        
        # 生成JSON格式的詳細報告
        json_filename = f"log/todolist_course_analysis_{timestamp}.json"
        with open(json_filename, 'w', encoding='utf-8') as json_file:
            json.dump(report_data, json_file, ensure_ascii=False, indent=2)
        
        log_file.write("="*80 + "\n")
        log_file.write("報告生成完成\n")
        log_file.write(f"JSON詳細報告: {json_filename}\n")
        
        print(f"📝 詳細統計報告已生成:")
        print(f"   📄 日誌檔案: {log_filename}")
        print(f"   📋 JSON報告: {json_filename}")
        
        return report_data

def find_header_positions(df):
    """動態查找表頭位置"""
    header_info = {
        'header_row': -1,
        'course_col': -1,
        'chapter_col': -1,
        'unit_cols': [],
        'activity_col': -1,
        'type_col': -1,
        'path_col': -1,
        'system_path_col': -1  # 新增系統路徑欄位
    }
    
    # 查找表頭行（包含"課程名稱"的行）
    for row_idx in range(min(15, len(df))):
        row_values = df.iloc[row_idx].astype(str).str.strip()
        if '課程名稱' in row_values.values:
            header_info['header_row'] = row_idx
            break
    
    if header_info['header_row'] == -1:
        print("❌ 找不到包含'課程名稱'的表頭行")
        return header_info
    
    # 查找各欄位位置
    header_row = df.iloc[header_info['header_row']].astype(str).str.strip()
    
    for col_idx, col_value in enumerate(header_row):
        if col_value == '課程名稱':
            header_info['course_col'] = col_idx
        elif col_value == '章節':
            header_info['chapter_col'] = col_idx
        elif col_value.startswith('單元'):
            header_info['unit_cols'].append(col_idx)
        elif col_value == '學習活動':
            header_info['activity_col'] = col_idx
    
    # 查找類型、路徑和系統路徑欄位（可能在其他行）
    for row_idx in range(min(15, len(df))):
        row_values = df.iloc[row_idx].astype(str).str.strip()
        for col_idx, col_value in enumerate(row_values):
            if col_value == '類型' and header_info['type_col'] == -1:
                header_info['type_col'] = col_idx
            elif col_value == '路徑' and header_info['path_col'] == -1:
                header_info['path_col'] = col_idx
            elif col_value == '系統路徑' and header_info['system_path_col'] == -1:
                header_info['system_path_col'] = col_idx
    
    print(f"  ✅ 表頭資訊: 表頭行={header_info['header_row']+1}, 課程列={header_info['course_col']+1}, "
          f"章節列={header_info['chapter_col']+1}, 單元列={[c+1 for c in header_info['unit_cols']]}, "
          f"活動列={header_info['activity_col']+1}, 類型列={header_info['type_col']+1}, "
          f"路徑列={header_info['path_col']+1}, 系統路徑列={header_info['system_path_col']+1}")
    
    return header_info
    """動態查找表頭位置"""
    header_info = {
        'header_row': -1,
        'course_col': -1,
        'chapter_col': -1,
        'unit_cols': [],
        'activity_col': -1,
        'type_col': -1,
        'path_col': -1,
        'system_path_col': -1  # 新增系統路徑欄位
    }
    
    # 查找表頭行（包含"課程名稱"的行）
    for row_idx in range(min(15, len(df))):
        row_values = df.iloc[row_idx].astype(str).str.strip()
        if '課程名稱' in row_values.values:
            header_info['header_row'] = row_idx
            break
    
    if header_info['header_row'] == -1:
        print("❌ 找不到包含'課程名稱'的表頭行")
        return header_info
    
    # 查找各欄位位置
    header_row = df.iloc[header_info['header_row']].astype(str).str.strip()
    
    for col_idx, col_value in enumerate(header_row):
        if col_value == '課程名稱':
            header_info['course_col'] = col_idx
        elif col_value == '章節':
            header_info['chapter_col'] = col_idx
        elif col_value.startswith('單元'):
            header_info['unit_cols'].append(col_idx)
        elif col_value == '學習活動':
            header_info['activity_col'] = col_idx
    
    # 查找類型、路徑和系統路徑欄位（可能在其他行）
    for row_idx in range(min(15, len(df))):
        row_values = df.iloc[row_idx].astype(str).str.strip()
        for col_idx, col_value in enumerate(row_values):
            if col_value == '類型' and header_info['type_col'] == -1:
                header_info['type_col'] = col_idx
            elif col_value == '路徑' and header_info['path_col'] == -1:
                header_info['path_col'] = col_idx
            elif col_value == '系統路徑' and header_info['system_path_col'] == -1:
                header_info['system_path_col'] = col_idx
    
    print(f"  ✅ 表頭資訊: 表頭行={header_info['header_row']+1}, 課程列={header_info['course_col']+1}, "
          f"章節列={header_info['chapter_col']+1}, 單元列={[c+1 for c in header_info['unit_cols']]}, "
          f"活動列={header_info['activity_col']+1}, 類型列={header_info['type_col']+1}, "
          f"路徑列={header_info['path_col']+1}, 系統路徑列={header_info['system_path_col']+1}")
    
    return header_info

def create_extracted_excel(source_file, selected_sheets, timestamp):
    """創建提取後的 Excel 檔案"""
    output_filename = os.path.join('to_be_executed', f"todolist_extracted_{timestamp}.xlsx")
    
    all_result_data = []
    
    # 重要：在處理所有 sheets 之前，重置全域重複名稱管理器和章節計數器
    from sub_todolist_result import duplicate_manager, course_chapter_counter
    duplicate_manager.__init__()  # 重置管理器
    course_chapter_counter.clear()  # 重置章節計數器
    
    # 處理每個選中的 sheet
    for sheet_name in selected_sheets:
        print(f"\n📋 正在處理 sheet: {sheet_name}")
        
        try:
            # 讀取 sheet 資料
            df = pd.read_excel(source_file, sheet_name=sheet_name, header=None)
            
            # 查找表頭位置
            header_info = find_header_positions(df)
            
            # 使用 sub_todolist_result 處理資料
            result_data = process_sheet_data(df, sheet_name, header_info)
            
            all_result_data.extend(result_data)
            
            print(f"  ✅ 處理完成: {len(result_data)} 條記錄")
            
        except Exception as e:
            print(f"  ❌ 處理 sheet {sheet_name} 時出錯: {e}")
    
    # 使用 sub_todolist_resource 從 result_data 中提取資源（去重版本）
    print(f"\n📦 正在提取資源檔案清單...")
    all_resource_data = extract_resources_from_result(all_result_data)
    
    # 獲取資源統計資訊
    resource_stats = get_resource_statistics(all_result_data)
    
    # 顯示統計資訊
    print(f"\n📊 資源處理統計:")
    print(f"  - 有檔案路徑的學習活動: {resource_stats['total_activities_with_files']} 個")
    print(f"  - 唯一檔案路徑: {resource_stats['unique_file_paths']} 個")
    print(f"  - 生成資源記錄: {len(all_resource_data)} 筆")
    
    if resource_stats['multiple_reference_files'] > 0:
        print(f"  - 重複引用的檔案: {resource_stats['multiple_reference_files']} 個")
        print(f"    💡 這些檔案只會上傳一次，但被多個學習活動引用")
    
    # 創建 Excel 檔案
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        # 1. 複製原始資料到 Ori_document sheet
        print(f"\n📄 正在複製原始資料...")
        all_ori_data = []
        
        for sheet_name in selected_sheets:
            try:
                # 使用 header=None 讀取，保持原始格式
                df = pd.read_excel(source_file, sheet_name=sheet_name, header=None)
                # 添加來源標記列
                df['原始Sheet'] = sheet_name
                all_ori_data.append(df)
            except Exception as e:
                print(f"  ❌ 複製 sheet {sheet_name} 時出錯: {e}")
        
        if all_ori_data:
            # 垂直合併，保持原始結構
            combined_df = pd.concat(all_ori_data, ignore_index=True, sort=False)
            combined_df.to_excel(writer, sheet_name='Ori_document', index=False, header=False)
            print(f"  ✅ 已保存原始資料 ({len(combined_df)} 行)")
        
        # 2. 保存 Result sheet
        if all_result_data:
            result_df = pd.DataFrame(all_result_data)
            result_df.to_excel(writer, sheet_name='Result', index=False)
            print(f"  ✅ 已保存 Result 資料 ({len(all_result_data)} 條記錄)")
        else:
            # 創建空的 Result sheet
            result_columns = ['類型', '名稱', 'ID', '所屬課程', '所屬課程ID', '所屬章節', '所屬章節ID', 
                            '所屬單元', '所屬單元ID', '學習活動類型', '網址路徑', '檔案路徑', '資源ID', '最後修改時間', '來源Sheet']
            result_df = pd.DataFrame(columns=pd.Index(result_columns))
            result_df.to_excel(writer, sheet_name='Result', index=False)
            print(f"  ✅ 已創建空的 Result sheet")
        
        # 3. 保存 Resource sheet
        if all_resource_data:
            resource_df = pd.DataFrame(all_resource_data)
            resource_df.to_excel(writer, sheet_name='Resource', index=False)
            print(f"  ✅ 已保存 Resource 資料 ({len(all_resource_data)} 個唯一資源)")
        else:
            # 創建空的 Resource sheet 但包含標題行
            resource_columns = ['檔案名稱', '檔案路徑', '資源ID', '最後修改時間', '來源Sheet', '引用學習活動數', '引用活動列表']
            resource_df = pd.DataFrame(columns=pd.Index(resource_columns))
            resource_df.to_excel(writer, sheet_name='Resource', index=False)
            print(f"  ✅ 已創建空的 Resource sheet")
    
    print(f"\n🎉 提取完成！檔案已生成: {output_filename}")
    
    # 顯示重複資源節省的資訊
    if resource_stats['total_activities_with_files'] > resource_stats['unique_file_paths']:
        saved_count = resource_stats['total_activities_with_files'] - resource_stats['unique_file_paths']
        print(f"\n💾 資源去重效果:")
        print(f"  - 原本需要: {resource_stats['total_activities_with_files']} 個資源上傳")
        print(f"  - 去重後需要: {resource_stats['unique_file_paths']} 個資源上傳")
        print(f"  - 節省上傳: {saved_count} 個重複資源")
    
    # 生成詳細的課程統計報告
    print(f"\n📋 正在生成詳細統計報告...")
    analyze_course_statistics(all_result_data, all_resource_data, timestamp)
    
    return output_filename

def main():
    """主函數 - GUI終端友善版本"""
    print("🚀 TronClass 課程結構資料提取器 - 資源去重版本")
    print("✨ 新功能：相同檔案路徑的資源只會生成一筆記錄")
    print("=" * 55)
    print("🔍 正在檢查 to_be_executed 目錄...")
    
    # 1. 獲取 analyzed 檔案
    files = get_analyzed_files()
    if not files:
        print("❌ 未找到 analyzed 檔案，腳本結束")
        return
    
    print(f"✅ 找到 {len(files)} 個 analyzed 檔案")
    
    # 2. 讓用戶選擇檔案
    selected_file = select_file(files)
    print(f"\n✅ 已選擇檔案: {os.path.basename(selected_file)}")
    
    # 3. 獲取 sheet 名稱
    print(f"\n🔍 正在讀取 Excel 檔案結構...")
    sheet_names = get_sheet_names(selected_file)
    if not sheet_names:
        print("❌ 無法讀取 Excel 檔案，腳本結束")
        return
    
    print(f"✅ 檔案包含 {len(sheet_names)} 個 sheet")
    
    # 4. 讓用戶選擇 sheet
    selected_sheets = select_sheet(sheet_names)
    print(f"\n✅ 已選擇 sheet: {', '.join(selected_sheets)}")
    
    # 5. 提取時間戳
    timestamp = extract_timestamp(selected_file)
    
    # 6. 創建提取後的 Excel 檔案
    print(f"\n🔄 正在處理數據並生成輸出檔案...")
    print("🔄 這可能需要一些時間，請耐心等待...")
    output_file = create_extracted_excel(selected_file, selected_sheets, timestamp)
    print(f"\n✅ 輸出檔案生成完成: {output_file}")
    
    print(f"\n📊 生成的檔案包含以下 sheet:")
    print("  1. Ori_document - 原始資料")
    print("  2. Result - 提取的結構化資料")
    print("  3. Resource - 需要上傳的資源檔案清單（已去重）")
    
    print(f"\n💡 改進說明:")
    print("  ✅ 已自動提取課程、章節、單元、學習活動的層級關係")
    print("  ✅ 已自動識別線上連結和檔案路徑")
    print("  ✅ 已自動生成資源上傳清單（相同檔案路徑去重）")
    print("  ✅ 支援動態欄位位置識別")
    print("  ✅ 支援多種檔案結構")
    print("  🆕 Resource 表新增欄位：引用學習活動數、引用活動列表")
    print("  🆕 相同檔案路徑的資源只生成一筆記錄，避免重複上傳")

if __name__ == "__main__":
    main()