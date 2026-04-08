import pandas as pd
import re
import json
from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn
from rich.prompt import Prompt, Confirm
from rich.table import Table
from rich import box
import os
import sys
from pathlib import Path
from copy import deepcopy
import openpyxl
from openpyxl.styles import numbers as openpyxl_numbers

console = Console()

# openpyxl 不允许写入的控制字符（会触发 IllegalCharacterError）
# 规则参考 openpyxl.cell.cell.ILLEGAL_CHARACTERS_RE
_ILLEGAL_EXCEL_CHAR_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def sanitize_excel_cell_value(value):
    """清理 Excel 单元格中 openpyxl 不允许的字符"""
    if value is None:
        return value

    # pandas 读入空值后可能是 NaN（float）
    try:
        if pd.isna(value):
            return value
    except Exception:
        pass

    text = str(value)
    return _ILLEGAL_EXCEL_CHAR_RE.sub("", text)

class DataFormat:
    """数据格式枚举"""
    JSON = "json"
    KEY_VALUE = "key_value"
    PLAIN_TEXT = "plain_text"

def detect_format(text):
    """检测文本格式"""
    if pd.isna(text) or text == 'nan':
        return None
    
    text = str(text).strip()
    
    # 尝试解析为JSON
    try:
        json.loads(text)
        return DataFormat.JSON
    except:
        pass
    
    # 检测是否为键值对格式 {key:value}
    if re.search(r'\{[^:}]+:.*?\}', text):
        return DataFormat.KEY_VALUE
    
    # 否则为纯文本
    return DataFormat.PLAIN_TEXT

def extract_keys_from_json(text):
    """从JSON格式提取键"""
    try:
        data = json.loads(str(text))
        if isinstance(data, dict):
            return list(data.keys())
    except:
        pass
    return []

def extract_keys_from_key_value(text):
    """从键值对格式提取键 - 修复版本"""
    if pd.isna(text) or text == 'nan':
        return []
    
    text = str(text)
    keys = []
    
    i = 0
    while i < len(text):
        if text[i] == '{':
            colon_pos = text.find(':', i)
            if colon_pos == -1:
                i += 1
                continue
            
            brace_count = 1
            j = i + 1
            right_brace_pos = -1
            
            while j < len(text) and brace_count > 0:
                if text[j] == '{':
                    brace_count += 1
                elif text[j] == '}':
                    brace_count -= 1
                    if brace_count == 0:
                        right_brace_pos = j
                        break
                j += 1
            
            if right_brace_pos != -1 and colon_pos < right_brace_pos:
                key = text[i+1:colon_pos].strip()
                if key:
                    keys.append(key)
                i = right_brace_pos + 1
            else:
                i += 1
        else:
            i += 1
    
    return keys

def scan_column_format(df, column_name):
    """扫描列的格式分布"""
    format_counts = {
        DataFormat.JSON: 0,
        DataFormat.KEY_VALUE: 0,
        DataFormat.PLAIN_TEXT: 0,
        None: 0
    }
    
    all_keys = set()
    
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console
    ) as progress:
        task = progress.add_task("[cyan]正在扫描数据格式...", total=len(df))
        
        for idx, value in enumerate(df[column_name]):
            fmt = detect_format(value)
            format_counts[fmt] = format_counts.get(fmt, 0) + 1
            
            if fmt == DataFormat.JSON:
                keys = extract_keys_from_json(value)
                all_keys.update(keys)
            elif fmt == DataFormat.KEY_VALUE:
                keys = extract_keys_from_key_value(value)
                all_keys.update(keys)
            
            progress.update(task, advance=1)
    
    return format_counts, sorted(list(all_keys))

def extract_value_from_json(text, key):
    """从JSON中提取值"""
    try:
        data = json.loads(str(text))
        if isinstance(data, dict) and key in data:
            return data[key]
    except:
        pass
    return ""

def extract_value_from_key_value(text, key):
    """从键值对中提取值"""
    if pd.isna(text) or text == 'nan':
        return ""
    
    text = str(text)
    escaped_key = re.escape(key)
    pattern = rf'\{{{escaped_key}:'
    match = re.search(pattern, text)
    
    if not match:
        return ""
    
    start_pos = match.end()
    brace_count = 1
    i = start_pos
    
    while i < len(text) and brace_count > 0:
        if text[i] == '{':
            brace_count += 1
        elif text[i] == '}':
            brace_count -= 1
            if brace_count == 0:
                value = text[start_pos:i].strip()
                return value
        i += 1
    
    return ""

def extract_key_value_pair(text, key):
    """从键值对中提取完整的键值对"""
    if pd.isna(text) or text == 'nan':
        return ""
    
    text = str(text)
    escaped_key = re.escape(key)
    pattern = rf'\{{{escaped_key}:'
    match = re.search(pattern, text)
    
    if not match:
        return ""
    
    start_pos = match.start()
    brace_count = 1
    i = match.end()
    
    while i < len(text) and brace_count > 0:
        if text[i] == '{':
            brace_count += 1
        elif text[i] == '}':
            brace_count -= 1
            if brace_count == 0:
                return text[start_pos:i+1]
        i += 1
    
    return ""

def remove_key_from_json(text, key):
    """从JSON中删除键"""
    try:
        data = json.loads(str(text))
        if isinstance(data, dict) and key in data:
            del data[key]
            return json.dumps(data, ensure_ascii=False)
    except:
        pass
    return text

def remove_key_from_key_value(text, key):
    """从键值对中删除键"""
    if pd.isna(text) or text == 'nan':
        return ""
    
    text = str(text)
    key_value_pair = extract_key_value_pair(text, key)
    if key_value_pair:
        text = text.replace(key_value_pair + '\n', '')
        text = text.replace('\n' + key_value_pair, '')
        text = text.replace(key_value_pair, '')
    
    return text.strip()

def process_escape_chars(text):
    """处理转义字符"""
    if pd.isna(text) or text == 'nan':
        return ""
    
    text = str(text)
    text = text.replace('\\n', '\n')
    text = text.replace('\\t', '\t')
    text = text.replace('\\r', '\r')
    text = text.replace('\\"', '"')
    text = text.replace("\\'", "'")
    text = text.replace('\\\\', '\\')
    
    return text

def show_format_distribution(format_counts):
    """显示格式分布"""
    table = Table(title="数据格式分布", box=box.ROUNDED)
    table.add_column("格式类型", style="cyan")
    table.add_column("数量", style="yellow", justify="right")
    table.add_column("占比", style="green", justify="right")
    
    total = sum(format_counts.values())
    
    format_names = {
        DataFormat.JSON: "JSON格式",
        DataFormat.KEY_VALUE: "键值对格式 {key:value}",
        DataFormat.PLAIN_TEXT: "纯文本",
        None: "空值"
    }
    
    for fmt, count in format_counts.items():
        if count > 0:
            percentage = f"{count / total * 100:.1f}%"
            table.add_row(format_names.get(fmt, "未知"), str(count), percentage)
    
    console.print(table)

def show_keys_table(keys, page_size=20):
    """显示键列表"""
    if not keys:
        console.print("[yellow]未检测到任何键[/yellow]")
        return
    
    total_pages = (len(keys) + page_size - 1) // page_size
    current_page = 0
    
    while True:
        console.clear()
        start_idx = current_page * page_size
        end_idx = min(start_idx + page_size, len(keys))
        page_keys = keys[start_idx:end_idx]
        
        table = Table(
            title=f"检测到的键 (共 {len(keys)} 个) - 第 {current_page + 1}/{total_pages} 页",
            box=box.ROUNDED
        )
        table.add_column("序号", style="cyan", justify="center", width=8)
        table.add_column("键名", style="green")
        
        for i, key in enumerate(page_keys, start=start_idx + 1):
            table.add_row(str(i), key)
        
        console.print(table)
        
        if total_pages > 1:
            console.print(f"\n[dim]提示: 输入 'n' 下一页, 'p' 上一页, 'q' 继续[/dim]")
            nav = Prompt.ask("导航", choices=["n", "p", "q"], default="q")
            
            if nav == "n" and current_page < total_pages - 1:
                current_page += 1
            elif nav == "p" and current_page > 0:
                current_page -= 1
            elif nav == "q":
                break
        else:
            break

def select_keys(all_keys):
    """让用户选择要提取的键"""
    if not all_keys:
        console.print("[yellow]没有可选择的键[/yellow]")
        return None
    
    console.print("\n[bold cyan]请选择要提取的键:[/bold cyan]")
    console.print("  [dim]输入格式: 可以输入序号(如: 1,3,5)或键名[/dim]")
    console.print("  [dim]输入 'all' 提取所有键[/dim]")
    console.print("  [dim]输入 'q' 返回上一步[/dim]\n")
    
    while True:
        selection = Prompt.ask("请输入选择")
        
        if selection.lower() == 'q':
            return None
        
        if selection.lower() == 'all':
            return all_keys
        
        selected_keys = []
        parts = [p.strip() for p in selection.split(',')]
        
        for part in parts:
            try:
                idx = int(part) - 1
                if 0 <= idx < len(all_keys):
                    selected_keys.append(all_keys[idx])
                else:
                    console.print(f"[yellow]⚠ 序号 {part} 超出范围[/yellow]")
            except ValueError:
                if part in all_keys:
                    selected_keys.append(part)
                else:
                    console.print(f"[yellow]⚠ 键名 '{part}' 不存在[/yellow]")
        
        if selected_keys:
            console.print(f"\n[green]✓ 已选择 {len(selected_keys)} 个键:[/green]")
            for key in selected_keys:
                console.print(f"  • {key}")
            
            if Confirm.ask("\n确认选择?", default=True):
                return selected_keys
        else:
            console.print("[red]✗ 没有选择任何有效的键[/red]")

def get_processing_options(has_keys):
    """获取处理选项"""
    options = {}
    
    console.print("\n[bold cyan]处理选项配置:[/bold cyan]")
    
    options['process_escape'] = Confirm.ask(
        "是否处理转义字符(如 \\n 转为换行)?",
        default=True
    )
    
    options['strip_values'] = Confirm.ask(
        "是否去除值的首尾空白字符(.strip())?",
        default=True
    )
    
    if has_keys:
        console.print("\n[bold]写入格式:[/bold]")
        console.print("  [cyan]1.[/cyan] 只写入值(键作为列名)")
        console.print("  [cyan]2.[/cyan] 写入完整键值对 {键:值}")
        
        write_format = Prompt.ask(
            "请选择",
            choices=["1", "2"],
            default="1"
        )
        options['write_key_value'] = (write_format == "2")
    else:
        options['write_key_value'] = False
    
    return options

def process_value(value, key, options, data_format):
    """根据选项处理值"""
    if pd.isna(value) or value == 'nan' or value == '':
        return ""
    
    value = str(value)
    
    if options.get('process_escape', False):
        value = process_escape_chars(value)
    
    if options.get('strip_values', False):
        value = value.strip()
    
    if options.get('write_key_value', False) and key:
        return f"{{{key}:{value}}}"
    
    return value

def show_operation_menu(current_column=None):
    """显示操作菜单"""
    console.print("\n[bold cyan]═══════════════════════════════════════[/bold cyan]")
    if current_column:
        console.print(f"[bold]当前处理列: [green]{current_column}[/green][/bold]")
    console.print("[bold]请选择操作:[/bold]")
    console.print("  [cyan]1.[/cyan] 提取键值对到新列")
    console.print("  [cyan]2.[/cyan] 删除指定键值对")
    console.print("  [cyan]3.[/cyan] 处理纯文本列")
    console.print("  [cyan]4.[/cyan] 切换到其他列")
    console.print("  [cyan]5.[/cyan] 保存并退出")
    console.print("  [cyan]6.[/cyan] 放弃修改并退出")
    
    return Prompt.ask("\n请输入选项", choices=["1", "2", "3", "4", "5", "6"], default="1")

def show_columns_menu(df, current_column=None):
    """显示列选择菜单"""
    console.print("\n[bold cyan]可用的列:[/bold cyan]")
    
    table = Table(box=box.ROUNDED)
    table.add_column("序号", style="cyan", justify="center", width=8)
    table.add_column("列名", style="green")
    table.add_column("状态", style="yellow", width=10)
    
    for idx, col in enumerate(df.columns):
        status = "当前列" if col == current_column else ""
        table.add_row(str(idx + 1), col, status)
    
    console.print(table)
    
    col_choice = Prompt.ask(
        f"\n请选择要处理的列 (1-{len(df.columns)}) 或输入 'q' 返回",
        default="q"
    )
    
    if col_choice.lower() == 'q':
        return None
    
    try:
        idx = int(col_choice) - 1
        if 0 <= idx < len(df.columns):
            return df.columns[idx]
        else:
            console.print("[red]✗ 无效的列编号[/red]")
            return None
    except ValueError:
        console.print("[red]✗ 请输入有效的数字[/red]")
        return None

def extract_keys_operation(df, column_name, all_keys, format_counts):
    """提取键值对操作"""
    console.print("\n[bold cyan]提取键值对到新列[/bold cyan]")

    show_keys_table(all_keys)

    selected_keys = select_keys(all_keys)
    if selected_keys is None:
        return df, False

    # 默认使用“键名作为列名”，同时支持用户自定义列名
    # 若自定义列名与已有列重名，后续直接覆盖（pandas 赋值行为）
    target_columns = {}
    custom_names = Confirm.ask("是否为提取结果自定义列名? (默认直接使用键名)", default=False)
    for key in selected_keys:
        if custom_names:
            col_name = Prompt.ask(f"键 [{key}] 的列名", default=key).strip()
            target_columns[key] = col_name if col_name else key
        else:
            target_columns[key] = key

    options = get_processing_options(has_keys=True)

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console
    ) as progress:
        task = progress.add_task(
            f"[cyan]正在提取 {len(selected_keys)} 个键值对...",
            total=len(selected_keys)
        )

        for key in selected_keys:
            values = []

            for idx, cell_value in enumerate(df[column_name]):
                fmt = detect_format(cell_value)

                if fmt == DataFormat.JSON:
                    raw_value = extract_value_from_json(cell_value, key)
                elif fmt == DataFormat.KEY_VALUE:
                    if options.get('write_key_value', False):
                        raw_value = extract_key_value_pair(cell_value, key)
                    else:
                        raw_value = extract_value_from_key_value(cell_value, key)
                else:
                    raw_value = ""

                processed = process_value(raw_value, key if not options.get('write_key_value', False) else None, options, fmt)
                values.append(processed)

            target_col = target_columns[key]
            df[target_col] = values
            progress.update(task, advance=1)

    console.print(f"[green]✓ 成功提取 {len(selected_keys)} 个键到新列[/green]")
    return df, True

def delete_keys_operation(df, column_name, all_keys):
    """删除键值对操作"""
    console.print("\n[bold cyan]删除指定键值对[/bold cyan]")
    
    show_keys_table(all_keys)
    
    selected_keys = select_keys(all_keys)
    if selected_keys is None:
        return df, False
    
    if not Confirm.ask(f"\n确认从源列中删除这 {len(selected_keys)} 个键值对?", default=False):
        return df, False
    
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console
    ) as progress:
        task = progress.add_task(
            f"[cyan]正在删除 {len(selected_keys)} 个键值对...",
            total=len(df)
        )
        
        new_values = []
        for cell_value in df[column_name]:
            fmt = detect_format(cell_value)
            current_value = cell_value
            
            for key in selected_keys:
                if fmt == DataFormat.JSON:
                    current_value = remove_key_from_json(current_value, key)
                elif fmt == DataFormat.KEY_VALUE:
                    current_value = remove_key_from_key_value(current_value, key)
            
            new_values.append(current_value)
            progress.update(task, advance=1)
        
        df[column_name] = new_values
    
    console.print(f"[green]✓ 成功删除 {len(selected_keys)} 个键值对[/green]")
    return df, True

def process_plain_text_operation(df, column_name):
    """处理纯文本操作"""
    console.print("\n[bold cyan]处理纯文本列[/bold cyan]")
    
    options = get_processing_options(has_keys=False)
    
    new_column_name = Prompt.ask(
        "请输入新列名",
        default=f"{column_name}_处理后"
    )
    
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console
    ) as progress:
        task = progress.add_task("[cyan]正在处理纯文本...", total=len(df))
        
        processed_values = []
        for value in df[column_name]:
            processed = process_value(value, None, options, DataFormat.PLAIN_TEXT)
            processed_values.append(processed)
            progress.update(task, advance=1)
        
        df[new_column_name] = processed_values
    
    console.print(f"[green]✓ 成功处理纯文本到新列: {new_column_name}[/green]")
    return df, True

def process_column_interactive(df, column_name):
    """交互式处理单个列"""
    format_counts, all_keys = scan_column_format(df, column_name)
    
    show_format_distribution(format_counts)
    
    main_format = max(format_counts.items(), key=lambda x: x[1] if x[0] is not None else 0)[0]
    
    if main_format == DataFormat.PLAIN_TEXT and not all_keys:
        console.print("\n[yellow]检测到主要为纯文本格式,无结构化键值对[/yellow]")
    
    modified = False
    
    while True:
        operation = show_operation_menu(column_name)
        
        if operation == "1":
            if not all_keys:
                console.print("[yellow]当前列没有检测到键值对[/yellow]")
                continue
            
            df, success = extract_keys_operation(df, column_name, all_keys, format_counts)
            if success:
                modified = True
                format_counts, all_keys = scan_column_format(df, column_name)
        
        elif operation == "2":
            if not all_keys:
                console.print("[yellow]当前列没有检测到键值对[/yellow]")
                continue
            
            df, success = delete_keys_operation(df, column_name, all_keys)
            if success:
                modified = True
                format_counts, all_keys = scan_column_format(df, column_name)
        
        elif operation == "3":
            df, success = process_plain_text_operation(df, column_name)
            if success:
                modified = True
        
        elif operation == "4":
            return df, modified, "switch_column"
        
        elif operation == "5":
            return df, modified, "save"
        
        elif operation == "6":
            return df, modified, "cancel"

def process_excel_interactive(file_path):
    """交互式处理Excel文件"""
    try:
        console.print(f"[cyan]正在读取文件: {file_path}[/cyan]")
        df = pd.read_excel(file_path, dtype=str)
        original_df = deepcopy(df)
        console.print(f"[green]✓[/green] 成功读取 {len(df)} 行数据")
        
        current_column = show_columns_menu(df)
        if current_column is None:
            return False, "用户取消操作", 0
        
        console.print(f"[cyan]处理列: {current_column}[/cyan]\n")
        
        global_modified = False
        
        while True:
            df, modified, action = process_column_interactive(df, current_column)
            
            if modified:
                global_modified = True
            
            if action == "switch_column":
                new_column = show_columns_menu(df, current_column)
                if new_column is None:
                    continue
                current_column = new_column
                console.print(f"[cyan]切换到列: {current_column}[/cyan]\n")
            
            elif action == "save":
                if not global_modified:
                    console.print("[yellow]没有进行任何修改[/yellow]")
                    if not Confirm.ask("确认退出?", default=True):
                        continue
                    return False, "未进行修改", 0
                
                output_file = file_path.replace(".xlsx", "_处理后.xlsx")
                
                if os.path.exists(output_file):
                    if not Confirm.ask(f"文件 {output_file} 已存在,是否覆盖?", default=True):
                        new_name = Prompt.ask("请输入新的文件名")
                        if not new_name.endswith('.xlsx'):
                            new_name += '.xlsx'
                        output_file = new_name
                
                console.print("\n[cyan]正在保存文件...[/cyan]")

                # 先在 DataFrame 层做一次统一清洗，避免 openpyxl 在写入阶段触发 IllegalCharacterError
                df_to_save = df.copy()
                df_to_save = df_to_save.apply(lambda col: col.map(sanitize_excel_cell_value))

                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    df_to_save.to_excel(writer, index=False, sheet_name='Sheet1')

                    worksheet = writer.sheets['Sheet1']
                    from openpyxl.styles import numbers

                    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row,
                                                  min_col=1, max_col=worksheet.max_column):
                        for cell in row:
                            cell.number_format = numbers.FORMAT_TEXT
                            if cell.value is not None:
                                # 兜底再清洗一次，确保不会因极端脏数据导致保存失败
                                cell.value = sanitize_excel_cell_value(cell.value)
                
                return True, output_file, len(df)
            
            elif action == "cancel":
                if global_modified:
                    if not Confirm.ask("确认放弃所有修改?", default=False):
                        continue
                
                return False, "用户取消操作", 0
    
    except Exception as e:
        import traceback
        return False, f"{str(e)}\n{traceback.format_exc()}", 0

# ────────────────────────────────────────────────────────────────────────────────
# 通用工具函数（用于新功能）
# ────────────────────────────────────────────────────────────────────────────────

def index_to_col_letter(index):
    """将0-based索引转换为列字母（如0->A，25->Z，26->AA）"""
    result = ""
    index += 1
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(ord('A') + remainder) + result
    return result


def col_letter_to_index(col_letter):
    """将列字母转换为0-based索引"""
    col_letter = col_letter.upper()
    index = 0
    for char in col_letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1


def is_valid_col_letter(s):
    """判断字符串是否为合法的列字母（仅含A-Z）"""
    return len(s) > 0 and all('A' <= c <= 'Z' for c in s.upper())


def get_file_size_str(filepath):
    """返回人类可读的文件大小字符串"""
    size = os.path.getsize(filepath)
    if size < 1024:
        return f"{size} B"
    elif size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    else:
        return f"{size / 1024 / 1024:.1f} MB"


def get_xlsx_files():
    """获取当前工作目录下的所有xlsx文件"""
    current_dir = Path.cwd()
    xlsx_files = sorted([f for f in current_dir.glob('*.xlsx') if not f.name.startswith('~$')])
    return xlsx_files, str(current_dir)


def wrap_text(text, max_width):
    """智能文本换行，支持长文件名"""
    if len(text) <= max_width:
        return text
    
    # 对于长文件名，智能截断并显示省略号
    if len(text) > max_width:
        # 保留前后部分，中间用省略号
        half = (max_width - 3) // 2
        return text[:half] + "..." + text[-(max_width - half - 3):]
    return text


def get_safe_filename(filepath_or_name):
    """获取安全的文件名显示（只显示文件名，不包含路径）"""
    if isinstance(filepath_or_name, Path):
        return filepath_or_name.name
    elif isinstance(filepath_or_name, str):
        return os.path.basename(filepath_or_name)
    return str(filepath_or_name)


def display_file_list_tui(xlsx_files, script_dir, title="📂 当前文件夹中的 Excel 文件"):
    """以表格形式展示 xlsx 文件列表，优化长文件名显示"""
    # 获取终端宽度，默认80
    try:
        import shutil
        terminal_width = shutil.get_terminal_size().columns
    except:
        terminal_width = 80
    
    # 计算文件名列的宽度（留出编号和大小的空间）
    col_widths = {
        '编号': 6,
        '大小': 12,
        '间隔': 4  # 列之间的间隔
    }
    filename_width = terminal_width - col_widths['编号'] - col_widths['大小'] - col_widths['间隔']
    filename_width = max(filename_width, 30)  # 至少30个字符宽
    
    table = Table(
        title=title,
        show_header=True,
        header_style="bold cyan",
        box=box.ROUNDED,
        show_lines=False,
        padding=(0, 1)
    )
    table.add_column("编号", style="bold yellow", justify="center", width=6)
    table.add_column("文件名", style="white", width=filename_width, overflow="fold")
    table.add_column("文件大小", style="dim", justify="right", width=10)

    for idx, file_item in enumerate(xlsx_files, start=1):
        # 统一获取文件名
        filename_str = get_safe_filename(file_item)
        
        # 构建完整路径用于计算文件大小
        filepath = os.path.join(script_dir, filename_str)
        
        # 如果文件名过长，进行智能处理
        if len(filename_str) > filename_width:
            display_name = wrap_text(filename_str, filename_width)
        else:
            display_name = filename_str
            
        table.add_row(str(idx), display_name, get_file_size_str(filepath))

    console.print(table)
    
    # 如果有长文件名被截断，显示提示
    long_files_count = sum(1 for f in xlsx_files if len(get_safe_filename(f)) > filename_width)
    if long_files_count > 0:
        console.print(f"\n[dim]💡 提示: 共有 {long_files_count} 个文件名超过显示宽度，已自动优化显示[/dim]")


def select_file_tui(xlsx_files, prompt_text="请输入要处理的文件编号"):
    """通用文件选择，返回 0-based 索引"""
    while True:
        user_input = Prompt.ask(f"\n[bold green]{prompt_text}[/bold green]").strip()

        if not user_input.isdigit():
            console.print("[bold red]❌ 请输入有效的数字编号！[/bold red]")
            continue

        idx = int(user_input)
        if idx < 1 or idx > len(xlsx_files):
            console.print(
                f"[bold red]❌ 编号超出范围，请输入 1 到 {len(xlsx_files)} 之间的数字！[/bold red]"
            )
            continue

        return idx - 1


def select_column_interactively(ws, prompt_title="请输入列号（如 A、B、P、AA 等）"):
    """展示工作表列预览，提示用户输入列字母，返回 (col_letter, col_index_0based)"""
    max_col = ws.max_column

    # 读取表头行
    headers = []
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col_idx)
        headers.append(str(cell.value) if cell.value is not None else "")

    # 展示列预览（最多 52 列）
    col_table = Table(title="📋 文件列信息预览", show_header=True, header_style="bold cyan", box=box.ROUNDED)
    col_table.add_column("列字母", style="bold yellow", justify="center")
    col_table.add_column("列名（表头）", style="white")

    preview_count = min(max_col, 52)
    for i in range(preview_count):
        col_table.add_row(index_to_col_letter(i), headers[i])
    if max_col > preview_count:
        col_table.add_row("...", f"（共 {max_col} 列）")

    console.print(col_table)

    while True:
        user_input = Prompt.ask(f"\n[bold green]{prompt_title}[/bold green]").strip()

        if not user_input:
            console.print("[bold red]❌ 输入不能为空，请重新输入！[/bold red]")
            continue

        # 自动处理小写
        if user_input != user_input.upper():
            console.print(
                f"[bold yellow]⚠️  检测到小写字母，已自动转换为大写："
                f"[bold white]{user_input.upper()}[/bold white][/bold yellow]"
            )
            user_input = user_input.upper()

        # 非法字符检测
        if not is_valid_col_letter(user_input):
            console.print(
                "[bold red]❌ 列号包含非法字符，只能输入字母（A-Z），请重新输入！[/bold red]"
            )
            continue

        col_index = col_letter_to_index(user_input)

        # 超出范围检测
        if col_index >= max_col:
            console.print(
                f"[bold red]❌ 列号 {user_input} 超出文件范围"
                f"（共 {max_col} 列，最大列为 {index_to_col_letter(max_col - 1)}），请重新输入！[/bold red]"
            )
            continue

        return user_input, col_index


def select_column_from_df(df, prompt_title="请输入列号（如 A、B、P、AA 等）"):
    """从DataFrame展示列预览，提示用户输入列字母，返回 (col_letter, col_index_0based)"""
    total_cols = len(df.columns)

    # 展示列预览（最多 52 列）
    col_table = Table(title="📋 文件列信息预览", show_header=True, header_style="bold cyan", box=box.ROUNDED)
    col_table.add_column("列字母", style="bold yellow", justify="center")
    col_table.add_column("列名（表头）", style="white")

    preview_count = min(total_cols, 52)
    for i in range(preview_count):
        col_table.add_row(index_to_col_letter(i), str(df.columns[i]))
    if total_cols > preview_count:
        col_table.add_row("...", f"（共 {total_cols} 列）")

    console.print(col_table)

    while True:
        user_input = Prompt.ask(f"\n[bold green]{prompt_title}[/bold green]").strip()

        if not user_input:
            console.print("[bold red]❌ 输入不能为空，请重新输入！[/bold red]")
            continue

        # 自动处理小写
        if user_input != user_input.upper():
            console.print(
                f"[bold yellow]⚠️  检测到小写字母，已自动转换为大写："
                f"[bold white]{user_input.upper()}[/bold white][/bold yellow]"
            )
            user_input = user_input.upper()

        # 非法字符检测
        if not is_valid_col_letter(user_input):
            console.print(
                "[bold red]❌ 列号包含非法字符，只能输入字母（A-Z），请重新输入！[/bold red]"
            )
            continue

        col_index = col_letter_to_index(user_input)

        # 超出范围检测
        if col_index >= total_cols:
            console.print(
                f"[bold red]❌ 列号 {user_input} 超出文件范围"
                f"（共 {total_cols} 列，最大列为 {index_to_col_letter(total_cols - 1)}），请重新输入！[/bold red]"
            )
            continue

        return user_input, col_index


# ────────────────────────────────────────────────────────────────────────────────
# 功能1：去除xlsx指定列的重复项
# ────────────────────────────────────────────────────────────────────────────────

def deduplicate_by_column():
    """根据指定列去重，保留首次出现的行"""
    console.print(Panel.fit(
        "[bold cyan]🔁 Excel 去重处理工具[/bold cyan]\n[dim]根据指定列去重，保留首次出现的行[/dim]",
        border_style="cyan",
        padding=(1, 4)
    ))
    console.print()  # 空行美化

    xlsx_files, script_dir = get_xlsx_files()

    if not xlsx_files:
        console.print(Panel(
            "[bold red]❌ 当前文件夹中没有找到任何 .xlsx 文件！\n请将需要处理的 Excel 文件放在本程序所在文件夹中。[/bold red]",
            border_style="red"
        ))
        return False

    # 显示文件列表
    display_file_list_tui(xlsx_files, script_dir)

    # 用户选择文件
    file_idx = select_file_tui(xlsx_files)
    selected_file = get_safe_filename(xlsx_files[file_idx])
    filepath = os.path.join(script_dir, selected_file)

    console.print(f"\n[bold green]✅ 已选择文件：[bold white]{selected_file}[/bold white][/bold green]")

    # 读取文件以获取列信息
    with console.status("[bold cyan]正在读取文件列信息...[/bold cyan]", spinner="dots"):
        df_preview = pd.read_excel(filepath, dtype=str, nrows=0)  # 只读表头

    # 用户选择列
    col_letter, col_index = select_column_from_df(df_preview)

    # 确认操作
    console.print()
    console.print(Panel(
        f"[bold]操作确认[/bold]\n\n"
        f"  📄 处理文件：[cyan]{selected_file}[/cyan]\n"
        f"  🔑 去重列：  [cyan]{col_letter} 列[/cyan]（{df_preview.columns[col_index]}）\n"
        f"  📁 输出目录：[cyan]{script_dir}[/cyan]",
        border_style="yellow",
        title="[yellow]请确认[/yellow]"
    ))

    confirm = Prompt.ask(
        "[bold yellow]确认开始处理？[/bold yellow]",
        choices=["y", "n"],
        default="y"
    )

    if confirm.lower() != 'y':
        console.print("[dim]已取消操作。[/dim]")
        return False

    # 执行处理
    console.print()
    with console.status(f"[bold cyan]正在读取文件...[/bold cyan]", spinner="dots"):
        df = pd.read_excel(filepath, dtype=str)

    original_count = len(df)
    console.print(f"[green]✅ 文件读取完成，共 [bold]{original_count}[/bold] 行数据[/green]")

    with console.status(f"[bold cyan]正在根据 {col_letter} 列去重...[/bold cyan]", spinner="dots"):
        df_dedup = df.drop_duplicates(subset=df.columns[col_index], keep='first')
        df_dedup = df_dedup.reset_index(drop=True)

    dedup_count = len(df_dedup)
    removed_count = original_count - dedup_count

    # 生成输出文件名
    base_name = os.path.splitext(selected_file)[0]
    output_filename = f"{base_name}_去重_{col_letter}列.xlsx"
    output_path = os.path.join(script_dir, output_filename)

    with console.status(f"[bold cyan]正在保存文件...[/bold cyan]", spinner="dots"):
        # 清理控制字符
        df_to_save = df_dedup.apply(lambda col: col.map(sanitize_excel_cell_value))
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_to_save.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.number_format = '@'

    # 显示结果
    result_table = Table(title="✅ 处理完成", show_header=True, header_style="bold green", box=box.ROUNDED)
    result_table.add_column("项目", style="bold")
    result_table.add_column("数值", style="cyan", justify="right")

    result_table.add_row("原始行数", str(original_count))
    result_table.add_row("去重后行数", str(dedup_count))
    result_table.add_row("删除重复行数", f"[red]{removed_count}[/red]")
    result_table.add_row("输出文件名", f"[green]{output_filename}[/green]")

    console.print(result_table)
    return True


# ────────────────────────────────────────────────────────────────────────────────
# 功能2：筛选交集（从数据文件筛选匹配行，顺序与筛选文件一致）
# ────────────────────────────────────────────────────────────────────────────────

def read_b_column(filepath, col_index_0based):
    """
    读取筛选文件指定列的所有值（从第 2 行起，跳过表头）。
    返回:
        b_order : list[str]  — 按行顺序排列的值（保留重复项用于定序，后续会去重）
        b_values: set[str]   — 用于快速查找的集合
    """
    col_1based = col_index_0based + 1
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active

    b_order = []   # 保留顺序（含重复）
    b_values = set()

    with Progress(
        SpinnerColumn(),
        TextColumn("[cyan]正在读取筛选文件...[/cyan]"),
        transient=True,
    ) as progress:
        progress.add_task("read_b", total=None)
        for row in ws.iter_rows(min_row=2, min_col=col_1based, max_col=col_1based):
            cell = row[0]
            if cell.value is not None:
                val = str(cell.value).strip()
                b_order.append(val)   # 记录原始顺序（含重复）
                b_values.add(val)

    wb.close()
    return b_order, b_values


def filter_and_write(file_a_path, col_a_index, b_order, b_values, output_path):
    """
    读取数据文件，将匹配行按照筛选文件的值顺序写入输出文件。
    返回: (original_rows, matched_rows)
    """

    # ── Step 1：将数据文件所有行读入内存，构建 key → [行数据, ...] 的映射 ──
    wb_a = openpyxl.load_workbook(file_a_path, read_only=True, data_only=True)
    ws_a = wb_a.active
    max_row = ws_a.max_row
    max_col = ws_a.max_column

    header_row = None
    key_to_rows: dict[str, list] = {}

    with Progress(
        SpinnerColumn(),
        TextColumn("[cyan]正在读取数据文件...[/cyan]"),
        BarColumn(),
        TaskProgressColumn(),
        transient=True,
    ) as progress:
        task = progress.add_task("read_a", total=max_row)
        for row in ws_a.iter_rows():
            row_num = row[0].row

            # 表头单独保存
            if row_num == 1:
                header_row = [
                    (str(cell.value).strip() if cell.value is not None else None)
                    for cell in row
                ]
                progress.advance(task)
                continue

            key_cell = row[col_a_index]
            if key_cell.value is None:
                progress.advance(task)
                continue

            key = str(key_cell.value).strip()

            # 只缓存在 b_values 中存在的行，节省内存
            if key in b_values:
                row_data = [
                    (str(cell.value).strip() if cell.value is not None else None)
                    for cell in row
                ]
                if key not in key_to_rows:
                    key_to_rows[key] = []
                key_to_rows[key].append(row_data)

            progress.advance(task)

    wb_a.close()

    original_rows = max_row - 1  # 不含表头

    # ── Step 2：按 b_order 的顺序拼接最终输出行列表 ──
    seen_keys = set()
    ordered_rows = []

    for val in b_order:
        if val in seen_keys:
            continue
        seen_keys.add(val)
        if val in key_to_rows:
            ordered_rows.extend(key_to_rows[val])   # 同 key 多行，组内保持原顺序

    matched_rows = len(ordered_rows)

    # ── Step 3：写入输出文件 ──
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active

    all_rows = [header_row] + ordered_rows  # 表头 + 数据

    with Progress(
        SpinnerColumn(),
        TextColumn("[cyan]正在写入输出文件...[/cyan]"),
        BarColumn(),
        TaskProgressColumn(),
        transient=True,
    ) as progress:
        task = progress.add_task("write", total=len(all_rows))
        for new_row_idx, row_data in enumerate(all_rows, start=1):
            for col_idx, value in enumerate(row_data, start=1):
                dst_cell = ws_out.cell(row=new_row_idx, column=col_idx)
                dst_cell.value = value
                dst_cell.number_format = openpyxl_numbers.FORMAT_TEXT
            progress.advance(task)

    wb_out.save(output_path)
    wb_out.close()

    return original_rows, matched_rows


def filter_intersection():
    """从数据文件中筛选匹配行，输出顺序与筛选文件保持一致"""
    console.print(Panel.fit(
        "[bold cyan]🔗 Excel 行筛选工具[/bold cyan]\n"
        "[dim]从数据文件中筛选匹配行，输出顺序与筛选文件保持一致[/dim]",
        border_style="cyan",
        padding=(1, 4),
    ))
    console.print()  # 空行美化

    xlsx_files, script_dir = get_xlsx_files()

    if not xlsx_files:
        console.print(Panel(
            "[bold red]❌ 当前文件夹中没有找到任何 .xlsx 文件！\n"
            "请将需要处理的 Excel 文件放在本程序所在文件夹中。[/bold red]",
            border_style="red",
        ))
        return False

    # ── 选择数据文件（A 文件）──
    console.rule("[bold cyan]第一步：选择数据文件（被筛选的文件）[/bold cyan]")
    display_file_list_tui(xlsx_files, script_dir)
    file_a_idx = select_file_tui(xlsx_files, prompt_text="请输入数据文件的编号")
    file_a_name = get_safe_filename(xlsx_files[file_a_idx])
    file_a_path = os.path.join(script_dir, file_a_name)
    console.print(f"[green]✅ 已选择数据文件：[bold white]{file_a_name}[/bold white][/green]")

    # ── 选择筛选文件（B 文件）──
    console.print()
    console.rule("[bold cyan]第二步：选择筛选文件（提供匹配值的文件）[/bold cyan]")
    display_file_list_tui(xlsx_files, script_dir)
    file_b_idx = select_file_tui(xlsx_files, prompt_text="请输入筛选文件的编号")
    file_b_name = get_safe_filename(xlsx_files[file_b_idx])
    file_b_path = os.path.join(script_dir, file_b_name)
    console.print(f"[green]✅ 已选择筛选文件：[bold white]{file_b_name}[/bold white][/green]")

    # ── 选择数据文件的匹配列 ──
    console.print()
    console.rule("[bold cyan]第三步：选择数据文件中用于匹配的列[/bold cyan]")
    with console.status("[cyan]正在读取数据文件列信息...[/cyan]", spinner="dots"):
        wb_a_tmp = openpyxl.load_workbook(file_a_path, read_only=True, data_only=True)
        ws_a_tmp = wb_a_tmp.active

    col_a_letter, col_a_index = select_column_interactively(
        ws_a_tmp,
        prompt_title="请输入数据文件中用于匹配的列号（如 A、B、P、AA 等）",
    )
    wb_a_tmp.close()
    console.print(f"[green]✅ 已选择数据文件匹配列：[bold white]{col_a_letter} 列[/bold white][/green]")

    # ── 选择筛选文件的匹配列 ──
    console.print()
    console.rule("[bold cyan]第四步：选择筛选文件中提供匹配值的列[/bold cyan]")
    with console.status("[cyan]正在读取筛选文件列信息...[/cyan]", spinner="dots"):
        wb_b_tmp = openpyxl.load_workbook(file_b_path, read_only=True, data_only=True)
        ws_b_tmp = wb_b_tmp.active

    col_b_letter, col_b_index = select_column_interactively(
        ws_b_tmp,
        prompt_title="请输入筛选文件中提供匹配值的列号（如 A、B、P、AA 等）",
    )
    wb_b_tmp.close()
    console.print(f"[green]✅ 已选择筛选文件匹配列：[bold white]{col_b_letter} 列[/bold white][/green]")

    # ── 生成输出文件名 ──
    base_name = os.path.splitext(file_a_name)[0]
    output_filename = f"{base_name}_筛选结果.xlsx"
    output_path = os.path.join(script_dir, output_filename)

    # ── 确认操作 ──
    console.print()
    console.print(Panel(
        f"[bold]操作确认[/bold]\n\n"
        f"  📄 数据文件：  [cyan]{file_a_name}[/cyan]（匹配列：{col_a_letter} 列）\n"
        f"  🔑 筛选文件：  [cyan]{file_b_name}[/cyan]（取值列：{col_b_letter} 列）\n"
        f"  🔀 输出顺序：  与筛选文件 {col_b_letter} 列顺序一致\n"
        f"  💾 输出文件：  [cyan]{output_filename}[/cyan]\n"
        f"  📁 输出目录：  [cyan]{script_dir}[/cyan]",
        border_style="yellow",
        title="[yellow]请确认[/yellow]",
    ))

    confirm = Prompt.ask(
        "[bold yellow]确认开始处理？[/bold yellow]",
        choices=["y", "n"],
        default="y",
    )
    if confirm.lower() != "y":
        console.print("[dim]已取消操作。[/dim]")
        return False

    # ── 执行处理 ──
    console.print()
    console.rule("[bold cyan]处理中[/bold cyan]")

    b_order, b_values = read_b_column(file_b_path, col_b_index)
    console.print(
        f"[green]✅ 筛选文件 {col_b_letter} 列共读取到 "
        f"[bold]{len(b_order)}[/bold] 个值（[bold]{len(b_values)}[/bold] 个唯一值）[/green]"
    )

    original_rows, matched_rows = filter_and_write(
        file_a_path, col_a_index, b_order, b_values, output_path
    )

    # ── 展示结果 ──
    console.print()
    result_table = Table(title="✅ 处理完成", show_header=True, header_style="bold green", box=box.ROUNDED)
    result_table.add_column("项目", style="bold")
    result_table.add_column("数值", style="cyan", justify="right")

    result_table.add_row("数据文件总行数（不含表头）", str(original_rows))
    result_table.add_row("筛选文件唯一值数量", str(len(b_values)))
    result_table.add_row("命中并保留的行数", f"[green]{matched_rows}[/green]")
    result_table.add_row("未命中丢弃的行数", f"[red]{original_rows - matched_rows}[/red]")
    result_table.add_row("输出顺序", f"与筛选文件 {col_b_letter} 列一致")
    result_table.add_row("输出文件名", f"[green]{output_filename}[/green]")

    console.print(result_table)
    return True


def show_banner():
    """显示欢迎横幅"""
    banner = """
    ╔═══════════════════════════════════════════════════════╗
    ║                                                       ║
    ║     📊 Excel 通用数据处理工具                         ║
    ║                                                       ║
    ║     支持 JSON / 键值对 / 纯文本 格式                  ║
    ║     多列处理 - 交互式多步骤操作                       ║
    ║     去重 / 筛选交集 功能                              ║
    ║                                                       ║
    ╚═══════════════════════════════════════════════════════╝
    """
    console.print(banner, style="bold cyan")

def list_excel_files():
    """列出当前工作目录下的Excel文件"""
    # 使用当前工作目录而不是脚本所在目录
    current_dir = Path.cwd()
    excel_files = list(current_dir.glob('*.xlsx'))
    excel_files = [f for f in excel_files if not f.name.startswith('~$')]
    return sorted(excel_files, key=lambda x: x.stat().st_mtime, reverse=True)

def show_file_list(files):
    """显示文件列表"""
    if not files:
        console.print("[yellow]⚠ 当前目录下没有找到Excel文件[/yellow]")
        console.print(f"[dim]当前工作目录: {Path.cwd()}[/dim]")
        return

    # 获取终端宽度
    try:
        import shutil
        terminal_width = shutil.get_terminal_size().columns
    except:
        terminal_width = 80
    
    # 计算文件名列的宽度
    col_widths = {'序号': 8, '大小': 12, '修改时间': 20, '间隔': 6}
    filename_width = terminal_width - sum(col_widths.values())
    filename_width = max(filename_width, 30)

    table = Table(title="可用的Excel文件", box=box.ROUNDED)
    table.add_column("序号", style="cyan", justify="center", width=8)
    table.add_column("文件名", style="green", width=filename_width, overflow="fold")
    table.add_column("大小", style="yellow", justify="right", width=12)
    table.add_column("修改时间", style="magenta", width=20)

    for idx, file in enumerate(files, 1):
        size = file.stat().st_size
        size_str = f"{size / 1024:.2f} KB" if size < 1024 * 1024 else f"{size / (1024 * 1024):.2f} MB"

        import datetime
        mtime = datetime.datetime.fromtimestamp(file.stat().st_mtime)
        mtime_str = mtime.strftime("%Y-%m-%d %H:%M:%S")

        # 只显示文件名
        display_name = get_safe_filename(file)
        if len(display_name) > filename_width:
            display_name = wrap_text(display_name, filename_width)

        table.add_row(str(idx), display_name, size_str, mtime_str)

    console.print(table)

def main():
    """主函数"""
    console.clear()
    show_banner()

    # 显示当前工作目录
    console.print(f"[dim]📁 当前工作目录: {Path.cwd()}[/dim]\n")

    # ── 功能选择菜单 ──
    console.print()
    menu_table = Table(box=box.ROUNDED, show_header=False, padding=(0, 2))
    menu_table.add_column("选项", style="cyan", width=6)
    menu_table.add_column("说明", style="white")
    menu_table.add_row("1.", "Excel 通用数据提取（JSON/键值对/纯文本）")
    menu_table.add_row("2.", "去除xlsx指定列的重复项")
    menu_table.add_row("3.", "筛选交集（按另一文件顺序筛选）")
    menu_table.add_row("4.", "退出程序")
    console.print(menu_table)

    func_choice = Prompt.ask("\n[bold green]请输入选项 [1/2/3/4][/bold green]", choices=["1", "2", "3", "4"], default="1")

    if func_choice == "4":
        console.print("\n[yellow]👋 感谢使用,再见![/yellow]")
        return

    if func_choice == "2":
        # 去重功能
        console.clear()
        deduplicate_by_column()
        # 完成后询问是否返回主菜单
        if Confirm.ask("\n是否返回主菜单?", default=True):
            console.clear()
            main()
            return
        return

    if func_choice == "3":
        # 筛选交集
        console.clear()
        filter_intersection()
        # 完成后询问是否返回主菜单
        if Confirm.ask("\n是否返回主菜单?", default=True):
            console.clear()
            main()
            return
        return

    # 默认进入原有的通用数据提取功能
    console.print("\n[bold cyan]✨ 已进入 Excel 通用数据提取功能[/bold cyan]")
    while True:
        console.print("\n[bold cyan]╔═══════════════════════════════════════╗[/bold cyan]")
        console.print("[bold cyan]║[/bold cyan] [bold]文件选择菜单[/bold]                          [bold cyan]║[/bold cyan]")
        console.print("[bold cyan]╠═══════════════════════════════════════╣[/bold cyan]")
        console.print("[bold cyan]║[/bold cyan] [cyan]1.[/cyan] 选择文件编号                      [bold cyan]║[/bold cyan]")
        console.print("[bold cyan]║[/bold cyan] [cyan]2.[/cyan] 手动输入文件路径                  [bold cyan]║[/bold cyan]")
        console.print("[bold cyan]║[/bold cyan] [cyan]3.[/cyan] 刷新文件列表                      [bold cyan]║[/bold cyan]")
        console.print("[bold cyan]║[/bold cyan] [cyan]4.[/cyan] 返回主菜单                        [bold cyan]║[/bold cyan]")
        console.print("[bold cyan]╚═══════════════════════════════════════╝[/bold cyan]")

        # 显示文件列表
        excel_files = list_excel_files()
        show_file_list(excel_files)

        choice = Prompt.ask("\n[bold green]请输入选项 [1/2/3/4][/bold green]", choices=["1", "2", "3", "4"], default="1")

        if choice == "4":
            console.print("\n[yellow]返回主菜单[/yellow]")
            return

        if choice == "3":
            console.clear()
            show_banner()
            console.print(f"[dim]当前工作目录: {Path.cwd()}[/dim]")
            continue

        file_path = None

        if choice == "1":
            if not excel_files:
                console.print("[red]✗ 没有可选择的文件[/red]")
                continue

            file_idx = Prompt.ask(
                f"请输入文件编号 (1-{len(excel_files)})",
                default="1"
            )

            try:
                idx = int(file_idx) - 1
                if 0 <= idx < len(excel_files):
                    file_path = str(excel_files[idx])
                else:
                    console.print("[red]✗ 无效的文件编号[/red]")
                    continue
            except ValueError:
                console.print("[red]✗ 请输入有效的数字[/red]")
                continue

        elif choice == "2":
            file_path = Prompt.ask("请输入文件路径")
            if not os.path.isabs(file_path):
                file_path = str(Path.cwd() / file_path)
            if not os.path.exists(file_path):
                console.print("[red]✗ 文件不存在[/red]")
                continue

        console.print(f"\n[bold cyan]开始处理文件...[/bold cyan]")
        success, result, row_count = process_excel_interactive(file_path)
        
        if success:
            panel = Panel(
                f"[green]✓ 处理成功![/green]\n\n"
                f"📁 输出文件: [cyan]{result}[/cyan]\n"
                f"📊 处理行数: [yellow]{row_count}[/yellow] 行",
                title="[bold green]处理完成[/bold green]",
                border_style="green"
            )
            console.print(panel)
        else:
            if result != "用户取消操作" and result != "未进行修改":
                panel = Panel(
                    f"[red]✗ 处理失败![/red]\n\n"
                    f"错误信息: {result}",
                    title="[bold red]错误[/bold red]",
                    border_style="red"
                )
                console.print(panel)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        console.print("\n\n[yellow]程序已被用户中断[/yellow]")
    except Exception as e:
        console.print(f"\n[red]发生错误: {e}[/red]")
        import traceback
        console.print(traceback.format_exc())