import pandas as pd
import re
import json
from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.prompt import Prompt, Confirm
from rich.table import Table
from rich import box
import os
from pathlib import Path
from copy import deepcopy

console = Console()

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
            
            df[key] = values
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
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                    worksheet = writer.sheets['Sheet1']
                    from openpyxl.styles import numbers
                    
                    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, 
                                                  min_col=1, max_col=worksheet.max_column):
                        for cell in row:
                            cell.number_format = numbers.FORMAT_TEXT
                            if cell.value is not None:
                                cell.value = str(cell.value)
                
                return True, output_file, len(df)
            
            elif action == "cancel":
                if global_modified:
                    if not Confirm.ask("确认放弃所有修改?", default=False):
                        continue
                
                return False, "用户取消操作", 0
    
    except Exception as e:
        import traceback
        return False, f"{str(e)}\n{traceback.format_exc()}", 0

def show_banner():
    """显示欢迎横幅"""
    banner = """
    ╔═══════════════════════════════════════════════════════╗
    ║                                                       ║
    ║     📊 Excel 通用数据提取工具                         ║
    ║                                                       ║
    ║     支持 JSON / 键值对 / 纯文本 格式                  ║
    ║     多列处理 - 交互式多步骤操作                       ║
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
    
    table = Table(title=f"可用的Excel文件 - {Path.cwd()}", box=box.ROUNDED)
    table.add_column("序号", style="cyan", justify="center", width=8)
    table.add_column("文件名", style="green")
    table.add_column("大小", style="yellow", justify="right", width=12)
    table.add_column("修改时间", style="magenta", width=20)
    
    for idx, file in enumerate(files, 1):
        size = file.stat().st_size
        size_str = f"{size / 1024:.2f} KB" if size < 1024 * 1024 else f"{size / (1024 * 1024):.2f} MB"
        
        import datetime
        mtime = datetime.datetime.fromtimestamp(file.stat().st_mtime)
        mtime_str = mtime.strftime("%Y-%m-%d %H:%M:%S")
        
        table.add_row(str(idx), file.name, size_str, mtime_str)
    
    console.print(table)

def main():
    """主函数"""
    console.clear()
    show_banner()
    
    # 显示当前工作目录
    console.print(f"[dim]当前工作目录: {Path.cwd()}[/dim]")
    
    while True:
        console.print("\n[bold cyan]═══════════════════════════════════════[/bold cyan]")
        
        excel_files = list_excel_files()
        show_file_list(excel_files)
        
        console.print("\n[bold]请选择操作:[/bold]")
        console.print("  [cyan]1.[/cyan] 选择文件编号")
        console.print("  [cyan]2.[/cyan] 手动输入文件路径")
        console.print("  [cyan]3.[/cyan] 刷新文件列表")
        console.print("  [cyan]4.[/cyan] 退出程序")
        
        choice = Prompt.ask("\n请输入选项", choices=["1", "2", "3", "4"], default="1")
        
        if choice == "4":
            console.print("\n[yellow]👋 感谢使用,再见![/yellow]")
            break
        
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