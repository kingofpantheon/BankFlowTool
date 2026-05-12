# -*- coding: utf-8 -*-
"""
银行流水梳理工具 V1.6.8
修复：区块ID块颜色绿蓝交替，跨文件汇总表同样颜色处理；删除说明区重复版本号；优化打包命名体验
功能：清理不可见字符，使金额列数据可被正常统计
"""

import pandas as pd
import re
import os
import threading
import tkinter as tk
from tkinter import messagebox, filedialog, BooleanVar, Entry, Label, Frame, Checkbutton, Button
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter
import tempfile
import datetime
import random
import json
import string

VERSION = "V1.6.8"

# ---------- 全局日志收集器 ----------
log_entries = []

def log(msg):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_msg = f"[{timestamp}] {msg}"
    print(log_msg)
    log_entries.append(log_msg)

def export_log():
    if not log_entries:
        messagebox.showinfo("提示", "暂无日志记录。")
        return
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
        title="保存日志文件"
    )
    if file_path:
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(f"银行流水梳理工具 {VERSION} 运行日志\n")
                f.write(f"导出时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 80 + "\n")
                f.write("\n".join(log_entries))
            messagebox.showinfo("完成", f"日志已保存至：{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存日志失败：{e}")

# ---------- 脱敏映射管理 ----------
class DesensitizeMapper:
    def __init__(self):
        self.orig_to_mask = {}
        self.mask_to_orig = {}
        self.charset_digit = list(string.digits)
        self.charset_lower = list(string.ascii_lowercase)
        self.charset_upper = list(string.ascii_uppercase)
        self.charset_ch = [chr(i) for i in range(0x4e00, 0x9fff) if chr(i).isalpha()]

    def _random_char(self, ch_type, seed_str):
        h = hash(seed_str) % 2**32
        random.seed(h)
        if ch_type == 'digit':
            result = random.choice(self.charset_digit)
        elif ch_type == 'lower':
            result = random.choice(self.charset_lower)
        elif ch_type == 'upper':
            result = random.choice(self.charset_upper)
        elif ch_type == 'ch':
            result = random.choice(self.charset_ch)
        else:
            result = seed_str
        random.seed()
        return result

    def _generate_masked(self, s):
        masked_chars = []
        for i, ch in enumerate(s):
            seed = f"{s}_{i}"
            if ch.isdigit():
                masked_chars.append(self._random_char('digit', seed))
            elif ch.islower():
                masked_chars.append(self._random_char('lower', seed))
            elif ch.isupper():
                masked_chars.append(self._random_char('upper', seed))
            elif '\u4e00' <= ch <= '\u9fff':
                masked_chars.append(self._random_char('ch', seed))
            else:
                masked_chars.append(ch)
        return ''.join(masked_chars)

    def mask(self, original):
        if original is None:
            return None
        s = str(original)
        if s not in self.orig_to_mask:
            masked = self._generate_masked(s)
            self.orig_to_mask[s] = masked
            self.mask_to_orig[masked] = s
        return self.orig_to_mask[s]

    def unmask(self, masked):
        if masked is None:
            return None
        return self.mask_to_orig.get(str(masked), masked)

    def save_mapping(self, file_path):
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump({
                'orig_to_mask': self.orig_to_mask,
                'mask_to_orig': self.mask_to_orig
            }, f, ensure_ascii=False, indent=2)

    def load_mapping(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self.orig_to_mask = data['orig_to_mask']
        self.mask_to_orig = data['mask_to_orig']

desensitize_mapper = DesensitizeMapper()

# ---------- 辅助函数 ----------
def is_numeric_cell(cell_value):
    if cell_value is None:
        return False
    s = str(cell_value).strip()
    if s == '':
        return False
    if 'e' in s.lower():
        try:
            num = float(s)
            if abs(num) > 1e12:
                return False
        except:
            pass
    cleaned = re.sub(r'[^\d.-]', '', s)
    if cleaned == '' or cleaned == '-':
        return False
    try:
        num = float(cleaned)
        if abs(num) > 1e12:
            return False
        return True
    except ValueError:
        return False

def clean_money_cell(cell_value):
    if cell_value is None:
        return None
    s = str(cell_value).strip()
    if s == '':
        return None
    try:
        return float(s)
    except ValueError:
        cleaned = re.sub(r'[^\d.-]', '', s)
        if cleaned == '' or cleaned == '-':
            return s
        try:
            return float(cleaned)
        except ValueError:
            return s

# ---------- 自动识别表头行 ----------
def detect_header_row(ws, max_row, max_col, max_scan_rows=30):
    header_keywords = [
        '用户ID', '交易单号', '大单号', '账号名称', '借贷类型', '交易业务类型',
        '交易用途类型', '交易时间', '金额', '交易额', '发生额', '交易金额',
        '账户余额', '银行卡号', '网银联单号', '对手', '对方', '对面',
        '第三方账户', '银行名称', '基金公司', '间联', '备注'
    ]
    best_row = 1
    best_score = 0
    scan_limit = min(max_row, max_scan_rows)
    for row in range(1, scan_limit + 1):
        score = 0
        for col in range(1, max_col + 1):
            cell_val = ws.cell(row, col).value
            if cell_val is None:
                continue
            cell_str = str(cell_val).strip()
            if not cell_str:
                continue
            for kw in header_keywords:
                if kw in cell_str:
                    score += 1
                    break
        if score > best_score:
            best_score = score
            best_row = row
        if best_score >= 8:
            break
    log(f"自动识别表头行: 第 {best_row} 行 (关键词命中数: {best_score})")
    return best_row

# ---------- 识别列（基于表头行） ----------
def identify_money_columns_auto(ws, max_row, max_col, header_row):
    candidate_cols = []
    for col in range(1, max_col + 1):
        total = 0
        numeric = 0
        for row in range(header_row + 1, max_row + 1):
            cell = ws.cell(row, col)
            if cell.value is None:
                continue
            total += 1
            if is_numeric_cell(cell.value):
                numeric += 1
        if total > 0 and numeric / total >= 0.3:
            candidate_cols.append(col)
    if not candidate_cols:
        return None
    col_names = {}
    for col in candidate_cols:
        val = ws.cell(header_row, col).value
        if val:
            col_names[col] = str(val).strip()
    money_keywords = ['金额', '交易额', '发生额', '交易金额', '发生金额']
    best_col = None
    best_score = 0
    for col in candidate_cols:
        name = col_names.get(col, '')
        score = 0
        for kw in money_keywords:
            if kw in name:
                score += 10
        if score > best_score:
            best_score = score
            best_col = col
    if best_col is not None:
        return [best_col]
    else:
        return [candidate_cols[0]]

def identify_counterparty_column_auto(ws, max_col, header_row):
    col_names = {}
    for col in range(1, max_col + 1):
        val = ws.cell(header_row, col).value
        if val:
            col_names[col] = str(val).strip()
    relation_keywords = ['对手', '对方', '对面']
    candidates = []
    for col, name in col_names.items():
        if any(kw in name for kw in relation_keywords):
            candidates.append((col, name))
    if not candidates:
        log("未找到包含“对手/对方/对面”的列，对手方识别失败")
        return None
    primary = []
    secondary = []
    for col, name in candidates:
        if '户名' in name or '名称' in name:
            primary.append((col, name))
        else:
            secondary.append((col, name))
    final_candidates = primary if primary else secondary
    if not final_candidates:
        return None
    def chinese_count(s):
        return sum(1 for ch in s if '\u4e00' <= ch <= '\u9fff')
    best = max(final_candidates, key=lambda x: (chinese_count(x[1]), -len(x[1])))
    log(f"识别到对手方名称列: {best[1]} (列索引 {best[0]})")
    return best[0]

def identify_counterparty_id_column_auto(ws, max_col, header_row):
    """自动识别对手方ID列"""
    col_names = {}
    for col in range(1, max_col + 1):
        val = ws.cell(header_row, col).value
        if val:
            col_names[col] = str(val).strip()
    id_keywords = ['对手方ID', '账户ID', '交易方ID', '对方ID', '对手ID']
    best_col = None
    for col, name in col_names.items():
        for kw in id_keywords:
            if kw == name or (kw in name and len(name) - len(kw) <= 2):
                best_col = col
                log(f"识别到对手方ID列: {name} (列索引 {col})")
                return best_col
    log("未找到对手方ID列，将不使用ID分项统计")
    return None

def detect_direction_column_auto(ws, max_col, header_row):
    direction_keywords = ['借贷类型', '交易方向', '借贷', '方向', '借/贷', '出入']
    for col in range(1, max_col + 1):
        val = ws.cell(header_row, col).value
        if val:
            name = str(val).strip()
            for kw in direction_keywords:
                if kw in name:
                    log(f"识别到方向列: {name} (列索引 {col})")
                    return col
    log("未找到方向列，将根据金额正负判断流向")
    return None

def get_direction_from_cell(cell_value):
    if cell_value is None:
        return None
    s = str(cell_value).strip()
    if s in ['入', '贷', '收入', '增加', '+']:
        return 'inflow'
    elif s in ['出', '借', '支出', '减少', '-']:
        return 'outflow'
    else:
        if any(k in s for k in ['入', '贷', '收入']):
            return 'inflow'
        elif any(k in s for k in ['出', '借', '支出']):
            return 'outflow'
    return None

def find_column_by_title(ws, max_col, header_row, title_list):
    col_names = {}
    for col in range(1, max_col + 1):
        val = ws.cell(header_row, col).value
        if val:
            col_names[col] = str(val).strip()
    for user_title in title_list:
        for col, name in col_names.items():
            if name == user_title or user_title in name:
                return col
    return None

# ---------- 区块处理 ----------
def find_blocks(ws, max_row, max_col, header_row):
    key_cols = []
    for col in range(1, min(max_col, 10) + 1):
        if ws.cell(header_row, col).value is not None:
            key_cols.append(col)
    if not key_cols:
        return [(header_row + 1, max_row, 1)]

    header_values = {}
    for col in key_cols:
        val = ws.cell(header_row, col).value
        header_values[col] = str(val).strip() if val is not None else ""

    candidate_rows = []
    for row in range(header_row + 1, max_row + 1):
        has_content = any(ws.cell(row, col).value is not None for col in range(1, max_col+1))
        if not has_content:
            continue
        match = True
        for col in key_cols:
            cell_val = ws.cell(row, col).value
            cur_val = str(cell_val).strip() if cell_val is not None else ""
            if cur_val != header_values[col]:
                match = False
                break
        if match:
            candidate_rows.append(row)

    if not candidate_rows:
        return [(header_row + 1, max_row, 1)]

    blocks = []
    prev = header_row
    for idx, row in enumerate(candidate_rows):
        start = prev + 1
        end = row - 1
        if start <= end:
            blocks.append((start, end, idx + 1))
        prev = row
    if prev < max_row:
        blocks.append((prev + 1, max_row, len(candidate_rows) + 1))
    return blocks

def process_block(ws, wb, block_start, block_end, block_idx, header_row, enable_sum, enable_internal_stats, collect_data,
                  manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                  enable_desensitize, desensitize_cols, enable_money_clean):
    max_col = ws.max_column
    if block_start > block_end:
        return []

    # 确定金额列
    if manual_mode and user_amount_titles:
        amount_col = find_column_by_title(ws, max_col, header_row, user_amount_titles)
        if amount_col is None:
            raise Exception(f"未找到用户指定的金额列标题: {user_amount_titles}")
        money_cols = [amount_col]
        log(f"  使用手动金额列: {user_amount_titles}")
    else:
        money_cols = identify_money_columns_auto(ws, block_end, max_col, header_row)
        if not money_cols:
            log("  未找到金额列，跳过该区块")
            return []
        log(f"  自动识别金额列: {ws.cell(header_row, money_cols[0]).value}")

    # 确定对手方名称列
    if manual_mode and user_counterparty_titles:
        counterparty_col = find_column_by_title(ws, max_col, header_row, user_counterparty_titles)
        if counterparty_col is None:
            raise Exception(f"未找到用户指定的对手方名称列标题: {user_counterparty_titles}")
        log(f"  使用手动对手方名称列: {user_counterparty_titles}")
    else:
        counterparty_col = identify_counterparty_column_auto(ws, max_col, header_row)
        if counterparty_col is None:
            log("  未找到对手方名称列，跳过该区块")
            return []
        log(f"  自动识别对手方名称列: {ws.cell(header_row, counterparty_col).value}")

    # 确定对手方ID列
    counterparty_id_col = None
    if manual_mode and user_counterparty_id_titles:
        counterparty_id_col = find_column_by_title(ws, max_col, header_row, user_counterparty_id_titles)
        if counterparty_id_col is not None:
            log(f"  使用手动对手方ID列: {user_counterparty_id_titles}")
        else:
            log(f"  警告：未找到用户指定的对手方ID列标题: {user_counterparty_id_titles}")
    else:
        counterparty_id_col = identify_counterparty_id_column_auto(ws, max_col, header_row)
        if counterparty_id_col is not None:
            log(f"  自动识别对手方ID列: {ws.cell(header_row, counterparty_id_col).value}")

    direction_col = detect_direction_column_auto(ws, max_col, header_row)

    # 清理金额
    if enable_money_clean:
        for col in money_cols:
            for row in range(block_start, block_end + 1):
                cell = ws.cell(row, col)
                if cell.value is None:
                    continue
                cleaned = clean_money_cell(cell.value)
                if cleaned is not None:
                    cell.value = cleaned

    # 脱敏
    if enable_desensitize and desensitize_cols:
        for col in desensitize_cols:
            if col is None:
                continue
            for row in range(block_start, block_end + 1):
                cell = ws.cell(row, col)
                if cell.value is None:
                    continue
                cell.value = desensitize_mapper.mask(cell.value)

    # 求和行（黄色）—— 查找最后一个有效数据行
    last_data_row = None
    if enable_sum:
        sums = {}
        amount_col = money_cols[0]
        for row in range(block_start, block_end + 1):
            cell = ws.cell(row, amount_col)
            if cell.value is not None:
                if isinstance(cell.value, (int, float)):
                    val = cell.value
                else:
                    val = clean_money_cell(cell.value)
                if isinstance(val, (int, float)):
                    last_data_row = row
        for col in money_cols:
            total = 0.0
            for row in range(block_start, block_end + 1):
                cell = ws.cell(row, col)
                if cell.value is None:
                    continue
                if isinstance(cell.value, (int, float)):
                    val = cell.value
                else:
                    val = clean_money_cell(cell.value)
                if isinstance(val, (int, float)):
                    total += abs(val)
            if total != 0:
                sums[col] = total
        if sums and last_data_row is not None:
            ws.insert_rows(last_data_row + 1)
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            for col, val in sums.items():
                cell = ws.cell(row=last_data_row + 1, column=col, value=val)
                cell.fill = yellow_fill

    # 收集对手方数据（含ID）
    collected = []
    if enable_internal_stats or collect_data:
        amount_col = money_cols[0]
        for row in range(block_start, block_end + 1):
            counterparty = ws.cell(row, counterparty_col).value
            if not counterparty:
                continue
            counterparty_id = ws.cell(row, counterparty_id_col).value if counterparty_id_col is not None else None
            amount_cell = ws.cell(row, amount_col)
            amount = amount_cell.value
            if amount is None:
                continue
            if isinstance(amount, (int, float)):
                val = amount
            else:
                val = clean_money_cell(amount)
            if not isinstance(val, (int, float)):
                continue
            inflow = outflow = 0.0
            if direction_col:
                dir_val = ws.cell(row, direction_col).value
                direction = get_direction_from_cell(dir_val)
                if direction == 'inflow':
                    inflow = abs(val)
                elif direction == 'outflow':
                    outflow = abs(val)
                else:
                    if val > 0:
                        inflow = val
                    elif val < 0:
                        outflow = abs(val)
            else:
                if val > 0:
                    inflow = val
                elif val < 0:
                    outflow = abs(val)
            if inflow == 0 and outflow == 0:
                continue
            collected.append({
                '对手方名称': str(counterparty).strip(),
                '对手方ID': str(counterparty_id).strip() if counterparty_id is not None else '',
                'inflow': inflow,
                'outflow': outflow,
                'abs_amount': abs(val),
                'block_index': block_idx
            })
    return collected

# ---------- 统计区块写入函数（符合打样样式，支持独立ID排序和颜色） ----------
def write_stat_block(ws, start_row, title, name_id_map, total_id_stats=None, is_total_section=False):
    current_row = start_row
    max_id_count = 0
    for id_dict in name_id_map.values():
        max_id_count = max(max_id_count, len(id_dict))
    id_col_groups = max_id_count

    # 第一行：大标题行
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
    title_cell = ws.cell(row=current_row, column=1, value=title)
    title_cell.font = Font(bold=True)
    title_cell.alignment = Alignment(horizontal='center')
    col_offset = 6
    for idx in range(1, id_col_groups + 1):
        start_col = col_offset
        end_col = col_offset + 4
        ws.merge_cells(start_row=current_row, start_column=start_col, end_row=current_row, end_column=end_col)
        cell = ws.cell(row=current_row, column=start_col, value=f"对手方账户ID{idx}")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        col_offset += 5
    current_row += 1

    # 第二行：子标题行
    left_headers = ['对手方名称', '交易次数', '流入金额', '流出金额', '金额合计']
    for col, h in enumerate(left_headers, start=1):
        cell = ws.cell(row=current_row, column=col, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    col_offset = 6
    sub_headers = ['对手方账户ID', '交易次数', '流入金额', '流出金额', '金额合计']
    for _ in range(id_col_groups):
        for sub_col, sub_h in enumerate(sub_headers, start=col_offset):
            cell = ws.cell(row=current_row, column=sub_col, value=sub_h)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
        col_offset += 5
    current_row += 1

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    blue_fill = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
    id_color_sequence = [green_fill, blue_fill]

    if total_id_stats is not None:
        total_count = sum(stats['交易次数'] for stats in total_id_stats.values())
        total_inflow = sum(stats['流入金额'] for stats in total_id_stats.values())
        total_outflow = sum(stats['流出金额'] for stats in total_id_stats.values())
        total_amount = sum(stats['金额合计'] for stats in total_id_stats.values())
        ws.cell(row=current_row, column=1, value="合计")
        ws.cell(row=current_row, column=2, value=total_count)
        ws.cell(row=current_row, column=3, value=total_inflow)
        ws.cell(row=current_row, column=4, value=total_outflow)
        ws.cell(row=current_row, column=5, value=total_amount)
        if is_total_section:
            for col in range(1, 6):
                ws.cell(row=current_row, column=col).fill = yellow_fill
        current_row += 1

    name_total = {name: sum(stats['金额合计'] for stats in id_dict.values()) for name, id_dict in name_id_map.items()}
    sorted_names = sorted(name_total.keys(), key=lambda x: name_total[x], reverse=True)
    for name in sorted_names:
        total_count = sum(stats['交易次数'] for stats in name_id_map[name].values())
        total_inflow = sum(stats['流入金额'] for stats in name_id_map[name].values())
        total_outflow = sum(stats['流出金额'] for stats in name_id_map[name].values())
        total_amount = sum(stats['金额合计'] for stats in name_id_map[name].values())
        ws.cell(row=current_row, column=1, value=name)
        ws.cell(row=current_row, column=2, value=total_count)
        ws.cell(row=current_row, column=3, value=total_inflow)
        ws.cell(row=current_row, column=4, value=total_outflow)
        ws.cell(row=current_row, column=5, value=total_amount)
        id_dict = name_id_map[name]
        sorted_ids = sorted(id_dict.keys(), key=lambda x: id_dict[x]['金额合计'], reverse=True)
        col_offset = 6
        for idx in range(max_id_count):
            start_col = col_offset
            if idx < len(sorted_ids):
                uid = sorted_ids[idx]
                stats = id_dict[uid]
                ws.cell(row=current_row, column=start_col, value=uid)
                ws.cell(row=current_row, column=start_col+1, value=stats['交易次数'])
                ws.cell(row=current_row, column=start_col+2, value=stats['流入金额'])
                ws.cell(row=current_row, column=start_col+3, value=stats['流出金额'])
                ws.cell(row=current_row, column=start_col+4, value=stats['金额合计'])
            col_offset += 5
        if is_total_section:
            for col in range(1, 6):
                ws.cell(row=current_row, column=col).fill = yellow_fill
        col_offset = 6
        for idx in range(max_id_count):
            color = id_color_sequence[idx % 2]
            for c in range(5):
                ws.cell(row=current_row, column=col_offset + c).fill = color
            col_offset += 5
        current_row += 1

    current_row += 1
    return current_row

def process_worksheet_multi(ws, wb, sheet_name, enable_sum, enable_internal_stats, collect_data,
                           manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                           enable_desensitize, desensitize_cols, enable_money_clean):
    max_row = ws.max_row
    max_col = ws.max_column
    if max_row <= 1 or max_col == 0:
        return []

    header_row = detect_header_row(ws, max_row, max_col)
    blocks = find_blocks(ws, max_row, max_col, header_row)
    if not blocks:
        return []

    all_collected = []
    for block_start, block_end, block_idx in reversed(blocks):
        data = process_block(ws, wb, block_start, block_end, block_idx, header_row, enable_sum, enable_internal_stats, collect_data,
                            manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                            enable_desensitize, desensitize_cols, enable_money_clean)
        if enable_internal_stats:
            all_collected.extend(data)
        elif collect_data:
            all_collected.extend(data)

    if enable_internal_stats and all_collected:
        stats_sheet_name = f"{sheet_name}对手方统计汇总"
        idx = 1
        while stats_sheet_name in wb.sheetnames:
            stats_sheet_name = f"{sheet_name}对手方统计汇总_{idx}"
            idx += 1
        stats_ws = wb.create_sheet(stats_sheet_name)
        current_row = 1

        block_dict = {}
        for item in all_collected:
            blk = item.pop('block_index')
            item['block_index'] = blk
            block_dict.setdefault(blk, []).append(item)

        total_data = all_collected
        total_name_id_map = {}
        total_id_stats = {}
        for d in total_data:
            name = d['对手方名称']
            uid = d['对手方ID']
            if name not in total_name_id_map:
                total_name_id_map[name] = {}
            if uid not in total_name_id_map[name]:
                total_name_id_map[name][uid] = {'交易次数': 0, '流入金额': 0, '流出金额': 0, '金额合计': 0}
            total_name_id_map[name][uid]['交易次数'] += 1
            total_name_id_map[name][uid]['流入金额'] += d['inflow']
            total_name_id_map[name][uid]['流出金额'] += d['outflow']
            total_name_id_map[name][uid]['金额合计'] += d['abs_amount']
            if uid not in total_id_stats:
                total_id_stats[uid] = {'交易次数': 0, '流入金额': 0, '流出金额': 0, '金额合计': 0}
            total_id_stats[uid]['交易次数'] += 1
            total_id_stats[uid]['流入金额'] += d['inflow']
            total_id_stats[uid]['流出金额'] += d['outflow']
            total_id_stats[uid]['金额合计'] += d['abs_amount']
        current_row = write_stat_block(stats_ws, current_row, "总计", total_name_id_map, total_id_stats=total_id_stats, is_total_section=True)

        for blk_idx in sorted(block_dict.keys()):
            block_data = block_dict[blk_idx]
            block_name_id_map = {}
            block_id_stats = {}
            for d in block_data:
                name = d['对手方名称']
                uid = d['对手方ID']
                if name not in block_name_id_map:
                    block_name_id_map[name] = {}
                if uid not in block_name_id_map[name]:
                    block_name_id_map[name][uid] = {'交易次数': 0, '流入金额': 0, '流出金额': 0, '金额合计': 0}
                block_name_id_map[name][uid]['交易次数'] += 1
                block_name_id_map[name][uid]['流入金额'] += d['inflow']
                block_name_id_map[name][uid]['流出金额'] += d['outflow']
                block_name_id_map[name][uid]['金额合计'] += d['abs_amount']
                if uid not in block_id_stats:
                    block_id_stats[uid] = {'交易次数': 0, '流入金额': 0, '流出金额': 0, '金额合计': 0}
                block_id_stats[uid]['交易次数'] += 1
                block_id_stats[uid]['流入金额'] += d['inflow']
                block_id_stats[uid]['流出金额'] += d['outflow']
                block_id_stats[uid]['金额合计'] += d['abs_amount']
            current_row = write_stat_block(stats_ws, current_row, f"区块{blk_idx}", block_name_id_map, total_id_stats=block_id_stats, is_total_section=False)

        for col in stats_ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            stats_ws.column_dimensions[col_letter].width = min(max_len+2, 20)

    if collect_data:
        return [{'对手方名称': d['对手方名称'], '对手方ID': d['对手方ID'], 'inflow': d['inflow'], 'outflow': d['outflow'], 'abs_amount': d['abs_amount']} for d in all_collected]
    else:
        return []

# ---------- 文件处理 ----------
def process_csv(file_path, enable_sum, enable_internal_stats, collect_data, temp_dir,
                manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                enable_desensitize, desensitize_titles, enable_money_clean):
    file_path = Path(file_path)
    try:
        df = pd.read_csv(file_path, dtype=str, keep_default_na=False, encoding='utf-8-sig')
    except:
        df = pd.read_csv(file_path, dtype=str, keep_default_na=False, encoding='gbk')
    temp_xlsx = Path(temp_dir) / (file_path.stem + "_temp.xlsx")
    with pd.ExcelWriter(temp_xlsx, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, header=False)
    wb = load_workbook(temp_xlsx, data_only=True)
    ws = wb.active
    header_row = detect_header_row(ws, ws.max_row, ws.max_column)
    desensitize_cols = []
    if enable_desensitize and desensitize_titles:
        for title in desensitize_titles:
            col = find_column_by_title(ws, ws.max_column, header_row, [title])
            if col is not None:
                desensitize_cols.append(col)
            else:
                log(f"警告：未找到脱敏列标题 '{title}'，跳过该列")
    data = process_worksheet_multi(ws, wb, "Sheet1", enable_sum, enable_internal_stats, collect_data,
                                   manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                   enable_desensitize, desensitize_cols, enable_money_clean)
    output_path = file_path.parent / f"{file_path.stem}_已处理.xlsx"
    wb.save(output_path)
    os.remove(temp_xlsx)
    return output_path, data

def process_old_xls(file_path, enable_sum, enable_internal_stats, collect_data, temp_dir,
                    manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                    enable_desensitize, desensitize_titles, enable_money_clean):
    file_path = Path(file_path)
    try:
        df = pd.read_excel(file_path, header=None, dtype=str, engine='xlrd')
    except:
        raise Exception("无法读取旧版 .xls 文件，请确保已安装 xlrd 库。")
    temp_xlsx = Path(temp_dir) / (file_path.stem + "_temp.xlsx")
    with pd.ExcelWriter(temp_xlsx, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, header=False)
    wb = load_workbook(temp_xlsx, data_only=True)
    ws = wb.active
    header_row = detect_header_row(ws, ws.max_row, ws.max_column)
    desensitize_cols = []
    if enable_desensitize and desensitize_titles:
        for title in desensitize_titles:
            col = find_column_by_title(ws, ws.max_column, header_row, [title])
            if col is not None:
                desensitize_cols.append(col)
            else:
                log(f"警告：未找到脱敏列标题 '{title}'，跳过该列")
    data = process_worksheet_multi(ws, wb, "Sheet1", enable_sum, enable_internal_stats, collect_data,
                                   manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                   enable_desensitize, desensitize_cols, enable_money_clean)
    output_path = file_path.parent / f"{file_path.stem}_已处理.xlsx"
    wb.save(output_path)
    os.remove(temp_xlsx)
    return output_path, data

def process_xlsx(file_path, enable_sum, enable_internal_stats, collect_data,
                 manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                 enable_desensitize, desensitize_titles, enable_money_clean):
    file_path = Path(file_path)
    wb = load_workbook(file_path, data_only=True)
    all_data = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_row = detect_header_row(ws, ws.max_row, ws.max_column)
        desensitize_cols = []
        if enable_desensitize and desensitize_titles:
            for title in desensitize_titles:
                col = find_column_by_title(ws, ws.max_column, header_row, [title])
                if col is not None:
                    desensitize_cols.append(col)
                else:
                    log(f"警告：未找到脱敏列标题 '{title}'，跳过该列")
        data = process_worksheet_multi(ws, wb, sheet_name, enable_sum, enable_internal_stats, collect_data,
                                       manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                       enable_desensitize, desensitize_cols, enable_money_clean)
        all_data.extend(data)
    output_path = file_path.parent / f"{file_path.stem}_已处理.xlsx"
    wb.save(output_path)
    return output_path, all_data

def fix_files(file_paths, enable_sum, enable_internal_stats, enable_cross_summary,
              manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
              enable_desensitize, desensitize_titles, enable_money_clean,
              progress_callback, finish_callback):
    log(f"开始处理，文件数量: {len(file_paths)}")
    log(f"选项状态: 求和={enable_sum}, 内部统计={enable_internal_stats}, 跨文件汇总={enable_cross_summary}")
    log(f"手动模式: {manual_mode}, 金额标题={user_amount_titles}, 对手方名称标题={user_counterparty_titles}, 对手方ID标题={user_counterparty_id_titles}")
    log(f"脱敏开关: {enable_desensitize}, 脱敏列标题={desensitize_titles}")
    log(f"金额清理开关: {enable_money_clean}")
    try:
        all_collected = []
        processed_paths = []
        total = len(file_paths)

        for idx, file_path_str in enumerate(file_paths):
            file_path = Path(file_path_str)
            log(f"处理文件 {idx+1}/{total}: {file_path.name}")
            progress_callback(idx + 1, total, file_path.name)
            ext = file_path.suffix.lower()
            if ext == '.csv':
                with tempfile.TemporaryDirectory() as tmpdir:
                    out_path, data = process_csv(file_path, enable_sum, enable_internal_stats, enable_cross_summary, tmpdir,
                                                 manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                                 enable_desensitize, desensitize_titles, enable_money_clean)
                    processed_paths.append(out_path)
                    if enable_cross_summary:
                        all_collected.extend(data)
                        log(f"  收集到 {len(data)} 条对手方明细")
            elif ext == '.xls':
                with tempfile.TemporaryDirectory() as tmpdir:
                    out_path, data = process_old_xls(file_path, enable_sum, enable_internal_stats, enable_cross_summary, tmpdir,
                                                     manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                                     enable_desensitize, desensitize_titles, enable_money_clean)
                    processed_paths.append(out_path)
                    if enable_cross_summary:
                        all_collected.extend(data)
                        log(f"  收集到 {len(data)} 条对手方明细")
            elif ext == '.xlsx':
                out_path, data = process_xlsx(file_path, enable_sum, enable_internal_stats, enable_cross_summary,
                                              manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                              enable_desensitize, desensitize_titles, enable_money_clean)
                processed_paths.append(out_path)
                if enable_cross_summary:
                    all_collected.extend(data)
                    log(f"  收集到 {len(data)} 条对手方明细")
            else:
                raise Exception(f"不支持的文件类型：{ext}")

        if enable_desensitize and desensitize_mapper.orig_to_mask:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            mapping_path = Path(file_paths[0]).parent / f"脱敏映射_{timestamp}.json"
            desensitize_mapper.save_mapping(mapping_path)
            log(f"脱敏映射已保存至: {mapping_path}")

        log(f"汇总数据总量: {len(all_collected)} 条")
        if enable_cross_summary and len(file_paths) > 1:
            if all_collected:
                cross_name_id_map = {}
                cross_id_stats = {}
                for d in all_collected:
                    name = d['对手方名称']
                    uid = d['对手方ID']
                    if name not in cross_name_id_map:
                        cross_name_id_map[name] = {}
                    if uid not in cross_name_id_map[name]:
                        cross_name_id_map[name][uid] = {'交易次数': 0, '流入金额': 0, '流出金额': 0, '金额合计': 0}
                    cross_name_id_map[name][uid]['交易次数'] += 1
                    cross_name_id_map[name][uid]['流入金额'] += d['inflow']
                    cross_name_id_map[name][uid]['流出金额'] += d['outflow']
                    cross_name_id_map[name][uid]['金额合计'] += d['abs_amount']
                    if uid not in cross_id_stats:
                        cross_id_stats[uid] = {'交易次数': 0, '流入金额': 0, '流出金额': 0, '金额合计': 0}
                    cross_id_stats[uid]['交易次数'] += 1
                    cross_id_stats[uid]['流入金额'] += d['inflow']
                    cross_id_stats[uid]['流出金额'] += d['outflow']
                    cross_id_stats[uid]['金额合计'] += d['abs_amount']

                first_dir = Path(file_paths[0]).parent
                summary_path = first_dir / "对手方汇总统计.xlsx"
                wb_summary = Workbook()
                ws_summary = wb_summary.active
                ws_summary.title = "对手方汇总"
                current_row = 1
                current_row = write_stat_block(ws_summary, current_row, "总计", cross_name_id_map, total_id_stats=cross_id_stats, is_total_section=True)
                for col in ws_summary.columns:
                    max_len = 0
                    col_letter = get_column_letter(col[0].column)
                    for cell in col:
                        if cell.value:
                            max_len = max(max_len, len(str(cell.value)))
                    ws_summary.column_dimensions[col_letter].width = min(max_len+2, 25)
                wb_summary.save(summary_path)
                log(f"成功生成跨文件汇总表: {summary_path}")
                finish_callback(True, f"所有文件处理完成！\n共处理 {len(processed_paths)} 个文件。\n汇总统计表已生成：{summary_path}\n脱敏映射已保存至同目录。", processed_paths)
                return
            else:
                log("警告：跨文件汇总已启用，但未收集到任何对手方数据，未生成汇总文件。")
                finish_callback(True, f"所有文件处理完成！\n共处理 {len(processed_paths)} 个文件。\n未收集到有效对手方数据，因此未生成汇总文件。", processed_paths)
                return
        elif enable_cross_summary and len(file_paths) == 1:
            log("仅选择单个文件，跨文件汇总已启用但未满足条件（需要至少2个文件）")
            finish_callback(True, f"仅选择单个文件，未生成跨文件汇总表。\n共处理 {len(processed_paths)} 个文件。", processed_paths)
        else:
            finish_callback(True, f"所有文件处理完成！\n共处理 {len(processed_paths)} 个文件。", processed_paths)
    except Exception as e:
        log(f"处理出错：{e}")
        finish_callback(False, f"处理出错：{e}", None)

# ---------- 还原功能 ----------
def restore_file():
    file_path = filedialog.askopenfilename(
        title="请选择已脱敏的 Excel 或 CSV 文件",
        filetypes=[("Excel/CSV 文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
    )
    if not file_path:
        return
    mapping_path = filedialog.askopenfilename(
        title="请选择对应的脱敏映射文件（.json）",
        filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")]
    )
    if not mapping_path:
        return

    try:
        mapper = DesensitizeMapper()
        mapper.load_mapping(mapping_path)

        file_ext = Path(file_path).suffix.lower()
        if file_ext == '.csv':
            df = pd.read_csv(file_path, dtype=str, keep_default_na=False, encoding='utf-8-sig')
            output_path = Path(file_path).parent / f"{Path(file_path).stem}_还原.xlsx"
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=True)
            wb = load_workbook(output_path)
        elif file_ext in ['.xlsx', '.xls']:
            wb = load_workbook(file_path, data_only=True)
            output_path = Path(file_path).parent / f"{Path(file_path).stem}_还原.xlsx"
        else:
            messagebox.showerror("错误", "不支持的文件格式")
            return

        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None and isinstance(cell.value, str):
                        original = mapper.unmask(cell.value)
                        if original != cell.value:
                            cell.value = original
        wb.save(output_path)
        messagebox.showinfo("完成", f"还原完成！\n已保存至：{output_path}")
    except Exception as e:
        messagebox.showerror("错误", f"还原失败：{e}")

# ---------- GUI ----------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"银行流水梳理工具 {VERSION}")
        self.root.geometry("680x850")
        self.root.configure(bg="#f0f0f0")

        self.enable_money_clean = BooleanVar(value=True)
        self.enable_sum = BooleanVar(value=False)
        self.enable_internal_stats = BooleanVar(value=False)
        self.enable_cross_summary = BooleanVar(value=False)
        self.manual_mode = BooleanVar(value=False)
        self.enable_desensitize = BooleanVar(value=False)
        self.user_amount_titles = tk.StringVar()
        self.user_counterparty_titles = tk.StringVar()
        self.user_counterparty_id_titles = tk.StringVar()
        self.desensitize_titles = tk.StringVar()
        self.processing = False

        info_frame = tk.Frame(root, bg="#f0f0f0")
        info_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        info_text = tk.Text(info_frame, height=8, width=65, wrap=tk.WORD, bg="#fafafa", font=("微软雅黑", 9), relief=tk.FLAT)
        info_text.insert(tk.END, "本软件不会去动原始表格的数据，而是生成一个处理后的表格。\n", "red_bold")
        info_text.insert(tk.END, "银行拉取的流水文件中，金额单元格常包含不可见格式（如货币符号、千分位、空格等），"
                                 "导致Excel公式无法正常求和。本软件可自动识别金额列，清理不可见字符，并转换为数值，"
                                 "确保统计准确。同时可生成对手方流水统计表（含流入/流出/合计）以及按对手方账户ID的细分统计。\n\n"
                                 "【信息脱敏】\n可对指定列进行确定性脱敏（数字→数字、中文→中文、英文→英文，长度不变），"
                                 "并保存映射文件，方便后续还原。\n\n"
                                 "【还原脱敏】\n点击下方按钮，分两步操作：先选择脱敏后的文件，再选择对应的映射文件，恢复原始数据。\n\n"
                                 "【导出运行日志】\n当遇到异常时，可点击“导出运行日志”按钮，将详细运行信息导出供分析。")
        info_text.tag_config("red_bold", foreground="red", font=("微软雅黑", 9, "bold"))
        info_text.config(state=tk.DISABLED)
        info_text.pack(fill=tk.BOTH, expand=True)

        options_frame = tk.Frame(root, bg="#f0f0f0")
        options_frame.pack(pady=5)

        row = 0
        cb_money_clean = tk.Checkbutton(options_frame, text="清理不可见字符，使金额列数据可被正常统计", variable=self.enable_money_clean, bg="#f0f0f0", font=("微软雅黑", 10))
        cb_money_clean.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        cb_sum = tk.Checkbutton(options_frame, text="在每个区块末尾插入绝对值求和行（标黄）", variable=self.enable_sum, bg="#f0f0f0", font=("微软雅黑", 10))
        cb_sum.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        cb_stats = tk.Checkbutton(options_frame, text="在生成表格内添加对手方流水统计表", variable=self.enable_internal_stats, bg="#f0f0f0", font=("微软雅黑", 10))
        cb_stats.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        cb_cross = tk.Checkbutton(options_frame, text="生成对手方流水统计表汇总文件（仅在勾选并选择多个文件时可用）", variable=self.enable_cross_summary, bg="#f0f0f0", font=("微软雅黑", 10))
        cb_cross.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        self.desensitize_frame = Frame(options_frame, bg="#f0f0f0")
        self.desensitize_frame.grid(row=row, column=0, sticky="w", padx=10, pady=10)
        desensitize_cb = Checkbutton(self.desensitize_frame, text="启用信息脱敏", variable=self.enable_desensitize, bg="#f0f0f0", font=("微软雅黑", 10), command=self.toggle_desensitize_input)
        desensitize_cb.pack(anchor="w")

        self.desensitize_label = Label(self.desensitize_frame, text="需要脱敏的列标题（多个用逗号分隔）:", bg="#f0f0f0", font=("微软雅黑", 9))
        self.desensitize_entry = Entry(self.desensitize_frame, textvariable=self.desensitize_titles, width=60, font=("微软雅黑", 9))
        self.desensitize_label.pack_forget()
        self.desensitize_entry.pack_forget()

        row += 1
        self.manual_frame = Frame(options_frame, bg="#f0f0f0")
        self.manual_frame.grid(row=row, column=0, sticky="w", padx=10, pady=10)
        manual_cb = Checkbutton(self.manual_frame, text="手动指定列（自动识别失败时使用）", variable=self.manual_mode, bg="#f0f0f0", font=("微软雅黑", 10), command=self.toggle_manual_input)
        manual_cb.pack(anchor="w")

        self.amount_label = Label(self.manual_frame, text="金额列标题（多个用逗号分隔）:", bg="#f0f0f0", font=("微软雅黑", 9))
        self.amount_entry = Entry(self.manual_frame, textvariable=self.user_amount_titles, width=50, font=("微软雅黑", 9))
        self.counterparty_label = Label(self.manual_frame, text="对手方名称列标题（多个用逗号分隔）:", bg="#f0f0f0", font=("微软雅黑", 9))
        self.counterparty_entry = Entry(self.manual_frame, textvariable=self.user_counterparty_titles, width=50, font=("微软雅黑", 9))
        self.counterparty_id_label = Label(self.manual_frame, text="对手方ID列标题（多个用逗号分隔）:", bg="#f0f0f0", font=("微软雅黑", 9))
        self.counterparty_id_entry = Entry(self.manual_frame, textvariable=self.user_counterparty_id_titles, width=50, font=("微软雅黑", 9))

        self.amount_label.pack_forget()
        self.amount_entry.pack_forget()
        self.counterparty_label.pack_forget()
        self.counterparty_entry.pack_forget()
        self.counterparty_id_label.pack_forget()
        self.counterparty_id_entry.pack_forget()

        note_label = tk.Label(root, text="提示：手动模式下，软件会优先使用您输入的标题查找对应列，若找不到则跳过该文件。脱敏列同样支持手动指定。",
                              bg="#f0f0f0", font=("微软雅黑", 9), fg="blue", wraplength=600)
        note_label.pack(pady=5)

        btn_frame = tk.Frame(root, bg="#f0f0f0")
        btn_frame.pack(pady=10)
        self.btn = tk.Button(btn_frame, text="选择文件（可多选）", command=self.select_files, font=("微软雅黑", 10))
        self.btn.pack()

        restore_btn = Button(btn_frame, text="还原脱敏（分两步：先选文件，再选映射文件）", command=restore_file, font=("微软雅黑", 10), bg="#e0e0e0")
        restore_btn.pack(pady=5)

        restore_hint = tk.Label(btn_frame, text="还原操作：依次选择脱敏后的Excel/CSV文件，然后选择对应的脱敏映射文件（.json）",
                                bg="#f0f0f0", font=("微软雅黑", 8), fg="gray")
        restore_hint.pack(pady=2)

        export_btn = tk.Button(btn_frame, text="导出运行日志", command=export_log, font=("微软雅黑", 10), bg="#e0e0e0")
        export_btn.pack(pady=5)

    def toggle_manual_input(self):
        if self.manual_mode.get():
            self.amount_label.pack(anchor="w", pady=(5,0))
            self.amount_entry.pack(anchor="w", pady=2)
            self.counterparty_label.pack(anchor="w", pady=(5,0))
            self.counterparty_entry.pack(anchor="w", pady=2)
            self.counterparty_id_label.pack(anchor="w", pady=(5,0))
            self.counterparty_id_entry.pack(anchor="w", pady=2)
        else:
            self.amount_label.pack_forget()
            self.amount_entry.pack_forget()
            self.counterparty_label.pack_forget()
            self.counterparty_entry.pack_forget()
            self.counterparty_id_label.pack_forget()
            self.counterparty_id_entry.pack_forget()

    def toggle_desensitize_input(self):
        if self.enable_desensitize.get():
            self.desensitize_label.pack(anchor="w", pady=(5,0))
            self.desensitize_entry.pack(anchor="w", pady=2)
        else:
            self.desensitize_label.pack_forget()
            self.desensitize_entry.pack_forget()

    def select_files(self):
        if self.processing:
            return
        file_paths = filedialog.askopenfilenames(
            title="选择要修复的文件（可多选）",
            filetypes=[("Excel/CSV 文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
        )
        if file_paths:
            self.start_processing(file_paths)

    def start_processing(self, file_paths):
        self.processing = True
        self.btn.config(state=tk.DISABLED)
        self.progress_win = tk.Toplevel(self.root)
        self.progress_win.title("处理中")
        self.progress_win.geometry("400x150")
        self.progress_win.transient(self.root)
        self.progress_win.grab_set()
        self.progress_win.protocol("WM_DELETE_WINDOW", lambda: None)

        self.progress_label = tk.Label(self.progress_win, text="正在处理文件，请稍候...\n", font=("微软雅黑", 10))
        self.progress_label.pack(pady=10)
        self.progress_bar = tk.Label(self.progress_win, text="", font=("微软雅黑", 9), fg="blue")
        self.progress_bar.pack()

        def parse_titles(txt):
            if not txt:
                return []
            parts = re.split(r'[，,、\s]+', txt)
            return [p.strip() for p in parts if p.strip()]

        user_amount = parse_titles(self.user_amount_titles.get())
        user_counterparty = parse_titles(self.user_counterparty_titles.get())
        user_counterparty_id = parse_titles(self.user_counterparty_id_titles.get())
        desensitize_titles = parse_titles(self.desensitize_titles.get()) if self.enable_desensitize.get() else []

        def update_progress(current, total, filename):
            self.progress_label.config(text=f"正在处理: {filename}\n({current}/{total})")
            self.progress_bar.config(text=f"进度: {current}/{total}")

        def finish_callback(success, message, extra):
            global desensitize_mapper
            desensitize_mapper = DesensitizeMapper()
            self.progress_win.destroy()
            self.btn.config(state=tk.NORMAL)
            self.processing = False
            if success:
                messagebox.showinfo("完成", message)
            else:
                messagebox.showerror("错误", message)

        thread = threading.Thread(target=fix_files, args=(
            file_paths, self.enable_sum.get(), self.enable_internal_stats.get(), self.enable_cross_summary.get(),
            self.manual_mode.get(), user_amount, user_counterparty, user_counterparty_id,
            self.enable_desensitize.get(), desensitize_titles,
            self.enable_money_clean.get(),
            update_progress, finish_callback
        ))
        thread.daemon = True
        thread.start()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()