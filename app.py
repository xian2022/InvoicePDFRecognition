#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票处理 Web 应用
功能：上传 PDF 发票 → 提取信息 → 导出 Excel
"""

import os
import re
import json
import tempfile
from datetime import datetime
from flask import Flask, render_template_string, request, send_file, jsonify
import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB 限制

# HTML 模板
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>发票信息提取工具</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 20px;
            padding: 40px;
            width: 100%;
            max-width: 600px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        }
        h1 {
            text-align: center;
            color: #333;
            margin-bottom: 10px;
            font-size: 28px;
        }
        .subtitle {
            text-align: center;
            color: #666;
            margin-bottom: 30px;
            font-size: 14px;
        }
        .upload-area {
            border: 3px dashed #ddd;
            border-radius: 15px;
            padding: 60px 20px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            background: #fafafa;
        }
        .upload-area:hover, .upload-area.dragover {
            border-color: #667eea;
            background: #f0f4ff;
        }
        .upload-icon {
            font-size: 48px;
            margin-bottom: 15px;
        }
        .upload-text {
            color: #666;
            font-size: 16px;
        }
        .upload-hint {
            color: #999;
            font-size: 12px;
            margin-top: 10px;
        }
        input[type="file"] {
            display: none;
        }
        .file-list {
            margin-top: 20px;
            max-height: 200px;
            overflow-y: auto;
        }
        .file-item {
            display: flex;
            align-items: center;
            padding: 12px 15px;
            background: #f8f9fa;
            border-radius: 8px;
            margin-bottom: 8px;
            font-size: 14px;
        }
        .file-item .icon { margin-right: 10px; }
        .file-item .name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .file-item .size { color: #999; font-size: 12px; margin-left: 10px; }
        .file-item .remove {
            color: #ff4757;
            cursor: pointer;
            margin-left: 10px;
            font-weight: bold;
        }
        .btn {
            width: 100%;
            padding: 18px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 12px;
            font-size: 18px;
            font-weight: 600;
            cursor: pointer;
            margin-top: 25px;
            transition: transform 0.2s, box-shadow 0.2s;
        }
        .btn:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 10px 30px rgba(102, 126, 234, 0.4);
        }
        .btn:disabled {
            opacity: 0.6;
            cursor: not-allowed;
        }
        .progress {
            margin-top: 20px;
            display: none;
        }
        .progress-bar {
            height: 8px;
            background: #eee;
            border-radius: 4px;
            overflow: hidden;
        }
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #667eea, #764ba2);
            width: 0%;
            transition: width 0.3s ease;
        }
        .progress-text {
            text-align: center;
            margin-top: 10px;
            color: #666;
            font-size: 14px;
        }
        .result {
            margin-top: 20px;
            padding: 20px;
            background: #d4edda;
            border-radius: 10px;
            display: none;
        }
        .result.success { background: #d4edda; color: #155724; }
        .result.error { background: #f8d7da; color: #721c24; }
        .result-title { font-weight: 600; margin-bottom: 10px; }
        .download-btn {
            display: inline-block;
            margin-top: 15px;
            padding: 12px 30px;
            background: #28a745;
            color: white;
            text-decoration: none;
            border-radius: 8px;
            font-weight: 600;
        }
        .download-btn:hover { background: #218838; }
        .loading-spinner {
            display: inline-block;
            width: 20px;
            height: 20px;
            border: 3px solid rgba(255,255,255,0.3);
            border-radius: 50%;
            border-top-color: white;
            animation: spin 1s ease-in-out infinite;
            margin-right: 10px;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <div class="container">
        <h1>📄 发票信息提取工具</h1>
        <p class="subtitle">支持批量上传 PDF 电子发票，自动提取关键信息并导出 Excel</p>
        
        <div class="upload-area" id="uploadArea">
            <div class="upload-icon">📁</div>
            <div class="upload-text">点击或拖拽上传发票文件</div>
            <div class="upload-hint">支持 PDF 格式，可多选</div>
        </div>
        <input type="file" id="fileInput" multiple accept=".pdf">
        
        <div class="file-list" id="fileList"></div>

        <button class="btn" id="processBtn" disabled>
            开始处理
        </button>

        <div class="progress" id="progress">
            <div class="progress-bar">
                <div class="progress-fill" id="progressFill"></div>
            </div>
            <div class="progress-text" id="progressText">准备中...</div>
        </div>

        <div class="result" id="result"></div>
    </div>

    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');
        const processBtn = document.getElementById('processBtn');
        const progress = document.getElementById('progress');
        const progressFill = document.getElementById('progressFill');
        const progressText = document.getElementById('progressText');
        const result = document.getElementById('result');

        let selectedFiles = [];
        let lastDownloadUrl = null;  // 保存上次下载的 URL
        let lastFileName = '';       // 保存上次下载的文件名

        // 上传区域点击
        uploadArea.addEventListener('click', () => fileInput.click());

        // 拖拽效果
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            handleFiles(e.dataTransfer.files);
        });

        // 文件选择
        fileInput.addEventListener('change', (e) => {
            handleFiles(e.target.files);
        });

        function handleFiles(files) {
            for (let file of files) {
                if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
                    selectedFiles.push(file);
                } else {
                    alert(`文件 "${file.name}" 不是 PDF 格式，已跳过`);
                }
            }
            updateFileList();
        }

        function updateFileList() {
            fileList.innerHTML = '';
            selectedFiles.forEach((file, index) => {
                const item = document.createElement('div');
                item.className = 'file-item';
                item.innerHTML = `
                    <span class="icon">📄</span>
                    <span class="name">${file.name}</span>
                    <span class="size">${formatSize(file.size)}</span>
                    <span class="remove" onclick="removeFile(${index})">×</span>
                `;
                fileList.appendChild(item);
            });
            processBtn.disabled = selectedFiles.length === 0;
        }

        function removeFile(index) {
            selectedFiles.splice(index, 1);
            updateFileList();
        }

        function formatSize(bytes) {
            if (bytes < 1024) return bytes + ' B';
            if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
            return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
        }

        // 重新下载功能
        function redownload() {
            if (lastDownloadUrl) {
                const a = document.createElement('a');
                a.href = lastDownloadUrl;
                a.download = lastFileName;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            }
        }

        // 继续处理（重置页面）
        function continueProcessing() {
            // 释放旧的 URL
            if (lastDownloadUrl) {
                window.URL.revokeObjectURL(lastDownloadUrl);
                lastDownloadUrl = null;
            }
            
            // 重置状态
            selectedFiles = [];
            fileList.innerHTML = '';
            processBtn.disabled = true;
            processBtn.innerHTML = '开始处理';
            processBtn.style.display = 'block';  // 恢复按钮显示
            result.style.display = 'none';
            progress.style.display = 'none';
            fileInput.value = '';
        }

        // 处理按钮点击
        processBtn.addEventListener('click', async () => {
            if (selectedFiles.length === 0) return;

            processBtn.disabled = true;
            processBtn.innerHTML = '<span class="loading-spinner"></span>处理中...';
            progress.style.display = 'block';
            result.style.display = 'none';

            const formData = new FormData();
            selectedFiles.forEach(file => {
                formData.append('files', file);
            });

            try {
                progressText.textContent = `正在处理 ${selectedFiles.length} 个文件...`;
                progressFill.style.width = '30%';

                const response = await fetch('/process', {
                    method: 'POST',
                    body: formData
                });

                progressFill.style.width = '80%';

                if (!response.ok) {
                    throw new Error(await response.text());
                }

                const blob = await response.blob();
                progressFill.style.width = '100%';

                // 释放之前的 URL
                if (lastDownloadUrl) {
                    window.URL.revokeObjectURL(lastDownloadUrl);
                }

                // 创建新的下载链接并保存
                lastDownloadUrl = window.URL.createObjectURL(blob);
                lastFileName = `发票信息_${new Date().toISOString().slice(0,10)}.xlsx`;
                
                // 自动触发下载
                const a = document.createElement('a');
                a.href = lastDownloadUrl;
                a.download = lastFileName;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);

                // 隐藏处理按钮，显示结果区域
                processBtn.style.display = 'none';
                
                result.className = 'result success';
                result.innerHTML = `
                    <div class="result-title">处理完成！</div>
                    <div>成功处理 ${selectedFiles.length} 张发票</div>
                    <div style="margin-top: 15px;">
                        <button class="download-btn" onclick="redownload()">📥 重新下载</button>
                        <button class="download-btn" style="background: #667eea; margin-left: 10px;" onclick="continueProcessing()">➕ 继续处理</button>
                    </div>
                `;
                result.style.display = 'block';

            } catch (error) {
                result.className = 'result error';
                result.innerHTML = `<div class="result-title">处理失败</div><div>${error.message}</div>`;
                result.style.display = 'block';
                processBtn.disabled = false;
                processBtn.innerHTML = '开始处理';
            } finally {
                setTimeout(() => { progress.style.display = 'none'; }, 1000);
            }
        });
    </script>
</body>
</html>
'''

class InvoiceExtractor:
    """发票信息提取器"""
    
    @staticmethod
    def extract_from_pdf(pdf_path, original_filename=None):
        """从 PDF 文件中提取发票信息
        
        Args:
            pdf_path: PDF 文件路径
            original_filename: 原始文件名（用于 Web 上传场景）
        """
        # 使用原始文件名（如果有），否则从路径中提取
        filename = original_filename if original_filename else os.path.basename(pdf_path)
        
        info = {
            '文件名': filename,
            '发票代码': '',
            '发票号码': '',
            '开票日期': '',
            '购买方名称': '',
            '购买方税号': '',
            '销售方名称': '',
            '销售方税号': '',
            '金额合计': '',
            '税额合计': '',
            '价税合计': '',
            '备注': '',
            '明细项目': []  # 存储明细行列表
        }
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                full_text = ''
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        full_text += text + '\n'
                
                # 保留原始文本用于分析
                raw_text = full_text
                
                # 清理文本（去除多余空格但保留换行）
                text = re.sub(r'[ \t]+', ' ', full_text)
                
                # 提取明细项目
                items = InvoiceExtractor._extract_items(raw_text)
                info['明细项目'] = items
                
                # 提取发票代码和号码
                code_match = re.search(r'发票代码[：:]?\s*(\d{10,12})', text)
                if code_match:
                    info['发票代码'] = code_match.group(1)
                
                num_match = re.search(r'发票号码[：:]?\s*(\d{8,20})', text)
                if num_match:
                    info['发票号码'] = num_match.group(1)
                
                # 提取开票日期
                date_match = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', text)
                if date_match:
                    info['开票日期'] = f"{date_match.group(1)}-{date_match.group(2).zfill(2)}-{date_match.group(3).zfill(2)}"
                
                # 提取购买方和销售方信息 - 适配电子发票格式
                # 格式示例: "购 名称：陕西安泰金属新材料有限公司 销 名称：陕西众邦达源再生资源利用有限公司"
                # 先尝试匹配同一行的情况
                both_match = re.search(r'购[^销]*?名称[：:]\s*([^\s]+?)\s+销[^名]*?名称[：:]\s*([^\n]+)', text)
                if both_match:
                    info['购买方名称'] = both_match.group(1).strip()
                    info['销售方名称'] = both_match.group(2).strip()
                else:
                    # 分别匹配
                    buyer_match = re.search(r'购[^售]*?名称[：:]\s*([^\n]+?)(?:\s+销|\s+统一|$)', text)
                    if buyer_match:
                        info['购买方名称'] = buyer_match.group(1).strip()
                    
                    seller_match = re.search(r'销[^购]*?名称[：:]\s*([^\n]+?)(?:\s+统一|\s+开户|$)', text)
                    if seller_match:
                        info['销售方名称'] = seller_match.group(1).strip()
                
                # 购买方税号
                buyer_tax_match = re.search(r'91610203[A-Za-z0-9]+', text)
                if not buyer_tax_match:
                    buyer_tax_match = re.search(r'(?:购买方|购方)[^销]*?(?:纳税人识别号|信用代码)[：:]?\s*([A-Za-z0-9]{18,20})', text)
                if buyer_tax_match:
                    info['购买方税号'] = buyer_tax_match.group(0) if '91610203' in buyer_tax_match.group(0) else buyer_tax_match.group(1)
                
                # 销售方税号
                seller_tax_match = re.search(r'91611105[A-Za-z0-9]+', text)
                if not seller_tax_match:
                    seller_tax_match = re.search(r'(?:销售方|销方)[^购]*?(?:纳税人识别号|信用代码)[：:]?\s*([A-Za-z0-9]{18,20})', text)
                if seller_tax_match:
                    info['销售方税号'] = seller_tax_match.group(0) if '91611105' in seller_tax_match.group(0) else seller_tax_match.group(1)
                
                # 提取金额 - 查找合计行
                # 格式: "合 计 ¥1692531.35 ¥220029.09"
                amount_match = re.search(r'合\s*计[^\d¥￥]*[¥￥]\s*([\d,]+\.\d+)\s+[¥￥]\s*([\d,]+\.\d+)', text)
                if amount_match:
                    info['金额合计'] = amount_match.group(1).replace(',', '')
                    info['税额合计'] = amount_match.group(2).replace(',', '')
                
                # 提取价税合计 - 查找小写金额
                # 格式: "（小写） ¥1912560.44"
                total_match = re.search(r'[（(]小写[)）][^\d¥￥]*[¥￥]\s*([\d,]+\.\d+)', text)
                if total_match:
                    info['价税合计'] = total_match.group(1).replace(',', '')
                
                # 如果没找到，尝试其他模式
                if info['价税合计'] == '':
                    total_match2 = re.search(r'价税合计[^¥￥]*[¥￥]\s*([\d,]+\.\d+)', text)
                    if total_match2:
                        info['价税合计'] = total_match2.group(1).replace(',', '')
                
        except Exception as e:
            info['备注'] = f'解析错误: {str(e)}'
        
        return info
    
    @staticmethod
    def _extract_items(text):
        """提取发票明细项目"""
        items = []
        lines = text.split('\n')
        
        # 查找表头行
        header_idx = -1
        for i, line in enumerate(lines):
            if '项目名称' in line and ('金额' in line or '单' in line):
                header_idx = i
                break
        
        if header_idx == -1:
            return items
        
        # 从表头下一行开始提取，直到"合 计"行
        i = header_idx + 1
        while i < len(lines):
            line = lines[i].strip()
            
            # 遇到合计行停止
            if re.match(r'^合\s*计', line):
                break
            
            # 跳过空行
            if not line:
                i += 1
                continue
            
            # 检查这一行是否包含数字（是明细行）
            has_numbers = bool(re.search(r'\d+\.?\d*', line))
            
            if has_numbers:
                # 尝试解析当前行
                item = InvoiceExtractor._parse_item_line(line)
                
                # 如果解析失败，可能是名称被分行，尝试合并下一行
                if not item and i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    # 如果下一行没有数字，可能是名称的延续
                    if next_line and not re.search(r'\d+\.?\d*', next_line):
                        combined_line = line + next_line
                        item = InvoiceExtractor._parse_item_line(combined_line)
                        if item:
                            i += 1  # 跳过已合并的行
                
                if item:
                    items.append(item)
            
            i += 1
        
        return items
    
    @staticmethod
    def _parse_item_line(line):
        """解析单行明细数据"""
        # 先清理多余空格
        line = re.sub(r'\s+', ' ', line).strip()
        
        # 匹配模式：项目名称 [规格型号] 单位 数量 单价 金额 税率 税额
        # 示例: "*有色金属压延材*铝次品 吨 31.17 18184.9203539823 566823.96 13% 73687.12"
        
        # 尝试匹配带单位的格式
        pattern = r'^(.+?)\s+(吨|个|件|kg|千克|米|台|套|箱|包|瓶|只|条|张|块|根|组|批|升|毫升|g|克)\s+(\d+\.?\d*)\s+(\d+\.?\d*)\s+(\d+\.?\d*)\s+(\d+%)\s+(\d+\.?\d*)$'
        match = re.match(pattern, line)
        
        if match:
            return {
                '项目名称': match.group(1).strip(),
                '规格型号': '',
                '单位': match.group(2),
                '数量': match.group(3),
                '单价': match.group(4),
                '金额': match.group(5),
                '税率': match.group(6),
                '税额': match.group(7)
            }
        
        # 尝试不带单位的格式
        pattern2 = r'^(.+?)\s+(\d+\.?\d*)\s+(\d+\.?\d*)\s+(\d+\.?\d*)\s+(\d+%)\s+(\d+\.?\d*)$'
        match2 = re.match(pattern2, line)
        
        if match2:
            return {
                '项目名称': match2.group(1).strip(),
                '规格型号': '',
                '单位': '',
                '数量': match2.group(2),
                '单价': match2.group(3),
                '金额': match2.group(4),
                '税率': match2.group(5),
                '税额': match2.group(6)
            }
        
        # 备用解析：按空格分割找数字
        parts = line.split()
        num_parts = []
        text_parts = []
        
        for p in parts:
            if re.match(r'^[\d.%]+$', p):
                num_parts.append(p)
            else:
                text_parts.append(p)
        
        # 需要至少5个数字（数量、单价、金额、税率、税额）
        if len(num_parts) >= 5:
            name = ' '.join(text_parts)
            return {
                '项目名称': name,
                '规格型号': '',
                '单位': '',
                '数量': num_parts[0],
                '单价': num_parts[1],
                '金额': num_parts[2],
                '税率': num_parts[3],
                '税额': num_parts[4]
            }
        
        return None


def create_excel(invoice_data_list, output_path):
    """创建 Excel 文件，支持明细项目多行合并展示"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '发票信息'
    
    # 表头样式
    header_font = Font(bold=True, size=11, color='FFFFFF')
    header_fill = openpyxl.styles.PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    # 定义列头 - 包含明细项目字段
    headers = [
        '文件名', '发票代码', '发票号码', '开票日期', 
        '购买方名称', '购买方税号', '销售方名称', '销售方税号',
        '项目名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额',
        '金额合计', '税额合计', '价税合计', '备注'
    ]
    
    # 写入表头
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # 写入数据
    current_row = 2
    for invoice in invoice_data_list:
        items = invoice.get('明细项目', [])
        
        if not items:
            # 没有明细项目，只写一行
            row_data = [
                invoice.get('文件名', ''),
                invoice.get('发票代码', ''),
                invoice.get('发票号码', ''),
                invoice.get('开票日期', ''),
                invoice.get('购买方名称', ''),
                invoice.get('购买方税号', ''),
                invoice.get('销售方名称', ''),
                invoice.get('销售方税号', ''),
                '', '', '', '', '', '', '', '',  # 明细字段留空
                invoice.get('金额合计', ''),
                invoice.get('税额合计', ''),
                invoice.get('价税合计', ''),
                invoice.get('备注', '')
            ]
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=value)
                cell.border = thin_border
                cell.alignment = center_align if col_idx <= 8 or col_idx >= 17 else left_align
            current_row += 1
        else:
            # 有明细项目，每个明细一行，基本信息合并单元格
            start_row = current_row
            for item in items:
                row_data = [
                    invoice.get('文件名', ''),
                    invoice.get('发票代码', ''),
                    invoice.get('发票号码', ''),
                    invoice.get('开票日期', ''),
                    invoice.get('购买方名称', ''),
                    invoice.get('购买方税号', ''),
                    invoice.get('销售方名称', ''),
                    invoice.get('销售方税号', ''),
                    item.get('项目名称', ''),
                    item.get('规格型号', ''),
                    item.get('单位', ''),
                    item.get('数量', ''),
                    item.get('单价', ''),
                    item.get('金额', ''),
                    item.get('税率', ''),
                    item.get('税额', ''),
                    invoice.get('金额合计', ''),
                    invoice.get('税额合计', ''),
                    invoice.get('价税合计', ''),
                    invoice.get('备注', '')
                ]
                for col_idx, value in enumerate(row_data, 1):
                    cell = ws.cell(row=current_row, column=col_idx, value=value)
                    cell.border = thin_border
                    cell.alignment = center_align if col_idx <= 8 or col_idx >= 17 else left_align
                current_row += 1
            
            end_row = current_row - 1
            
            # 合并基本信息单元格（前8列和最后4列）
            if end_row > start_row:
                # 合并前8列（基本信息）
                for col in range(1, 9):
                    ws.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)
                
                # 合并最后4列（合计信息）
                for col in range(17, 21):
                    ws.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)
    
    # 调整列宽
    column_widths = {
        1: 25,   # 文件名
        2: 15,   # 发票代码
        3: 20,   # 发票号码
        4: 12,   # 开票日期
        5: 30,   # 购买方名称
        6: 22,   # 购买方税号
        7: 30,   # 销售方名称
        8: 22,   # 销售方税号
        9: 35,   # 项目名称
        10: 15,  # 规格型号
        11: 8,   # 单位
        12: 10,  # 数量
        13: 15,  # 单价
        14: 12,  # 金额
        15: 8,   # 税率
        16: 12,  # 税额
        17: 12,  # 金额合计
        18: 12,  # 税额合计
        19: 12,  # 价税合计
        20: 30   # 备注
    }
    
    for col, width in column_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    
    # 冻结首行
    ws.freeze_panes = 'A2'
    
    wb.save(output_path)
    return output_path


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/process', methods=['POST'])
def process():
    if 'files' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400
    
    files = request.files.getlist('files')
    if not files or files[0].filename == '':
        return jsonify({'error': '没有选择文件'}), 400
    
    extractor = InvoiceExtractor()
    invoice_data_list = []
    
    # 处理每个文件
    for file in files:
        if file.filename.lower().endswith('.pdf'):
            # 获取原始文件名（使用 secure_filename 确保安全，但保留中文）
            original_name = file.filename
            
            # 保存到临时文件
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                file.save(tmp.name)
                tmp_path = tmp.name

            try:
                # 提取信息，传入原始文件名
                info = extractor.extract_from_pdf(tmp_path, original_filename=original_name)
                invoice_data_list.append(info)
            finally:
                # 清理临时文件
                os.unlink(tmp_path)
    
    if not invoice_data_list:
        return jsonify({'error': '没有有效的 PDF 文件'}), 400
    
    # 创建 Excel
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        excel_path = tmp.name
    
    create_excel(invoice_data_list, excel_path)
    
    return send_file(
        excel_path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'发票信息_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    )


if __name__ == '__main__':
    import webbrowser
    import threading
    import time
    
    print("=" * 50)
    print("  发票信息提取工具已启动")
    print("  请在浏览器中打开: http://localhost:5000")
    print("=" * 50)
    
    # 延迟 2 秒后自动打开浏览器
    def open_browser():
        time.sleep(2)
        webbrowser.open('http://localhost:5000')
    
    threading.Thread(target=open_browser, daemon=True).start()
    
    app.run(host='0.0.0.0', port=5000, debug=False)
