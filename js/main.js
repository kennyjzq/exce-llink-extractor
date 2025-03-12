/**
 * Excel链接导出工具 - 主要JavaScript文件
 * 实现Excel文件处理和链接提取功能
 */

document.addEventListener('DOMContentLoaded', () => {
    // 获取DOM元素
    const dropArea = document.getElementById('dropArea');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    const resultsSection = document.getElementById('resultsSection');
    const linksTableBody = document.getElementById('linksTableBody');
    const exportTxtBtn = document.getElementById('exportTxt');
    const exportExcelBtn = document.getElementById('exportExcel');
    const loadingOverlay = document.getElementById('loadingOverlay');
    
    // 存储提取的链接
    let extractedLinks = [];
    
    // 拖放功能
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, () => {
            dropArea.classList.add('drag-over');
        }, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, () => {
            dropArea.classList.remove('drag-over');
        }, false);
    });
    
    // 处理文件拖放
    dropArea.addEventListener('drop', handleDrop, false);
    
    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length) {
            handleFiles(files);
        }
    }
    
    // 处理文件选择
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFiles(e.target.files);
        }
    });
    
    // 点击上传区域触发文件选择
    dropArea.addEventListener('click', () => {
        fileInput.click();
    });
    
    // 处理文件
    function handleFiles(files) {
        const file = files[0];
        
        // 检查文件类型
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            alert('请上传Excel文件 (.xlsx 或 .xls)');
            return;
        }
        
        // 显示文件信息
        fileInfo.textContent = `文件名: ${file.name} | 大小: ${formatFileSize(file.size)}`;
        
        // 显示加载动画
        loadingOverlay.style.display = 'flex';
        
        // 读取文件
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                processExcel(data, file.name);
            } catch (error) {
                console.error('处理Excel文件时出错:', error);
                alert('处理文件时出错: ' + error.message);
                loadingOverlay.style.display = 'none';
            }
        };
        
        reader.onerror = function() {
            alert('读取文件时出错');
            loadingOverlay.style.display = 'none';
        };
        
        reader.readAsArrayBuffer(file);
    }
    
    // 处理Excel文件并提取链接
    function processExcel(data, fileName) {
        try {
            // 使用SheetJS库解析Excel文件，启用所有选项以确保能够提取所有类型的链接
            const workbook = XLSX.read(data, { 
                type: 'array', 
                cellFormula: true, 
                cellHtml: true,
                cellStyles: true,
                cellDates: true,
                cellNF: true,
                cellText: true,
                cellHyperlinks: true
            });
            
            console.log('工作簿信息:', workbook);
            
            // 清空之前的链接
            extractedLinks = [];
            
            // 遍历所有工作表
            workbook.SheetNames.forEach(sheetName => {
                console.log(`处理工作表: ${sheetName}`);
                const worksheet = workbook.Sheets[sheetName];
                
                // 检查工作表是否有范围定义
                if (!worksheet['!ref']) {
                    console.log(`工作表 ${sheetName} 没有范围定义`);
                    return;
                }
                
                const range = XLSX.utils.decode_range(worksheet['!ref']);
                console.log(`工作表范围: ${JSON.stringify(range)}`);
                
                // 方法1: 检查工作表中的超链接属性
                if (worksheet['!hyperlinks'] && Array.isArray(worksheet['!hyperlinks'])) {
                    console.log(`工作表 ${sheetName} 有 ${worksheet['!hyperlinks'].length} 个超链接`);
                    
                    worksheet['!hyperlinks'].forEach((hyperlink, idx) => {
                        console.log(`处理超链接 #${idx}:`, hyperlink);
                        
                        if (hyperlink && hyperlink.target) {
                            const { r, c } = hyperlink;
                            const cellAddress = XLSX.utils.encode_cell({ r, c });
                            const cell = worksheet[cellAddress];
                            
                            console.log(`单元格 ${cellAddress} 内容:`, cell);
                            
                            extractedLinks.push({
                                text: cell ? (cell.v || '') : '',
                                url: hyperlink.target || '',
                                sheet: sheetName,
                                cell: cellAddress,
                                type: 'direct'
                            });
                        }
                    });
                } else {
                    console.log(`工作表 ${sheetName} 没有 !hyperlinks 属性或不是数组`);
                }
                
                // 方法2: 遍历所有单元格查找超链接
                for (let R = range.s.r; R <= range.e.r; ++R) {
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                        const cell = worksheet[cellAddress];
                        
                        if (!cell) continue;
                        
                        // 检查单元格是否有超链接属性
                        if (cell.l && cell.l.Target) {
                            console.log(`单元格 ${cellAddress} 有直接超链接:`, cell.l);
                            
                            extractedLinks.push({
                                text: cell.v || '',
                                url: cell.l.Target,
                                sheet: sheetName,
                                cell: cellAddress,
                                type: 'cell_link'
                            });
                        }
                        
                        // 检查HYPERLINK函数
                        if (cell.f && cell.f.toString().toUpperCase().includes('HYPERLINK')) {
                            console.log(`单元格 ${cellAddress} 有HYPERLINK函数:`, cell.f);
                            
                            // 尝试从HYPERLINK函数中提取URL
                            const formula = cell.f.toString();
                            // 匹配各种格式的HYPERLINK函数
                            const matches = formula.match(/HYPERLINK\s*\(\s*"([^"]+)"/i) || 
                                           formula.match(/HYPERLINK\s*\(\s*'([^']+)'/i) ||
                                           formula.match(/HYPERLINK\s*\(\s*([^,"']+)/i);
                            
                            if (matches && matches[1]) {
                                extractedLinks.push({
                                    text: cell.v || '',
                                    url: matches[1],
                                    sheet: sheetName,
                                    cell: cellAddress,
                                    type: 'formula'
                                });
                            }
                        }
                        
                        // 检查HTML内容
                        if (cell.h) {
                            console.log(`单元格 ${cellAddress} 有HTML内容:`, cell.h);
                            
                            const html = cell.h.toString();
                            const matches = html.match(/<a[^>]+href=["']([^"']+)["'][^>]*>/g);
                            
                            if (matches) {
                                matches.forEach(match => {
                                    const urlMatch = match.match(/href=["']([^"']+)["']/);
                                    if (urlMatch && urlMatch[1]) {
                                        extractedLinks.push({
                                            text: cell.v || '',
                                            url: urlMatch[1],
                                            sheet: sheetName,
                                            cell: cellAddress,
                                            type: 'html'
                                        });
                                    }
                                });
                            }
                        }
                        
                        // 检查单元格值是否为URL格式
                        if (cell.v && typeof cell.v === 'string') {
                            const urlRegex = /^(https?:\/\/|www\.)[^\s]+\.[^\s]+/i;
                            if (urlRegex.test(cell.v)) {
                                console.log(`单元格 ${cellAddress} 包含URL文本:`, cell.v);
                                
                                let url = cell.v;
                                if (url.startsWith('www.')) {
                                    url = 'http://' + url;
                                }
                                
                                extractedLinks.push({
                                    text: cell.v,
                                    url: url,
                                    sheet: sheetName,
                                    cell: cellAddress,
                                    type: 'text_url'
                                });
                            }
                        }
                    }
                }
            });
            
            console.log('提取的链接总数:', extractedLinks.length);
            console.log('提取的链接:', extractedLinks);
            
            // 显示提取的链接
            displayLinks();
            
            // 隐藏加载动画
            loadingOverlay.style.display = 'none';
            
        } catch (error) {
            console.error('解析Excel文件时出错:', error);
            alert('解析文件时出错: ' + error.message);
            loadingOverlay.style.display = 'none';
        }
    }
    
    // 显示提取的链接
    function displayLinks() {
        // 清空表格
        linksTableBody.innerHTML = '';
        
        if (extractedLinks.length === 0) {
            // 没有找到链接
            linksTableBody.innerHTML = `
                <tr>
                    <td colspan="5" style="text-align: center; padding: 2rem;">
                        未找到任何超链接
                    </td>
                </tr>
            `;
            resultsSection.style.display = 'block';
            return;
        }
        
        // 添加链接到表格
        extractedLinks.forEach((link, index) => {
            const row = document.createElement('tr');
            
            row.innerHTML = `
                <td>${index + 1}</td>
                <td>${escapeHtml(link.text)}</td>
                <td><a href="${link.url}" target="_blank" rel="noopener noreferrer">${escapeHtml(link.url)}</a></td>
                <td>${escapeHtml(link.sheet)}</td>
                <td>${escapeHtml(link.cell)}</td>
            `;
            
            linksTableBody.appendChild(row);
        });
        
        // 显示结果区域
        resultsSection.style.display = 'block';
    }
    
    // 导出为TXT文件
    exportTxtBtn.addEventListener('click', () => {
        if (extractedLinks.length === 0) {
            alert('没有可导出的链接');
            return;
        }
        
        let txtContent = '提取的链接:\n\n';
        
        extractedLinks.forEach((link, index) => {
            txtContent += `${index + 1}. 文本: ${link.text}\n`;
            txtContent += `   URL: ${link.url}\n`;
            txtContent += `   工作表: ${link.sheet}\n`;
            txtContent += `   单元格: ${link.cell}\n\n`;
        });
        
        downloadFile(txtContent, 'excel_links.txt', 'text/plain');
    });
    
    // 导出为Excel文件
    exportExcelBtn.addEventListener('click', () => {
        if (extractedLinks.length === 0) {
            alert('没有可导出的链接');
            return;
        }
        
        // 创建工作簿
        const wb = XLSX.utils.book_new();
        
        // 准备数据
        const data = [
            ['序号', '链接文本', 'URL', '工作表', '单元格']
        ];
        
        extractedLinks.forEach((link, index) => {
            data.push([
                index + 1,
                link.text,
                link.url,
                link.sheet,
                link.cell
            ]);
        });
        
        // 创建工作表
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        // 添加超链接
        if (!ws['!hyperlinks']) ws['!hyperlinks'] = [];
        
        extractedLinks.forEach((link, index) => {
            ws['!hyperlinks'].push({
                r: index + 1,
                c: 2,
                target: link.url
            });
        });
        
        // 将工作表添加到工作簿
        XLSX.utils.book_append_sheet(wb, ws, '提取的链接');
        
        // 导出Excel文件
        XLSX.writeFile(wb, 'excel_links.xlsx');
    });
    
    // 辅助函数：格式化文件大小
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
    
    // 辅助函数：HTML转义
    function escapeHtml(text) {
        if (typeof text !== 'string') {
            return '';
        }
        
        return text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }
    
    // 辅助函数：下载文件
    function downloadFile(content, fileName, contentType) {
        const blob = new Blob([content], { type: contentType });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        a.click();
        
        setTimeout(() => {
            URL.revokeObjectURL(url);
        }, 100);
    }
}); 