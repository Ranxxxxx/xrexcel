import { Injectable } from '@angular/core';
import * as ExcelJS from 'exceljs';
import { TableStyleConfig } from '../models/table-style.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelUtilsService {

  /**
   * 读取Excel文件
   */
  async readExcelFile(file: File): Promise<{
    headers: string[];
    rawData: any[][];
    originalWorkbook: ExcelJS.Workbook;
    originalSheet: ExcelJS.Worksheet;
  }> {
    try {
      // 检查文件类型
      const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
        'application/vnd.ms-excel', // .xls
        'application/octet-stream' // 某些浏览器可能返回这个
      ];

      if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        throw new Error('不支持的文件格式，请上传 .xlsx 或 .xls 文件');
      }

      const arrayBuffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);

      // 读取第一个sheet
      const firstSheet = workbook.worksheets[0];
      if (!firstSheet) {
        throw new Error('Excel文件中没有找到工作表');
      }

      // 检查sheet是否有数据
      if (firstSheet.rowCount === 0) {
        throw new Error('工作表为空，没有数据');
      }

      // 读取第一行作为表头，并保留其物理列索引
      const headerRow = firstSheet.getRow(1);
      const headers: string[] = [];
      const colToHeaderMap = new Map<number, string>(); // 物理列号 -> 表头名

      headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        let headerText = this.getCellValue(cell);
        if (headerText) {
          headerText = String(headerText).trim();
          headers.push(headerText);
          colToHeaderMap.set(colNumber, headerText);
        }
      });

      if (headers.length === 0) {
        throw new Error('未找到有效的表头，请确保第一行包含表头数据');
      }

      // 读取所有数据行
      const allData: any[][] = [];
      // 添加表头行
      allData.push(headers);

      // 读取数据行（从第2行开始）
      for (let rowNumber = 2; rowNumber <= firstSheet.rowCount; rowNumber++) {
        const row = firstSheet.getRow(rowNumber);
        if (!row || row.cellCount === 0) continue;

        const rowData: any[] = [];
        let hasData = false;

        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          const cellValue = cell.value;
          let cellText = '';

          // 优先检查是否有公式：如果有公式，使用公式而不是计算结果
          if (cell.formula) {
            // 单元格包含公式，使用公式字符串
            cellText = cell.formula;
          } else if (cellValue !== null && cellValue !== undefined) {
            if (typeof cellValue === 'string') {
              cellText = cellValue.trim();
            } else if (typeof cellValue === 'number') {
              cellText = String(cellValue);
            } else if (cellValue instanceof Date) {
              cellText = cellValue.toLocaleDateString();
            } else if (typeof cellValue === 'object') {
              // 处理对象类型的值（如公式结果、超链接等）
              // 检查对象中是否包含公式
              if ('formula' in cellValue && (cellValue as any).formula) {
                cellText = String((cellValue as any).formula);
              } else if (cell.text) {
                // 如果没有公式，使用文本值（可能是超链接文本）
                cellText = cell.text.trim();
              } else if ('text' in cellValue) {
                // 如果是超链接对象，尝试获取文本
                cellText = String((cellValue as any).text || '').trim();
              } else {
                // 最后才使用计算结果（如果没有公式）
                if (cell.result !== null && cell.result !== undefined) {
                  if (typeof cell.result === 'number') {
                    cellText = String(cell.result);
                  } else if (typeof cell.result === 'string') {
                    cellText = cell.result.trim();
                  } else {
                    cellText = String(cell.result);
                  }
                } else {
                  cellText = '';
                }
              }
            } else {
              cellText = String(cellValue).trim();
            }
          } else {
            // 如果value为空，检查是否有公式
            if (cell.formula) {
              cellText = cell.formula;
            } else if (cell.text) {
              cellText = cell.text.trim();
            } else {
              cellText = '';
            }
          }

          rowData.push(cellText);
          if (cellText) hasData = true;
        });

        // 如果行数据不足表头数量，补齐空值
        while (rowData.length < headers.length) {
          rowData.push('');
        }

        // 只添加有数据的行
        if (hasData) {
          allData.push(rowData);
        }
      }

      return {
        headers,
        rawData: allData,
        originalWorkbook: workbook,
        originalSheet: firstSheet
      };
    } catch (error: any) {
      // 重新抛出错误以便上层处理
      if (error instanceof Error) {
        throw error;
      }
      throw new Error(`读取文件时发生错误：${String(error)}`);
    }
  }

  /**
   * 获取单元格值
   */
  getCellValue(cell: ExcelJS.Cell): any {
    if (cell.value === null || cell.value === undefined) {
      return cell.text || '';
    }
    if (typeof cell.value === 'object' && 'text' in cell.value) {
      return (cell.value as any).text || '';
    }
    return cell.value;
  }

  /**
   * 将十六进制颜色转换为ARGB格式
   */
  hexToArgb(hex: string): string {
    hex = hex.replace('#', '');
    if (hex.length === 3) {
      hex = hex.split('').map(c => c + c).join('');
    }
    return 'FF' + hex.toUpperCase();
  }

  /**
   * 获取Excel列名（1 -> A, 2 -> B, 27 -> AA）
   */
  getExcelColumnName(columnNumber: number): string {
    let result = '';
    while (columnNumber > 0) {
      columnNumber--;
      result = String.fromCharCode(65 + (columnNumber % 26)) + result;
      columnNumber = Math.floor(columnNumber / 26);
    }
    return result;
  }

  /**
   * 自动调整列宽，确保内容不折叠、不超出
   */
  autoFitColumns(sheet: ExcelJS.Worksheet, columnCount: number) {
    // Excel列宽单位：1个字符宽度 ≈ 7像素（对于默认字体）
    // 左右各20像素 = 40像素 ≈ 5.7个字符宽度
    const paddingChars = 40 / 7; // 左右各20像素的字符宽度

    for (let col = 1; col <= columnCount; col++) {
      const column = sheet.getColumn(col);
      let maxLength = 0;

      // 遍历该列的所有单元格，找出最大内容长度
      sheet.eachRow((row, rowNumber) => {
        const cell = row.getCell(col);
        let cellText = '';
        let isNumber = false;

        if (cell.value !== null && cell.value !== undefined) {
          if (typeof cell.value === 'number') {
            // 数字类型：按数字格式计算宽度
            cellText = String(cell.value);
            isNumber = true;
          } else if (typeof cell.value === 'string') {
            cellText = cell.value;
          } else if (cell.value instanceof Date) {
            cellText = cell.value.toLocaleDateString();
          } else if (typeof cell.value === 'object' && 'text' in cell.value) {
            cellText = (cell.value as any).text || '';
          } else {
            cellText = String(cell.value);
          }
        } else if (cell.text) {
          cellText = cell.text;
        }

        // 计算文本长度
        let textLength: number;
        if (isNumber) {
          // 数字类型：按实际字符数计算
          textLength = cellText.length;
        } else {
          // 文本类型：中文字符按2个字符宽度计算
          textLength = this.calculateTextWidth(cellText);
        }
        maxLength = Math.max(maxLength, textLength);
      });

      // 设置列宽 = 最大内容宽度 + 左右各20像素的空间
      const baseWidth = maxLength + paddingChars;
      column.width = Math.max(8, baseWidth); // 最小宽度8
    }
  }

  /**
   * 计算文本宽度（中文字符按2个字符宽度计算）
   */
  calculateTextWidth(text: string): number {
    if (!text) return 0;
    let width = 0;
    for (let i = 0; i < text.length; i++) {
      const char = text[i];
      // 判断是否为中文字符、日文、韩文等宽字符
      if (/[\u4e00-\u9fa5\u3040-\u309f\u30a0-\u30ff\uac00-\ud7af]/.test(char)) {
        width += 2; // 中文字符占2个字符宽度
      } else {
        width += 1; // 英文字符占1个字符宽度
      }
    }
    return width;
  }

  /**
   * 转换公式：将原始数据中的公式引用转换为分类表中的引用
   * 重要：只转换当前工作表的单元格引用，不转换其他工作表的引用（如 '汇总表'!A1）
   */
  convertFormula(
    formula: string,
    originalRow: number,
    categoryRow: number,
    originalCol: number,
    categoryCol: number,
    categoryHeaders: string[],
    headerIndexMap: Map<string, number>,
    originalSheet: ExcelJS.Worksheet | null,
    categoryDataRowCount: number = 0 // 分类表数据行数，用于限制行号范围
  ): string | null {
    try {
      // 匹配单元格引用模式：$?列字母$?行号 或 范围引用
      // 但不匹配前面有工作表引用的（如 '工作表名'!A1）
      const cellRefPattern = /(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|\$?[A-Z]+:\$?[A-Z]+|\$?\d+:\$?\d+)/gi;

      let convertedFormula = formula;
      const matches = Array.from(formula.matchAll(cellRefPattern));

      // 从后往前处理，避免索引变化
      for (let i = matches.length - 1; i >= 0; i--) {
        const match = matches[i];
        const fullRef = match[1];
        const matchIndex = match.index!;

        // 检查前面是否有工作表引用（如 '汇总表'! 或 工作表名!）
        // 向前查找最多50个字符，查找 '...'! 或 工作表名! 模式
        const beforeMatch = formula.substring(Math.max(0, matchIndex - 50), matchIndex);

        // 检查是否有工作表引用模式：'工作表名'! 或 工作表名!（工作表名不包含特殊字符）
        const hasSheetRef = /'[^']*'!\s*$/.test(beforeMatch) || /[A-Za-z0-9_]+!\s*$/.test(beforeMatch);

        if (hasSheetRef) {
          // 如果前面有工作表引用，跳过这个匹配（不转换其他工作表的引用）
          continue;
        }

        // 处理范围引用（如 A1:B2）
        if (fullRef.includes(':')) {
          const parts = fullRef.split(':');
          const convertedParts = parts.map(part => this.convertSingleCellRef(
            part,
            originalRow,
            categoryRow,
            categoryHeaders,
            headerIndexMap,
            originalSheet,
            categoryDataRowCount
          ));

          if (convertedParts.some(p => p === null)) return null;

          convertedFormula = convertedFormula.substring(0, matchIndex) +
                             convertedParts.join(':') +
                             convertedFormula.substring(matchIndex + fullRef.length);
        } else {
          // 处理单个单元格引用
          const converted = this.convertSingleCellRef(
            fullRef,
            originalRow,
            categoryRow,
            categoryHeaders,
            headerIndexMap,
            originalSheet,
            categoryDataRowCount
          );
          if (converted === null) return null;

          convertedFormula = convertedFormula.substring(0, matchIndex) +
                             converted +
                             convertedFormula.substring(matchIndex + fullRef.length);
        }
      }
      return convertedFormula;
    } catch (e) {
      console.error('转换公式失败:', e, formula);
      return null;
    }
  }

  /**
   * 转换单个单元格引用
   */
  private convertSingleCellRef(
    ref: string,
    originalRow: number,
    categoryRow: number,
    categoryHeaders: string[],
    headerIndexMap: Map<string, number>,
    originalSheet: ExcelJS.Worksheet | null,
    categoryDataRowCount: number = 0
  ): string | null {
    // 提取列和行的绝对/相对引用标记
    const isAbsoluteCol = ref.startsWith('$');
    const parts = ref.split('$');
    let colPart = '';
    let rowPart = '';

    // 解析列部分（字母）
    const colMatch = ref.match(/[A-Z]+/i);
    if (colMatch) {
      colPart = colMatch[0].toUpperCase();
      // 检查列部分是否有绝对引用标记
      const colIndex = ref.indexOf(colPart);
      const beforeCol = ref.substring(0, colIndex);
      const isColAbsolute = beforeCol.includes('$') || beforeCol === '$';

      // 将列字母转换为列号
      let originalColNum = 0;
      for (let i = 0; i < colPart.length; i++) {
        originalColNum = originalColNum * 26 + (colPart.charCodeAt(i) - 64);
      }

      // 从原表获取该列对应的表头名称
      let headerName = '';
      if (originalSheet && originalColNum > 0) {
        const headerCell = originalSheet.getRow(1).getCell(originalColNum);
        if (headerCell && headerCell.value !== null && headerCell.value !== undefined) {
          headerName = String(headerCell.value).trim();
        }
      }

      // 如果无法从原表获取表头名，尝试通过 headerIndexMap 查找
      if (!headerName && headerIndexMap) {
        const headerIndex = originalColNum - 1; // 转换为0-based索引
        for (const [header, index] of headerIndexMap.entries()) {
          if (index === headerIndex) {
            headerName = header;
            break;
          }
        }
      }

      // 在分类表中查找该表头对应的列位置
      const newColIdx = categoryHeaders.indexOf(headerName);
      if (newColIdx === -1) {
        // 该表头不在分类表中，返回 null 表示无法转换
        return null;
      }

      // 构建新的列引用
      const newColRef = this.getExcelColumnName(newColIdx + 1);
      colPart = (isColAbsolute ? '$' : '') + newColRef;
    }

    // 解析行部分（数字）
    const rowMatch = ref.match(/\d+/);
    if (rowMatch) {
      const originalRowNum = parseInt(rowMatch[0]);
      const rowIndex = ref.indexOf(rowMatch[0]);
      const beforeRow = ref.substring(0, rowIndex);
      const isRowAbsolute = beforeRow.includes('$') || (beforeRow.endsWith('$') && !beforeRow.endsWith('$$'));

      let newRowNum: number;
      if (originalRowNum === 1) {
        // 引用表头行 -> 分类表第3行（表头行）
        newRowNum = 3;
      } else {
        // 引用数据行 -> 根据相对偏移计算
        const rowOffset = originalRowNum - originalRow;
        newRowNum = categoryRow + rowOffset;
        // 确保行号在有效范围内：
        // - 最小行号：第3行（表头行）
        // - 最大行号：第(3+categoryDataRowCount)行（最后一行数据，不包括表尾行）
        if (newRowNum < 3) {
          return null;
        }
        // 如果转换后的行号超出了数据行范围，返回 null（避免引用表尾行造成循环引用）
        if (categoryDataRowCount > 0 && newRowNum > 3 + categoryDataRowCount) {
          return null;
        }
      }
      rowPart = (isRowAbsolute ? '$' : '') + newRowNum;
    }

    // 组合新的单元格引用
    return colPart + rowPart;
  }

  /**
   * 应用表格样式到单元格
   */
  applyCellStyle(
    cell: ExcelJS.Cell,
    style: TableStyleConfig,
    cellType: 'title' | 'header' | 'data' | 'total'
  ) {
    const fontConfig = this.getFontConfig(style, cellType);
    const fillConfig = this.getFillConfig(style, cellType);
    const alignmentConfig = this.getAlignmentConfig(cellType);

    cell.font = fontConfig;
    if (fillConfig) {
      cell.fill = fillConfig;
    }
    cell.alignment = alignmentConfig;
    cell.border = {
      top: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
      left: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
      bottom: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
      right: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } }
    };
  }

  /**
   * 获取字体配置
   */
  private getFontConfig(style: TableStyleConfig, cellType: 'title' | 'header' | 'data' | 'total') {
    switch (cellType) {
      case 'title':
        return {
          name: style.fontFamily,
          size: style.titleFontSize,
          bold: style.titleFontBold,
          color: { argb: this.hexToArgb(style.titleFontColor) }
        };
      case 'header':
        return {
          name: style.fontFamily,
          size: style.headerFontSize,
          bold: style.headerFontBold,
          color: { argb: this.hexToArgb(style.headerFontColor) }
        };
      case 'data':
      case 'total':
        return {
          name: style.fontFamily,
          size: style.dataFontSize,
          bold: cellType === 'total' ? true : style.dataFontBold,
          color: { argb: this.hexToArgb(style.dataFontColor) }
        };
    }
  }

  /**
   * 获取填充配置
   */
  private getFillConfig(style: TableStyleConfig, cellType: 'title' | 'header' | 'data' | 'total') {
    switch (cellType) {
      case 'title':
        return {
          type: 'pattern' as const,
          pattern: 'solid' as const,
          fgColor: { argb: this.hexToArgb(style.titleColor) }
        };
      case 'header':
        return {
          type: 'pattern' as const,
          pattern: 'solid' as const,
          fgColor: { argb: this.hexToArgb(style.headerColor) }
        };
      case 'total':
        return {
          type: 'pattern' as const,
          pattern: 'solid' as const,
          fgColor: { argb: this.hexToArgb(style.totalColor) }
        };
      case 'data':
        return undefined;
    }
  }

  /**
   * 获取对齐配置
   */
  private getAlignmentConfig(cellType: 'title' | 'header' | 'data' | 'total') {
    switch (cellType) {
      case 'title':
        return { horizontal: 'center' as const, vertical: 'middle' as const, wrapText: false };
      case 'header':
        return { horizontal: 'center' as const, vertical: 'middle' as const, wrapText: false };
      case 'data':
        return { horizontal: 'left' as const, vertical: 'middle' as const, wrapText: false };
      case 'total':
        return { horizontal: 'right' as const, vertical: 'middle' as const, wrapText: false };
    }
  }
}

