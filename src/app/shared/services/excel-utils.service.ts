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

      const isXlsx = file.name.match(/\.xlsx$/i);
      const isXls = file.name.match(/\.xls$/i);

      // 检查文件大小
      if (file.size === 0) {
        throw new Error('文件为空，请选择一个有效的Excel文件');
      }

      const arrayBuffer = await file.arrayBuffer();

      // 验证文件签名（.xlsx文件是ZIP格式，以PK开头）
      if (isXlsx) {
        const uint8Array = new Uint8Array(arrayBuffer);
        // .xlsx文件应该以ZIP文件签名开头：PK (50 4B)
        if (uint8Array.length < 4 || uint8Array[0] !== 0x50 || uint8Array[1] !== 0x4B) {
          throw new Error('文件格式不正确：该文件不是有效的 .xlsx 文件。请确保文件未损坏，并且确实是Excel文件。');
        }
      }

      const workbook = new ExcelJS.Workbook();

      try {
        await workbook.xlsx.load(arrayBuffer);
      } catch (loadError: any) {
        // 处理ExcelJS加载错误
        if (loadError?.message?.includes('zip') || loadError?.message?.includes('central directory')) {
          throw new Error('无法读取Excel文件：文件可能已损坏或格式不正确。\n\n请尝试：\n1. 在Excel中打开并重新保存文件\n2. 确保文件是 .xlsx 格式（不是 .xls）\n3. 检查文件是否完整下载');
        }
        if (isXls && !isXlsx) {
          throw new Error('检测到 .xls 格式文件。ExcelJS主要支持 .xlsx 格式。\n\n请将文件转换为 .xlsx 格式：\n1. 在Excel中打开文件\n2. 选择"另存为"\n3. 选择"Excel工作簿(.xlsx)"格式');
        }
        throw loadError;
      }

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
   * 读取Excel工作表的表头
   * @param sheet Excel工作表
   * @param headerRow 表头所在行号（从1开始）
   * @param includeEmpty 是否包含空单元格
   * @returns 表头数组
   */
  readSheetHeaders(sheet: ExcelJS.Worksheet, headerRow: number = 1, includeEmpty: boolean = false): string[] {
    const headers: string[] = [];
    const headerRowObj = sheet.getRow(headerRow);
    if (!headerRowObj) return headers;

    headerRowObj.eachCell({ includeEmpty }, (cell) => {
      // 对于表头，使用cell.text获取显示的文本，而不是cell.value（避免日期对象等被错误转换）
      const headerText = cell.text || (cell.value ? String(cell.value).trim() : '');
      if (headerText && headerText.trim() && !headers.includes(headerText.trim())) {
        headers.push(headerText.trim());
      }
    });

    return headers;
  }

  /**
   * 读取多个Excel工作表的表头（合并去重）
   * @param workbook Excel工作簿
   * @param sheetNames 工作表名称数组
   * @param headerRow 表头所在行号（从1开始）
   * @param includeEmpty 是否包含空单元格
   * @returns 表头数组
   */
  readMultipleSheetHeaders(
    workbook: ExcelJS.Workbook,
    sheetNames: string[],
    headerRow: number = 1,
    includeEmpty: boolean = false
  ): string[] {
    const headers: string[] = [];
    const headerSet = new Set<string>();

    for (const sheetName of sheetNames) {
      const sheet = workbook.getWorksheet(sheetName);
      if (sheet) {
        const sheetHeaders = this.readSheetHeaders(sheet, headerRow, includeEmpty);
        sheetHeaders.forEach(header => {
          if (!headerSet.has(header)) {
            headerSet.add(header);
            headers.push(header);
          }
        });
      }
    }

    return headers;
  }

  /**
   * 读取表头并创建列号到表头名的映射
   * @param sheet Excel工作表
   * @param headerRow 表头所在行号（从1开始）
   * @param includeEmpty 是否包含空单元格
   * @returns 列号到表头名的映射（Map<列号, 表头名>）
   */
  readHeadersMap(
    sheet: ExcelJS.Worksheet,
    headerRow: number = 1,
    includeEmpty: boolean = true
  ): Map<number, string> {
    const headersMap = new Map<number, string>();
    const headerRowObj = sheet.getRow(headerRow);
    if (!headerRowObj) return headersMap;

    headerRowObj.eachCell({ includeEmpty }, (cell, colNumber) => {
      // 对于表头，使用cell.text获取显示的文本，而不是cell.value（避免日期对象等被错误转换）
      const headerText = cell.text || (cell.value ? String(cell.value).trim() : '');
      if (headerText && headerText.trim()) {
        headersMap.set(colNumber, headerText.trim());
      }
    });

    return headersMap;
  }

  /**
   * 获取单元格值（包括公式）
   * @param cell Excel单元格
   * @param sheet Excel工作表（可选，用于公式计算）
   * @param rowNumber 行号（可选）
   * @param colNumber 列号（可选）
   * @returns 单元格值或公式对象
   */
  getCellValueWithFormula(
    cell: ExcelJS.Cell,
    sheet?: ExcelJS.Worksheet,
    rowNumber?: number,
    colNumber?: number
  ): any {
    // 检查是否是公式
    if (cell.formula) {
      return { formula: cell.formula };
    }

    // 检查value是否是公式对象
    if (typeof cell.value === 'object' && cell.value !== null && 'formula' in cell.value) {
      return cell.value;
    }

    // 返回普通值
    return this.getCellValue(cell);
  }

  /**
   * 从单元格引用中提取表头名称
   * @param ref 单元格引用（如 "A1", "$B$2"）
   * @param headersMap 列号到表头名的映射
   * @returns 表头名称，如果找不到则返回null
   */
  extractHeaderNameFromRef(ref: string, headersMap: Map<number, string>): string | null {
    const colMatch = ref.match(/[A-Z]+/i);
    if (!colMatch) return null;

    const colPart = colMatch[0].toUpperCase();
    // 将列字母转换为列号（1-based）
    let originalColNum = 0;
    for (let i = 0; i < colPart.length; i++) {
      originalColNum = originalColNum * 26 + (colPart.charCodeAt(i) - 64);
    }

    // 查找原始列对应的表头名称
    if (headersMap.has(originalColNum)) {
      return headersMap.get(originalColNum)!;
    }

    return null;
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
    categoryDataRowCount: number = 0, // 分类表数据行数，用于限制行号范围
    categoryStartRow: number = 0 // 当前分类的起始行号（单表输出模式使用）
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
            categoryDataRowCount,
            categoryStartRow
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
            categoryDataRowCount,
            categoryStartRow
          );
          if (converted === null) return null;

          convertedFormula = convertedFormula.substring(0, matchIndex) +
                             converted +
                             convertedFormula.substring(matchIndex + fullRef.length);
        }
      }
      return convertedFormula;
    } catch (e) {
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
    categoryDataRowCount: number = 0,
    categoryStartRow: number = 0
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

        // 确定有效行号范围
        let minRow = 3; // 最小行号：表头行
        let maxRow = 0; // 最大行号：根据模式确定

        if (categoryStartRow > 0) {
          // 单表输出模式：使用当前分类的起始行号和结束行号
          minRow = categoryStartRow;
          maxRow = categoryStartRow + categoryDataRowCount - 1;
        } else {
          // 多表输出模式：使用固定的起始行号3
          maxRow = 3 + categoryDataRowCount;
        }

        // 确保行号在有效范围内
        if (newRowNum < minRow) {
          return null;
        }
        // 如果转换后的行号超出了数据行范围，返回 null（避免引用表尾行造成循环引用）
        if (categoryDataRowCount > 0 && maxRow > 0 && newRowNum > maxRow) {
          return null;
        }
      }
      rowPart = (isRowAbsolute ? '$' : '') + newRowNum;
    }

    // 组合新的单元格引用
    return colPart + rowPart;
  }

  /**
   * 提取并转换公式：从原始Excel单元格中提取公式并转换为新表格的公式引用
   * @param originalCell 原始Excel单元格
   * @param originalRowNumber 原始行号
   * @param currentRow 当前新表格的行号
   * @param headerIndex 表头在原始数据中的索引
   * @param colIndex 列在新表格中的索引（1-based）
   * @param categoryHeaders 分类表表头数组
   * @param headerIndexMap 表头索引映射
   * @param originalSheet 原始工作表（用于公式转换）
   * @param categoryDataRowCount 分类表数据行数
   * @param categoryStartRow 当前分类的起始行号（单表输出模式使用）
   * @returns 转换后的公式字符串，如果不是公式或转换失败则返回 null
   */
  extractAndConvertFormula(
    originalCell: ExcelJS.Cell,
    originalRowNumber: number,
    currentRow: number,
    headerIndex: number,
    colIndex: number,
    categoryHeaders: string[],
    headerIndexMap: Map<string, number>,
    originalSheet: ExcelJS.Worksheet | null,
    categoryDataRowCount: number = 0,
    categoryStartRow: number = 0
  ): string | null {
    // 检查是否是公式
    let formulaText = '';
    if (originalCell.formula) {
      formulaText = originalCell.formula;
    } else if (typeof originalCell.value === 'object' && originalCell.value !== null && 'formula' in originalCell.value) {
      formulaText = (originalCell.value as any).formula;
    }

    // 如果不是公式，返回 null
    if (!formulaText) {
      return null;
    }

    // 转换公式引用
    return this.convertFormula(
      formulaText,
      originalRowNumber,
      currentRow,
      headerIndex,
      colIndex,
      categoryHeaders,
      headerIndexMap,
      originalSheet,
      categoryDataRowCount,
      categoryStartRow
    );
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

  /**
   * 读取Excel文件（多Sheet版本，用于单表格处理）
   */
  async readExcelFileMultiSheet(file: File): Promise<{
    sheetNames: string[];
    originalWorkbook: ExcelJS.Workbook;
  }> {
    try {
      // 检查文件大小
      if (file.size === 0) {
        throw new Error('文件为空，请选择一个有效的Excel文件');
      }

      const isXlsx = file.name.match(/\.xlsx$/i);
      const isXls = file.name.match(/\.xls$/i);

      const arrayBuffer = await file.arrayBuffer();

      // 验证文件签名（.xlsx文件是ZIP格式，以PK开头）
      if (isXlsx) {
        const uint8Array = new Uint8Array(arrayBuffer);
        // .xlsx文件应该以ZIP文件签名开头：PK (50 4B)
        if (uint8Array.length < 4 || uint8Array[0] !== 0x50 || uint8Array[1] !== 0x4B) {
          throw new Error('文件格式不正确：该文件不是有效的 .xlsx 文件。请确保文件未损坏，并且确实是Excel文件。');
        }
      }

      const workbook = new ExcelJS.Workbook();

      try {
        await workbook.xlsx.load(arrayBuffer);
      } catch (loadError: any) {
        // 处理ExcelJS加载错误
        if (loadError?.message?.includes('zip') || loadError?.message?.includes('central directory')) {
          throw new Error('无法读取Excel文件：文件可能已损坏或格式不正确。\n\n请尝试：\n1. 在Excel中打开并重新保存文件\n2. 确保文件是 .xlsx 格式（不是 .xls）\n3. 检查文件是否完整下载');
        }
        if (isXls && !isXlsx) {
          throw new Error('检测到 .xls 格式文件。ExcelJS主要支持 .xlsx 格式。\n\n请将文件转换为 .xlsx 格式：\n1. 在Excel中打开文件\n2. 选择"另存为"\n3. 选择"Excel工作簿(.xlsx)"格式');
        }
        throw loadError;
      }

      // 获取所有sheet名称
      const sheets = workbook.worksheets.map(sheet => sheet.name);
      if (sheets.length === 0) {
        throw new Error('Excel文件中没有找到工作表');
      }

      return {
        sheetNames: sheets,
        originalWorkbook: workbook
      };
    } catch (error: any) {
      if (error instanceof Error) {
        throw error;
      }
      throw new Error(`读取文件时发生错误：${String(error)}`);
    }
  }

  /**
   * 生成默认预览数据
   */
  getDefaultPreviewData(): any[][] {
    const headers = ['列1', '列2', '列3', '列4'];
    const previewData: any[][] = [];

    // 添加标题行（第一行，跨所有列）
    const titleRow = ['示例表格标题', '', '', ''];
    previewData.push(titleRow);

    // 添加表头行
    previewData.push(headers);

    // 添加示例数据行
    for (let i = 0; i < 3; i++) {
      const row = headers.map((_, index) => `示例数据${index + 1}-${i + 1}`);
      previewData.push(row);
    }

    // 添加合计行
    const totalRow = ['合计', ...headers.slice(1).map(() => '100.00')];
    previewData.push(totalRow);

    return previewData;
  }

  /**
   * 生成带进度条的文件下载
   */
  async downloadFileWithProgress(
    generateFn: () => Promise<Blob>,
    fileName: string,
    progressCallback: (progress: number) => void
  ): Promise<void> {
    const startTime = Date.now();
    const minDuration = 1000; // 最少1秒

    // 开始生成Excel（返回blob，不直接下载）
    const generatePromise = generateFn();

    // 更新进度条
    const progressInterval = setInterval(() => {
      const elapsed = Date.now() - startTime;
      if (elapsed < minDuration) {
        // 在1秒内，进度条从0到90%
        const progress = Math.min(90, (elapsed / minDuration) * 90);
        progressCallback(progress);
      } else {
        // 1秒后，等待生成完成
        progressCallback(95);
      }
    }, 50);

    // 等待生成完成
    const blob = await generatePromise;

    // 确保至少1秒
    const elapsed = Date.now() - startTime;
    if (elapsed < minDuration) {
      await new Promise(resolve => setTimeout(resolve, minDuration - elapsed));
    }

    // 完成进度条
    clearInterval(progressInterval);
    progressCallback(100);

    // 等待进度条动画完成后再下载（等待一小段时间确保进度条显示100%）
    await new Promise(resolve => setTimeout(resolve, 200));

    // 现在触发下载
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${fileName}.xlsx`;
    link.click();
    window.URL.revokeObjectURL(url);
  }

  /**
   * 生成默认输出文件名（基于原始文件名和时间戳）
   */
  generateDefaultFileName(originalFileName: string, suffix: string = ''): string {
    const nameWithoutExt = originalFileName.replace(/\.[^/.]+$/, '');
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    const timeString = `${year}${month}${day}_${hours}${minutes}${seconds}`;
    return suffix ? `${nameWithoutExt}_${suffix}_${timeString}` : `${nameWithoutExt}_${timeString}`;
  }
}

