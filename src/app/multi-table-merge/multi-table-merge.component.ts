import { Component, signal, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { RouterLink } from '@angular/router';
import { MatToolbarModule } from '@angular/material/toolbar';
import { MatIconModule } from '@angular/material/icon';
import { MatButtonModule } from '@angular/material/button';
import { MatCardModule } from '@angular/material/card';
import { MatChipsModule } from '@angular/material/chips';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatSelectModule } from '@angular/material/select';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { MatDividerModule } from '@angular/material/divider';
import { MatStepperModule, MatStepper } from '@angular/material/stepper';
import { MatProgressBarModule } from '@angular/material/progress-bar';
import { DragDropModule, CdkDragDrop, moveItemInArray } from '@angular/cdk/drag-drop';
import { TableStyleConfig, DEFAULT_TABLE_STYLE } from '../shared/models/table-style.model';
import { TableStylePreviewComponent } from '../shared/components/table-style-preview/table-style-preview.component';
import { TableStyleStorageService } from '../shared/services/table-style-storage.service';
import { FileUploadComponent } from '../shared/components/file-upload/file-upload.component';
import { ConfirmDialogComponent } from '../shared/components/confirm-dialog/confirm-dialog.component';
import { PrivacyNoticeComponent } from '../shared/components/privacy-notice/privacy-notice.component';
import { ExcelUtilsService } from '../shared/services/excel-utils.service';
import * as ExcelJS from 'exceljs';

@Component({
  selector: 'app-multi-table-merge',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    RouterLink,
    MatToolbarModule,
    MatIconModule,
    MatButtonModule,
    MatCardModule,
    MatChipsModule,
    MatFormFieldModule,
    MatInputModule,
    MatSelectModule,
    MatCheckboxModule,
    MatDividerModule,
    MatStepperModule,
    MatProgressBarModule,
    DragDropModule,
    TableStylePreviewComponent,
    FileUploadComponent,
    ConfirmDialogComponent,
    PrivacyNoticeComponent
  ],
  templateUrl: './multi-table-merge.component.html',
  styleUrl: './multi-table-merge.component.scss'
})
export class MultiTableMergeComponent {
  @ViewChild('stepper', { static: false }) stepper?: MatStepper;

  tableStyle = signal<TableStyleConfig>({ ...DEFAULT_TABLE_STYLE });
  previewExpanded = signal<boolean>(false); // 预览区域展开状态，默认折叠
  previewData = signal<any[][]>([]); // 预览数据，待后续实现多表合并功能时填充
  showResetButton = signal<boolean>(false); // 是否显示重置按钮

  // Step 1: 文件选择相关
  baseFile = signal<File | null>(null); // 基础文件
  dataSourceFile = signal<File | null>(null); // 数据源文件
  isUploadingBaseFile = signal<boolean>(false);
  isUploadingDataSourceFile = signal<boolean>(false);
  baseWorkbook = signal<ExcelJS.Workbook | null>(null); // 基础文件工作簿
  dataSourceWorkbook = signal<ExcelJS.Workbook | null>(null); // 数据源文件工作簿

  // Step 2: 设置合并目标
  baseSheetNames = signal<string[]>([]); // 基础文件的sheet列表
  dataSourceSheetNames = signal<string[]>([]); // 数据源文件的sheet列表
  selectedBaseSheets = signal<string[]>([]); // 选中的基础文件sheet（多选）
  selectedDataSourceSheets = signal<string[]>([]); // 选中的数据源文件sheet（多选）
  baseHeaderRow = signal<number>(1); // 基础文件表头所在行数
  dataSourceHeaderRow = signal<number>(1); // 数据源文件表头所在行数

  // 表头相关
  baseHeaders = signal<string[]>([]); // 基础文件读取到的表头列表
  dataSourceHeaders = signal<string[]>([]); // 数据源文件读取到的表头列表
  selectedHeaders = signal<string[]>([]); // 选中的表头（用于排序）
  sortedHeaders = signal<string[]>([]); // 排序后的表头列表

  // Step3: 表尾功能配置
  footerFunctions = signal<Array<{ type: '合计' | '平均值', header: string, id: string }>>([]);
  newFooterType = signal<'合计' | '平均值'>('合计'); // Step3新增表尾功能类型
  newFooterHeader = signal<string>(''); // Step3新增表尾功能关联的表头

  // 确认对话框相关
  showConfirmDialog = signal<boolean>(false); // 显示确认对话框
  outputFileName = signal<string>(''); // 输出文件名
  generationProgress = signal<number>(0); // 生成进度（0-100）
  isProcessing = signal<boolean>(false); // 处理中状态


  constructor(
    private tableStyleStorage: TableStyleStorageService,
    private excelUtils: ExcelUtilsService
  ) {
    // 从 localStorage 加载表格风格配置（所有模块共享）
    this.loadTableStyleFromStorage();
  }

  updateStyle(key: keyof TableStyleConfig, value: any) {
    this.tableStyle.update(config => {
      const newConfig = { ...config, [key]: value };
      // 保存到 localStorage（所有模块共享）
      this.tableStyleStorage.saveTableStyle(newConfig);
      this.showResetButton.set(true);
      return newConfig;
    });
  }

  // 从 localStorage 加载表格风格配置
  private loadTableStyleFromStorage() {
    const savedStyle = this.tableStyleStorage.loadTableStyle();
    if (savedStyle) {
      this.tableStyle.set(savedStyle);
      this.showResetButton.set(true);
    } else {
      this.tableStyle.set({ ...DEFAULT_TABLE_STYLE });
      this.showResetButton.set(false);
    }
  }

  // 重置为默认配置
  resetTableStyleToDefault() {
    this.tableStyleStorage.resetToDefault();
    this.tableStyle.set({ ...DEFAULT_TABLE_STYLE });
    this.showResetButton.set(false);
  }

  // Step 1: 处理基础文件选择
  async onBaseFileSelected(file: File) {
    this.baseFile.set(file);
    this.isUploadingBaseFile.set(true);
    try {
      const result = await this.excelUtils.readExcelFileMultiSheet(file);
      this.baseSheetNames.set(result.sheetNames);
      this.baseWorkbook.set(result.originalWorkbook);
      // 重置step2的选择状态
      this.selectedBaseSheets.set([]);
      this.baseHeaderRow.set(1);
      this.baseHeaders.set([]);
      this.selectedHeaders.set([]);
      this.sortedHeaders.set([]);
    } catch (error: any) {
      console.error('读取基础文件失败:', error);
      alert(error.message || '读取基础文件失败');
      this.baseFile.set(null);
      this.baseSheetNames.set([]);
      this.baseWorkbook.set(null);
    } finally {
      this.isUploadingBaseFile.set(false);
    }
  }

  // Step 1: 处理数据源文件选择
  async onDataSourceFileSelected(file: File) {
    this.dataSourceFile.set(file);
    this.isUploadingDataSourceFile.set(true);
    try {
      const result = await this.excelUtils.readExcelFileMultiSheet(file);
      this.dataSourceSheetNames.set(result.sheetNames);
      this.dataSourceWorkbook.set(result.originalWorkbook);
      // 重置step2的选择状态
      this.selectedDataSourceSheets.set([]);
      this.dataSourceHeaderRow.set(1);
      this.dataSourceHeaders.set([]);
      this.selectedHeaders.set([]);
      this.sortedHeaders.set([]);
    } catch (error: any) {
      console.error('读取数据源文件失败:', error);
      alert(error.message || '读取数据源文件失败');
      this.dataSourceFile.set(null);
      this.dataSourceSheetNames.set([]);
      this.dataSourceWorkbook.set(null);
    } finally {
      this.isUploadingDataSourceFile.set(false);
    }
  }

  // 检查是否可以进入step2（两个文件都已选择）
  canProceedToStep2(): boolean {
    return this.baseFile() !== null && this.dataSourceFile() !== null;
  }

  // Step 2: 切换基础文件sheet选择
  toggleBaseSheet(sheetName: string) {
    const current = this.selectedBaseSheets();
    if (current.includes(sheetName)) {
      this.selectedBaseSheets.set(current.filter(s => s !== sheetName));
    } else {
      this.selectedBaseSheets.set([...current, sheetName]);
    }
    // 更新表头和显示数据源区域
    this.onBaseSheetsChange();
  }


  // Step 2: 全选/取消全选基础文件sheet
  toggleSelectAllBaseSheets() {
    const allSheets = this.baseSheetNames();
    const selected = this.selectedBaseSheets();
    const allSelected = allSheets.length > 0 && allSheets.every(sheet => selected.includes(sheet));

    if (allSelected) {
      this.selectedBaseSheets.set([]);
    } else {
      this.selectedBaseSheets.set([...allSheets]);
    }
    this.onBaseSheetsChange();
  }

  // Step 2: 全选/取消全选数据源文件sheet
  toggleSelectAllDataSourceSheets() {
    const allSheets = this.dataSourceSheetNames();
    const selected = this.selectedDataSourceSheets();
    const allSelected = allSheets.length > 0 && allSheets.every(sheet => selected.includes(sheet));

    if (allSelected) {
      this.selectedDataSourceSheets.set([]);
    } else {
      this.selectedDataSourceSheets.set([...allSheets]);
    }
    this.onDataSourceSheetsChange();
  }

  // 检查是否所有基础文件sheet都被选中
  isAllBaseSheetsSelected(): boolean {
    const allSheets = this.baseSheetNames();
    const selected = this.selectedBaseSheets();
    return allSheets.length > 0 && allSheets.every(sheet => selected.includes(sheet));
  }

  // 检查是否所有数据源文件sheet都被选中
  isAllDataSourceSheetsSelected(): boolean {
    const allSheets = this.dataSourceSheetNames();
    const selected = this.selectedDataSourceSheets();
    return allSheets.length > 0 && allSheets.every(sheet => selected.includes(sheet));
  }

  // Step 2: 更新基础文件表头行数
  onBaseHeaderRowChange() {
    // 如果数据源已选择，同步更新数据源表头行数
    if (this.selectedDataSourceSheets().length > 0) {
      this.dataSourceHeaderRow.set(this.baseHeaderRow());
    }
    this.loadBaseHeaders();
  }

  // Step 2: 更新数据源文件表头行数
  onDataSourceHeaderRowChange() {
    this.loadDataSourceHeaders();
  }

  // Step 2: 加载基础文件表头
  private async loadBaseHeaders() {
    const baseSheets = this.selectedBaseSheets();
    if (baseSheets.length === 0 || !this.baseWorkbook()) {
      this.baseHeaders.set([]);
      return;
    }

    try {
      const headers: string[] = [];
      // 读取第一个选中的基础文件sheet的表头
      const firstSheet = this.baseWorkbook()!.getWorksheet(baseSheets[0]);
      if (firstSheet) {
        const headerRow = firstSheet.getRow(this.baseHeaderRow());
        headerRow.eachCell({ includeEmpty: false }, (cell) => {
          const value = this.excelUtils.getCellValue(cell);
          if (value) {
            const headerText = String(value).trim();
            if (headerText && !headers.includes(headerText)) {
              headers.push(headerText);
            }
          }
        });
      }
      this.baseHeaders.set(headers);
    } catch (error) {
      console.error('读取基础文件表头失败:', error);
      this.baseHeaders.set([]);
    }
  }

  // Step 2: 加载数据源文件表头
  private async loadDataSourceHeaders() {
    const dataSourceSheets = this.selectedDataSourceSheets();
    if (dataSourceSheets.length === 0 || !this.dataSourceWorkbook()) {
      this.dataSourceHeaders.set([]);
      return;
    }

    try {
      const headers: string[] = [];
      // 读取所有选中的数据源文件sheet的表头
      for (const sheetName of dataSourceSheets) {
        const sheet = this.dataSourceWorkbook()!.getWorksheet(sheetName);
        if (sheet) {
          const headerRow = sheet.getRow(this.dataSourceHeaderRow());
          headerRow.eachCell({ includeEmpty: false }, (cell) => {
            const value = this.excelUtils.getCellValue(cell);
            if (value) {
              const headerText = String(value).trim();
              if (headerText && !headers.includes(headerText)) {
                headers.push(headerText);
              }
            }
          });
        }
      }
      this.dataSourceHeaders.set(headers);
    } catch (error) {
      console.error('读取数据源文件表头失败:', error);
      this.dataSourceHeaders.set([]);
    }
  }

  // Step 2: 获取所有可用的表头（基础文件 + 数据源文件）
  getAllAvailableHeaders(): string[] {
    const baseHeaders = this.baseHeaders();
    const dataSourceHeaders = this.dataSourceHeaders();
    const allHeaders: string[] = [];

    baseHeaders.forEach(header => {
      if (!allHeaders.includes(header)) {
        allHeaders.push(header);
      }
    });

    dataSourceHeaders.forEach(header => {
      if (!allHeaders.includes(header)) {
        allHeaders.push(header);
      }
    });

    return allHeaders;
  }

  // Step 2: 切换表头选择
  toggleHeader(header: string) {
    const selected = this.selectedHeaders();
    if (selected.includes(header)) {
      this.selectedHeaders.set(selected.filter(h => h !== header));
    } else {
      this.selectedHeaders.set([...selected, header]);
    }
    this.updateSortedHeaders();
  }

  // Step 2: 更新排序列表（将选中的表头添加到排序列表）
  private updateSortedHeaders() {
    const selected = this.selectedHeaders();
    const sorted = this.sortedHeaders();

    // 添加新选中的表头到排序列表末尾
    selected.forEach(header => {
      if (!sorted.includes(header)) {
        sorted.push(header);
      }
    });

    // 移除未选中的表头
    const newSorted = sorted.filter(header => selected.includes(header));
    this.sortedHeaders.set(newSorted);
  }

  // Step 2: 从排序列表中删除表头
  removeHeaderFromSorted(header: string) {
    // 从排序列表中移除
    const sorted = this.sortedHeaders().filter(h => h !== header);
    this.sortedHeaders.set(sorted);

    // 从选中列表中移除
    const selected = this.selectedHeaders().filter(h => h !== header);
    this.selectedHeaders.set(selected);
  }

  // Step 2: 拖拽排序表头
  dropHeader(event: CdkDragDrop<string[]>) {
    const headers = [...this.sortedHeaders()];
    moveItemInArray(headers, event.previousIndex, event.currentIndex);
    this.sortedHeaders.set(headers);
  }

  // Step3: 添加表尾功能
  addFooterFunction() {
    const header = this.newFooterHeader().trim();
    const type = this.newFooterType();

    if (!header) {
      return; // 如果没有选择表头，不添加
    }

    // 检查该表头是否已经配置了相同类型的表尾功能
    const existing = this.footerFunctions().find(
      f => f.header === header && f.type === type
    );

    if (existing) {
      return; // 如果已经存在，不重复添加
    }

    // 添加新的表尾功能配置
    const newFooter = {
      type: type,
      header: header,
      id: `${type}-${header}-${Date.now()}` // 生成唯一ID
    };

    this.footerFunctions.set([...this.footerFunctions(), newFooter]);
    this.newFooterHeader.set(''); // 重置选择
  }

  // Step3: 移除表尾功能
  removeFooterFunction(id: string) {
    this.footerFunctions.set(
      this.footerFunctions().filter(f => f.id !== id)
    );
  }

  // Step3: 获取表尾功能显示文本
  getFooterFunctionLabel(footer: { type: '合计' | '平均值', header: string }): string {
    return `${footer.type}(${footer.header})`;
  }

  // Step3: 获取可用的表头列表（用于表尾功能选择）
  getAvailableHeadersForFooter(): string[] {
    return this.sortedHeaders();
  }

  // Step 2: 当基础文件sheet选择变化时，更新数据源表头行数默认值
  onBaseSheetsChange() {
    const baseSheets = this.selectedBaseSheets();
    // 如果选择了基础sheet，且数据源还未选择，设置数据源表头行数默认值为基础文件的值
    if (baseSheets.length > 0 && this.selectedDataSourceSheets().length === 0) {
      this.dataSourceHeaderRow.set(this.baseHeaderRow());
    }
    // 如果选择了数据源sheet，同步更新表头行数
    if (baseSheets.length > 0 && this.selectedDataSourceSheets().length > 0) {
      this.dataSourceHeaderRow.set(this.baseHeaderRow());
    }
    // 加载基础文件表头
    this.loadBaseHeaders();
  }

  // Step 2: 当数据源文件sheet选择变化时
  onDataSourceSheetsChange() {
    // 加载数据源文件表头
    this.loadDataSourceHeaders();
  }

  // 检查是否显示数据源选择区域（当基础文件选择了多个sheet时）
  shouldShowDataSourceSection(): boolean {
    return this.selectedBaseSheets().length > 0;
  }

  // 检查是否显示表头排序区域（当有选中的表头时）
  shouldShowHeaderSortSection(): boolean {
    return this.sortedHeaders().length > 0;
  }

  // 基础文件sheet选择变化
  onBaseSheetsSelectionChange(selected: string[]) {
    this.selectedBaseSheets.set(selected);
    this.onBaseSheetsChange();
  }

  // 数据源文件sheet选择变化
  onDataSourceSheetsSelectionChange(selected: string[]) {
    const previousLength = this.selectedDataSourceSheets().length;
    this.selectedDataSourceSheets.set(selected);
    // 首次选择数据源sheet时，同步表头行数
    if (selected.length > 0 && previousLength === 0) {
      this.dataSourceHeaderRow.set(this.baseHeaderRow());
    }
    this.onDataSourceSheetsChange();
  }

  // 完成按钮点击处理
  onComplete() {
    if (!this.baseFile() || !this.dataSourceFile()) {
      alert('请先选择基础文件和数据源文件');
      return;
    }

    if (this.selectedBaseSheets().length === 0) {
      alert('请先选择基础文件的Sheet');
      return;
    }

    if (this.selectedDataSourceSheets().length === 0) {
      alert('请先选择数据源文件的Sheet');
      return;
    }

    if (this.sortedHeaders().length === 0) {
      alert('请先选择并排序表头');
      return;
    }

    // 设置默认文件名（去除扩展名）
    const defaultFileName = this.excelUtils.generateDefaultFileName(this.baseFile()!.name, '合并表');
    this.outputFileName.set(defaultFileName);
    this.showConfirmDialog.set(true);
  }

  // 关闭确认对话框
  closeConfirmDialog() {
    this.showConfirmDialog.set(false);
    this.generationProgress.set(0);
  }

  // 开始生成文件
  async startGeneration() {
    const fileName = this.outputFileName().trim() || '合并表';
    this.generationProgress.set(0);
    this.isProcessing.set(true);

    try {
      await this.excelUtils.downloadFileWithProgress(
        () => this.generateExcel(fileName),
        fileName,
        (progress) => this.generationProgress.set(progress)
      );

      // 等待一小段时间后关闭对话框
      setTimeout(() => {
        this.closeConfirmDialog();
        this.isProcessing.set(false);
      }, 300);
    } catch (error: any) {
      console.error('生成Excel失败:', error);
      alert(`生成Excel失败：${error?.message || '未知错误'}`);
      this.closeConfirmDialog();
      this.isProcessing.set(false);
    }
  }

  // 生成Excel文件
  async generateExcel(fileName: string = '合并表') {
    const baseWorkbook = this.baseWorkbook()!;
    const dataSourceWorkbook = this.dataSourceWorkbook()!;
    const selectedBaseSheets = this.selectedBaseSheets();
    const selectedDataSourceSheets = this.selectedDataSourceSheets();
    const sortedHeaders = this.sortedHeaders();
    const baseHeaderRow = this.baseHeaderRow();
    const dataSourceHeaderRow = this.dataSourceHeaderRow();
    const style = this.tableStyle();

    // 创建新的工作簿
    const newWorkbook = new ExcelJS.Workbook();
    const unmatchedSheets: string[] = []; // 记录未匹配的sheet

    // 按照数据源的sheet循环
    for (const dataSourceSheetName of selectedDataSourceSheets) {
      const dataSourceSheet = dataSourceWorkbook.getWorksheet(dataSourceSheetName);
      if (!dataSourceSheet) continue;

      // 查找对应的基础表sheet（按名称匹配）
      let matchedBaseSheet: ExcelJS.Worksheet | null = null;
      let matchedBaseSheetName: string | null = null;

      for (const baseSheetName of selectedBaseSheets) {
        if (baseSheetName === dataSourceSheetName) {
          const sheet = baseWorkbook.getWorksheet(baseSheetName);
          if (sheet) {
            matchedBaseSheet = sheet;
            matchedBaseSheetName = baseSheetName;
            break;
          }
        }
      }

      if (!matchedBaseSheet || !matchedBaseSheetName) {
        // 没有找到对应的基础表sheet
        unmatchedSheets.push(dataSourceSheetName);
        continue;
      }

      // 创建新的sheet
      const newSheet = newWorkbook.addWorksheet(dataSourceSheetName);

      // 读取基础表的表头
      const baseHeaderRowObj = matchedBaseSheet.getRow(baseHeaderRow);
      const baseHeadersMap = new Map<number, string>(); // 列号 -> 表头名
      baseHeaderRowObj.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const headerText = this.excelUtils.getCellValue(cell);
        if (headerText) {
          baseHeadersMap.set(colNumber, String(headerText).trim());
        }
      });

      // 读取数据源表的表头
      const dataSourceHeaderRowObj = dataSourceSheet.getRow(dataSourceHeaderRow);
      const dataSourceHeadersMap = new Map<number, string>(); // 列号 -> 表头名
      dataSourceHeaderRowObj.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const headerText = this.excelUtils.getCellValue(cell);
        if (headerText) {
          dataSourceHeadersMap.set(colNumber, String(headerText).trim());
        }
      });

      // 创建表头映射：基础表表头 -> 数据源表列号
      const headerToDataSourceColMap = new Map<string, number>();
      dataSourceHeadersMap.forEach((header, colNumber) => {
        headerToDataSourceColMap.set(header, colNumber);
      });

      // 添加标题行（使用基础表的标题，如果有）
      const titleRow = newSheet.addRow([`${dataSourceSheetName}`]);
      titleRow.height = 25 * 0.75;
      this.excelUtils.applyCellStyle(titleRow.getCell(1), style, 'title');
      newSheet.mergeCells(1, 1, 1, sortedHeaders.length);

      // 添加表头行（使用排序后的表头）
      const headerRow = newSheet.addRow(sortedHeaders);
      headerRow.height = 22 * 0.75;
      headerRow.eachCell((cell) => {
        this.excelUtils.applyCellStyle(cell, style, 'header');
      });

      // 读取基础表的数据（从表头行之后开始）
      const baseDataRows: Array<{ rowNumber: number, data: Map<string, any> }> = [];
      for (let rowNum = baseHeaderRow + 1; rowNum <= matchedBaseSheet.rowCount; rowNum++) {
        const row = matchedBaseSheet.getRow(rowNum);
        if (!row || row.cellCount === 0) continue;

        const rowData = new Map<string, any>();
        let hasData = false;

        baseHeadersMap.forEach((header, colNumber) => {
          const cell = row.getCell(colNumber);
          const cellValue = this.getCellValueWithFormula(cell, matchedBaseSheet, rowNum, colNumber);
          if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
            rowData.set(header, cellValue);
            hasData = true;
          }
        });

        if (hasData) {
          baseDataRows.push({ rowNumber: rowNum, data: rowData });
        }
      }

      // 读取数据源表的数据（从表头行之后开始）
      const dataSourceDataRows: Array<{ rowNumber: number, data: Map<string, any> }> = [];
      for (let rowNum = dataSourceHeaderRow + 1; rowNum <= dataSourceSheet.rowCount; rowNum++) {
        const row = dataSourceSheet.getRow(rowNum);
        if (!row || row.cellCount === 0) continue;

        const rowData = new Map<string, any>();
        let hasData = false;

        dataSourceHeadersMap.forEach((header, colNumber) => {
          const cell = row.getCell(colNumber);
          const cellValue = this.getCellValueWithFormula(cell, dataSourceSheet, rowNum, colNumber);
          if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
            rowData.set(header, cellValue);
            hasData = true;
          }
        });

        if (hasData) {
          dataSourceDataRows.push({ rowNumber: rowNum, data: rowData });
        }
      }

      // 合并数据：基础表数据 + 数据源表数据
      // 取基础表和数据源表中行数较多的作为基准
      const maxRows = Math.max(baseDataRows.length, dataSourceDataRows.length);
      let currentRow = 3; // 第1行是标题，第2行是表头，数据从第3行开始

      for (let i = 0; i < maxRows; i++) {
        const dataRow: any[] = [];
        const formulaMap = new Map<number, string>(); // 列索引 -> 公式

        for (let colIndex = 0; colIndex < sortedHeaders.length; colIndex++) {
          const header = sortedHeaders[colIndex];
          let cellValue: any = '';

          // 优先从数据源表获取数据
          if (i < dataSourceDataRows.length) {
            const dataSourceRowData = dataSourceDataRows[i].data;
            if (dataSourceRowData.has(header)) {
              cellValue = dataSourceRowData.get(header);
            }
          }

          // 如果数据源表没有，从基础表获取
          if ((cellValue === null || cellValue === undefined || cellValue === '') && i < baseDataRows.length) {
            const baseRowData = baseDataRows[i].data;
            if (baseRowData.has(header)) {
              cellValue = baseRowData.get(header);
            }
          }

          // 如果cellValue是公式对象，提取公式
          if (cellValue && typeof cellValue === 'object' && 'formula' in cellValue) {
            const formula = (cellValue as any).formula;
            // 转换公式引用（按照新的列位置转换）
            const convertResult = this.convertFormulaForMerge(
              formula,
              i < baseDataRows.length ? baseDataRows[i].rowNumber : (i < dataSourceDataRows.length ? dataSourceDataRows[i].rowNumber : currentRow),
              currentRow, // 当前行号
              colIndex + 1, // 当前列号
              sortedHeaders,
              matchedBaseSheet,
              dataSourceSheet,
              baseHeadersMap,
              dataSourceHeadersMap
            );
            if (convertResult && convertResult.missingHeaders.length === 0) {
              // 公式转换成功且没有缺少表头
              formulaMap.set(colIndex + 1, convertResult.formula);
              dataRow.push(null);
            } else {
              // 公式转换失败或缺少表头
              if (convertResult && convertResult.missingHeaders.length > 0) {
                // 有缺少的表头，用"缺少xxx"填充
                const missingHeadersText = convertResult.missingHeaders.join('、');
                dataRow.push(`缺少${missingHeadersText}`);
              } else {
                // 转换失败，尝试提取缺少的表头
                const missingHeaders = this.extractMissingHeadersFromFormula(
                  formula,
                  baseHeadersMap,
                  dataSourceHeadersMap,
                  sortedHeaders
                );
                if (missingHeaders.length > 0) {
                  const missingHeadersText = missingHeaders.join('、');
                  dataRow.push(`缺少${missingHeadersText}`);
                } else {
                  dataRow.push('缺少表头');
                }
              }
            }
          } else {
            dataRow.push(cellValue || '');
          }
        }

        const sheetRow = newSheet.addRow(dataRow);
        sheetRow.height = 20 * 0.75;
        for (let colNumber = 1; colNumber <= sortedHeaders.length; colNumber++) {
          const cell = sheetRow.getCell(colNumber);
          if (formulaMap.has(colNumber)) {
            const formula = formulaMap.get(colNumber)!;
            cell.value = { formula: formula };
          }
          this.excelUtils.applyCellStyle(cell, style, 'data');
        }
        currentRow++;
      }

      // 添加表尾功能行
      const footerFunctions = this.footerFunctions();
      if (footerFunctions.length > 0) {
        const footerRow: any[] = [];
        const dataStartRow = 3; // 数据从第3行开始
        const dataEndRow = currentRow - 1; // 数据结束行

        for (const header of sortedHeaders) {
          const footer = footerFunctions.find(f => f.header === header);
          if (footer) {
            const colName = this.excelUtils.getExcelColumnName(sortedHeaders.indexOf(header) + 1);
            if (footer.type === '合计') {
              footerRow.push({ formula: `SUM(${colName}${dataStartRow}:${colName}${dataEndRow})` });
            } else if (footer.type === '平均值') {
              footerRow.push({ formula: `AVERAGE(${colName}${dataStartRow}:${colName}${dataEndRow})` });
            }
          } else {
            footerRow.push('');
          }
        }

        const row = newSheet.addRow(footerRow);
        row.height = 22 * 0.75;
        row.eachCell((cell, colNumber) => {
          this.excelUtils.applyCellStyle(cell, style, 'total');
        });
        currentRow++;
      }

      // 自动调整列宽
      this.excelUtils.autoFitColumns(newSheet, sortedHeaders.length);

      // 设置表头筛选器
      newSheet.autoFilter = {
        from: { row: 2, column: 1 },
        to: { row: currentRow - 1, column: sortedHeaders.length }
      };
    }

    // 如果有未匹配的sheet，生成补充文件
    if (unmatchedSheets.length > 0) {
      const supplementWorkbook = new ExcelJS.Workbook();
      const baseFileName = this.baseFile()!.name.replace(/\.[^/.]+$/, '');
      const supplementFileName = `${baseFileName}-未匹配上的补充文件`;

      for (const unmatchedSheetName of unmatchedSheets) {
        const dataSourceSheet = dataSourceWorkbook.getWorksheet(unmatchedSheetName);
        if (!dataSourceSheet) continue;

        // 创建新的sheet
        const newSheet = supplementWorkbook.addWorksheet(unmatchedSheetName);

        // 读取数据源表的表头
        const dataSourceHeaderRowObj = dataSourceSheet.getRow(dataSourceHeaderRow);
        const dataSourceHeaders: string[] = [];
        dataSourceHeaderRowObj.eachCell({ includeEmpty: true }, (cell) => {
          const headerText = this.excelUtils.getCellValue(cell);
          if (headerText) {
            dataSourceHeaders.push(String(headerText).trim());
          }
        });

        if (dataSourceHeaders.length === 0) continue;

        // 添加标题行
        const titleRow = newSheet.addRow([unmatchedSheetName]);
        titleRow.height = 25 * 0.75;
        this.excelUtils.applyCellStyle(titleRow.getCell(1), style, 'title');
        newSheet.mergeCells(1, 1, 1, dataSourceHeaders.length);

        // 添加表头行
        const headerRow = newSheet.addRow(dataSourceHeaders);
        headerRow.height = 22 * 0.75;
        headerRow.eachCell((cell) => {
          this.excelUtils.applyCellStyle(cell, style, 'header');
        });

        // 复制数据源表的数据
        let currentRow = 3;
        for (let rowNum = dataSourceHeaderRow + 1; rowNum <= dataSourceSheet.rowCount; rowNum++) {
          const row = dataSourceSheet.getRow(rowNum);
          if (!row || row.cellCount === 0) continue;

          const dataRow: any[] = [];
          for (let colIndex = 0; colIndex < dataSourceHeaders.length; colIndex++) {
            const colNumber = colIndex + 1;
            const cell = row.getCell(colNumber);
            const cellValue = this.getCellValueWithFormula(cell, dataSourceSheet, rowNum, colNumber);
            dataRow.push(cellValue || '');
          }

          const sheetRow = newSheet.addRow(dataRow);
          sheetRow.height = 20 * 0.75;
          sheetRow.eachCell((cell) => {
            this.excelUtils.applyCellStyle(cell, style, 'data');
          });
          currentRow++;
        }

        // 自动调整列宽
        this.excelUtils.autoFitColumns(newSheet, dataSourceHeaders.length);

        // 设置表头筛选器
        newSheet.autoFilter = {
          from: { row: 2, column: 1 },
          to: { row: currentRow - 1, column: dataSourceHeaders.length }
        };
      }

      // 生成补充文件
      const supplementBuffer = await supplementWorkbook.xlsx.writeBuffer();
      const supplementBlob = new Blob([supplementBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const supplementUrl = window.URL.createObjectURL(supplementBlob);
      const supplementLink = document.createElement('a');
      supplementLink.href = supplementUrl;
      supplementLink.download = `${supplementFileName}.xlsx`;
      supplementLink.click();
      window.URL.revokeObjectURL(supplementUrl);
    }

    // 生成主文件
    const buffer = await newWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return blob;
  }

  // 获取单元格值（包括公式）
  private getCellValueWithFormula(
    cell: ExcelJS.Cell,
    sheet: ExcelJS.Worksheet,
    rowNumber: number,
    colNumber: number
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
    return this.excelUtils.getCellValue(cell);
  }

  // 转换公式引用（用于合并）- 按照新的列位置转换
  // 返回结果包含转换后的公式和缺少的表头列表
  private convertFormulaForMerge(
    formula: string,
    originalRow: number,
    currentRow: number,
    currentCol: number,
    sortedHeaders: string[],
    baseSheet: ExcelJS.Worksheet | null,
    dataSourceSheet: ExcelJS.Worksheet | null,
    baseHeadersMap: Map<number, string>,
    dataSourceHeadersMap: Map<number, string>
  ): { formula: string; missingHeaders: string[] } | null {
    try {
      // 匹配单元格引用模式：$?列字母$?行号 或 范围引用
      const cellRefPattern = /(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|\$?[A-Z]+:\$?[A-Z]+|\$?\d+:\$?\d+)/gi;
      let convertedFormula = formula;
      const matches = Array.from(formula.matchAll(cellRefPattern));
      const missingHeadersSet = new Set<string>();

      // 从后往前处理，避免索引变化
      for (let i = matches.length - 1; i >= 0; i--) {
        const match = matches[i];
        const fullRef = match[1];
        const matchIndex = match.index!;

        // 检查前面是否有工作表引用（如 '汇总表'! 或 工作表名!）
        const beforeMatch = formula.substring(Math.max(0, matchIndex - 50), matchIndex);
        const hasSheetRef = /'[^']*'!\s*$/.test(beforeMatch) || /[A-Za-z0-9_]+!\s*$/.test(beforeMatch);
        if (hasSheetRef) {
          continue; // 跳过其他工作表的引用
        }

        // 处理范围引用（如 A1:B2）
        if (fullRef.includes(':')) {
          const parts = fullRef.split(':');
          const convertedParts: (string | null)[] = [];
          for (const part of parts) {
            const result = this.convertSingleCellRefForMerge(
              part,
              originalRow,
              currentRow,
              sortedHeaders,
              baseSheet,
              dataSourceSheet,
              baseHeadersMap,
              dataSourceHeadersMap
            );
            if (result === null) {
              convertedParts.push(null);
            } else if (typeof result === 'string') {
              convertedParts.push(result);
            } else {
              // result是对象，包含converted和missingHeaders
              convertedParts.push(result.converted);
              result.missingHeaders.forEach(h => missingHeadersSet.add(h));
            }
          }
          if (convertedParts.some(p => p === null)) {
            // 如果范围引用中有部分失败，尝试提取缺少的表头
            parts.forEach(part => {
              const headerName = this.extractHeaderNameFromRef(part, baseHeadersMap, dataSourceHeadersMap);
              if (headerName && !sortedHeaders.includes(headerName)) {
                missingHeadersSet.add(headerName);
              }
            });
            return null;
          }
          convertedFormula = convertedFormula.substring(0, matchIndex) +
                             convertedParts.join(':') +
                             convertedFormula.substring(matchIndex + fullRef.length);
        } else {
          // 处理单个单元格引用
          const result = this.convertSingleCellRefForMerge(
            fullRef,
            originalRow,
            currentRow,
            sortedHeaders,
            baseSheet,
            dataSourceSheet,
            baseHeadersMap,
            dataSourceHeadersMap
          );
          if (result === null) {
            // 提取缺少的表头
            const headerName = this.extractHeaderNameFromRef(fullRef, baseHeadersMap, dataSourceHeadersMap);
            if (headerName && !sortedHeaders.includes(headerName)) {
              missingHeadersSet.add(headerName);
            }
            return null;
          } else if (typeof result === 'string') {
            convertedFormula = convertedFormula.substring(0, matchIndex) +
                               result +
                               convertedFormula.substring(matchIndex + fullRef.length);
          } else {
            // result是对象
            convertedFormula = convertedFormula.substring(0, matchIndex) +
                               result.converted +
                               convertedFormula.substring(matchIndex + fullRef.length);
            result.missingHeaders.forEach(h => missingHeadersSet.add(h));
          }
        }
      }
      return {
        formula: convertedFormula,
        missingHeaders: Array.from(missingHeadersSet)
      };
    } catch (e) {
      console.error('转换公式失败:', e, formula);
      // 即使转换失败，也尝试提取缺少的表头
      const missingHeaders = this.extractMissingHeadersFromFormula(
        formula,
        baseHeadersMap,
        dataSourceHeadersMap,
        sortedHeaders
      );
      if (missingHeaders.length > 0) {
        return {
          formula: '',
          missingHeaders: missingHeaders
        };
      }
      return null;
    }
  }

  // 从单元格引用中提取表头名称
  private extractHeaderNameFromRef(
    ref: string,
    baseHeadersMap: Map<number, string>,
    dataSourceHeadersMap: Map<number, string>
  ): string | null {
    const colMatch = ref.match(/[A-Z]+/i);
    if (!colMatch) return null;

    const colPart = colMatch[0].toUpperCase();
    // 将列字母转换为列号（1-based）
    let originalColNum = 0;
    for (let i = 0; i < colPart.length; i++) {
      originalColNum = originalColNum * 26 + (colPart.charCodeAt(i) - 64);
    }

    // 查找原始列对应的表头名称
    if (baseHeadersMap.has(originalColNum)) {
      return baseHeadersMap.get(originalColNum)!;
    } else if (dataSourceHeadersMap.has(originalColNum)) {
      return dataSourceHeadersMap.get(originalColNum)!;
    }

    return null;
  }

  // 转换单个单元格引用（用于合并）
  // 返回转换后的引用字符串，或包含缺少表头信息的对象，或null
  private convertSingleCellRefForMerge(
    ref: string,
    originalRow: number,
    currentRow: number,
    sortedHeaders: string[],
    baseSheet: ExcelJS.Worksheet | null,
    dataSourceSheet: ExcelJS.Worksheet | null,
    baseHeadersMap: Map<number, string>,
    dataSourceHeadersMap: Map<number, string>
  ): string | { converted: string; missingHeaders: string[] } | null {
    // 提取列和行的绝对/相对引用标记
    const colMatch = ref.match(/[A-Z]+/i);
    if (!colMatch) return null;

    const colPart = colMatch[0].toUpperCase();
    const colIndex = ref.indexOf(colPart);
    const beforeCol = ref.substring(0, colIndex);
    const isColAbsolute = beforeCol.includes('$') || beforeCol === '$';

    // 将列字母转换为列号（1-based）
    let originalColNum = 0;
    for (let i = 0; i < colPart.length; i++) {
      originalColNum = originalColNum * 26 + (colPart.charCodeAt(i) - 64);
    }

    // 查找原始列对应的表头名称
    let headerName = '';
    if (baseSheet && baseHeadersMap.has(originalColNum)) {
      headerName = baseHeadersMap.get(originalColNum)!;
    } else if (dataSourceSheet && dataSourceHeadersMap.has(originalColNum)) {
      headerName = dataSourceHeadersMap.get(originalColNum)!;
    }

    if (!headerName) {
      return null; // 找不到对应的表头
    }

    // 在新表头列表中找到该表头的位置
    const newColIdx = sortedHeaders.indexOf(headerName);
    if (newColIdx === -1) {
      // 该表头不在新表中，返回缺少的表头信息
      return { converted: '', missingHeaders: [headerName] };
    }

    // 构建新的列引用
    const newColRef = this.excelUtils.getExcelColumnName(newColIdx + 1);
    const newColPart = (isColAbsolute ? '$' : '') + newColRef;

    // 解析行部分（数字）
    const rowMatch = ref.match(/\d+/);
    if (rowMatch) {
      const originalRowNum = parseInt(rowMatch[0]);
      const rowIndex = ref.indexOf(rowMatch[0]);
      const beforeRow = ref.substring(0, rowIndex);
      const isRowAbsolute = beforeRow.includes('$') || (beforeRow.endsWith('$') && !beforeRow.endsWith('$$'));

      let newRowNum: number;
      if (originalRowNum === 1 || originalRowNum === 2) {
        // 引用表头行 -> 新表第2行（表头行）
        newRowNum = 2;
      } else {
        // 引用数据行 -> 根据相对偏移计算
        const rowOffset = currentRow - originalRow;
        newRowNum = originalRowNum + rowOffset;
      }
      const newRowPart = (isRowAbsolute ? '$' : '') + newRowNum;
      return newColPart + newRowPart;
    }

    return newColPart;
  }

  // 从公式中提取缺少的表头
  private extractMissingHeadersFromFormula(
    formula: string,
    baseHeadersMap: Map<number, string>,
    dataSourceHeadersMap: Map<number, string>,
    sortedHeaders: string[]
  ): string[] {
    const missingHeaders = new Set<string>();
    // 匹配单元格引用模式：$?列字母$?行号
    const cellRefPattern = /(\$?[A-Z]+)/gi;
    const matches = Array.from(formula.matchAll(cellRefPattern));

    for (const match of matches) {
      const colPart = match[1].replace(/\$/g, '').toUpperCase();
      // 将列字母转换为列号（1-based）
      let originalColNum = 0;
      for (let i = 0; i < colPart.length; i++) {
        originalColNum = originalColNum * 26 + (colPart.charCodeAt(i) - 64);
      }

      // 查找原始列对应的表头名称
      let headerName = '';
      if (baseHeadersMap.has(originalColNum)) {
        headerName = baseHeadersMap.get(originalColNum)!;
      } else if (dataSourceHeadersMap.has(originalColNum)) {
        headerName = dataSourceHeadersMap.get(originalColNum)!;
      }

      // 如果找到了表头名称，但不在新表头列表中，则记录为缺少的表头
      if (headerName && !sortedHeaders.includes(headerName)) {
        missingHeaders.add(headerName);
      }
    }

    return Array.from(missingHeaders);
  }
}

