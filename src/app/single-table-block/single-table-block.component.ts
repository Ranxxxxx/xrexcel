import { Component, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { RouterLink } from '@angular/router';
import { MatToolbarModule } from '@angular/material/toolbar';
import { MatIconModule } from '@angular/material/icon';
import { MatButtonModule } from '@angular/material/button';
import { MatCardModule } from '@angular/material/card';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatSelectModule } from '@angular/material/select';
import { MatInputModule } from '@angular/material/input';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { MatExpansionModule } from '@angular/material/expansion';
import { MatStepperModule } from '@angular/material/stepper';
import { MatChipsModule } from '@angular/material/chips';
import { MatProgressBarModule } from '@angular/material/progress-bar';
import { MatRadioModule } from '@angular/material/radio';
import { DragDropModule, CdkDragDrop, moveItemInArray } from '@angular/cdk/drag-drop';
import { MatListModule } from '@angular/material/list';
import { TableStyleConfig, DEFAULT_TABLE_STYLE } from '../shared/models/table-style.model';
import { TableStylePreviewComponent } from '../shared/components/table-style-preview/table-style-preview.component';
import { ExcelUtilsService } from '../shared/services/excel-utils.service';
import * as ExcelJS from 'exceljs';

// 功能配置类型
type FunctionConfig =
  | { type: 'modifyHeader'; id: string; modifiedHeaders: string[] }
  | { type: 'modifyFooter'; id: string; footerHeaders: string[] }
  | { type: 'categorySummary'; id: string; categoryHeader: string; summaryHeaders: string[] }
  | { type: 'sort'; id: string; sortHeaders: Array<{ header: string; order: 'asc' | 'desc' }> };

@Component({
  selector: 'app-single-table-block',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    RouterLink,
    MatToolbarModule,
    MatIconModule,
    MatButtonModule,
    MatCardModule,
    MatFormFieldModule,
    MatSelectModule,
    MatInputModule,
    MatCheckboxModule,
    MatExpansionModule,
    MatStepperModule,
    MatChipsModule,
    MatProgressBarModule,
    MatRadioModule,
    DragDropModule,
    MatListModule,
    TableStylePreviewComponent
  ],
  templateUrl: './single-table-block.component.html',
  styleUrl: './single-table-block.component.scss'
})
export class SingleTableBlockComponent {
  tableStyle = signal<TableStyleConfig>({ ...DEFAULT_TABLE_STYLE });
  previewExpanded = signal<boolean>(false);

  styleOptions = ['商务风格', '简约风格', '经典风格', '现代风格'];

  fontOptions = ['微软雅黑', '宋体', '黑体', 'Arial', 'Times New Roman'];

  fontSizeOptions = [8, 9, 10, 11, 12, 14, 16, 18, 20];

  borderStyleOptions = [
    { value: 'thin', label: '细边框' },
    { value: 'medium', label: '中等边框' },
    { value: 'thick', label: '粗边框' }
  ];

  // Step相关数据
  selectedFile: File | null = null;
  isUploading = signal<boolean>(false);
  sheetNames = signal<string[]>([]);
  originalWorkbook: ExcelJS.Workbook | null = null;
  isBatchMode = signal<boolean>(false); // 是否批量修改
  selectedSheets = signal<string[]>([]); // 选中的Sheet列表
  headerRowCount = signal<number>(1); // 表头行数
  headersData = signal<{ sheetName: string; headers: string[][] }[]>([]); // 表头数据

  // Step4相关数据
  availableFunctions = [
    { value: 'modifyHeader', label: '修改表头' },
    { value: 'modifyFooter', label: '修改表尾' },
    { value: 'categorySummary', label: '分类合计' },
    { value: 'sort', label: '排序' }
  ];

  functionConfigs = signal<FunctionConfig[]>([]); // 功能配置列表
  nextFunctionId = 1; // 下一个功能ID

  // 生成文件相关
  isProcessing = signal<boolean>(false); // 处理中状态
  showConfirmDialog = signal<boolean>(false); // 显示确认对话框
  outputFileName = signal<string>(''); // 输出文件名
  generationProgress = signal<number>(0); // 生成进度（0-100）

  // 当前可用的表头列表（从修改表头功能中获取，如果没有则使用原始表头）
  get currentAvailableHeaders(): string[] {
    // 查找最新的修改表头功能配置
    const modifyHeaderConfigs = this.functionConfigs().filter(f => f.type === 'modifyHeader') as Array<{ type: 'modifyHeader'; id: string; modifiedHeaders: string[] }>;
    if (modifyHeaderConfigs.length > 0) {
      // 返回最后一个修改表头配置的表头
      return modifyHeaderConfigs[modifyHeaderConfigs.length - 1].modifiedHeaders;
    }
    // 如果没有修改表头，返回原始表头的第一行
    if (this.headersData().length > 0 && this.headersData()[0].headers.length > 0) {
      return this.headersData()[0].headers[0];
    }
    return [];
  }

  constructor(private excelUtils: ExcelUtilsService) {}

  updateStyle(key: keyof TableStyleConfig, value: any) {
    this.tableStyle.update(config => ({ ...config, [key]: value }));
  }

  async onFileSelected(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      this.selectedFile = input.files[0];
      this.isUploading.set(true);

      try {
        await this.readExcelFile(this.selectedFile);
      } catch (error: any) {
        console.error('读取Excel文件失败:', error);
        const errorMessage = error?.message || '未知错误';
        alert(`读取Excel文件失败：${errorMessage}\n\n请确保：\n1. 文件格式为 .xlsx 或 .xls\n2. 文件未损坏\n3. 文件包含至少一个工作表`);
      } finally {
        this.isUploading.set(false);
        // 重置文件输入，允许重新选择同一文件
        input.value = '';
      }
    }
  }

  async readExcelFile(file: File) {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);

      // 获取所有sheet名称
      const sheets = workbook.worksheets.map(sheet => sheet.name);
      if (sheets.length === 0) {
        throw new Error('Excel文件中没有找到工作表');
      }

      this.sheetNames.set(sheets);
      this.originalWorkbook = workbook;
      // 重置选择
      this.selectedSheets.set([]);
      this.isBatchMode.set(false);
    } catch (error: any) {
      if (error instanceof Error) {
        throw error;
      }
      throw new Error(`读取文件时发生错误：${String(error)}`);
    }
  }

  toggleSheet(sheetName: string) {
    if (this.isBatchMode()) {
      // 批量模式：多选
      const current = this.selectedSheets();
      if (current.includes(sheetName)) {
        this.selectedSheets.set(current.filter(s => s !== sheetName));
      } else {
        this.selectedSheets.set([...current, sheetName]);
      }
    } else {
      // 单选模式：只能选一个
      this.selectedSheets.set([sheetName]);
    }
  }

  toggleSelectAll() {
    if (this.isSelectAll()) {
      this.selectedSheets.set([]);
    } else {
      this.selectedSheets.set([...this.sheetNames()]);
    }
  }

  isSelectAll(): boolean {
    return this.selectedSheets().length === this.sheetNames().length && this.sheetNames().length > 0;
  }

  isSheetSelected(sheetName: string): boolean {
    return this.selectedSheets().includes(sheetName);
  }

  onBatchModeChange() {
    // 切换批量模式时，清空选择
    this.selectedSheets.set([]);
  }

  onBatchModeRadioChange(event: any) {
    const isBatch = event.value === 'batch';
    this.isBatchMode.set(isBatch);
    this.onBatchModeChange();
  }

  // 读取表头数据
  async loadHeaders() {
    if (!this.originalWorkbook || this.selectedSheets().length === 0) {
      return;
    }

    const rowCount = this.headerRowCount();
    if (rowCount < 1) {
      alert('表头行数必须大于0');
      return;
    }

    const headersData: { sheetName: string; headers: string[][] }[] = [];

    // 如果是批量修改，只读取第一个sheet的表头
    // 如果不是批量修改，读取所有选中sheet的表头
    const sheetsToProcess = this.isBatchMode()
      ? [this.selectedSheets()[0]]
      : this.selectedSheets();

    for (const sheetName of sheetsToProcess) {
      const sheet = this.originalWorkbook.getWorksheet(sheetName);
      if (!sheet) {
        continue;
      }

      const headers: string[][] = [];

      // 读取指定行数的表头
      for (let rowNum = 1; rowNum <= rowCount; rowNum++) {
        const row = sheet.getRow(rowNum);
        const headerRow: string[] = [];

        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const cellValue = this.excelUtils.getCellValue(cell);
          headerRow.push(cellValue ? String(cellValue).trim() : '');
        });

        headers.push(headerRow);
      }

      headersData.push({ sheetName, headers });
    }

    this.headersData.set(headersData);
  }

  // 当表头行数改变时，重新加载表头
  onHeaderRowCountChange() {
    if (this.selectedSheets().length > 0 && this.headerRowCount() > 0) {
      this.loadHeaders();
    }
  }

  // 当进入Step3时，如果已选择sheet，自动加载表头
  onStep3Enter() {
    if (this.selectedSheets().length > 0 && this.headerRowCount() > 0) {
      this.loadHeaders();
    }
  }

  // Step4: 添加功能
  addFunction(functionType: string) {
    const id = `func-${this.nextFunctionId++}`;
    let newConfig: FunctionConfig;

    switch (functionType) {
      case 'modifyHeader':
        // 初始化修改表头功能，使用当前可用的表头
        newConfig = {
          type: 'modifyHeader',
          id,
          modifiedHeaders: [...this.currentAvailableHeaders]
        };
        break;
      case 'modifyFooter':
        // 初始化修改表尾功能，使用当前可用的表头
        newConfig = {
          type: 'modifyFooter',
          id,
          footerHeaders: []
        };
        break;
      case 'categorySummary':
        // 初始化分类合计功能
        newConfig = {
          type: 'categorySummary',
          id,
          categoryHeader: '',
          summaryHeaders: []
        };
        break;
      case 'sort':
        // 初始化排序功能
        newConfig = {
          type: 'sort',
          id,
          sortHeaders: []
        };
        break;
      default:
        return;
    }

    this.functionConfigs.update(configs => [...configs, newConfig]);
  }

  // Step4: 删除功能
  removeFunction(functionId: string) {
    this.functionConfigs.update(configs => configs.filter(c => c.id !== functionId));
    // 如果删除的是修改表头功能，需要更新后续功能的表头引用
    this.updateHeadersAfterModifyHeaderChange();
  }

  // Step4: 修改表头 - 添加表头
  addHeaderToModifyHeader(functionId: string) {
    this.functionConfigs.update(configs => {
      return configs.map(config => {
        if (config.id === functionId && config.type === 'modifyHeader') {
          return {
            ...config,
            modifiedHeaders: [...config.modifiedHeaders, '新表头']
          };
        }
        return config;
      });
    });
    this.updateHeadersAfterModifyHeaderChange();
  }

  // Step4: 修改表头 - 删除表头
  removeHeaderFromModifyHeader(functionId: string, index: number) {
    this.functionConfigs.update(configs => {
      return configs.map(config => {
        if (config.id === functionId && config.type === 'modifyHeader') {
          const newHeaders = [...config.modifiedHeaders];
          newHeaders.splice(index, 1);
          return {
            ...config,
            modifiedHeaders: newHeaders
          };
        }
        return config;
      });
    });
    this.updateHeadersAfterModifyHeaderChange();
  }

  // Step4: 修改表头 - 更新表头文本
  updateHeaderInModifyHeader(functionId: string, index: number, value: string) {
    this.functionConfigs.update(configs => {
      return configs.map(config => {
        if (config.id === functionId && config.type === 'modifyHeader') {
          const newHeaders = [...config.modifiedHeaders];
          newHeaders[index] = value;
          return {
            ...config,
            modifiedHeaders: newHeaders
          };
        }
        return config;
      });
    });
    this.updateHeadersAfterModifyHeaderChange();
  }

  // Step4: 修改表头 - 拖拽排序
  dropHeaderInModifyHeader(event: CdkDragDrop<string[]>, functionId: string) {
    this.functionConfigs.update(configs => {
      return configs.map(config => {
        if (config.id === functionId && config.type === 'modifyHeader') {
          const newHeaders = [...config.modifiedHeaders];
          moveItemInArray(newHeaders, event.previousIndex, event.currentIndex);
          return {
            ...config,
            modifiedHeaders: newHeaders
          };
        }
        return config;
      });
    });
    this.updateHeadersAfterModifyHeaderChange();
  }

  // Step4: 修改表头后，更新后续功能的表头引用
  updateHeadersAfterModifyHeaderChange() {
    // 获取最新的修改表头配置
    const modifyHeaderConfigs = this.functionConfigs().filter(f => f.type === 'modifyHeader') as Array<{ type: 'modifyHeader'; id: string; modifiedHeaders: string[] }>;
    const latestHeaders = modifyHeaderConfigs.length > 0
      ? modifyHeaderConfigs[modifyHeaderConfigs.length - 1].modifiedHeaders
      : this.currentAvailableHeaders;

    // 更新修改表尾功能中的表头选项（不改变已选择的表尾）
    this.functionConfigs.update(configs => {
      return configs.map(config => {
        if (config.type === 'modifyFooter') {
          // 只保留仍然存在的表尾表头
          return {
            ...config,
            footerHeaders: config.footerHeaders.filter(h => latestHeaders.includes(h))
          };
        }
        if (config.type === 'categorySummary') {
          // 如果分类依据表头不存在了，清空
          if (config.categoryHeader && !latestHeaders.includes(config.categoryHeader)) {
            return {
              ...config,
              categoryHeader: '',
              summaryHeaders: []
            };
          }
          // 只保留仍然存在的合计表头
          return {
            ...config,
            summaryHeaders: config.summaryHeaders.filter(h => latestHeaders.includes(h))
          };
        }
        if (config.type === 'sort') {
          // 只保留仍然存在的排序表头
          return {
            ...config,
            sortHeaders: config.sortHeaders.filter(s => latestHeaders.includes(s.header))
          };
        }
        return config;
      });
    });
  }

  // Step4: 修改表尾 - 切换表头选择
  toggleFooterHeader(functionId: string, header: string) {
    this.functionConfigs.update(configs => {
      return configs.map(config => {
        if (config.id === functionId && config.type === 'modifyFooter') {
          const index = config.footerHeaders.indexOf(header);
          if (index >= 0) {
            return {
              ...config,
              footerHeaders: config.footerHeaders.filter(h => h !== header)
            };
          } else {
            return {
              ...config,
              footerHeaders: [...config.footerHeaders, header]
            };
          }
        }
        return config;
      });
    });
  }

  // Step4: 分类合计 - 设置分类依据
  setCategoryHeader(functionId: string, header: string) {
    this.functionConfigs.update(configs => {
      return configs.map(config => {
        if (config.id === functionId && config.type === 'categorySummary') {
          return {
            ...config,
            categoryHeader: header
          };
        }
        return config;
      });
    });
  }

  // Step4: 分类合计 - 切换合计表头
  toggleSummaryHeader(functionId: string, header: string) {
    this.functionConfigs.update(configs => {
      return configs.map(config => {
        if (config.id === functionId && config.type === 'categorySummary') {
          const index = config.summaryHeaders.indexOf(header);
          if (index >= 0) {
            return {
              ...config,
              summaryHeaders: config.summaryHeaders.filter(h => h !== header)
            };
          } else {
            return {
              ...config,
              summaryHeaders: [...config.summaryHeaders, header]
            };
          }
        }
        return config;
      });
    });
  }

  // Step4: 排序 - 添加排序规则
  addSortRule(functionId: string) {
    this.functionConfigs.update(configs => {
      return configs.map(config => {
        if (config.id === functionId && config.type === 'sort') {
          const availableHeaders = this.currentAvailableHeaders.filter(
            h => !config.sortHeaders.some(s => s.header === h)
          );
          if (availableHeaders.length > 0) {
            return {
              ...config,
              sortHeaders: [...config.sortHeaders, { header: availableHeaders[0], order: 'asc' }]
            };
          }
        }
        return config;
      });
    });
  }

  // Step4: 排序 - 删除排序规则
  removeSortRule(functionId: string, index: number) {
    this.functionConfigs.update(configs => {
      return configs.map(config => {
        if (config.id === functionId && config.type === 'sort') {
          const newSortHeaders = [...config.sortHeaders];
          newSortHeaders.splice(index, 1);
          return {
            ...config,
            sortHeaders: newSortHeaders
          };
        }
        return config;
      });
    });
  }

  // Step4: 排序 - 更新排序规则
  updateSortRule(functionId: string, index: number, header: string, order: 'asc' | 'desc') {
    this.functionConfigs.update(configs => {
      return configs.map(config => {
        if (config.id === functionId && config.type === 'sort') {
          const newSortHeaders = [...config.sortHeaders];
          newSortHeaders[index] = { header, order };
          return {
            ...config,
            sortHeaders: newSortHeaders
          };
        }
        return config;
      });
    });
  }

  // Step4: 获取功能标签
  getFunctionLabel(functionType: string): string {
    const func = this.availableFunctions.find(f => f.value === functionType);
    return func ? func.label : '';
  }

  // Step4: 获取可用于选择的功能列表（排除已添加的功能）
  get availableFunctionsForSelection() {
    const addedTypes = this.functionConfigs().map(c => c.type) as string[];
    return this.availableFunctions.filter(f => !addedTypes.includes(f.value));
  }

  // Step4: 检查表头是否在表尾中
  isHeaderInFooter(functionId: string, header: string): boolean {
    const config = this.functionConfigs().find(c => c.id === functionId);
    return !!(config && config.type === 'modifyFooter' && config.footerHeaders.includes(header));
  }

  // Step4: 检查表头是否在合计中
  isHeaderInSummary(functionId: string, header: string): boolean {
    const config = this.functionConfigs().find(c => c.id === functionId);
    return !!(config && config.type === 'categorySummary' && config.summaryHeaders.includes(header));
  }

  // 获取预览数据
  get previewData(): any[][] {
    if (this.selectedSheets().length === 0) {
      return this.getDefaultPreviewData();
    }
    // 如果有选中的sheet，返回示例预览数据
    return this.getDefaultPreviewData();
  }

  // 获取默认预览数据
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

  // Step4: 完成按钮点击处理
  onComplete() {
    if (!this.selectedFile) {
      alert('请先上传Excel文件');
      return;
    }

    if (this.selectedSheets().length === 0) {
      alert('请先选择要修改的Sheet');
      return;
    }

    if (this.functionConfigs().length === 0) {
      alert('请先添加功能');
      return;
    }

    // 设置默认文件名（去除扩展名）
    const originalFileName = this.selectedFile.name.replace(/\.[^/.]+$/, '');
    this.outputFileName.set(originalFileName);
    this.showConfirmDialog.set(true);
  }

  // 关闭确认对话框
  closeConfirmDialog() {
    this.showConfirmDialog.set(false);
    this.generationProgress.set(0);
  }

  // 开始生成文件
  async startGeneration() {
    const fileName = this.outputFileName().trim() || '单表格处理';
    this.generationProgress.set(0);
    this.isProcessing.set(true);

    try {
      // 模拟进度条，最快1秒完成
      const startTime = Date.now();
      const minDuration = 1000; // 最少1秒

      // 开始生成Excel（返回blob，不直接下载）
      const generatePromise = this.generateExcel(fileName);

      // 更新进度条
      const progressInterval = setInterval(() => {
        const elapsed = Date.now() - startTime;
        if (elapsed < minDuration) {
          // 在1秒内，进度条从0到90%
          const progress = Math.min(90, (elapsed / minDuration) * 90);
          this.generationProgress.set(progress);
        } else {
          // 1秒后，等待生成完成
          this.generationProgress.set(95);
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
      this.generationProgress.set(100);

      // 等待进度条动画完成后再下载（等待一小段时间确保进度条显示100%）
      await new Promise(resolve => setTimeout(resolve, 200));

      // 现在触发下载
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${fileName}.xlsx`;
      link.click();
      window.URL.revokeObjectURL(url);

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
  async generateExcel(fileName: string = '单表格处理') {
    if (!this.originalWorkbook) {
      throw new Error('原始工作簿不存在');
    }

    const workbook = new ExcelJS.Workbook();
    const style = this.tableStyle();
    const selectedSheets = this.selectedSheets();
    const functionConfigs = this.functionConfigs();
    const headerRowCount = this.headerRowCount();

    // 获取最新的表头配置（从修改表头功能中获取）
    const modifyHeaderConfigs = functionConfigs.filter(f => f.type === 'modifyHeader') as Array<{ type: 'modifyHeader'; id: string; modifiedHeaders: string[] }>;
    const hasModifyHeader = modifyHeaderConfigs.length > 0;
    const finalHeaders = hasModifyHeader
      ? modifyHeaderConfigs[modifyHeaderConfigs.length - 1].modifiedHeaders
      : (this.headersData().length > 0 && this.headersData()[0].headers.length > 0
        ? this.headersData()[0].headers[0]
        : []);

    // 创建原始表头到新表头的映射
    const originalHeaders = this.headersData().length > 0 && this.headersData()[0].headers.length > 0
      ? this.headersData()[0].headers[0]
      : [];
    const headerMapping = new Map<string, number>(); // 原始表头 -> 新表头索引
    finalHeaders.forEach((header, index) => {
      const originalIndex = originalHeaders.indexOf(header);
      if (originalIndex >= 0) {
        headerMapping.set(header, index);
      }
    });

    // 如果没有修改表头，标记使用原始顺序
    const useOriginalOrder = !hasModifyHeader;

    // 检查是否有分类合计功能
    const categorySummaryConfig = functionConfigs.find(f => f.type === 'categorySummary') as { type: 'categorySummary'; id: string; categoryHeader: string; summaryHeaders: string[] } | undefined;
    const hasCategorySummary = categorySummaryConfig && categorySummaryConfig.categoryHeader && categorySummaryConfig.summaryHeaders.length > 0;

    // 处理每个选中的Sheet
    for (const sheetName of selectedSheets) {
      const originalSheet = this.originalWorkbook.getWorksheet(sheetName);
      if (!originalSheet) {
        continue;
      }

      // 读取原始数据
      const originalData: any[][] = [];
      for (let rowNum = 1; rowNum <= originalSheet.rowCount; rowNum++) {
        const row = originalSheet.getRow(rowNum);
        const rowData: any[] = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          rowData.push(cell);
        });
        originalData.push(rowData);
      }

      if (hasCategorySummary) {
        // 如果设置了分类合计，为每个分类创建单独的sheet
        await this.createCategorySheets(
          workbook,
          originalSheet,
          originalData,
          functionConfigs,
          finalHeaders,
          headerMapping,
          headerRowCount,
          style,
          useOriginalOrder,
          categorySummaryConfig!,
          sheetName
        );
      } else {
        // 如果没有分类合计，按原来的方式处理
        const newSheet = workbook.addWorksheet(sheetName);
        await this.applyFunctionConfigs(
          newSheet,
          originalSheet,
          originalData,
          functionConfigs,
          finalHeaders,
          headerMapping,
          headerRowCount,
          style,
          useOriginalOrder
        );
      }
    }

    // 生成Excel文件并返回blob
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return blob;
  }

  // 应用功能配置到工作表
  private async applyFunctionConfigs(
    newSheet: ExcelJS.Worksheet,
    originalSheet: ExcelJS.Worksheet,
    originalData: any[][],
    functionConfigs: FunctionConfig[],
    finalHeaders: string[],
    headerMapping: Map<string, number>,
    headerRowCount: number,
    style: TableStyleConfig,
    useOriginalOrder: boolean = false
  ) {
    // 先处理表头
    const headerRows: any[][] = [];
    for (let rowNum = 1; rowNum <= headerRowCount; rowNum++) {
      const headerRow: any[] = [];
      if (useOriginalOrder) {
        // 如果没有修改表头，直接按照原始列的顺序读取
        const originalRow = originalSheet.getRow(rowNum);
        originalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const cellValue = this.excelUtils.getCellValue(cell);
          headerRow.push(cellValue ? String(cellValue).trim() : '');
        });
      } else {
        // 如果修改了表头，按照finalHeaders的顺序查找
        finalHeaders.forEach((header) => {
          // 查找原始表头中对应的列
          const originalColIndex = this.findOriginalColumnIndex(originalSheet, header, rowNum);
          if (originalColIndex >= 0 && originalData[rowNum - 1] && originalData[rowNum - 1][originalColIndex]) {
            const originalCell = originalData[rowNum - 1][originalColIndex];
            headerRow.push(this.excelUtils.getCellValue(originalCell));
          } else {
            headerRow.push(header);
          }
        });
      }
      headerRows.push(headerRow);
    }

    // 读取数据行（从表头行之后开始）
    let dataRows: Array<{ cells: ExcelJS.Cell[], originalRowNum: number }> = [];
    for (let rowNum = headerRowCount + 1; rowNum <= originalSheet.rowCount; rowNum++) {
      const row = originalSheet.getRow(rowNum);
      const rowData: ExcelJS.Cell[] = [];
      if (useOriginalOrder) {
        // 如果没有修改表头，直接按照原始列的顺序读取
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          rowData.push(cell);
        });
      } else {
        // 如果修改了表头，按照finalHeaders的顺序查找
        finalHeaders.forEach((header) => {
          const originalColIndex = this.findOriginalColumnIndex(originalSheet, header, 1);
          if (originalColIndex >= 0) {
            const cell = row.getCell(originalColIndex);
            rowData.push(cell);
          } else {
            rowData.push(null as any);
          }
        });
      }
      dataRows.push({ cells: rowData, originalRowNum: rowNum });
    }

    // 应用排序功能
    const sortConfig = functionConfigs.find(f => f.type === 'sort') as { type: 'sort'; id: string; sortHeaders: Array<{ header: string; order: 'asc' | 'desc' }> } | undefined;
    if (sortConfig && sortConfig.sortHeaders.length > 0) {
      dataRows = this.applySort(dataRows, sortConfig.sortHeaders, finalHeaders);
    }

    // 应用分类合计功能
    const categorySummaryConfig = functionConfigs.find(f => f.type === 'categorySummary') as { type: 'categorySummary'; id: string; categoryHeader: string; summaryHeaders: string[] } | undefined;
    if (categorySummaryConfig && categorySummaryConfig.categoryHeader && categorySummaryConfig.summaryHeaders.length > 0) {
      dataRows = this.applyCategorySummary(dataRows, categorySummaryConfig, finalHeaders);
    }

    // 写入表头行
    headerRows.forEach((headerRow) => {
      const row = newSheet.addRow(headerRow);
      row.height = 22 * 0.75;
      row.eachCell((cell, colNumber) => {
        this.excelUtils.applyCellStyle(cell, style, 'header');
      });
    });

    // 写入数据行
    const dataStartRow = headerRowCount + 1;
    const originalHeaderIndexMap = new Map<string, number>();
    const originalHeaders = this.headersData().length > 0 && this.headersData()[0].headers.length > 0
      ? this.headersData()[0].headers[0]
      : [];
    originalHeaders.forEach((header, index) => {
      originalHeaderIndexMap.set(header, index);
    });

    // 记录每个分类组的数据行范围（Excel行号）
    const categoryDataRanges = new Map<string, { startRow: number, endRow: number }>();
    let currentCategoryValue: string | null = null;
    let currentCategoryStartRow: number | null = null;

    dataRows.forEach((rowData, rowIndex) => {
      const rowValues: any[] = [];
      const formulaMap = new Map<number, string>();

      // 检查是否是合计行、分类标题行或空行
      const isSummaryRow = (rowData as any).isSummary === true;
      const isCategoryTitle = (rowData as any).isCategoryTitle === true;
      const isBlankRow = (rowData as any).isBlankRow === true;
      const categoryValue = (rowData as any).categoryValue;

      // 处理空行
      if (isBlankRow) {
        // 如果之前有分类组，记录其结束行
        if (currentCategoryValue !== null && currentCategoryStartRow !== null) {
          categoryDataRanges.set(currentCategoryValue, {
            startRow: currentCategoryStartRow,
            endRow: newSheet.rowCount
          });
          currentCategoryValue = null;
          currentCategoryStartRow = null;
        }
        finalHeaders.forEach(() => {
          rowValues.push('');
        });
        const row = newSheet.addRow(rowValues);
        row.height = 20 * 0.75;
        row.eachCell((cell, colNumber) => {
          this.excelUtils.applyCellStyle(cell, style, 'data');
        });
        return;
      }

      // 处理分类标题行
      if (isCategoryTitle) {
        // 如果之前有分类组，记录其结束行
        if (currentCategoryValue !== null && currentCategoryStartRow !== null) {
          categoryDataRanges.set(currentCategoryValue, {
            startRow: currentCategoryStartRow,
            endRow: newSheet.rowCount - 1
          });
        }
        // 开始新的分类组
        currentCategoryValue = categoryValue || '';
        currentCategoryStartRow = newSheet.rowCount + 1; // 下一行是数据开始行
        // 分类标题行：第一列显示分类值，其他列为空（用于合并）
        finalHeaders.forEach((header, colIndex) => {
          if (colIndex === 0) {
            // 第一列显示分类值
            rowValues.push(categoryValue || '');
          } else {
            rowValues.push('');
          }
        });
        const row = newSheet.addRow(rowValues);
        row.height = 22 * 0.75;
        // 合并所有列来显示分类标题
        if (finalHeaders.length > 1) {
          newSheet.mergeCells(
            newSheet.rowCount,
            1,
            newSheet.rowCount,
            finalHeaders.length
          );
        }
        // 应用标题样式
        const firstCell = row.getCell(1);
        this.excelUtils.applyCellStyle(firstCell, style, 'title');
        return;
      }

      // 获取当前行的原始行号（用于从原始sheet获取公式）
      const currentOriginalRowNum = rowData.originalRowNum;
      const currentNewRowNum = newSheet.rowCount + 1; // 即将写入的行号

      // 如果是合计行，先计算并记录该分类的数据行范围
      if (isSummaryRow && currentCategoryValue !== null && currentCategoryStartRow !== null) {
        // 在写入合计行之前，先记录当前分类的数据行范围
        categoryDataRanges.set(currentCategoryValue, {
          startRow: currentCategoryStartRow,
          endRow: newSheet.rowCount // 当前行号（合计行之前）
        });
      }

      rowData.cells.forEach((cell, colIndex) => {
        if (isSummaryRow) {
          // 合计行处理
          if (finalHeaders[colIndex] === categorySummaryConfig?.categoryHeader) {
            // 分类列显示分类值
            rowValues.push(categoryValue || '');
            return;
          } else if (categorySummaryConfig && categorySummaryConfig.summaryHeaders.includes(finalHeaders[colIndex])) {
            // 合计列生成公式 - 只对选择的表头进行合计
            const colName = this.excelUtils.getExcelColumnName(colIndex + 1);
            // 从记录的范围中获取该分类的数据行范围
            const range = categoryDataRanges.get(categoryValue || '');
            if (range && range.startRow <= range.endRow) {
              formulaMap.set(colIndex + 1, `SUM(${colName}${range.startRow}:${colName}${range.endRow})`);
            }
            rowValues.push(null);
            return;
          } else {
            rowValues.push('');
            return;
          }
        }

        if (!cell) {
          rowValues.push('');
          return;
        }

        // 从原始sheet获取公式（参照汇总页面的方式）
        let formulaText = '';
        let cellValue: any = null;

        if (currentOriginalRowNum > 0) {
          let originalCell: ExcelJS.Cell | null = null;

          if (useOriginalOrder) {
            // 如果没有修改表头，直接从原始sheet获取对应列
            originalCell = originalSheet.getRow(currentOriginalRowNum).getCell(colIndex + 1);
          } else {
            // 如果修改了表头，需要找到对应的原始列
            const header = finalHeaders[colIndex];
            const originalColIndex = this.findOriginalColumnIndex(originalSheet, header, 1);
            if (originalColIndex >= 0) {
              originalCell = originalSheet.getRow(currentOriginalRowNum).getCell(originalColIndex);
            }
          }

          if (originalCell) {
            // 检查是否是公式
            if (originalCell.formula) {
              formulaText = originalCell.formula;
            } else if (typeof originalCell.value === 'object' && originalCell.value !== null && 'formula' in originalCell.value) {
              formulaText = (originalCell.value as any).formula;
            }
            // 获取单元格值（保持数据类型）
            cellValue = originalCell.value;
            // 如果value为空，尝试使用result
            if (cellValue === null || cellValue === undefined) {
              if (originalCell.result !== null && originalCell.result !== undefined) {
                cellValue = originalCell.result;
              }
            }
          } else {
            // 如果找不到原始单元格，使用cell的值
            if (cell.formula) {
              formulaText = cell.formula;
            } else if (typeof cell.value === 'object' && cell.value !== null && 'formula' in cell.value) {
              formulaText = (cell.value as any).formula;
            }
            cellValue = cell.value !== null && cell.value !== undefined ? cell.value : (cell.result || '');
          }
        } else {
          // 如果找不到原始行号，使用cell的值
          if (cell.formula) {
            formulaText = cell.formula;
          } else if (typeof cell.value === 'object' && cell.value !== null && 'formula' in cell.value) {
            formulaText = (cell.value as any).formula;
          }
          cellValue = cell.value !== null && cell.value !== undefined ? cell.value : (cell.result || '');
        }

        if (formulaText) {
          // 转换公式
          const convertedFormula = this.convertFormulaForNewHeaders(
            formulaText,
            currentOriginalRowNum > 0 ? currentOriginalRowNum : rowIndex + dataStartRow,
            currentNewRowNum,
            colIndex + 1,
            finalHeaders,
            originalHeaderIndexMap,
            originalSheet,
            dataRows.filter(r => !(r as any).isSummary && !(r as any).isCategoryTitle && !(r as any).isBlankRow).length
          );
          if (convertedFormula) {
            formulaMap.set(colIndex + 1, convertedFormula);
            rowValues.push(null);
          } else {
            // 公式转换失败，使用计算结果或值
            if (cellValue !== null && cellValue !== undefined) {
              rowValues.push(cellValue);
            } else if (cell.result !== null && cell.result !== undefined) {
              rowValues.push(cell.result);
            } else {
              rowValues.push('');
            }
          }
        } else {
          // 没有公式，直接使用值（保持数据类型）
          if (cellValue !== null && cellValue !== undefined) {
            if (typeof cellValue === 'number') {
              rowValues.push(cellValue);
            } else if (typeof cellValue === 'object' && 'text' in cellValue) {
              // 超链接对象
              rowValues.push((cellValue as any).text || '');
            } else if (cellValue instanceof Date) {
              rowValues.push(cellValue);
            } else {
              rowValues.push(cellValue);
            }
          } else if (cell.result !== null && cell.result !== undefined) {
            rowValues.push(cell.result);
          } else {
            rowValues.push('');
          }
        }
      });

      const row = newSheet.addRow(rowValues);
      row.height = 20 * 0.75;
      row.eachCell((cell, colNumber) => {
        // 设置公式
        if (formulaMap.has(colNumber)) {
          const formula = formulaMap.get(colNumber)!;
          cell.value = { formula: formula };
        }
        this.excelUtils.applyCellStyle(cell, style, isSummaryRow ? 'total' : 'data');
      });

      // 如果是普通数据行（不是合计行、标题行或空行），更新分类数据行的范围
      if (!isSummaryRow && !isCategoryTitle && !isBlankRow && currentCategoryValue !== null) {
        const currentRowNum = newSheet.rowCount;
        if (currentCategoryStartRow === null) {
          currentCategoryStartRow = currentRowNum;
        }
        // 更新结束行号
        categoryDataRanges.set(currentCategoryValue, {
          startRow: currentCategoryStartRow!,
          endRow: currentRowNum
        });
      }

      // 如果是合计行，记录该分类组的结束并准备下一个分类组
      if (isSummaryRow && currentCategoryValue !== null) {
        // 合计行已经写入，该分类组完成（范围已在写入前记录）
        currentCategoryValue = null;
        currentCategoryStartRow = null;
      }

      // 如果是空行，结束上一个分类
      if (isBlankRow && currentCategoryValue !== null && currentCategoryStartRow !== null) {
        categoryDataRanges.set(currentCategoryValue, {
          startRow: currentCategoryStartRow,
          endRow: newSheet.rowCount - 1
        });
        currentCategoryValue = null;
        currentCategoryStartRow = null;
      }
    });

    // 应用修改表尾功能
    // 如果设置了分类合计，不需要表尾合计（每个分类都有自己的合计行）
    const modifyFooterConfig = functionConfigs.find(f => f.type === 'modifyFooter') as { type: 'modifyFooter'; id: string; footerHeaders: string[] } | undefined;

    // 确定表尾需要合计的列
    let footerHeadersToSum: string[] = [];
    if (categorySummaryConfig && categorySummaryConfig.summaryHeaders.length > 0) {
      // 如果设置了分类合计，不需要表尾合计（每个分类都有自己的合计行）
      footerHeadersToSum = [];
    } else if (modifyFooterConfig && modifyFooterConfig.footerHeaders.length > 0) {
      // 如果没有分类合计，使用修改表尾功能的配置
      footerHeadersToSum = modifyFooterConfig.footerHeaders;
    }

    if (footerHeadersToSum.length > 0) {
      const footerRow: any[] = [];
      finalHeaders.forEach((header) => {
        if (footerHeadersToSum.includes(header)) {
          const colIndex = finalHeaders.indexOf(header);
          const colName = this.excelUtils.getExcelColumnName(colIndex + 1);
          // 计算数据行的范围（排除分类标题行、空行和合计行）
          let startRow = headerRowCount + 1;
          let endRow = newSheet.rowCount;

          // 如果设置了分类合计，需要排除分类标题行、空行和合计行
          if (categorySummaryConfig) {
            // 找到第一个和最后一个数据行（排除分类标题行、空行和合计行）
            let foundFirstDataRow = false;
            let lastDataRow = headerRowCount;

            // 遍历所有数据行，找到实际的数据行范围
            for (let i = 0; i < dataRows.length; i++) {
              const rowData = dataRows[i];
              const isSummary = (rowData as any).isSummary === true;
              const isCategoryTitle = (rowData as any).isCategoryTitle === true;
              const isBlank = (rowData as any).isBlankRow === true;

              // 只计算普通数据行
              if (!isSummary && !isCategoryTitle && !isBlank) {
                const actualRowNum = headerRowCount + 1 + i;
                if (!foundFirstDataRow) {
                  startRow = actualRowNum;
                  foundFirstDataRow = true;
                }
                lastDataRow = actualRowNum;
              }
            }

            if (foundFirstDataRow) {
              endRow = lastDataRow;
            } else {
              // 如果没有找到数据行，使用表头行之后的第一行
              startRow = headerRowCount + 1;
              endRow = headerRowCount + 1;
            }
          }

          footerRow.push({ formula: `SUM(${colName}${startRow}:${colName}${endRow})` });
        } else {
          footerRow.push('');
        }
      });
      const footerRowObj = newSheet.addRow(footerRow);
      footerRowObj.height = 22 * 0.75;
      footerRowObj.eachCell((cell, colNumber) => {
        if (footerRow[colNumber - 1] && typeof footerRow[colNumber - 1] === 'object' && 'formula' in footerRow[colNumber - 1]) {
          cell.value = { formula: (footerRow[colNumber - 1] as any).formula };
        }
        this.excelUtils.applyCellStyle(cell, style, 'total');
      });
    }

    // 自动调整列宽
    this.excelUtils.autoFitColumns(newSheet, finalHeaders.length);
  }

  // 查找原始列索引
  private findOriginalColumnIndex(sheet: ExcelJS.Worksheet, headerName: string, headerRowNum: number): number {
    const headerRow = sheet.getRow(headerRowNum);
    let colIndex = -1;
    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const cellValue = this.excelUtils.getCellValue(cell);
      if (String(cellValue || '').trim() === headerName) {
        colIndex = colNumber;
      }
    });
    return colIndex;
  }

  // 应用排序
  private applySort(
    dataRows: Array<{ cells: ExcelJS.Cell[], originalRowNum: number }>,
    sortHeaders: Array<{ header: string; order: 'asc' | 'desc' }>,
    finalHeaders: string[]
  ): Array<{ cells: ExcelJS.Cell[], originalRowNum: number }> {
    const sortedRows = [...dataRows];
    sortedRows.sort((a, b) => {
      for (const sortRule of sortHeaders) {
        const colIndex = finalHeaders.indexOf(sortRule.header);
        if (colIndex < 0) continue;

        const aCell = a.cells[colIndex];
        const bCell = b.cells[colIndex];
        const aVal = aCell ? this.excelUtils.getCellValue(aCell) : '';
        const bVal = bCell ? this.excelUtils.getCellValue(bCell) : '';

        let comparison = 0;
        if (aVal < bVal) comparison = -1;
        else if (aVal > bVal) comparison = 1;

        if (comparison !== 0) {
          return sortRule.order === 'asc' ? comparison : -comparison;
        }
      }
      return 0;
    });
    return sortedRows;
  }

  // 应用分类合计
  private applyCategorySummary(
    dataRows: Array<{ cells: ExcelJS.Cell[], originalRowNum: number }>,
    categorySummaryConfig: { type: 'categorySummary'; id: string; categoryHeader: string; summaryHeaders: string[] },
    finalHeaders: string[]
  ): Array<{ cells: ExcelJS.Cell[], originalRowNum: number, isSummary?: boolean, categoryValue?: string, isCategoryTitle?: boolean, isBlankRow?: boolean }> {
    const categoryIndex = finalHeaders.indexOf(categorySummaryConfig.categoryHeader);
    if (categoryIndex < 0) return dataRows;

    // 按分类分组
    const grouped = new Map<string, Array<{ cells: ExcelJS.Cell[], originalRowNum: number }>>();
    dataRows.forEach(row => {
      const categoryCell = row.cells[categoryIndex];
      const categoryValue = categoryCell ? String(this.excelUtils.getCellValue(categoryCell) || '') : '';
      if (!grouped.has(categoryValue)) {
        grouped.set(categoryValue, []);
      }
      grouped.get(categoryValue)!.push(row);
    });

    // 重新组合：每组数据 + 合计行，并在组之间添加空行
    const result: Array<{ cells: ExcelJS.Cell[], originalRowNum: number, isSummary?: boolean, categoryValue?: string, isCategoryTitle?: boolean, isBlankRow?: boolean }> = [];
    const categoryValues = Array.from(grouped.keys());

    categoryValues.forEach((categoryValue, groupIndex) => {
      const rows = grouped.get(categoryValue)!;

      // 如果不是第一组，添加空行分隔
      if (groupIndex > 0) {
        const blankRow: { cells: ExcelJS.Cell[], originalRowNum: number, isBlankRow: boolean } = {
          cells: [],
          originalRowNum: -1,
          isBlankRow: true
        };
        finalHeaders.forEach(() => {
          blankRow.cells.push(null as any);
        });
        result.push(blankRow);
      }

      // 添加分类标题行（在数据上方）
      const titleRow: { cells: ExcelJS.Cell[], originalRowNum: number, isCategoryTitle: boolean, categoryValue: string } = {
        cells: [],
        originalRowNum: -1,
        isCategoryTitle: true,
        categoryValue: categoryValue
      };
      finalHeaders.forEach(() => {
        titleRow.cells.push(null as any);
      });
      result.push(titleRow);

      // 添加数据行
      result.push(...rows);

      // 添加合计行标记（在每个分类的数据后面）
      const summaryRow: { cells: ExcelJS.Cell[], originalRowNum: number, isSummary: boolean, categoryValue: string } = {
        cells: [],
        originalRowNum: -1,
        isSummary: true,
        categoryValue: categoryValue
      };
      finalHeaders.forEach((header, colIndex) => {
        summaryRow.cells.push(null as any);
      });
      result.push(summaryRow);
    });

    return result;
  }

  // 转换公式以适应新的表头结构
  private convertFormulaForNewHeaders(
    formula: string,
    originalRow: number,
    newRow: number,
    newCol: number,
    newHeaders: string[],
    originalHeaderIndexMap: Map<string, number>,
    originalSheet: ExcelJS.Worksheet,
    dataRowCount: number
  ): string {
    // 使用 ExcelUtilsService 的 convertFormula 方法
    const converted = this.excelUtils.convertFormula(
      formula,
      originalRow,
      newRow,
      1,
      newCol,
      newHeaders,
      originalHeaderIndexMap,
      originalSheet,
      dataRowCount
    );

    return converted || formula;
  }

  // 为每个分类创建单独的sheet
  private async createCategorySheets(
    workbook: ExcelJS.Workbook,
    originalSheet: ExcelJS.Worksheet,
    originalData: any[][],
    functionConfigs: FunctionConfig[],
    finalHeaders: string[],
    headerMapping: Map<string, number>,
    headerRowCount: number,
    style: TableStyleConfig,
    useOriginalOrder: boolean,
    categorySummaryConfig: { type: 'categorySummary'; id: string; categoryHeader: string; summaryHeaders: string[] },
    baseSheetName: string
  ) {
    // 先处理表头
    const headerRows: any[][] = [];
    for (let rowNum = 1; rowNum <= headerRowCount; rowNum++) {
      const headerRow: any[] = [];
      if (useOriginalOrder) {
        // 如果没有修改表头，直接按照原始列的顺序读取
        const originalRow = originalSheet.getRow(rowNum);
        originalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const cellValue = this.excelUtils.getCellValue(cell);
          headerRow.push(cellValue ? String(cellValue).trim() : '');
        });
      } else {
        // 如果修改了表头，按照finalHeaders的顺序查找
        finalHeaders.forEach((header) => {
          const originalColIndex = this.findOriginalColumnIndex(originalSheet, header, rowNum);
          if (originalColIndex >= 0 && originalData[rowNum - 1] && originalData[rowNum - 1][originalColIndex]) {
            const originalCell = originalData[rowNum - 1][originalColIndex];
            headerRow.push(this.excelUtils.getCellValue(originalCell));
          } else {
            headerRow.push(header);
          }
        });
      }
      headerRows.push(headerRow);
    }

    // 读取数据行（从表头行之后开始）
    let dataRows: Array<{ cells: ExcelJS.Cell[], originalRowNum: number }> = [];
    for (let rowNum = headerRowCount + 1; rowNum <= originalSheet.rowCount; rowNum++) {
      const row = originalSheet.getRow(rowNum);
      const rowData: ExcelJS.Cell[] = [];
      if (useOriginalOrder) {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          rowData.push(cell);
        });
      } else {
        finalHeaders.forEach((header) => {
          const originalColIndex = this.findOriginalColumnIndex(originalSheet, header, 1);
          if (originalColIndex >= 0) {
            const cell = row.getCell(originalColIndex);
            rowData.push(cell);
          } else {
            rowData.push(null as any);
          }
        });
      }
      dataRows.push({ cells: rowData, originalRowNum: rowNum });
    }

    // 应用排序功能
    const sortConfig = functionConfigs.find(f => f.type === 'sort') as { type: 'sort'; id: string; sortHeaders: Array<{ header: string; order: 'asc' | 'desc' }> } | undefined;
    if (sortConfig && sortConfig.sortHeaders.length > 0) {
      dataRows = this.applySort(dataRows, sortConfig.sortHeaders, finalHeaders);
    }

    // 按分类分组
    const categoryIndex = finalHeaders.indexOf(categorySummaryConfig.categoryHeader);
    if (categoryIndex < 0) return;

    const grouped = new Map<string, Array<{ cells: ExcelJS.Cell[], originalRowNum: number }>>();
    dataRows.forEach(row => {
      const categoryCell = row.cells[categoryIndex];
      const categoryValue = categoryCell ? String(this.excelUtils.getCellValue(categoryCell) || '') : '';
      if (!grouped.has(categoryValue)) {
        grouped.set(categoryValue, []);
      }
      grouped.get(categoryValue)!.push(row);
    });

    // 为每个分类创建sheet
    const categoryValues = Array.from(grouped.keys());
    for (const categoryValue of categoryValues) {
      const categoryDataRows = grouped.get(categoryValue)!;

      // 生成sheet名称（Excel工作表名称限制31个字符，不能包含: \ / ? * [ ]）
      let sheetName = String(categoryValue || '未分类');
      sheetName = sheetName.replace(/[:\\\/\?\*\[\]]/g, '_');
      sheetName = sheetName.substring(0, 31);

      // 如果baseSheetName不是默认名称，可以加上前缀
      if (baseSheetName && baseSheetName !== 'Sheet1') {
        const prefix = baseSheetName.substring(0, 15);
        sheetName = `${prefix}_${sheetName}`.substring(0, 31);
      }

      // 确保sheet名称唯一
      let finalSheetName = sheetName;
      let counter = 1;
      while (workbook.getWorksheet(finalSheetName)) {
        finalSheetName = `${sheetName}_${counter}`.substring(0, 31);
        counter++;
      }

      const categorySheet = workbook.addWorksheet(finalSheetName);

      // 添加分类标题行
      const categoryHeader = categorySummaryConfig.categoryHeader;
      const titleRow = categorySheet.addRow([`${categoryHeader}: ${categoryValue}`]);
      titleRow.height = 25 * 0.75;
      this.excelUtils.applyCellStyle(titleRow.getCell(1), style, 'title');
      categorySheet.mergeCells(1, 1, 1, finalHeaders.length);

      // 写入表头行
      headerRows.forEach((headerRow) => {
        const row = categorySheet.addRow(headerRow);
        row.height = 22 * 0.75;
        row.eachCell((cell, colNumber) => {
          this.excelUtils.applyCellStyle(cell, style, 'header');
        });
      });

      // 写入数据行（数据从第 headerRowCount + 2 行开始，因为第1行是标题，第2行开始是表头）
      const dataStartRow = headerRowCount + 2;
      const originalHeaderIndexMap = new Map<string, number>();
      const originalHeaders = this.headersData().length > 0 && this.headersData()[0].headers.length > 0
        ? this.headersData()[0].headers[0]
        : [];
      originalHeaders.forEach((header, index) => {
        originalHeaderIndexMap.set(header, index);
      });

      let currentCategoryRow = dataStartRow; // 分类表数据从dataStartRow开始
      for (let rowIndex = 0; rowIndex < categoryDataRows.length; rowIndex++) {
        const rowData = categoryDataRows[rowIndex];
        const rowValues: any[] = [];
        const formulaMap = new Map<number, string>();
        const currentOriginalRowNum = rowData.originalRowNum;

        for (let colIndex = 0; colIndex < finalHeaders.length; colIndex++) {
          const header = finalHeaders[colIndex];
          const headerIndex = originalHeaderIndexMap.get(header);

          if (headerIndex !== undefined && currentOriginalRowNum > 0) {
            // 核心逻辑：从 originalSheet 直接获取单元格对象（参照汇总分类页面）
            // 找到该表头在原表中的物理列号
            let originalColNumber = -1;
            originalSheet.getRow(1).eachCell({ includeEmpty: true }, (cell, col) => {
              if (String(cell.value || '').trim() === header) {
                originalColNumber = col;
              }
            });

            if (originalColNumber !== -1) {
              const originalCell = originalSheet.getRow(currentOriginalRowNum).getCell(originalColNumber);

              // 检查是否是公式（处理对象形式或直接属性）
              let formulaText = '';
              if (originalCell.formula) {
                formulaText = originalCell.formula;
              } else if (typeof originalCell.value === 'object' && originalCell.value !== null && 'formula' in originalCell.value) {
                formulaText = (originalCell.value as any).formula;
              }

              if (formulaText) {
                // 解析并转换公式（参照汇总分类页面）
                // 传入分类表数据行数，确保不会引用表尾行
                const convertedFormula = this.convertFormulaForNewHeaders(
                  formulaText,
                  currentOriginalRowNum,
                  currentCategoryRow,
                  colIndex + 1,
                  finalHeaders,
                  originalHeaderIndexMap,
                  originalSheet,
                  categoryDataRows.length // 分类表数据行数
                );

                if (convertedFormula) {
                  formulaMap.set(colIndex + 1, convertedFormula);
                  rowValues.push(null); // 占位，稍后会被公式替换
                  continue; // 跳过后续的普通值处理
                } else {
                  // 公式转换失败，记录日志以便调试
                  rowValues.push('F-Null'); // 公式依赖缺失
                  continue;
                }
              } else {
                // 不是公式，获取原始值并保持数据类型
                const cellValue = originalCell.value;
                // 优先使用 value（原始值），如果是 number 类型则保持为 number
                let finalValue: any = null;
                if (cellValue !== null && cellValue !== undefined) {
                  if (typeof cellValue === 'number') {
                    // 数字类型，直接保持为 number
                    finalValue = cellValue;
                  } else if (typeof cellValue === 'object' && 'text' in cellValue) {
                    // 超链接对象，使用 text
                    finalValue = (cellValue as any).text || '';
                  } else if (cellValue instanceof Date) {
                    // 日期类型，保持 Date 对象
                    finalValue = cellValue;
                  } else {
                    // 其他类型（string等），直接使用
                    finalValue = cellValue;
                  }
                } else if (originalCell.result !== null && originalCell.result !== undefined) {
                  // 如果 value 为空，尝试使用 result
                  // 如果 result 是 number，保持为 number
                  if (typeof originalCell.result === 'number') {
                    finalValue = originalCell.result;
                  } else {
                    finalValue = originalCell.result;
                  }
                } else {
                  // 都为空，使用空字符串
                  finalValue = '';
                }
                rowValues.push(finalValue);
                continue; // 跳过后续的普通值处理
              }
            } else {
              // 如果找不到原始列号，使用空值
              rowValues.push('');
            }
          } else {
            // 如果找不到原始列号或行号，使用空值
            rowValues.push('');
          }
        }

        const row = categorySheet.addRow(rowValues);
        row.height = 20 * 0.75;
        // 遍历所有单元格（包括空值），确保公式能被设置
        for (let colNumber = 1; colNumber <= finalHeaders.length; colNumber++) {
          const cell = row.getCell(colNumber);
          // 如果这个单元格有公式，设置公式
          if (formulaMap.has(colNumber)) {
            const formula = formulaMap.get(colNumber)!;
            cell.value = { formula: formula };
          }
          this.excelUtils.applyCellStyle(cell, style, 'data');
        }
        currentCategoryRow++;
      }

      // 添加表尾合计行
      const footerRow: any[] = [];
      finalHeaders.forEach((header) => {
        if (categorySummaryConfig.summaryHeaders.includes(header)) {
          const colIndex = finalHeaders.indexOf(header);
          const colName = this.excelUtils.getExcelColumnName(colIndex + 1);
          const startRow = dataStartRow; // 数据开始行
          const endRow = categorySheet.rowCount; // 当前最后一行（数据结束行）
          footerRow.push({ formula: `SUM(${colName}${startRow}:${colName}${endRow})` });
        } else {
          footerRow.push('');
        }
      });
      const footerRowObj = categorySheet.addRow(footerRow);
      footerRowObj.height = 22 * 0.75;
      footerRowObj.eachCell((cell, colNumber) => {
        if (footerRow[colNumber - 1] && typeof footerRow[colNumber - 1] === 'object' && 'formula' in footerRow[colNumber - 1]) {
          cell.value = { formula: (footerRow[colNumber - 1] as any).formula };
        }
        this.excelUtils.applyCellStyle(cell, style, 'total');
      });

      // 自动调整列宽
      this.excelUtils.autoFitColumns(categorySheet, finalHeaders.length);
    }
  }
}

