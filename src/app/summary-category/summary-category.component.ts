import { Component, signal, AfterViewInit, ViewChild, ElementRef, effect } from '@angular/core';
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
import { MatStepperModule, MatStepper } from '@angular/material/stepper';
import { MatChipsModule } from '@angular/material/chips';
import { MatProgressBarModule } from '@angular/material/progress-bar';
import { MatListModule } from '@angular/material/list';
import { MatRadioModule } from '@angular/material/radio';
import { MatTooltipModule } from '@angular/material/tooltip';
import { DragDropModule, CdkDragDrop, moveItemInArray } from '@angular/cdk/drag-drop';
import { TableStyleConfig, DEFAULT_TABLE_STYLE } from '../shared/models/table-style.model';
import * as ExcelJS from 'exceljs';

@Component({
  selector: 'app-summary-category',
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
    MatListModule,
    MatRadioModule,
    MatTooltipModule,
    DragDropModule
  ],
  templateUrl: './summary-category.component.html',
  styleUrl: './summary-category.component.scss'
})
export class SummaryCategoryComponent implements AfterViewInit {
  @ViewChild('connectionArea', { static: false }) connectionAreaRef?: ElementRef<HTMLDivElement>;
  @ViewChild('stepper', { static: false }) stepper?: MatStepper;
  private viewInitialized = signal<boolean>(false);
  tableStyle = signal<TableStyleConfig>({ ...DEFAULT_TABLE_STYLE });
  previewExpanded = signal<boolean>(false); // 预览区域展开状态，默认折叠

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
  headers = signal<string[]>([]);
  rawData: any[][] = []; // 存储原始Excel数据（包括表头和数据行）
  originalWorkbook: ExcelJS.Workbook | null = null; // 保存原始Excel工作簿，用于读取公式和超链接
  originalSheet: ExcelJS.Worksheet | null = null; // 保存原始工作表
  categoryHeader = signal<string>(''); // Step2: 分类依据表头（单选，不能新增）
  summaryHeaders = signal<string[]>([]); // Step3: 汇总表表头（可新增）
  categoryTableHeaders = signal<string[]>([]); // Step4: 分类表表头（可新增）
  newHeaderName = signal<string>(''); // Step3和Step4共用
  newHeaderNameStep3 = signal<string>(''); // Step3新增表头
  isUploading = signal<boolean>(false);
  isProcessing = signal<boolean>(false); // 处理中状态
  showConfirmDialog = signal<boolean>(false); // 显示确认对话框
  outputFileName = signal<string>(''); // 输出文件名
  generationProgress = signal<number>(0); // 生成进度（0-100）
  // Step5: 表头关联映射（汇总表表头 -> 分类表表头）
  headerMappings = signal<Map<string, string>>(new Map());
  selectedSummaryHeaderForMapping = signal<string>(''); // 当前选中的汇总表表头（用于关联）

  // Step3: 表尾功能配置
  footerFunctions = signal<Array<{ type: '合计' | '平均值', header: string, id: string }>>([]);
  newFooterType = signal<'合计' | '平均值'>('合计'); // Step3新增表尾功能类型
  newFooterHeader = signal<string>(''); // Step3新增表尾功能关联的表头

  // Step4: 表尾功能配置
  categoryFooterFunctions = signal<Array<{ type: '合计' | '平均值', header: string, id: string }>>([]);
  newCategoryFooterType = signal<'合计' | '平均值'>('合计'); // Step4新增表尾功能类型
  newCategoryFooterHeader = signal<string>(''); // Step4新增表尾功能关联的表头

  constructor() {
    // 监听映射变化，更新连接线位置
    effect(() => {
      const mappings = this.headerMappings();
      if (this.viewInitialized()) {
        setTimeout(() => this.updateConnectionLines(), 100);
      }
    });

  }

  updateStyle(key: keyof TableStyleConfig, value: any) {
    this.tableStyle.update(config => ({ ...config, [key]: value }));
  }

  // 预览相关方法
  getPreviewHeaders(): string[] {
    // 如果有已选择的汇总表表头，使用它们；否则使用前3个表头或默认表头
    if (this.summaryHeaders().length > 0) {
      return this.summaryHeaders().slice(0, 4); // 最多显示4列
    }
    if (this.headers().length > 0) {
      return this.headers().slice(0, 4);
    }
    return ['列1', '列2', '列3', '列4'];
  }

  getPreviewData(): string[][] {
    // 生成示例数据行
    const headers = this.getPreviewHeaders();
    return [
      headers.map((_, i) => `示例数据${i + 1}-1`),
      headers.map((_, i) => `示例数据${i + 1}-2`),
      headers.map((_, i) => `示例数据${i + 1}-3`)
    ];
  }

  getPreviewTotalValue(): string {
    return '100.00';
  }

  getPreviewTotalCells(): number[] {
    // 返回合计行需要的单元格数量（除了第一列的"合计"标签）
    const headers = this.getPreviewHeaders();
    return Array.from({ length: headers.length - 1 }, (_, i) => i);
  }

  getBorderStyle(): string {
    const style = this.tableStyle().borderStyle;
    const color = this.tableStyle().borderColor;
    const widthMap: Record<string, string> = {
      'thin': '1px',
      'medium': '2px',
      'thick': '3px'
    };
    const width = widthMap[style] || '1px';
    return `${width} solid ${color}`;
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

  // 步骤变化处理
  onStepChange(event: any) {
    // 预览会自动更新，因为使用了响应式数据
  }

  async readExcelFile(file: File) {
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

      // 保存原始工作簿和工作表
      this.originalWorkbook = workbook;

      // 读取第一个sheet
      const firstSheet = workbook.worksheets[0];
      this.originalSheet = firstSheet;
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

      this.headers.set(headers);

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

      this.rawData = allData;
    } catch (error: any) {
      // 重新抛出错误以便上层处理
      if (error instanceof Error) {
        throw error;
      }
      throw new Error(`读取文件时发生错误：${String(error)}`);
    }
  }

  getRemainingHeadersForSummary(): string[] {
    // Step3: 汇总表的表头 = 所有表头 - Step2选择的分类依据表头
    const categoryHeader = this.categoryHeader();
    if (!categoryHeader) return this.headers();
    return this.headers().filter(h => h !== categoryHeader);
  }

  getRemainingHeadersForCategory(): string[] {
    // Step4: 分类表的表头 = 所有表头 - Step2选择的分类依据表头
    const categoryHeader = this.categoryHeader();
    if (!categoryHeader) return this.headers();
    return this.headers().filter(h => h !== categoryHeader);
  }

  selectCategoryHeader(header: string) {
    // Step2: 选择分类依据表头（单选）
    this.categoryHeader.set(header);
  }

  toggleSummaryHeader(header: string) {
    // Step3: 切换汇总表表头
    const current = this.summaryHeaders();
    if (current.includes(header)) {
      this.summaryHeaders.set(current.filter(h => h !== header));
    } else {
      this.summaryHeaders.set([...current, header]);
    }
  }

  toggleCategoryTableHeader(header: string) {
    // Step4: 切换分类表表头
    const current = this.categoryTableHeaders();
    if (current.includes(header)) {
      this.categoryTableHeaders.set(current.filter(h => h !== header));
    } else {
      this.categoryTableHeaders.set([...current, header]);
    }
  }

  addNewHeaderStep3() {
    const name = this.newHeaderNameStep3().trim();
    if (name && !this.summaryHeaders().includes(name)) {
      this.summaryHeaders.set([...this.summaryHeaders(), name]);
      this.newHeaderNameStep3.set('');
    }
  }

  addNewHeaderStep4() {
    const name = this.newHeaderName().trim();
    if (name && !this.categoryTableHeaders().includes(name)) {
      this.categoryTableHeaders.set([...this.categoryTableHeaders(), name]);
      this.newHeaderName.set('');
    }
  }

  clearCategoryHeader() {
    // Step2: 清除分类依据表头
    this.categoryHeader.set('');
  }

  removeSummaryHeader(header: string) {
    // Step3: 移除汇总表表头
    this.summaryHeaders.set(this.summaryHeaders().filter(h => h !== header));
    // 同时移除该表头关联的所有表尾功能
    this.footerFunctions.set(
      this.footerFunctions().filter(f => f.header !== header)
    );
  }

  removeCategoryTableHeader(header: string) {
    // Step4: 移除分类表表头
    this.categoryTableHeaders.set(this.categoryTableHeaders().filter(h => h !== header));
    // 同时移除该表头关联的所有表尾功能
    this.categoryFooterFunctions.set(
      this.categoryFooterFunctions().filter(f => f.header !== header)
    );
  }

  dropSummaryHeader(event: CdkDragDrop<string[]>) {
    // Step3: 拖拽排序汇总表表头
    if (event.previousIndex === event.currentIndex) {
      return; // 位置没有变化，不需要更新
    }
    const headers = [...this.summaryHeaders()];
    moveItemInArray(headers, event.previousIndex, event.currentIndex);
    this.summaryHeaders.set(headers);
  }

  dropCategoryTableHeader(event: CdkDragDrop<string[]>) {
    // Step4: 拖拽排序分类表表头
    if (event.previousIndex === event.currentIndex) {
      return; // 位置没有变化，不需要更新
    }
    const headers = [...this.categoryTableHeaders()];
    moveItemInArray(headers, event.previousIndex, event.currentIndex);
    this.categoryTableHeaders.set(headers);
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

  // Step4: 添加表尾功能
  addCategoryFooterFunction() {
    const header = this.newCategoryFooterHeader().trim();
    const type = this.newCategoryFooterType();

    if (!header) {
      return; // 如果没有选择表头，不添加
    }

    // 检查该表头是否已经配置了相同类型的表尾功能
    const existing = this.categoryFooterFunctions().find(
      f => f.header === header && f.type === type
    );

    if (existing) {
      return; // 如果已经存在，不重复添加
    }

    // 添加新的表尾功能配置
    const newFooter = {
      type: type,
      header: header,
      id: `category-${type}-${header}-${Date.now()}` // 生成唯一ID
    };

    this.categoryFooterFunctions.set([...this.categoryFooterFunctions(), newFooter]);
    this.newCategoryFooterHeader.set(''); // 重置选择
  }

  // Step4: 移除表尾功能
  removeCategoryFooterFunction(id: string) {
    this.categoryFooterFunctions.set(
      this.categoryFooterFunctions().filter(f => f.id !== id)
    );
  }

  // Step4: 获取表尾功能显示文本
  getCategoryFooterFunctionLabel(footer: { type: '合计' | '平均值', header: string }): string {
    return `${footer.type}(${footer.header})`;
  }

  // Step5: 选择汇总表表头进行关联
  selectSummaryHeaderForMapping(header: string) {
    const currentSelected = this.selectedSummaryHeaderForMapping();
    const mappings = this.headerMappings();

    // 如果点击的是已选中的表头，则取消选择
    if (currentSelected === header) {
      this.selectedSummaryHeaderForMapping.set('');
      setTimeout(() => this.updateConnectionLines(), 50);
      return;
    }

    // 如果该表头已经有关联，取消关联
    if (mappings.has(header)) {
      const newMappings = new Map(mappings);
      newMappings.delete(header);
      this.headerMappings.set(newMappings);
      this.selectedSummaryHeaderForMapping.set('');
      setTimeout(() => this.updateConnectionLines(), 50);
      return;
    }

    // 选择新的表头
    this.selectedSummaryHeaderForMapping.set(header);
  }

  // Step5: 获取分类表的所有选项（只包括表尾功能：合计或平均值）
  getCategoryTableOptions(): Array<{ id: string, label: string, type: 'header' | 'footer', icon?: string }> {
    const options: Array<{ id: string, label: string, type: 'header' | 'footer', icon?: string }> = [];

    // 只添加表尾功能（合计或平均值），不添加表头选项
    for (const footer of this.categoryFooterFunctions()) {
      options.push({
        id: `footer:${footer.id}`,
        label: this.getCategoryFooterFunctionLabel(footer),
        type: 'footer',
        icon: footer.type === '合计' ? 'calculate' : 'trending_up'
      });
    }

    return options;
  }

  // Step5: 选择分类表选项（表头或表尾功能）进行关联
  selectCategoryOptionForMapping(optionId: string) {
    const selectedSummary = this.selectedSummaryHeaderForMapping();
    if (!selectedSummary) {
      return; // 如果没有选中汇总表表头，直接返回
    }

    const mappings = new Map(this.headerMappings());

    // 如果该分类表选项已经被其他汇总表表头关联，先取消之前的关联
    for (const [summaryHeader, categoryOptionId] of mappings.entries()) {
      if (categoryOptionId === optionId && summaryHeader !== selectedSummary) {
        mappings.delete(summaryHeader);
        break;
      }
    }

    // 如果点击的是已关联的选项，则取消关联
    if (mappings.get(selectedSummary) === optionId) {
      mappings.delete(selectedSummary);
      this.selectedSummaryHeaderForMapping.set('');
    } else {
      // 建立新的关联
      mappings.set(selectedSummary, optionId);
      this.selectedSummaryHeaderForMapping.set('');
    }

    this.headerMappings.set(mappings);
    setTimeout(() => this.updateConnectionLines(), 50);
  }

  // Step5: 检查汇总表表头是否已选中用于关联
  isSummaryHeaderSelectedForMapping(header: string): boolean {
    return this.selectedSummaryHeaderForMapping() === header;
  }

  // Step5: 检查汇总表表头是否已关联
  isSummaryHeaderMapped(header: string): boolean {
    return this.headerMappings().has(header);
  }

  // Step5: 获取汇总表表头对应的分类表表头
  getMappedCategoryHeader(summaryHeader: string): string | undefined {
    return this.headerMappings().get(summaryHeader);
  }

  // Step5: 检查分类表选项（表头或表尾功能）是否已关联
  isCategoryOptionMapped(optionId: string): boolean {
    const mappings = this.headerMappings();
    for (const categoryOptionId of mappings.values()) {
      if (categoryOptionId === optionId) {
        return true;
      }
    }
    return false;
  }

  // Step5: 获取关联到指定分类表选项的汇总表表头
  getMappedSummaryHeader(categoryOptionId: string): string | undefined {
    const mappings = this.headerMappings();
    for (const [summaryHeader, mappedCategoryOptionId] of mappings.entries()) {
      if (mappedCategoryOptionId === categoryOptionId) {
        return summaryHeader;
      }
    }
    return undefined;
  }

  // Step5: 移除关联
  removeMapping(summaryHeader: string) {
    const mappings = new Map(this.headerMappings());
    mappings.delete(summaryHeader);
    this.headerMappings.set(mappings);
    setTimeout(() => this.updateConnectionLines(), 50);
  }

  // Step5: 获取表头在列表中的索引位置（用于计算连接线位置）
  getSummaryHeaderIndex(header: string): number {
    return this.summaryHeaders().indexOf(header);
  }

  getCategoryHeaderIndex(header: string): number {
    return this.categoryTableHeaders().indexOf(header);
  }

  // Step5: 根据选项ID获取显示标签
  getCategoryOptionLabel(optionId: string): string {
    if (optionId.startsWith('header:')) {
      return optionId.replace('header:', '');
    } else if (optionId.startsWith('footer:')) {
      const footerId = optionId.replace('footer:', '');
      const footer = this.categoryFooterFunctions().find(f => f.id === footerId);
      return footer ? this.getCategoryFooterFunctionLabel(footer) : optionId;
    }
    return optionId;
  }

  ngAfterViewInit() {
    this.viewInitialized.set(true);
    // 初始化时更新一次连接线
    setTimeout(() => this.updateConnectionLines(), 100);

    // 监听窗口resize，更新连接线位置
    window.addEventListener('resize', () => {
      this.updateConnectionLines();
    });
  }

  // 根据当前步骤获取预览数据
  getPreviewDataForCurrentStep(): any[][] {
    const currentStep = this.getCurrentStepIndex();

    // Step 1: 显示原始 Excel 数据（添加标题行）
    if (currentStep >= 0 && this.rawData.length > 0) {
      const rawDataRows = this.rawData.slice(0, Math.min(11, this.rawData.length));
      // 添加标题行
      const headers = rawDataRows[0] || [];
      const titleRow = ['示例表格标题', ...headers.slice(1).map(() => '')];
      return [titleRow, ...rawDataRows];
    }

    // Step 2+: 根据已选择的表头显示预览
    if (this.summaryHeaders().length > 0) {
      return this.getPreviewDataForSummaryTable();
    }

    // 未上传文件时，返回示例预览数据
    return this.getDefaultPreviewData();
  }

  // 获取默认预览数据（未上传文件时使用）
  getDefaultPreviewData(): any[][] {
    // 创建示例表头和数据
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

  // 获取汇总表预览数据
  getPreviewDataForSummaryTable(): any[][] {
    const headers = this.summaryHeaders();
    if (headers.length === 0) return [];

    // 构建预览数据
    const previewData: any[][] = [];

    // 添加标题行（第一行，跨所有列）
    const titleRow = ['示例表格标题', ...headers.slice(1).map(() => '')];
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

  // 获取当前步骤索引
  getCurrentStepIndex(): number {
    if (!this.stepper) return -1;
    return this.stepper.selectedIndex ?? -1;
  }

  // 获取预览数据（供模板使用）
  get previewData(): any[][] {
    return this.getPreviewDataForCurrentStep();
  }

  updateConnectionLines() {
    if (!this.connectionAreaRef) return;

    const connectionArea = this.connectionAreaRef.nativeElement;
    const container = connectionArea.parentElement;
    if (!container) return;

    const summaryColumn = container.querySelector('.summary-headers-column .header-list-mapping');
    const categoryColumn = container.querySelector('.category-headers-column .header-list-mapping');

    if (!summaryColumn || !categoryColumn) return;

    const mappings = this.headerMappings();
    const lines = connectionArea.querySelectorAll('.connection-line');

    lines.forEach((line) => {
      const lineElement = line as HTMLElement;
      const summaryHeader = lineElement.getAttribute('data-summary');
      const categoryHeader = lineElement.getAttribute('data-category');

      if (!summaryHeader || !categoryHeader) return;

      const summaryItem = summaryColumn.querySelector(`[data-header-id="summary-${summaryHeader}"]`) as HTMLElement;
      const categoryItem = categoryColumn.querySelector(`[data-header-id="category-${categoryHeader}"]`) as HTMLElement;

      if (summaryItem && categoryItem) {
        const containerRect = container.getBoundingClientRect();
        const summaryRect = summaryItem.getBoundingClientRect();
        const categoryRect = categoryItem.getBoundingClientRect();

        const startX = summaryRect.right - containerRect.left;
        const startY = summaryRect.top + summaryRect.height / 2 - containerRect.top;
        const endX = categoryRect.left - containerRect.left;
        const endY = categoryRect.top + categoryRect.height / 2 - containerRect.top;

        const dx = endX - startX;
        const dy = endY - startY;
        const length = Math.sqrt(dx * dx + dy * dy);
        const angle = Math.atan2(dy, dx) * 180 / Math.PI;

        lineElement.style.left = `${startX}px`;
        lineElement.style.top = `${startY}px`;
        lineElement.style.width = `${length}px`;
        lineElement.style.transform = `rotate(${angle}deg)`;
        lineElement.style.transformOrigin = 'left center';
      }
    });
  }

  // Step5: 完成按钮点击处理
  onComplete() {
    if (!this.selectedFile) {
      alert('请先上传Excel文件');
      return;
    }

    if (!this.categoryHeader()) {
      alert('请先选择分类依据');
      return;
    }

    if (this.summaryHeaders().length === 0) {
      alert('请先设置汇总表表头');
      return;
    }

    if (this.categoryTableHeaders().length === 0) {
      alert('请先设置分类表表头');
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
    const fileName = this.outputFileName().trim() || '汇总分类表';
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
  async generateExcel(fileName: string = '汇总分类表') {
    const workbook = new ExcelJS.Workbook();
    const style = this.tableStyle();

    // 获取表头索引映射
    const headerIndexMap = new Map<string, number>();
    this.headers().forEach((header, index) => {
      headerIndexMap.set(header, index);
    });

    const categoryHeaderIndex = headerIndexMap.get(this.categoryHeader()!)!;

    // 数据去重：锁定原始物理行号
    const uniqueRows = new Set<string>();
    const deduplicatedRawData: Array<{ data: any[], originalIndex: number }> = [];
    for (let i = 1; i < this.rawData.length; i++) {
      const row = this.rawData[i];
      const rowJson = JSON.stringify(row);
      if (!uniqueRows.has(rowJson)) {
        uniqueRows.add(rowJson);
        deduplicatedRawData.push({ data: row, originalIndex: i + 1 }); // i=1 是第2行
      }
    }

    // 按分类依据分组数据（保持行号关联）
    const groupedData = new Map<string, Array<{ data: any[], originalIndex: number }>>();

    for (const item of deduplicatedRawData) {
      const categoryValue = item.data[categoryHeaderIndex] || '';

      if (!groupedData.has(categoryValue)) {
        groupedData.set(categoryValue, []);
      }
      groupedData.get(categoryValue)!.push(item);
    }

    // 创建汇总表工作表
    const summarySheet = workbook.addWorksheet('汇总表');

    // 为每个分类创建分类表工作表（先创建，以便汇总表可以引用）
    // 汇总表的分类需要去重：使用 Set 确保每个分类值只出现一次
    const uniqueCategoryValues = Array.from(new Set(Array.from(groupedData.keys()))).sort();
    const categoryValues = uniqueCategoryValues;
    const categorySheetMap = new Map<string, ExcelJS.Worksheet>(); // 存储分类表映射

    for (const categoryValue of categoryValues) {
      const categoryData = groupedData.get(categoryValue)!;
      let sheetName = String(categoryValue || '未分类');
      // Excel工作表名称限制31个字符，不能包含: \ / ? * [ ]
      sheetName = sheetName.replace(/[:\\\/\?\*\[\]]/g, '_');
      sheetName = sheetName.substring(0, 31);
      const categorySheet = workbook.addWorksheet(sheetName);
      categorySheetMap.set(categoryValue, categorySheet);
    }

    // 创建汇总表（传入分类表映射以便添加超链接和引用）
    await this.createSummarySheet(summarySheet, groupedData, style, categorySheetMap, workbook);

    // 创建分类表（传入汇总表以便添加返回链接）
    for (const categoryValue of categoryValues) {
      const categoryData = groupedData.get(categoryValue)!;
      const categorySheet = categorySheetMap.get(categoryValue)!;
      await this.createCategorySheet(categorySheet, categoryValue, categoryData, style, summarySheet);
    }

    // 生成Excel文件并返回blob（不直接下载）
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return blob;
  }

  // 创建汇总表
  async createSummarySheet(
    sheet: ExcelJS.Worksheet,
    groupedData: Map<string, Array<{ data: any[], originalIndex: number }>>,
    style: TableStyleConfig,
    categorySheetMap: Map<string, ExcelJS.Worksheet>,
    workbook: ExcelJS.Workbook
  ) {
    const summaryHeaders = this.summaryHeaders();
    const categoryHeader = this.categoryHeader()!;
    const headerIndexMap = new Map<string, number>();
    this.headers().forEach((header, index) => {
      headerIndexMap.set(header, index);
    });

    // 添加标题行
    const titleRow = sheet.addRow([`汇总表（按${categoryHeader}分类）`]);
    titleRow.height = 25 * 0.75; // 25像素转换为磅（1像素 ≈ 0.75磅）
    titleRow.getCell(1).font = {
      name: style.fontFamily,
      size: style.titleFontSize,
      bold: style.titleFontBold,
      color: { argb: this.hexToArgb(style.titleFontColor) }
    };
    titleRow.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: this.hexToArgb(style.titleColor) }
    };
    titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
    sheet.mergeCells(1, 1, 1, summaryHeaders.length);

    // 添加表头行
    const headerRow = sheet.addRow(summaryHeaders);
    headerRow.height = 22 * 0.75; // 22像素转换为磅
    headerRow.eachCell((cell, colNumber) => {
      cell.font = {
        name: style.fontFamily,
        size: style.headerFontSize,
        bold: style.headerFontBold,
        color: { argb: this.hexToArgb(style.headerFontColor) }
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: this.hexToArgb(style.headerColor) }
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      cell.border = {
        top: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
        left: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
        bottom: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
        right: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } }
      };
    });

    // 添加数据行
    const categoryValues = Array.from(groupedData.keys()).sort();
    const mappings = this.headerMappings();
    const categoryFooterFunctions = this.categoryFooterFunctions();
    const categoryHeaderColIndex = summaryHeaders.indexOf(categoryHeader); // 分类依据列索引

    for (let rowIndex = 0; rowIndex < categoryValues.length; rowIndex++) {
      const categoryValue = categoryValues[rowIndex];
      const categoryData = groupedData.get(categoryValue)!;
      const categorySheet = categorySheetMap.get(categoryValue)!;
      const dataRow: any[] = [];
      const dataRowIndex = rowIndex + 3; // 第1行是标题，第2行是表头，数据从第3行开始

      // 存储每列的公式信息
      const formulaMap = new Map<number, string>(); // 列索引 -> 公式

      for (let colIndex = 0; colIndex < summaryHeaders.length; colIndex++) {
        const summaryHeader = summaryHeaders[colIndex];
        let cellValue: any = '';
        let cellFormula: string | null = null;

        // 检查是否有映射到分类表（只能映射到表尾功能：合计或平均值）
        const mappedCategoryOptionId = mappings.get(summaryHeader);
        const categoryTableHeaders = this.categoryTableHeaders(); // 获取分类表表头
        if (mappedCategoryOptionId) {
          // 汇总表只能关联分类表的表尾功能（合计或平均值），不能关联数据单元格
          if (mappedCategoryOptionId.startsWith('footer:')) {
            // 映射到分类表表尾功能 - 使用公式引用分类表的表尾合计单元格
            const footerId = mappedCategoryOptionId.replace('footer:', '');
            const footer = categoryFooterFunctions.find(f => f.id === footerId);
            if (footer) {
              const mappedHeader = footer.header;
              const mappedHeaderIndex = categoryTableHeaders.indexOf(mappedHeader);
              if (mappedHeaderIndex >= 0) {
                // 引用分类表的表尾行（数据行数 + 4，因为第1行是返回链接，第2行是标题，第3行是表头，数据从第4行开始）
                const footerRow = categoryData.length + 4;
                const targetCol = this.getExcelColumnName(mappedHeaderIndex + 1);
                // 工作表名称必须用单引号包裹（Excel要求）
                const sheetName = `'${categorySheet.name}'`;
                cellFormula = `${sheetName}!${targetCol}${footerRow}`;
                cellValue = null; // 不在前端计算，让 Excel 处理
              } else {
                // 如果分类表中没有该表头，填充 F-Null
                cellValue = 'F-Null';
              }
            } else {
              // 如果找不到对应的表尾功能配置，填充 F-Null
              cellValue = 'F-Null';
            }
          } else {
            // 如果映射ID格式不正确，填充 F-Null
            cellValue = 'F-Null';
          }
        } else {
          // 没有映射，使用汇总表表头对应的原始数据
          const summaryHeaderIndex = headerIndexMap.get(summaryHeader);
          if (summaryHeaderIndex !== undefined) {
            // 如果是分类依据表头，显示分类值并添加超链接
            if (summaryHeader === categoryHeader) {
              cellValue = categoryValue;
            } else {
              // 检查是否有表尾功能配置
              const footer = this.footerFunctions().find(f => f.header === summaryHeader);
              if (footer) {
                // 如果汇总表该列有合计/平均值功能，但没有手动映射，
                // 我们尝试通过公式引用对应分类表的整列数据进行计算
                const catHeaderIdx = categoryTableHeaders.indexOf(summaryHeader);
                if (catHeaderIdx >= 0) {
                  const targetCol = this.getExcelColumnName(catHeaderIdx + 1);
                  const startRow = 4;
                  const endRow = 3 + categoryData.length;
                  const func = footer.type === '合计' ? 'SUM' : 'AVERAGE';
                  cellFormula = `${func}('${categorySheet.name}'!${targetCol}${startRow}:${targetCol}${endRow})`;
                  cellValue = null;
                } else {
                  cellValue = 'F-Null';
                }
              } else {
                // 如果没有表尾功能，但分类表有该表头，则通过公式引用分类表的第一行数据
                const catHeaderIdx = categoryTableHeaders.indexOf(summaryHeader);
                if (catHeaderIdx >= 0) {
                  const targetCol = this.getExcelColumnName(catHeaderIdx + 1);
                  cellFormula = `'${categorySheet.name}'!${targetCol}4`; // 引用分类表第一行数据
                  cellValue = null;
                } else {
                  cellValue = 'F-Null';
                }
              }
            }
          } else {
            cellValue = 'F-Null';
          }
        }

        // 如果有公式，存储到 formulaMap 中
        if (cellFormula) {
          formulaMap.set(colIndex + 1, cellFormula);
          // 如果有公式，推入 null 作为占位符，稍后会被公式替换
          dataRow.push(null);
        } else {
          dataRow.push(cellValue);
        }
      }

      const row = sheet.addRow(dataRow);
      row.height = 20 * 0.75; // 20像素转换为磅
      // 遍历所有单元格（包括空值），确保公式能被设置
      for (let colNumber = 1; colNumber <= summaryHeaders.length; colNumber++) {
        const cell = row.getCell(colNumber);
        const summaryHeader = summaryHeaders[colNumber - 1];

        // 设置公式或超链接
        if (formulaMap.has(colNumber)) {
          const formula = formulaMap.get(colNumber)!;
          cell.value = { formula: formula };
        } else if (summaryHeader === categoryHeader) {
          // 分类列添加超链接
          cell.value = {
            text: categoryValue,
            hyperlink: `#${categorySheet.name}!A1`
          };
          cell.font = {
            name: style.fontFamily,
            size: style.dataFontSize,
            bold: style.dataFontBold,
            underline: true,
            color: { argb: 'FF0000FF' } // 蓝色
          };
        }

        cell.font = cell.font || {
          name: style.fontFamily,
          size: style.dataFontSize,
          bold: style.dataFontBold,
          color: { argb: this.hexToArgb(style.dataFontColor) }
        };
        cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };
        cell.border = {
          top: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          left: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          bottom: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          right: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } }
        };
      }
    }

    // 添加表尾功能行
    const footerFunctions = this.footerFunctions();
    if (footerFunctions.length > 0) {
      const footerRow: any[] = [];

      for (const summaryHeader of summaryHeaders) {
        const footer = footerFunctions.find(f => f.header === summaryHeader);
        if (footer) {
          // 检查该表头是否映射到分类表的表头或表尾
          const mappedCategoryOptionId = mappings.get(summaryHeader);
          const categoryTableHeaders = this.categoryTableHeaders();

          if (mappedCategoryOptionId && mappedCategoryOptionId.startsWith('footer:')) {
            // 如果映射到分类表的表尾功能，汇总表的表尾应该汇总所有分类表的表尾
            const footerId = mappedCategoryOptionId.replace('footer:', '');
            const categoryFooter = categoryFooterFunctions.find(f => f.id === footerId);
            if (categoryFooter) {
              const mappedHeader = categoryFooter.header;
              const mappedHeaderIndex = categoryTableHeaders.indexOf(mappedHeader);
              if (mappedHeaderIndex >= 0) {
                const targetCol = this.getExcelColumnName(mappedHeaderIndex + 1);
                // 构建引用所有分类表表尾的公式
                const footerRefs: string[] = [];
                for (const categoryValue of categoryValues) {
                  const categorySheet = categorySheetMap.get(categoryValue)!;
                  const categoryData = groupedData.get(categoryValue)!;
                  const footerRowNum = categoryData.length + 4; // 分类表表尾行号
                  footerRefs.push(`'${categorySheet.name}'!${targetCol}${footerRowNum}`);
                }
                if (footerRefs.length > 0) {
                  if (footer.type === '合计') {
                    footerRow.push({ formula: `SUM(${footerRefs.join(',')})` });
                  } else if (footer.type === '平均值') {
                    footerRow.push({ formula: `AVERAGE(${footerRefs.join(',')})` });
                  }
                } else {
                  footerRow.push('');
                }
              } else {
                footerRow.push('');
              }
            } else {
              footerRow.push('');
            }
          } else {
            // 如果没有映射或映射到表头，检查分类表中是否有该表头
            const catHeaderIdx = categoryTableHeaders.indexOf(summaryHeader);
            if (catHeaderIdx >= 0) {
              // 汇总所有分类表中该列的数据
              const targetCol = this.getExcelColumnName(catHeaderIdx + 1);
              const dataRefs: string[] = [];
              for (const categoryValue of categoryValues) {
                const categorySheet = categorySheetMap.get(categoryValue)!;
                const categoryData = groupedData.get(categoryValue)!;
                const startRow = 4; // 分类表数据从第4行开始
                const endRow = 3 + categoryData.length;
                dataRefs.push(`'${categorySheet.name}'!${targetCol}${startRow}:${targetCol}${endRow}`);
              }
              if (dataRefs.length > 0) {
                if (footer.type === '合计') {
                  footerRow.push({ formula: `SUM(${dataRefs.join(',')})` });
                } else if (footer.type === '平均值') {
                  // 对于平均值，需要先计算总和，再除以总行数
                  const totalRows = Array.from(groupedData.values()).reduce((sum, data) => sum + data.length, 0);
                  footerRow.push({ formula: `SUM(${dataRefs.join(',')})/${totalRows}` });
                }
              } else {
                footerRow.push('');
              }
            } else {
              // 如果分类表中没有该表头，且该表头不在原始数据中（新增的表头）
              // 这种情况下，表尾功能无法计算，填充空值
              footerRow.push('');
            }
          }
        } else {
          footerRow.push('');
        }
      }

      const row = sheet.addRow(footerRow);
      row.height = 22 * 0.75; // 22像素转换为磅
      row.eachCell((cell, colNumber) => {
        cell.font = {
          name: style.fontFamily,
          size: style.dataFontSize,
          bold: true,
          color: { argb: this.hexToArgb(style.dataFontColor) }
        };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: this.hexToArgb(style.totalColor) }
        };
        cell.alignment = { horizontal: 'right', vertical: 'middle', wrapText: false };
        cell.border = {
          top: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          left: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          bottom: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          right: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } }
        };
      });
    }

    // 自动调整列宽
    this.autoFitColumns(sheet, summaryHeaders.length);

    // 设置表头筛选器（从第2行表头行开始，到数据结束行）
    sheet.autoFilter = {
      from: { row: 2, column: 1 },
      to: { row: 2 + categoryValues.length, column: summaryHeaders.length }
    };
  }

  // 创建分类表
  async createCategorySheet(
    sheet: ExcelJS.Worksheet,
    categoryValue: string,
    categoryData: Array<{ data: any[], originalIndex: number }>,
    style: TableStyleConfig,
    summarySheet: ExcelJS.Worksheet
  ) {
    const categoryHeaders = this.categoryTableHeaders();
    const categoryHeader = this.categoryHeader()!;
    const headerIndexMap = new Map<string, number>();
    this.headers().forEach((header, index) => {
      headerIndexMap.set(header, index);
    });

    // 添加第一行：返回汇总表的链接（没有背景色）
    const returnRow = sheet.addRow(['返回汇总表']);
    returnRow.height = 25 * 0.75; // 25像素转换为磅
    const returnCell = returnRow.getCell(1);
    returnCell.value = {
      text: '返回汇总表',
      hyperlink: `#汇总表!A1`
    };
    returnCell.font = {
      name: style.fontFamily,
      size: style.titleFontSize - 2,
      bold: true,
      underline: true,
      color: { argb: 'FF0000FF' } // 蓝色
    };
    returnCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };
    // 不设置背景色

    // 添加标题行
    const titleRow = sheet.addRow([`${categoryHeader}: ${categoryValue}`]);
    titleRow.height = 25 * 0.75; // 25像素转换为磅
    titleRow.getCell(1).font = {
      name: style.fontFamily,
      size: style.titleFontSize,
      bold: style.titleFontBold,
      color: { argb: this.hexToArgb(style.titleFontColor) }
    };
    titleRow.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: this.hexToArgb(style.titleColor) }
    };
    titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
    sheet.mergeCells(2, 1, 2, categoryHeaders.length);

    // 添加表头行
    const headerRow = sheet.addRow(categoryHeaders);
    headerRow.height = 22 * 0.75; // 22像素转换为磅
    headerRow.eachCell((cell, colNumber) => {
      cell.font = {
        name: style.fontFamily,
        size: style.headerFontSize,
        bold: style.headerFontBold,
        color: { argb: this.hexToArgb(style.headerFontColor) }
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: this.hexToArgb(style.headerColor) }
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      cell.border = {
        top: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
        left: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
        bottom: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
        right: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } }
      };
    });

    let currentCategoryRow = 4; // 分类表数据从第4行开始
    for (const item of categoryData) {
      const originalRowData = item.data;
      const originalRowNumber = item.originalIndex;
      const dataRow: any[] = [];
      const formulaMap = new Map<number, string>();

      for (let colIndex = 0; colIndex < categoryHeaders.length; colIndex++) {
        const header = categoryHeaders[colIndex];
        const headerIndex = headerIndexMap.get(header);

        if (headerIndex !== undefined) {
          // 核心逻辑：从 originalSheet 直接获取单元格对象
          if (this.originalSheet && originalRowNumber > 0) {
            // 找到该表头在原表中的物理列号
            let originalColNumber = -1;
            this.originalSheet.getRow(1).eachCell({ includeEmpty: true }, (cell, col) => {
              if (String(cell.value || '').trim() === header) {
                originalColNumber = col;
              }
            });

            if (originalColNumber !== -1) {
              const originalCell = this.originalSheet.getRow(originalRowNumber).getCell(originalColNumber);

              // 检查是否是公式（处理对象形式或直接属性）
              let formulaText = '';
              if (originalCell.formula) {
                formulaText = originalCell.formula;
              } else if (typeof originalCell.value === 'object' && originalCell.value !== null && 'formula' in originalCell.value) {
                formulaText = (originalCell.value as any).formula;
              }

              if (formulaText) {
                // 解析并转换公式
                // 传入分类表数据行数，确保不会引用表尾行
                const convertedFormula = this.convertFormula(
                  formulaText,
                  originalRowNumber,
                  currentCategoryRow,
                  headerIndex,
                  colIndex + 1,
                  categoryHeaders,
                  headerIndexMap,
                  categoryData.length // 分类表数据行数
                );

                if (convertedFormula) {
                  formulaMap.set(colIndex + 1, convertedFormula);
                  dataRow.push(null); // 占位，稍后会被公式替换
                  continue; // 跳过后续的普通值处理
                } else {
                  // 公式转换失败，记录日志以便调试
                  dataRow.push('F-Null'); // 公式依赖缺失
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
                dataRow.push(finalValue);
                continue; // 跳过后续的普通值处理
              }
            } else {
              // 如果找不到原始列号，尝试从 originalSheet 获取原始值（作为后备）
              // 先尝试从 originalSheet 获取
              if (this.originalSheet && originalRowNumber > 0) {
                // 尝试通过 headerIndex 找到对应的列
                let foundCell: ExcelJS.Cell | null = null;
                this.originalSheet.getRow(1).eachCell({ includeEmpty: true }, (cell, col) => {
                  if (String(cell.value || '').trim() === header) {
                    foundCell = this.originalSheet!.getRow(originalRowNumber).getCell(col);
                  }
                });
                if (foundCell !== null) {
                  const cellValue = (foundCell as ExcelJS.Cell).value;
                  if (typeof cellValue === 'number') {
                    dataRow.push(cellValue);
                  } else if (cellValue !== null && cellValue !== undefined) {
                    dataRow.push(cellValue);
                  } else {
                    dataRow.push(originalRowData[headerIndex] || '');
                  }
                } else {
                  // 如果还是找不到，使用 rawData 中的值（作为最后的后备）
                  const rawValue = originalRowData[headerIndex];
                  // 尝试将字符串数字转换为数字
                  if (typeof rawValue === 'string' && rawValue.trim() !== '' && !isNaN(Number(rawValue)) && !isNaN(parseFloat(rawValue))) {
                    const numValue = parseFloat(rawValue);
                    // 检查是否是整数
                    if (Number.isInteger(numValue)) {
                      dataRow.push(Math.floor(numValue));
                    } else {
                      dataRow.push(numValue);
                    }
                  } else {
                    dataRow.push(rawValue || '');
                  }
                }
              } else {
                // 如果找不到原始列号，使用 rawData 中的值（作为后备）
                const rawValue = originalRowData[headerIndex];
                // 尝试将字符串数字转换为数字
                if (typeof rawValue === 'string' && rawValue.trim() !== '' && !isNaN(Number(rawValue)) && !isNaN(parseFloat(rawValue))) {
                  const numValue = parseFloat(rawValue);
                  // 检查是否是整数
                  if (Number.isInteger(numValue)) {
                    dataRow.push(Math.floor(numValue));
                  } else {
                    dataRow.push(numValue);
                  }
                } else {
                  dataRow.push(rawValue || '');
                }
              }
            }
          } else {
            // 如果找不到原始列号，使用 rawData 中的值（作为后备）
            const rawValue = originalRowData[headerIndex];
            // 尝试将字符串数字转换为数字
            if (typeof rawValue === 'string' && rawValue.trim() !== '' && !isNaN(Number(rawValue)) && !isNaN(parseFloat(rawValue))) {
              const numValue = parseFloat(rawValue);
              // 检查是否是整数
              if (Number.isInteger(numValue)) {
                dataRow.push(Math.floor(numValue));
              } else {
                dataRow.push(numValue);
              }
            } else {
              dataRow.push(rawValue || '');
            }
          }
        } else {
          dataRow.push('F-Null');
        }
      }

      const sheetRow = sheet.addRow(dataRow);
      sheetRow.height = 20 * 0.75; // 20像素转换为磅
      // 遍历所有单元格（包括空值），确保公式能被设置
      for (let colNumber = 1; colNumber <= categoryHeaders.length; colNumber++) {
        const cell = sheetRow.getCell(colNumber);
        // 如果这个单元格有公式，设置公式
        if (formulaMap.has(colNumber)) {
          const formula = formulaMap.get(colNumber)!;
          cell.value = { formula: formula };
        }

        cell.font = {
          name: style.fontFamily,
          size: style.dataFontSize,
          bold: style.dataFontBold,
          color: { argb: this.hexToArgb(style.dataFontColor) }
        };
        cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };
        cell.border = {
          top: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          left: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          bottom: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          right: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } }
        };
      }
      currentCategoryRow++;
    }

    // 添加表尾功能行
    const categoryFooterFunctions = this.categoryFooterFunctions();
    if (categoryFooterFunctions.length > 0) {
      const footerRow: any[] = [];

      for (const header of categoryHeaders) {
        const footer = categoryFooterFunctions.find(f => f.header === header);
        if (footer) {
          const colName = this.getExcelColumnName(categoryHeaders.indexOf(header) + 1);
          const startRow = 4; // 数据从第4行开始（1:返回, 2:标题, 3:表头）
          const endRow = 3 + categoryData.length;

          if (footer.type === '合计') {
            footerRow.push({ formula: `SUM(${colName}${startRow}:${colName}${endRow})` });
          } else if (footer.type === '平均值') {
            footerRow.push({ formula: `AVERAGE(${colName}${startRow}:${colName}${endRow})` });
          }
        } else {
          footerRow.push('');
        }
      }

      const row = sheet.addRow(footerRow);
      row.height = 22 * 0.75; // 22像素转换为磅
      row.eachCell((cell, colNumber) => {
        cell.font = {
          name: style.fontFamily,
          size: style.dataFontSize,
          bold: true,
          color: { argb: this.hexToArgb(style.dataFontColor) }
        };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: this.hexToArgb(style.totalColor) }
        };
        cell.alignment = { horizontal: 'right', vertical: 'middle', wrapText: false };
        cell.border = {
          top: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          left: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          bottom: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } },
          right: { style: style.borderStyle as any, color: { argb: this.hexToArgb(style.borderColor) } }
        };
      });
    }

    // 自动调整列宽
    this.autoFitColumns(sheet, categoryHeaders.length);

    // 设置表头筛选器（从第3行表头行开始，到数据结束行）
    sheet.autoFilter = {
      from: { row: 3, column: 1 },
      to: { row: 3 + categoryData.length, column: categoryHeaders.length }
    };
  }

  // 自动调整列宽，确保内容不折叠、不超出
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

  // 计算文本宽度（中文字符按2个字符宽度计算）
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

  // 将十六进制颜色转换为ARGB格式
  hexToArgb(hex: string): string {
    hex = hex.replace('#', '');
    if (hex.length === 3) {
      hex = hex.split('').map(c => c + c).join('');
    }
    return 'FF' + hex.toUpperCase();
  }


  // 获取Excel列名（1 -> A, 2 -> B, 27 -> AA）
  getExcelColumnName(columnNumber: number): string {
    let result = '';
    while (columnNumber > 0) {
      columnNumber--;
      result = String.fromCharCode(65 + (columnNumber % 26)) + result;
      columnNumber = Math.floor(columnNumber / 26);
    }
    return result;
  }

  // 获取单元格值
  getCellValue(cell: ExcelJS.Cell): any {
    if (cell.value === null || cell.value === undefined) {
      return cell.text || '';
    }
    if (typeof cell.value === 'object' && 'text' in cell.value) {
      return (cell.value as any).text || '';
    }
    return cell.value;
  }

  // 转换公式：将原始数据中的公式引用转换为分类表中的引用
  // 重要：只转换当前工作表的单元格引用，不转换其他工作表的引用（如 '汇总表'!A1）
  convertFormula(
    formula: string,
    originalRow: number,
    categoryRow: number,
    originalCol: number,
    categoryCol: number,
    categoryHeaders: string[],
    headerIndexMap: Map<string, number>,
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
          const convertedParts = parts.map(part => this.convertSingleCellRef(part, originalRow, categoryRow, categoryHeaders, headerIndexMap, categoryDataRowCount));

          if (convertedParts.some(p => p === null)) return null;

          convertedFormula = convertedFormula.substring(0, matchIndex) +
                             convertedParts.join(':') +
                             convertedFormula.substring(matchIndex + fullRef.length);
        } else {
          // 处理单个单元格引用
          const converted = this.convertSingleCellRef(fullRef, originalRow, categoryRow, categoryHeaders, headerIndexMap, categoryDataRowCount);
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

  private convertSingleCellRef(ref: string, originalRow: number, categoryRow: number, categoryHeaders: string[], headerIndexMap: Map<string, number>, categoryDataRowCount: number = 0): string | null {
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
      if (this.originalSheet && originalColNum > 0) {
        const headerCell = this.originalSheet.getRow(1).getCell(originalColNum);
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
}

