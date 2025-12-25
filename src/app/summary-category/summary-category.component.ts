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
import { TableStylePreviewComponent } from '../shared/components/table-style-preview/table-style-preview.component';
import { ConfirmDialogComponent } from '../shared/components/confirm-dialog/confirm-dialog.component';
import { FileUploadComponent } from '../shared/components/file-upload/file-upload.component';
import { PrivacyNoticeComponent } from '../shared/components/privacy-notice/privacy-notice.component';
import { ExcelUtilsService } from '../shared/services/excel-utils.service';
import { TableStyleStorageService } from '../shared/services/table-style-storage.service';
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
    DragDropModule,
    TableStylePreviewComponent,
    ConfirmDialogComponent,
    FileUploadComponent,
    PrivacyNoticeComponent
  ],
  templateUrl: './summary-category.component.html',
  styleUrl: './summary-category.component.scss'
})
export class SummaryCategoryComponent implements AfterViewInit {
  @ViewChild('connectionArea', { static: false }) connectionAreaRef?: ElementRef<HTMLDivElement>;
  @ViewChild('stepper', { static: false }) stepper?: MatStepper;
  @ViewChild('summaryHeadersContainer', { static: false }) summaryHeadersContainerRef?: ElementRef<HTMLDivElement>;
  @ViewChild('categoryHeadersContainer', { static: false }) categoryHeadersContainerRef?: ElementRef<HTMLDivElement>;
  private viewInitialized = signal<boolean>(false);
  tableStyle = signal<TableStyleConfig>({ ...DEFAULT_TABLE_STYLE });
  previewExpanded = signal<boolean>(false); // 预览区域展开状态，默认折叠
  showResetButton = signal<boolean>(false); // 是否显示重置按钮

  // Step相关数据
  selectedFile: File | null = null;
  headers = signal<string[]>([]);
  rawData: any[][] = []; // 存储原始Excel数据（包括表头和数据行）
  originalWorkbook: ExcelJS.Workbook | null = null; // 保存原始Excel工作簿，用于读取公式和超链接
  originalSheet: ExcelJS.Worksheet | null = null; // 保存原始工作表
  categoryHeader = signal<string>(''); // Step2: 分类依据表头（单选，不能新增）
  outputFormat = signal<'single' | 'multiple'>('multiple'); // Step2: 输出格式（单表输出/拆分为多个表）
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

  constructor(
    private excelUtils: ExcelUtilsService,
    private tableStyleStorage: TableStyleStorageService
  ) {
    // 从 localStorage 加载表格风格配置
    this.loadTableStyleFromStorage();

    // 监听映射变化，更新连接线位置
    effect(() => {
      const mappings = this.headerMappings();
      if (this.viewInitialized()) {
        setTimeout(() => this.updateConnectionLines(), 100);
      }
    });
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

  async onFileSelected(file: File) {
    this.selectedFile = file;
    this.isUploading.set(true);

    try {
      await this.readExcelFile(this.selectedFile);
    } catch (error: any) {
      console.error('读取Excel文件失败:', error);
      const errorMessage = error?.message || '未知错误';
      alert(`读取Excel文件失败：${errorMessage}\n\n请确保：\n1. 文件格式为 .xlsx\n2. 文件未损坏\n3. 文件包含至少一个工作表`);
    } finally {
      this.isUploading.set(false);
    }
  }

  // 步骤变化处理
  onStepChange(event: any) {
    // 预览会自动更新，因为使用了响应式数据
    // 如果进入step4且是单表输出，自动设置分类表表头
    if (event.selectedIndex === 3 && this.outputFormat() === 'single') {
      const categoryHeader = this.categoryHeader();
      const allHeaders = this.headers();
      const remainingHeaders = categoryHeader
        ? allHeaders.filter(h => h !== categoryHeader)
        : allHeaders;
      if (this.categoryTableHeaders().length === 0) {
        this.categoryTableHeaders.set(remainingHeaders);
      }
    }
  }

  async readExcelFile(file: File) {
    try {
      const result = await this.excelUtils.readExcelFile(file);
      this.headers.set(result.headers);
      this.rawData = result.rawData;
      this.originalWorkbook = result.originalWorkbook;
      this.originalSheet = result.originalSheet;
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

  getCategoryFooterHeaders(): string[] {
    // Step4: 表尾功能可用的表头
    // 单表输出：使用step3选择的汇总表表头
    // 多表输出：使用step4选择的分类表表头
    if (this.outputFormat() === 'single') {
      return this.summaryHeaders();
    } else {
      // 多表输出时，如果有选择的分类表表头，使用它们；否则使用所有表头（除分类依据外）
      const categoryTableHeaders = this.categoryTableHeaders();
      if (categoryTableHeaders.length > 0) {
        return categoryTableHeaders;
      }
      return this.getRemainingHeadersForCategory();
    }
  }

  selectCategoryHeader(header: string) {
    // Step2: 选择分类依据表头（单选）
    // 如果点击的是已选中的表头，则取消选择；否则选择该表头
    if (this.categoryHeader() === header) {
      this.categoryHeader.set('');
    } else {
      this.categoryHeader.set(header);
    }
  }

  onOutputFormatChange(event: any) {
    // Step2: 处理输出格式变化
    this.outputFormat.set(event.value);

    // 如果选择单表输出，自动设置分类表表头为所有表头（除了分类依据表头）
    if (event.value === 'single') {
      const categoryHeader = this.categoryHeader();
      const allHeaders = this.headers();
      const remainingHeaders = categoryHeader
        ? allHeaders.filter(h => h !== categoryHeader)
        : allHeaders;
      this.categoryTableHeaders.set(remainingHeaders);
    }
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
    const headers = [...this.summaryHeaders()];
    moveItemInArray(headers, event.previousIndex, event.currentIndex);
    this.summaryHeaders.set(headers);
  }

  dropCategoryTableHeader(event: CdkDragDrop<string[]>) {
    // Step4: 拖拽排序分类表表头
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

    // 未选择文件时，返回示例预览数据
    return this.getDefaultPreviewData();
  }

  // 获取默认预览数据（未选择文件时使用）
  getDefaultPreviewData(): any[][] {
    return this.excelUtils.getDefaultPreviewData();
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
      alert('请先选择Excel文件');
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
    const defaultFileName = this.excelUtils.generateDefaultFileName(this.selectedFile.name, '汇总分类表');
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
    const fileName = this.outputFileName().trim() || '汇总分类表';
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
    const categoryValueTypes = new Map<string, 'date' | 'string'>(); // 存储每个分类值的数据类型

    for (const item of deduplicatedRawData) {
      const categoryValue = item.data[categoryHeaderIndex] || '';

      // 检测分类值的数据类型（检查原始 Excel 单元格）
      if (!categoryValueTypes.has(categoryValue)) {
        const isDate = this.isCategoryValueDate(categoryValue, item.originalIndex, categoryHeaderIndex);
        categoryValueTypes.set(categoryValue, isDate ? 'date' : 'string');
      }

      if (!groupedData.has(categoryValue)) {
        groupedData.set(categoryValue, []);
      }
      groupedData.get(categoryValue)!.push(item);
    }

    // 根据输出格式选择不同的生成逻辑
    if (this.outputFormat() === 'single') {
      // 单表输出：创建一个工作表，按分类分组显示，每个分类有小计，最后有总合计
      await this.createSingleSheetOutput(workbook, groupedData, style, categoryValueTypes);
    } else {
      // 多表输出：创建汇总表和多个分类表
      // 创建汇总表工作表
      const summarySheet = workbook.addWorksheet('汇总表');

      // 为每个分类创建分类表工作表（先创建，以便汇总表可以引用）
      // 汇总表的分类需要去重：使用 Set 确保每个分类值只出现一次
      // 根据数据类型进行排序：日期按日期排序，其他按字母排序
      const uniqueCategoryValues = Array.from(new Set(Array.from(groupedData.keys())));
      const categoryValues = this.sortCategoryValues(uniqueCategoryValues, categoryValueTypes);
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
      await this.createSummarySheet(summarySheet, groupedData, style, categorySheetMap, workbook, categoryValueTypes);

      // 创建分类表（传入汇总表以便添加返回链接）
      for (const categoryValue of categoryValues) {
        const categoryData = groupedData.get(categoryValue)!;
        const categorySheet = categorySheetMap.get(categoryValue)!;
        await this.createCategorySheet(categorySheet, categoryValue, categoryData, style, summarySheet);
      }
    }

    // 生成Excel文件并返回blob（不直接下载）
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return blob;
  }

  // 创建单表输出（按分类分组，每个分类有小计，最后有总合计）
  async createSingleSheetOutput(
    workbook: ExcelJS.Workbook,
    groupedData: Map<string, Array<{ data: any[], originalIndex: number }>>,
    style: TableStyleConfig,
    categoryValueTypes: Map<string, 'date' | 'string'>
  ) {
    const sheet = workbook.addWorksheet('汇总分类表');
    // 单表输出时，使用所有表头（除了分类依据表头）
    const categoryHeaders = this.categoryTableHeaders().length > 0
      ? this.categoryTableHeaders()
      : this.getRemainingHeadersForCategory();
    const categoryHeader = this.categoryHeader()!;
    const headerIndexMap = new Map<string, number>();
    this.headers().forEach((header, index) => {
      headerIndexMap.set(header, index);
    });

    // 添加标题行
    const titleRow = sheet.addRow([`汇总分类表（按${categoryHeader}分类）`]);
    titleRow.height = 25 * 0.75;
    this.excelUtils.applyCellStyle(titleRow.getCell(1), style, 'title');
    sheet.mergeCells(1, 1, 1, categoryHeaders.length);

    // 添加表头行
    const headerRow = sheet.addRow(categoryHeaders);
    headerRow.height = 22 * 0.75;
    headerRow.eachCell((cell, colNumber) => {
      this.excelUtils.applyCellStyle(cell, style, 'header');
    });

    // 根据数据类型进行排序
    const uniqueCategoryValues = Array.from(new Set(Array.from(groupedData.keys())));
    const categoryValues = this.sortCategoryValues(uniqueCategoryValues, categoryValueTypes);

    let currentRow = 3; // 第1行是标题，第2行是表头，数据从第3行开始
    const categoryFooterFunctions = this.categoryFooterFunctions();
    const hasCategoryFooter = categoryFooterFunctions.length > 0;
    const categorySubtotalRows: number[] = []; // 存储每个分类小计行的行号

    // 遍历每个分类
    for (let categoryIndex = 0; categoryIndex < categoryValues.length; categoryIndex++) {
      const categoryValue = categoryValues[categoryIndex];
      const categoryData = groupedData.get(categoryValue)!;

      // 添加分类标题行
      const categoryTitleRow = sheet.addRow([categoryValue]);
      categoryTitleRow.height = 22 * 0.75;
      const titleCell = categoryTitleRow.getCell(1);
      this.excelUtils.applyCellStyle(titleCell, style, 'header');
      // 合并单元格
      sheet.mergeCells(currentRow, 1, currentRow, categoryHeaders.length);
      currentRow++;

      const categoryStartRow = currentRow; // 该分类的第一行数据

      // 添加分类数据行
      for (const item of categoryData) {
        const originalRowData = item.data;
        const originalRowNumber = item.originalIndex;
        const dataRow: any[] = [];
        const formulaMap = new Map<number, string>();

        for (let colIndex = 0; colIndex < categoryHeaders.length; colIndex++) {
          const header = categoryHeaders[colIndex];
          const headerIndex = headerIndexMap.get(header);

          if (headerIndex !== undefined) {
            // 从原始Excel获取单元格值
            if (this.originalSheet && originalRowNumber > 0) {
              let originalColNumber = -1;
              this.originalSheet.getRow(1).eachCell({ includeEmpty: true }, (cell, col) => {
                if (String(cell.value || '').trim() === header) {
                  originalColNumber = col;
                }
              });

              if (originalColNumber !== -1) {
                const originalCell = this.originalSheet.getRow(originalRowNumber).getCell(originalColNumber);

                // 提取并转换公式
                const convertedFormula = this.excelUtils.extractAndConvertFormula(
                  originalCell,
                  originalRowNumber,
                  currentRow,
                  headerIndex!,
                  colIndex + 1,
                  categoryHeaders,
                  headerIndexMap,
                  this.originalSheet,
                  categoryData.length,
                  categoryStartRow
                );

                if (convertedFormula) {
                  formulaMap.set(colIndex + 1, convertedFormula);
                  dataRow.push(null);
                  continue;
                } else if (convertedFormula === null && (originalCell.formula || (typeof originalCell.value === 'object' && originalCell.value !== null && 'formula' in originalCell.value))) {
                  // 如果是公式但转换失败，填充 F-Null
                  dataRow.push('F-Null');
                  continue;
                } else {
                  // 不是公式，获取原始值
                  const cellValue = originalCell.value;
                  let finalValue: any = null;
                  if (cellValue !== null && cellValue !== undefined) {
                    if (typeof cellValue === 'number') {
                      finalValue = cellValue;
                    } else if (typeof cellValue === 'object' && 'text' in cellValue) {
                      finalValue = (cellValue as any).text || '';
                    } else if (cellValue instanceof Date) {
                      finalValue = cellValue;
                    } else {
                      finalValue = cellValue;
                    }
                  } else if (originalCell.result !== null && originalCell.result !== undefined) {
                    if (typeof originalCell.result === 'number') {
                      finalValue = originalCell.result;
                    } else {
                      finalValue = originalCell.result;
                    }
                  } else {
                    finalValue = '';
                  }
                  dataRow.push(finalValue);
                  continue;
                }
              } else {
                // 如果找不到原始列号，使用rawData中的值
                const rawValue = originalRowData[headerIndex];
                if (typeof rawValue === 'string' && rawValue.trim() !== '' && !isNaN(Number(rawValue)) && !isNaN(parseFloat(rawValue))) {
                  const numValue = parseFloat(rawValue);
                  dataRow.push(Number.isInteger(numValue) ? Math.floor(numValue) : numValue);
                } else {
                  dataRow.push(rawValue || '');
                }
              }
            } else {
              // 如果找不到原始列号，使用rawData中的值
              const rawValue = originalRowData[headerIndex];
              if (typeof rawValue === 'string' && rawValue.trim() !== '' && !isNaN(Number(rawValue)) && !isNaN(parseFloat(rawValue))) {
                const numValue = parseFloat(rawValue);
                dataRow.push(Number.isInteger(numValue) ? Math.floor(numValue) : numValue);
              } else {
                dataRow.push(rawValue || '');
              }
            }
          } else {
            dataRow.push('F-Null');
          }
        }

        const sheetRow = sheet.addRow(dataRow);
        sheetRow.height = 20 * 0.75;
        for (let colNumber = 1; colNumber <= categoryHeaders.length; colNumber++) {
          const cell = sheetRow.getCell(colNumber);
          if (formulaMap.has(colNumber)) {
            const formula = formulaMap.get(colNumber)!;
            cell.value = { formula: formula };
          }
          this.excelUtils.applyCellStyle(cell, style, 'data');
        }
        currentRow++;
      }

      // 添加分类小计行（使用step4的表尾功能）
      if (hasCategoryFooter) {
        const footerRow: any[] = [];
        const categoryEndRow = currentRow - 1; // 该分类的最后一行数据

        for (const header of categoryHeaders) {
          const footer = categoryFooterFunctions.find(f => f.header === header);
          if (footer) {
            const colName = this.excelUtils.getExcelColumnName(categoryHeaders.indexOf(header) + 1);
            if (footer.type === '合计') {
              footerRow.push({ formula: `SUM(${colName}${categoryStartRow}:${colName}${categoryEndRow})` });
            } else if (footer.type === '平均值') {
              footerRow.push({ formula: `AVERAGE(${colName}${categoryStartRow}:${colName}${categoryEndRow})` });
            }
          } else {
            footerRow.push('');
          }
        }

        // 在分类名称列显示"合计"
        const categoryHeaderColIndex = categoryHeaders.indexOf(categoryHeader);
        if (categoryHeaderColIndex >= 0) {
          footerRow[categoryHeaderColIndex] = '合计';
        }

        const row = sheet.addRow(footerRow);
        row.height = 22 * 0.75;
        row.eachCell((cell, colNumber) => {
          this.excelUtils.applyCellStyle(cell, style, 'total');
        });
        categorySubtotalRows.push(currentRow); // 记录小计行号
        currentRow++;
      }

      // 如果不是最后一个分类，添加间隔行（合并单元格）
      if (categoryIndex < categoryValues.length - 1) {
        const spacerRow = sheet.addRow(['']);
        spacerRow.height = 10 * 0.75;
        const spacerCell = spacerRow.getCell(1);
        spacerCell.style = {
          fill: {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' } // 白色背景
          }
        };
        // 合并单元格
        sheet.mergeCells(currentRow, 1, currentRow, categoryHeaders.length);
        currentRow++;
      }
    }

    // 添加总合计行（使用step3的表尾功能）
    const footerFunctions = this.footerFunctions();
    if (footerFunctions.length > 0) {
      const totalFooterRow: any[] = [];

      for (const header of categoryHeaders) {
        const footer = footerFunctions.find(f => f.header === header);
        if (footer) {
          const colName = this.excelUtils.getExcelColumnName(categoryHeaders.indexOf(header) + 1);
          if (hasCategoryFooter && categorySubtotalRows.length > 0) {
            // 如果有分类小计行，总合计应该是对所有小计行的合计
            const subtotalRefs = categorySubtotalRows.map(row => `${colName}${row}`);
            if (footer.type === '合计') {
              totalFooterRow.push({ formula: `SUM(${subtotalRefs.join(',')})` });
            } else if (footer.type === '平均值') {
              totalFooterRow.push({ formula: `AVERAGE(${subtotalRefs.join(',')})` });
            }
          } else {
            // 如果没有分类小计行，总合计是对所有数据行的合计
            const dataStartRow = 3; // 数据从第3行开始
            const dataEndRow = currentRow - 1; // 数据结束行
            if (footer.type === '合计') {
              totalFooterRow.push({ formula: `SUM(${colName}${dataStartRow}:${colName}${dataEndRow})` });
            } else if (footer.type === '平均值') {
              totalFooterRow.push({ formula: `AVERAGE(${colName}${dataStartRow}:${colName}${dataEndRow})` });
            }
          }
        } else {
          totalFooterRow.push('');
        }
      }

      // 在分类名称列显示"总合计"
      const categoryHeaderColIndex = categoryHeaders.indexOf(categoryHeader);
      if (categoryHeaderColIndex >= 0) {
        totalFooterRow[categoryHeaderColIndex] = '总合计';
      }

      const row = sheet.addRow(totalFooterRow);
      row.height = 22 * 0.75;
      row.eachCell((cell, colNumber) => {
        this.excelUtils.applyCellStyle(cell, style, 'total');
      });
    }

    // 自动调整列宽
    this.excelUtils.autoFitColumns(sheet, categoryHeaders.length);

    // 设置表头筛选器
    sheet.autoFilter = {
      from: { row: 2, column: 1 },
      to: { row: currentRow, column: categoryHeaders.length }
    };
  }

  // 创建汇总表
  async createSummarySheet(
    sheet: ExcelJS.Worksheet,
    groupedData: Map<string, Array<{ data: any[], originalIndex: number }>>,
    style: TableStyleConfig,
    categorySheetMap: Map<string, ExcelJS.Worksheet>,
    workbook: ExcelJS.Workbook,
    categoryValueTypes?: Map<string, 'date' | 'string'>
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
    this.excelUtils.applyCellStyle(titleRow.getCell(1), style, 'title');
    sheet.mergeCells(1, 1, 1, summaryHeaders.length);

    // 添加表头行
    const headerRow = sheet.addRow(summaryHeaders);
    headerRow.height = 22 * 0.75; // 22像素转换为磅
    headerRow.eachCell((cell, colNumber) => {
      this.excelUtils.applyCellStyle(cell, style, 'header');
    });

    // 添加数据行
    // 根据数据类型进行排序：日期按日期排序，其他按字母排序
    let types = categoryValueTypes;
    if (!types) {
      // 如果没有传入类型映射，则检测类型
      types = new Map<string, 'date' | 'string'>();
      const allCategoryValues = Array.from(groupedData.keys());
      const categoryHeaderIndex = headerIndexMap.get(categoryHeader)!;

      for (const categoryValue of allCategoryValues) {
        if (!types.has(categoryValue)) {
          // 从分组数据中查找第一个匹配的值来检测类型
          const firstItem = groupedData.get(categoryValue)?.[0];
          if (firstItem) {
            const isDate = this.isCategoryValueDate(categoryValue, firstItem.originalIndex, categoryHeaderIndex);
            types.set(categoryValue, isDate ? 'date' : 'string');
          } else {
            types.set(categoryValue, 'string');
          }
        }
      }
    }

    const allCategoryValues = Array.from(groupedData.keys());
    const categoryValues = this.sortCategoryValues(allCategoryValues, types);
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
                const targetCol = this.excelUtils.getExcelColumnName(mappedHeaderIndex + 1);
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
                  const targetCol = this.excelUtils.getExcelColumnName(catHeaderIdx + 1);
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
                  const targetCol = this.excelUtils.getExcelColumnName(catHeaderIdx + 1);
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

        if (!cell.font) {
          this.excelUtils.applyCellStyle(cell, style, 'data');
        } else {
          // 如果已经有字体设置（如超链接），只应用边框
          cell.border = {
            top: { style: style.borderStyle as any, color: { argb: this.excelUtils.hexToArgb(style.borderColor) } },
            left: { style: style.borderStyle as any, color: { argb: this.excelUtils.hexToArgb(style.borderColor) } },
            bottom: { style: style.borderStyle as any, color: { argb: this.excelUtils.hexToArgb(style.borderColor) } },
            right: { style: style.borderStyle as any, color: { argb: this.excelUtils.hexToArgb(style.borderColor) } }
          };
        }
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
                const targetCol = this.excelUtils.getExcelColumnName(mappedHeaderIndex + 1);
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
              const targetCol = this.excelUtils.getExcelColumnName(catHeaderIdx + 1);
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
        this.excelUtils.applyCellStyle(cell, style, 'total');
      });
    }

    // 自动调整列宽
    this.excelUtils.autoFitColumns(sheet, summaryHeaders.length);

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
    this.excelUtils.applyCellStyle(titleRow.getCell(1), style, 'title');
    sheet.mergeCells(2, 1, 2, categoryHeaders.length);

    // 添加表头行
    const headerRow = sheet.addRow(categoryHeaders);
    headerRow.height = 22 * 0.75; // 22像素转换为磅
    headerRow.eachCell((cell, colNumber) => {
      this.excelUtils.applyCellStyle(cell, style, 'header');
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

              // 提取并转换公式
              const convertedFormula = this.excelUtils.extractAndConvertFormula(
                originalCell,
                originalRowNumber,
                currentCategoryRow,
                headerIndex!,
                colIndex + 1,
                categoryHeaders,
                headerIndexMap,
                this.originalSheet,
                categoryData.length // 分类表数据行数
              );

              if (convertedFormula) {
                formulaMap.set(colIndex + 1, convertedFormula);
                dataRow.push(null); // 占位，稍后会被公式替换
                continue; // 跳过后续的普通值处理
              } else if (convertedFormula === null && (originalCell.formula || (typeof originalCell.value === 'object' && originalCell.value !== null && 'formula' in originalCell.value))) {
                // 如果是公式但转换失败，填充 F-Null
                dataRow.push('F-Null'); // 公式依赖缺失
                continue;
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

        this.excelUtils.applyCellStyle(cell, style, 'data');
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
          const colName = this.excelUtils.getExcelColumnName(categoryHeaders.indexOf(header) + 1);
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
        this.excelUtils.applyCellStyle(cell, style, 'total');
      });
    }

    // 自动调整列宽
    this.excelUtils.autoFitColumns(sheet, categoryHeaders.length);

    // 设置表头筛选器（从第3行表头行开始，到数据结束行）
    sheet.autoFilter = {
      from: { row: 3, column: 1 },
      to: { row: 3 + categoryData.length, column: categoryHeaders.length }
    };
  }


  /**
   * 检测分类值是否为日期类型
   * 通过检查原始 Excel 单元格的值类型来判断
   */
  private isCategoryValueDate(categoryValue: string, originalRowIndex: number, categoryHeaderIndex: number): boolean {
    // 如果值为空，返回 false
    if (!categoryValue || categoryValue.trim() === '') {
      return false;
    }

    // 方法1: 检查原始 Excel 单元格的值类型
    if (this.originalSheet && originalRowIndex > 0) {
      try {
        // 找到该表头在原表中的物理列号
        let originalColNumber = -1;
        const categoryHeader = this.categoryHeader();
        this.originalSheet.getRow(1).eachCell({ includeEmpty: true }, (cell, col) => {
          const headerValue = String(cell.value || '').trim();
          if (headerValue === categoryHeader) {
            originalColNumber = col;
          }
        });

        if (originalColNumber !== -1) {
          const cell = this.originalSheet.getRow(originalRowIndex).getCell(originalColNumber);
          const cellValue = cell.value;

          // 检查是否为 Date 对象
          if (cellValue instanceof Date) {
            return true;
          }

          // 检查单元格的样式格式是否包含日期格式
          if (cell.numFmt) {
            // Excel 日期格式通常包含 d, m, y 等字符
            const dateFormats = ['d', 'm', 'y', 'h', 's', 'D', 'M', 'Y', 'H', 'S'];
            const hasDateFormat = dateFormats.some(char => cell.numFmt.includes(char));
            if (hasDateFormat && typeof cellValue === 'number') {
              // Excel 日期是数字，但格式是日期格式
              return true;
            }
          }
        }
      } catch (e) {
        // 如果出错，继续使用字符串检测方法
        console.debug('检测日期类型时出错:', e);
      }
    }

    // 方法2: 通过字符串格式检测日期
    // 常见的日期格式：YYYY-MM-DD, YYYY/MM/DD, MM/DD/YYYY, DD/MM/YYYY 等
    const datePatterns = [
      /^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/, // YYYY-MM-DD 或 YYYY/MM/DD
      /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/, // MM/DD/YYYY 或 DD/MM/YYYY
      /^\d{4}\.\d{1,2}\.\d{1,2}$/, // YYYY.MM.DD
      /^\d{1,2}\.\d{1,2}\.\d{4}$/, // DD.MM.YYYY
    ];

    // 检查是否匹配日期格式
    const matchesDatePattern = datePatterns.some(pattern => pattern.test(categoryValue.trim()));

    if (matchesDatePattern) {
      // 尝试解析为日期，如果能解析且有效，则认为是日期
      const parsedDate = new Date(categoryValue);
      if (!isNaN(parsedDate.getTime())) {
        return true;
      }
    }

    return false;
  }

  /**
   * 对分类值进行排序
   * 如果是日期类型，按日期排序；否则按字母排序
   */
  private sortCategoryValues(
    values: string[],
    valueTypes: Map<string, 'date' | 'string'>
  ): string[] {
    return [...values].sort((a, b) => {
      const typeA = valueTypes.get(a) || 'string';
      const typeB = valueTypes.get(b) || 'string';

      // 如果都是日期类型，按日期排序
      if (typeA === 'date' && typeB === 'date') {
        const dateA = new Date(a);
        const dateB = new Date(b);

        // 如果日期无效，按字符串排序
        if (isNaN(dateA.getTime()) || isNaN(dateB.getTime())) {
          return a.localeCompare(b, 'zh-CN');
        }

        return dateA.getTime() - dateB.getTime();
      }

      // 如果一个是日期一个是字符串，日期排在前面
      if (typeA === 'date' && typeB === 'string') {
        return -1;
      }
      if (typeA === 'string' && typeB === 'date') {
        return 1;
      }

      // 都是字符串，按字母排序（支持中文）
      return a.localeCompare(b, 'zh-CN');
    });
  }
}

