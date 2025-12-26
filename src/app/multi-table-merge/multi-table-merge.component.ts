import { Component, signal, ViewChild, AfterViewInit, ElementRef, effect } from '@angular/core';
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
import { MatRadioModule } from '@angular/material/radio';
import { MatTooltipModule } from '@angular/material/tooltip';
import { DragDropModule, CdkDragDrop, moveItemInArray } from '@angular/cdk/drag-drop';
import { TableStyleConfig, DEFAULT_TABLE_STYLE } from '../shared/models/table-style.model';
import { TableStylePreviewComponent } from '../shared/components/table-style-preview/table-style-preview.component';
import { TableStyleStorageService } from '../shared/services/table-style-storage.service';
import { FileUploadComponent } from '../shared/components/file-upload/file-upload.component';
import { ConfirmDialogComponent } from '../shared/components/confirm-dialog/confirm-dialog.component';
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
    MatRadioModule,
    MatTooltipModule,
    DragDropModule,
    TableStylePreviewComponent,
    FileUploadComponent,
    ConfirmDialogComponent
  ],
  templateUrl: './multi-table-merge.component.html',
  styleUrl: './multi-table-merge.component.scss'
})
export class MultiTableMergeComponent implements AfterViewInit {
  @ViewChild('stepper', { static: false }) stepper?: MatStepper;
  @ViewChild('connectionArea', { static: false }) connectionAreaRef?: ElementRef<HTMLDivElement>;
  private viewInitialized = signal<boolean>(false);

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

  // Step4: 关联更新汇总表
  updateSummaryTable = signal<boolean>(false); // 是否更新汇总表
  selectedSummarySheet = signal<string>(''); // 选中的汇总表sheet
  summaryHeaderRow = signal<number>(1); // 汇总表表头所在行数
  summaryHeaders = signal<string[]>([]); // 汇总表表头列表（可新增）
  newSummaryHeaderName = signal<string>(''); // 新增汇总表表头名称
  // Step4: 表头关联映射（汇总表表头 -> 合并表表头）
  headerMappings = signal<Map<string, string>>(new Map());
  selectedSummaryHeaderForMapping = signal<string>(''); // 当前选中的汇总表表头（用于关联）

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

    // 监听映射变化，更新连接线位置
    effect(() => {
      const mappings = this.headerMappings();
      if (this.viewInitialized()) {
        setTimeout(() => this.updateConnectionLines(), 100);
      }
    });
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
      // 处理Excel文件
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
    // 只加载基础文件表头，不同步更新数据源表头行数
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
      // 读取第一个选中的基础文件sheet的表头
      const firstSheet = this.baseWorkbook()!.getWorksheet(baseSheets[0]);
      if (firstSheet) {
        const headers = this.excelUtils.readSheetHeaders(firstSheet, this.baseHeaderRow(), false);
        this.baseHeaders.set(headers);
      } else {
        this.baseHeaders.set([]);
      }
    } catch (error) {
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
      const headers = this.excelUtils.readMultipleSheetHeaders(
        this.dataSourceWorkbook()!,
        dataSourceSheets,
        this.dataSourceHeaderRow(),
        false
      );
      this.dataSourceHeaders.set(headers);
    } catch (error) {
      this.dataSourceHeaders.set([]);
    }
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
    // 如果选择了基础sheet，且数据源还未选择，设置数据源表头行数默认值为基础文件的值（仅首次设置默认值）
    if (baseSheets.length > 0 && this.selectedDataSourceSheets().length === 0) {
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

  // 检查step2是否应该显示"下一步"按钮
  shouldShowNextInStep2(): boolean {
    return this.shouldShowHeaderSortSection();
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

  // Step4: 汇总表sheet选择变化
  onSummarySheetChange() {
    this.loadSummaryHeaders();
  }

  // Step4: 加载汇总表表头
  private async loadSummaryHeaders() {
    const summarySheetName = this.selectedSummarySheet();
    if (!summarySheetName || !this.baseWorkbook()) {
      this.summaryHeaders.set([]);
      return;
    }

    try {
      const sheet = this.baseWorkbook()!.getWorksheet(summarySheetName);
      if (sheet) {
        const headers = this.excelUtils.readSheetHeaders(sheet, this.summaryHeaderRow(), false);
        this.summaryHeaders.set(headers);
      } else {
        this.summaryHeaders.set([]);
      }
    } catch (error) {
      this.summaryHeaders.set([]);
    }
  }

  // Step4: 新增汇总表表头
  addNewSummaryHeader() {
    const name = this.newSummaryHeaderName().trim();
    if (name && !this.summaryHeaders().includes(name)) {
      this.summaryHeaders.set([...this.summaryHeaders(), name]);
      this.newSummaryHeaderName.set('');
    }
  }

  // Step4: 删除汇总表表头
  removeSummaryHeader(header: string) {
    this.summaryHeaders.set(this.summaryHeaders().filter(h => h !== header));
    // 同时移除该表头关联的所有映射
    const mappings = new Map(this.headerMappings());
    if (mappings.has(header)) {
      mappings.delete(header);
      this.headerMappings.set(mappings);
      setTimeout(() => this.updateConnectionLines(), 50);
    }
  }

  // Step4: 拖拽排序汇总表表头
  dropSummaryHeader(event: CdkDragDrop<string[]>) {
    const headers = [...this.summaryHeaders()];
    moveItemInArray(headers, event.previousIndex, event.currentIndex);
    this.summaryHeaders.set(headers);
  }

  // Step4: 选择汇总表表头进行关联
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

  // Step4: 获取合并表的所有选项（只包括表尾功能：合计或平均值）
  getMergedTableOptions(): Array<{ id: string, label: string, type: 'header' | 'footer', icon?: string }> {
    const options: Array<{ id: string, label: string, type: 'header' | 'footer', icon?: string }> = [];

    // 只添加表尾功能（合计或平均值），不添加表头选项
    for (const footer of this.footerFunctions()) {
      options.push({
        id: `footer:${footer.id}`,
        label: this.getFooterFunctionLabel(footer),
        type: 'footer',
        icon: footer.type === '合计' ? 'calculate' : 'trending_up'
      });
    }

    return options;
  }

  // Step4: 选择合并表选项（表头或表尾功能）进行关联
  selectMergedTableOptionForMapping(optionId: string) {
    const selectedSummary = this.selectedSummaryHeaderForMapping();
    if (!selectedSummary) {
      return; // 如果没有选中汇总表表头，直接返回
    }

    const mappings = new Map(this.headerMappings());

    // 如果该合并表选项已经被其他汇总表表头关联，先取消之前的关联
    for (const [summaryHeader, mergedOptionId] of mappings.entries()) {
      if (mergedOptionId === optionId && summaryHeader !== selectedSummary) {
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

  // Step4: 检查汇总表表头是否已选中用于关联
  isSummaryHeaderSelectedForMapping(header: string): boolean {
    return this.selectedSummaryHeaderForMapping() === header;
  }

  // Step4: 检查汇总表表头是否已关联
  isSummaryHeaderMapped(header: string): boolean {
    return this.headerMappings().has(header);
  }

  // Step4: 检查合并表选项是否已关联
  isMergedTableOptionMapped(optionId: string): boolean {
    const mappings = this.headerMappings();
    for (const mergedOptionId of mappings.values()) {
      if (mergedOptionId === optionId) {
        return true;
      }
    }
    return false;
  }

  // Step4: 移除关联
  removeMapping(summaryHeader: string) {
    const mappings = new Map(this.headerMappings());
    mappings.delete(summaryHeader);
    this.headerMappings.set(mappings);
    setTimeout(() => this.updateConnectionLines(), 50);
  }

  // Step4: 根据选项ID获取显示标签
  getMergedTableOptionLabel(optionId: string): string {
    if (optionId.startsWith('header:')) {
      return optionId.replace('header:', '');
    } else if (optionId.startsWith('footer:')) {
      const footerId = optionId.replace('footer:', '');
      const footer = this.footerFunctions().find(f => f.id === footerId);
      return footer ? this.getFooterFunctionLabel(footer) : optionId;
    }
    return optionId;
  }

  // Step4: 更新连接线位置
  updateConnectionLines() {
    if (!this.connectionAreaRef) return;

    const connectionArea = this.connectionAreaRef.nativeElement;
    const container = connectionArea.parentElement;
    if (!container) return;

    const summaryColumn = container.querySelector('.summary-headers-column .header-list-mapping');
    const mergedColumn = container.querySelector('.merged-headers-column .header-list-mapping');

    if (!summaryColumn || !mergedColumn) return;

    const mappings = this.headerMappings();
    const lines = connectionArea.querySelectorAll('.connection-line');

    lines.forEach((line) => {
      const lineElement = line as HTMLElement;
      const summaryHeader = lineElement.getAttribute('data-summary');
      const mergedOptionId = lineElement.getAttribute('data-merged');

      if (!summaryHeader || !mergedOptionId) return;

      const summaryItem = summaryColumn.querySelector(`[data-header-id="summary-${summaryHeader}"]`) as HTMLElement;
      const mergedItem = mergedColumn.querySelector(`[data-header-id="merged-${mergedOptionId}"]`) as HTMLElement;

      if (summaryItem && mergedItem) {
        const containerRect = container.getBoundingClientRect();
        const summaryRect = summaryItem.getBoundingClientRect();
        const mergedRect = mergedItem.getBoundingClientRect();

        const startX = summaryRect.right - containerRect.left;
        const startY = summaryRect.top + summaryRect.height / 2 - containerRect.top;
        const endX = mergedRect.left - containerRect.left;
        const endY = mergedRect.top + mergedRect.height / 2 - containerRect.top;

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
    const updateSummaryTable = this.updateSummaryTable();
    const summarySheetName = this.selectedSummarySheet();
    const summaryHeaders = this.summaryHeaders();
    const headerMappings = this.headerMappings();

    // 创建新的工作簿
    const newWorkbook = new ExcelJS.Workbook();
    const unmatchedSheets: string[] = []; // 记录未匹配的sheet
    const mergedSheetMap = new Map<string, ExcelJS.Worksheet>(); // 存储合并表映射，用于汇总表引用
    const mergedSheetDataRowCountMap = new Map<string, number>(); // 存储每个合并表的数据行数

    // 如果更新汇总表，先创建汇总表（作为第一个sheet）
    let summarySheet: ExcelJS.Worksheet | null = null;
    if (updateSummaryTable && summarySheetName && summaryHeaders.length > 0) {
      summarySheet = newWorkbook.addWorksheet('汇总表');
      // 添加标题行和表头行（数据行稍后填充）
      const titleRow = summarySheet.addRow(['汇总表']);
      titleRow.height = 25 * 0.75;
      this.excelUtils.applyCellStyle(titleRow.getCell(1), style, 'title');
      summarySheet.mergeCells(1, 1, 1, summaryHeaders.length);

      const headerRow = summarySheet.addRow(summaryHeaders);
      headerRow.height = 22 * 0.75;
      headerRow.eachCell((cell: ExcelJS.Cell) => {
        this.excelUtils.applyCellStyle(cell, style, 'header');
      });
    }

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
      mergedSheetMap.set(dataSourceSheetName, newSheet); // 存储到映射中，供汇总表引用

      // 读取基础表的表头映射
      const baseHeadersMap = this.excelUtils.readHeadersMap(matchedBaseSheet, baseHeaderRow, true);

      // 读取数据源表的表头映射
      const dataSourceHeadersMap = this.excelUtils.readHeadersMap(dataSourceSheet, dataSourceHeaderRow, true);

      // 创建表头映射：基础表表头 -> 数据源表列号
      const headerToDataSourceColMap = new Map<string, number>();
      dataSourceHeadersMap.forEach((header, colNumber) => {
        headerToDataSourceColMap.set(header, colNumber);
      });

      let currentRow = 1; // 当前行号

      // 如果更新汇总表，添加返回汇总表的链接
      if (updateSummaryTable && summarySheetName) {
        const returnRow = newSheet.addRow(['返回汇总表']);
        returnRow.height = 25 * 0.75;
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
        currentRow++;
      }

      // 添加标题行（使用基础表的标题，如果有）
      const titleRow = newSheet.addRow([`${dataSourceSheetName}`]);
      titleRow.height = 25 * 0.75;
      this.excelUtils.applyCellStyle(titleRow.getCell(1), style, 'title');
      newSheet.mergeCells(currentRow, 1, currentRow, sortedHeaders.length);
      currentRow++;

      // 添加表头行（使用排序后的表头）
      const headerRow = newSheet.addRow(sortedHeaders);
      headerRow.height = 22 * 0.75;
      headerRow.eachCell((cell) => {
        this.excelUtils.applyCellStyle(cell, style, 'header');
      });
      currentRow++;

      // 读取基础表的数据（从表头行之后开始）
      const baseDataRows: Array<{ rowNumber: number, data: Map<string, any> }> = [];
      for (let rowNum = baseHeaderRow + 1; rowNum <= matchedBaseSheet.rowCount; rowNum++) {
        const row = matchedBaseSheet.getRow(rowNum);
        if (!row || row.cellCount === 0) continue;

        const rowData = new Map<string, any>();
        let hasData = false;

        baseHeadersMap.forEach((header, colNumber) => {
          const cell = row.getCell(colNumber);
          const cellValue = this.excelUtils.getCellValueWithFormula(cell, matchedBaseSheet, rowNum, colNumber);
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
          const cellValue = this.excelUtils.getCellValueWithFormula(cell, dataSourceSheet, rowNum, colNumber);
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
      // 数据起始行：如果更新汇总表，第1行是返回链接，第2行是标题，第3行是表头，数据从第4行开始
      // 否则，第1行是标题，第2行是表头，数据从第3行开始
      const dataStartRow = updateSummaryTable ? 4 : 3;
      let dataCurrentRow = dataStartRow; // 数据当前行号

      for (let i = 0; i < maxRows; i++) {
        const dataRow: any[] = [];
        const formulaMap = new Map<number, string>(); // 列索引 -> 公式

        for (let colIndex = 0; colIndex < sortedHeaders.length; colIndex++) {
          const header = sortedHeaders[colIndex];
          let cellValue: any = '';

          // 优先从数据源表获取数据
          if (i < dataSourceDataRows.length && dataSourceDataRows[i].data.has(header)) {
            cellValue = dataSourceDataRows[i].data.get(header);
          }

          // 如果数据源表没有，从基础表获取
          if ((cellValue === null || cellValue === undefined || cellValue === '') &&
              i < baseDataRows.length && baseDataRows[i].data.has(header)) {
            cellValue = baseDataRows[i].data.get(header);
          }

          // 如果cellValue是公式对象，提取公式
          if (cellValue && typeof cellValue === 'object' && 'formula' in cellValue) {
            const formula = (cellValue as any).formula;
            // 判断公式来自哪个表：优先从数据源表获取，如果数据源表没有，则来自基础表
            const isFromDataSource = i < dataSourceDataRows.length &&
                                     dataSourceDataRows[i].data.has(header);
            const isFromBase = !isFromDataSource && i < baseDataRows.length &&
                               baseDataRows[i].data.has(header);
            // 转换公式引用（按照新的列位置转换）
            const originalRow = i < baseDataRows.length ? baseDataRows[i].rowNumber : (i < dataSourceDataRows.length ? dataSourceDataRows[i].rowNumber : currentRow);
            const convertResult = this.convertFormulaForMerge(
              formula,
              originalRow,
              dataCurrentRow, // 当前行号
              colIndex + 1, // 当前列号
              sortedHeaders,
              matchedBaseSheet,
              dataSourceSheet,
              baseHeadersMap,
              dataSourceHeadersMap,
              isFromDataSource ? 'dataSource' : (isFromBase ? 'base' : 'unknown'), // 标识公式来源
              baseHeaderRow, // 基础表表头行数
              dataSourceHeaderRow, // 数据源表表头行数
              updateSummaryTable // 是否更新汇总表
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
                  sortedHeaders,
                  isFromDataSource ? 'dataSource' : (isFromBase ? 'base' : 'unknown')
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
        dataCurrentRow++;
      }

      // 存储数据行数（用于汇总表引用）
      const dataRowCount = dataCurrentRow - dataStartRow;
      mergedSheetDataRowCountMap.set(dataSourceSheetName, dataRowCount);

      // 添加表尾功能行
      const footerFunctions = this.footerFunctions();
      if (footerFunctions.length > 0) {
        const footerRow: any[] = [];
        const dataEndRow = dataCurrentRow - 1; // 数据结束行

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
        dataCurrentRow++;
      }

      // 自动调整列宽
      this.excelUtils.autoFitColumns(newSheet, sortedHeaders.length);

      // 设置表头筛选器
      const headerRowNum = updateSummaryTable ? 3 : 2; // 表头行号
      newSheet.autoFilter = {
        from: { row: headerRowNum, column: 1 },
        to: { row: dataCurrentRow - 1, column: sortedHeaders.length }
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
            const cellValue = this.excelUtils.getCellValueWithFormula(cell, dataSourceSheet, rowNum, colNumber);
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

    // 如果更新汇总表，填充汇总表数据
    if (updateSummaryTable && summarySheetName && summaryHeaders.length > 0 && summarySheet) {
      await this.fillSummarySheet(summarySheet, mergedSheetMap, mergedSheetDataRowCountMap, style);
    }

    // 生成主文件
    const buffer = await newWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return blob;
  }

  // 填充汇总表数据
  private async fillSummarySheet(
    summarySheet: ExcelJS.Worksheet,
    mergedSheetMap: Map<string, ExcelJS.Worksheet>,
    mergedSheetDataRowCountMap: Map<string, number>,
    style: TableStyleConfig
  ) {
    const summaryHeaders = this.summaryHeaders();
    const headerMappings = this.headerMappings();
    const footerFunctions = this.footerFunctions();
    const sortedHeaders = this.sortedHeaders();
    const selectedDataSourceSheets = this.selectedDataSourceSheets();
    const updateSummaryTable = this.updateSummaryTable();

    // 添加数据行（每个合并表对应一行）
    for (let rowIndex = 0; rowIndex < selectedDataSourceSheets.length; rowIndex++) {
      const mergedSheetName = selectedDataSourceSheets[rowIndex];
      const mergedSheet = mergedSheetMap.get(mergedSheetName);
      if (!mergedSheet) continue;

      const dataRow: any[] = [];
      // 汇总表的数据行号：第1行是标题，第2行是表头，数据从第3行开始
      const dataRowIndex = rowIndex + 3;
      const formulaMap = new Map<number, string>(); // 列索引 -> 公式

      for (let colIndex = 0; colIndex < summaryHeaders.length; colIndex++) {
        const summaryHeader = summaryHeaders[colIndex];
        let cellValue: any = '';
        let cellFormula: string | null = null;

        // 检查是否有映射到合并表的表尾功能
        const mappedOptionId = headerMappings.get(summaryHeader);
        if (mappedOptionId && mappedOptionId.startsWith('footer:')) {
          // 映射到合并表表尾功能 - 使用公式引用合并表的表尾合计单元格
          const footerId = mappedOptionId.replace('footer:', '');
          const footer = footerFunctions.find(f => f.id === footerId);
          if (footer) {
            const mappedHeader = footer.header;
            const mappedHeaderIndex = sortedHeaders.indexOf(mappedHeader);
            if (mappedHeaderIndex >= 0) {
              // 计算合并表表尾行号：
              // 如果更新汇总表，合并表第1行是返回按钮，第2行是标题，第3行是表头，第4行开始是数据
              // 所以表尾在数据行数+4行（第4行是第一个数据行，第4+dataRowCount-1行是最后一个数据行，第4+dataRowCount行是表尾）
              // 否则，合并表第1行是标题，第2行是表头，第3行开始是数据，表尾在数据行数+3行
              const dataRowCount = mergedSheetDataRowCountMap.get(mergedSheetName) || 0;
              const footerRow = updateSummaryTable ? (dataRowCount + 4) : (dataRowCount + 3);
              const targetCol = this.excelUtils.getExcelColumnName(mappedHeaderIndex + 1);
              const sheetName = `'${mergedSheet.name}'`;
              cellFormula = `${sheetName}!${targetCol}${footerRow}`;
              cellValue = null;
            } else {
              cellValue = 'F-Null';
            }
          } else {
            cellValue = 'F-Null';
          }
        } else {
          // 没有映射，填充空值
          cellValue = '';
        }

        // 如果有公式，存储到 formulaMap 中
        if (cellFormula) {
          formulaMap.set(colIndex + 1, cellFormula);
          dataRow.push(null);
        } else {
          dataRow.push(cellValue);
        }
      }

      const row = summarySheet.addRow(dataRow);
      row.height = 20 * 0.75;
      for (let colNumber = 1; colNumber <= summaryHeaders.length; colNumber++) {
        const cell = row.getCell(colNumber);
        const summaryHeader = summaryHeaders[colNumber - 1];

        // 设置公式或超链接
        if (formulaMap.has(colNumber)) {
          const formula = formulaMap.get(colNumber)!;
          cell.value = { formula: formula };
        } else if (summaryHeader && mergedSheetName) {
          // 第一个表头列添加超链接到合并表
          if (colNumber === 1) {
            cell.value = {
              text: mergedSheetName,
              hyperlink: `#${mergedSheet.name}!A1`
            };
            cell.font = {
              name: style.fontFamily,
              size: style.dataFontSize,
              bold: style.dataFontBold,
              underline: true,
              color: { argb: 'FF0000FF' } // 蓝色
            };
          }
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
    const summaryFooterFunctions = this.footerFunctions();
    if (summaryFooterFunctions.length > 0) {
      const footerRow: any[] = [];

      for (const summaryHeader of summaryHeaders) {
        const footer = summaryFooterFunctions.find(f => f.header === summaryHeader);
        if (footer) {
          // 检查该表头是否映射到合并表的表尾功能
          const mappedOptionId = headerMappings.get(summaryHeader);
          if (mappedOptionId && mappedOptionId.startsWith('footer:')) {
            // 如果映射到合并表的表尾功能，汇总表的表尾应该汇总所有合并表的表尾
            const footerId = mappedOptionId.replace('footer:', '');
            const mappedFooter = footerFunctions.find(f => f.id === footerId);
            if (mappedFooter) {
              const mappedHeader = mappedFooter.header;
              const mappedHeaderIndex = sortedHeaders.indexOf(mappedHeader);
              if (mappedHeaderIndex >= 0) {
                const targetCol = this.excelUtils.getExcelColumnName(mappedHeaderIndex + 1);
                // 构建引用所有合并表表尾的公式
                const footerRefs: string[] = [];
                for (const mergedSheetName of selectedDataSourceSheets) {
                  const mergedSheet = mergedSheetMap.get(mergedSheetName);
                  if (mergedSheet) {
                    const dataRowCount = mergedSheetDataRowCountMap.get(mergedSheetName) || 0;
                    const footerRowNum = updateSummaryTable ? (dataRowCount + 4) : (dataRowCount + 3);
                    footerRefs.push(`'${mergedSheet.name}'!${targetCol}${footerRowNum}`);
                  }
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
            // 如果没有映射，填充空值
            footerRow.push('');
          }
        } else {
          footerRow.push('');
        }
      }

      const row = summarySheet.addRow(footerRow);
      row.height = 22 * 0.75;
      row.eachCell((cell: ExcelJS.Cell, colNumber: number) => {
        this.excelUtils.applyCellStyle(cell, style, 'total');
      });
    }

    // 自动调整列宽
    this.excelUtils.autoFitColumns(summarySheet, summaryHeaders.length);

    // 设置表头筛选器
    const dataRowCount = selectedDataSourceSheets.length;
    summarySheet.autoFilter = {
      from: { row: 2, column: 1 },
      to: { row: 2 + dataRowCount, column: summaryHeaders.length }
    };
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
    dataSourceHeadersMap: Map<number, string>,
    sourceTable: 'base' | 'dataSource' | 'unknown' = 'unknown', // 标识公式来自哪个表
    baseHeaderRow: number = 1, // 基础表表头行数
    dataSourceHeaderRow: number = 1, // 数据源表表头行数
    updateSummaryTable: boolean = false // 是否更新汇总表
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
          let hasMissingHeadersInRange = false;

          for (const part of parts) {
            const result = this.convertSingleCellRefForMerge(
              part,
              originalRow,
              currentRow,
              sortedHeaders,
              baseSheet,
              dataSourceSheet,
              baseHeadersMap,
              dataSourceHeadersMap,
              sourceTable,
              baseHeaderRow,
              dataSourceHeaderRow,
              updateSummaryTable
            );
            if (result === null) {
              convertedParts.push(null);
            } else if (typeof result === 'string') {
              convertedParts.push(result);
            } else {
              // result是对象，包含converted和missingHeaders
              if (result.missingHeaders && result.missingHeaders.length > 0) {
                // 如果缺少表头，标记并收集缺少的表头
                hasMissingHeadersInRange = true;
                result.missingHeaders.forEach(h => missingHeadersSet.add(h));
                convertedParts.push(null); // 标记为失败
              } else {
                // 转换成功
                convertedParts.push(result.converted);
              }
            }
          }

          // 检查是否有转换失败或缺少表头的情况
          if (convertedParts.some(p => p === null) || hasMissingHeadersInRange) {
            // 如果范围引用中有部分失败或缺少表头，尝试提取缺少的表头
            parts.forEach(part => {
              const headerName = this.extractHeaderNameFromRef(part, baseHeadersMap, dataSourceHeadersMap, sourceTable);
              if (headerName && !sortedHeaders.includes(headerName)) {
                missingHeadersSet.add(headerName);
              }
            });
            return null;
          }

          // 所有部分都转换成功，拼接结果
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
            dataSourceHeadersMap,
            sourceTable,
            baseHeaderRow,
            dataSourceHeaderRow,
            updateSummaryTable
          );
          if (result === null) {
            // 提取缺少的表头
            const headerName = this.extractHeaderNameFromRef(fullRef, baseHeadersMap, dataSourceHeadersMap, sourceTable);
            if (headerName && !sortedHeaders.includes(headerName)) {
              missingHeadersSet.add(headerName);
            }
            return null;
          } else if (typeof result === 'string') {
            convertedFormula = convertedFormula.substring(0, matchIndex) +
                               result +
                               convertedFormula.substring(matchIndex + fullRef.length);
          } else {
            // result是对象，包含missingHeaders表示缺少表头
            // 如果缺少表头，应该标记整个公式转换失败
            if (result.missingHeaders && result.missingHeaders.length > 0) {
              result.missingHeaders.forEach(h => missingHeadersSet.add(h));
              // 不替换引用，直接返回失败
              return null;
            }
            // 如果没有缺少表头，正常替换
            convertedFormula = convertedFormula.substring(0, matchIndex) +
                               result.converted +
                               convertedFormula.substring(matchIndex + fullRef.length);
          }
        }
      }
      return {
        formula: convertedFormula,
        missingHeaders: Array.from(missingHeadersSet)
      };
    } catch (e) {
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
    dataSourceHeadersMap: Map<number, string>,
    sourceTable: 'base' | 'dataSource' | 'unknown' = 'unknown' // 标识公式来自哪个表
  ): string | null {
    // 根据公式来源表，只在对应的表头映射中查找
    if (sourceTable === 'base') {
      // 公式来自基础表，只在基础表的表头映射中查找
      return this.excelUtils.extractHeaderNameFromRef(ref, baseHeadersMap);
    } else if (sourceTable === 'dataSource') {
      // 公式来自数据源表，只在数据源表的表头映射中查找
      return this.excelUtils.extractHeaderNameFromRef(ref, dataSourceHeadersMap);
    } else {
      // 未知来源，先尝试基础表，再尝试数据源表（兼容旧逻辑）
      return this.excelUtils.extractHeaderNameFromRef(ref, baseHeadersMap) ||
             this.excelUtils.extractHeaderNameFromRef(ref, dataSourceHeadersMap);
    }
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
    dataSourceHeadersMap: Map<number, string>,
    sourceTable: 'base' | 'dataSource' | 'unknown' = 'unknown', // 标识公式来自哪个表
    baseHeaderRow: number = 1, // 基础表表头行数
    dataSourceHeaderRow: number = 1, // 数据源表表头行数
    updateSummaryTable: boolean = false // 是否更新汇总表
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
    // 根据公式来源表，只在对应的表头映射中查找
    let headerName = '';
    if (sourceTable === 'base') {
      // 公式来自基础表，只在基础表的表头映射中查找
      if (baseSheet && baseHeadersMap.has(originalColNum)) {
        headerName = baseHeadersMap.get(originalColNum)!;
      }
    } else if (sourceTable === 'dataSource') {
      // 公式来自数据源表，只在数据源表的表头映射中查找
      if (dataSourceSheet && dataSourceHeadersMap.has(originalColNum)) {
        headerName = dataSourceHeadersMap.get(originalColNum)!;
      }
    } else {
      // 未知来源，先尝试基础表，再尝试数据源表（兼容旧逻辑）
      if (baseSheet && baseHeadersMap.has(originalColNum)) {
        headerName = baseHeadersMap.get(originalColNum)!;
      } else if (dataSourceSheet && dataSourceHeadersMap.has(originalColNum)) {
        headerName = dataSourceHeadersMap.get(originalColNum)!;
      }
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

      // 根据公式来源表确定原始表的表头行数
      const originalHeaderRow = sourceTable === 'base' ? baseHeaderRow :
                               (sourceTable === 'dataSource' ? dataSourceHeaderRow : 1);

      // 新表的表头行和数据起始行：如果更新汇总表，第1行是返回链接，第2行是标题，第3行是表头，数据从第4行开始
      // 否则，第1行是标题，第2行是表头，数据从第3行开始
      const NEW_HEADER_ROW = updateSummaryTable ? 3 : 2;
      const NEW_DATA_START_ROW = updateSummaryTable ? 4 : 3;


      let newRowNum: number;
      if (originalRowNum === originalHeaderRow) {
        // 引用表头行 -> 新表表头行
        newRowNum = NEW_HEADER_ROW;
      } else {
        // 引用数据行 -> 根据表头行偏移重新转换
        // 新表的表头行和数据起始行根据updateSummaryTable确定
        // 原始表的表头行是originalHeaderRow，数据从originalHeaderRow+1开始

        if (isRowAbsolute) {
          // 绝对引用：直接计算相对于表头行的偏移
          // 原始表中引用行相对于表头行的偏移
          const relativeOffset = originalRowNum - originalHeaderRow;
          // 新表中的行号 = 新表表头行 + 相对偏移
          newRowNum = NEW_HEADER_ROW + relativeOffset;
          // 确保不会小于表头行
          if (newRowNum < NEW_HEADER_ROW) {
            newRowNum = NEW_HEADER_ROW;
          }
        } else {
          // 相对引用：保持相对位置，基于当前行计算
          // 判断引用的是表头行还是数据行
          const isReferencingHeaderRow = originalRowNum === originalHeaderRow;
          const isReferencingDataRow = originalRowNum > originalHeaderRow;


          if (isReferencingHeaderRow) {
            // 引用的是表头行，转换为新表的表头行
            newRowNum = NEW_HEADER_ROW;
          } else if (isReferencingDataRow) {
            // 引用的是数据行
            // 计算原始表中引用行相对于表头行的偏移（数据行索引，从1开始）
            const originalRefDataRowIndex = originalRowNum - originalHeaderRow;
            // 计算公式行相对于表头行的偏移（数据行索引，从1开始）
            const originalFormulaDataRowIndex = originalRow - originalHeaderRow;
            // 计算当前行相对于新表表头行的偏移（数据行索引，从1开始）
            const currentDataRowIndex = currentRow - NEW_HEADER_ROW;


            // 如果公式引用的行号等于公式所在行，不需要转换（引用当前行）
            if (originalRowNum === originalRow) {
              // 引用当前行，在新表中也引用当前行
              newRowNum = currentRow;

            } else {
              // 如果有偏差，根据偏差转换
              // 计算原始表中引用行相对于公式行的偏移（行号差）
              const relativeRowOffset = originalRowNum - originalRow;
              // 新表中的行号 = 当前行 + 相对行偏移
              newRowNum = currentRow + relativeRowOffset;


              // 确保不会小于表头行
              if (newRowNum < NEW_HEADER_ROW) {
                // 如果计算结果小于表头行，说明引用的是表头行之前的行
                // 但如果是相对引用，且引用的是数据行，应该保持为数据行的第1行
                if (originalRefDataRowIndex >= 1) {
                  // 引用的是数据行，应该保持为数据行的第1行
                  newRowNum = NEW_DATA_START_ROW;
                } else {
                  // 引用的是表头行之前的行，转换为表头行
                  newRowNum = NEW_HEADER_ROW;
                }
              }
              // 额外检查：如果原始引用行号等于当前行对应的原始行号，应该保持引用当前行
              // 这处理了当originalRow计算不正确时的情况
              // 如果originalRowNum等于currentRow对应的原始行号（通过数据行索引计算），应该引用currentRow
              const currentRowOriginalRowNum = originalHeaderRow + currentDataRowIndex;
              if (originalRowNum === currentRowOriginalRowNum) {
                // 引用的是当前行对应的原始行号，应该保持引用当前行
                newRowNum = currentRow;
              }

            }

          } else {
            // 引用的是表头行之前的行，转换为表头行
            newRowNum = NEW_HEADER_ROW;
          }
        }

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
    sortedHeaders: string[],
    sourceTable: 'base' | 'dataSource' | 'unknown' = 'unknown' // 标识公式来自哪个表
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
      // 根据公式来源表，只在对应的表头映射中查找
      let headerName = '';
      if (sourceTable === 'base') {
        // 公式来自基础表，只在基础表的表头映射中查找
        if (baseHeadersMap.has(originalColNum)) {
          headerName = baseHeadersMap.get(originalColNum)!;
        }
      } else if (sourceTable === 'dataSource') {
        // 公式来自数据源表，只在数据源表的表头映射中查找
        if (dataSourceHeadersMap.has(originalColNum)) {
          headerName = dataSourceHeadersMap.get(originalColNum)!;
        }
      } else {
        // 未知来源，先尝试基础表，再尝试数据源表（兼容旧逻辑）
        if (baseHeadersMap.has(originalColNum)) {
          headerName = baseHeadersMap.get(originalColNum)!;
        } else if (dataSourceHeadersMap.has(originalColNum)) {
          headerName = dataSourceHeadersMap.get(originalColNum)!;
        }
      }

      // 如果找到了表头名称，但不在新表头列表中，则记录为缺少的表头
      if (headerName && !sortedHeaders.includes(headerName)) {
        missingHeaders.add(headerName);
      }
    }

    return Array.from(missingHeaders);
  }
}

