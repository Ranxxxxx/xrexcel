import { Component, input, output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { MatCardModule } from '@angular/material/card';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatSelectModule } from '@angular/material/select';
import { MatInputModule } from '@angular/material/input';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { MatExpansionModule } from '@angular/material/expansion';
import { MatIconModule } from '@angular/material/icon';
import { TableStyleConfig } from '../../models/table-style.model';

@Component({
  selector: 'app-table-style-preview',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatCardModule,
    MatFormFieldModule,
    MatSelectModule,
    MatInputModule,
    MatCheckboxModule,
    MatExpansionModule,
    MatIconModule
  ],
  templateUrl: './table-style-preview.component.html',
  styleUrl: './table-style-preview.component.scss'
})
export class TableStylePreviewComponent {
  // 输入：表格风格配置
  tableStyle = input.required<TableStyleConfig>();

  // 输入：预览数据（二维数组，第一行是标题，第二行是表头，后续是数据行，最后一行是合计）
  previewData = input<any[][]>([]);

  // 输入：预览区域是否展开
  previewExpanded = input<boolean>(false);

  // 输出：风格配置变化事件
  styleChange = output<{ key: keyof TableStyleConfig; value: any }>();

  // 输出：预览区域展开状态变化
  previewExpandedChange = output<boolean>();

  // 选项配置
  styleOptions = ['商务风格', '简约风格', '经典风格', '现代风格'];
  fontOptions = ['微软雅黑', '宋体', '黑体', 'Arial', 'Times New Roman'];
  fontSizeOptions = [8, 9, 10, 11, 12, 14, 16, 18, 20];
  borderStyleOptions = [
    { value: 'thin', label: '细边框' },
    { value: 'medium', label: '中等边框' },
    { value: 'thick', label: '粗边框' }
  ];

  updateStyle(key: keyof TableStyleConfig, value: any) {
    this.styleChange.emit({ key, value });
  }

  onPreviewExpandedChange(expanded: boolean) {
    this.previewExpandedChange.emit(expanded);
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
}

