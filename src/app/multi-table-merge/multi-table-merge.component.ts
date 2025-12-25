import { Component, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { RouterLink } from '@angular/router';
import { MatToolbarModule } from '@angular/material/toolbar';
import { MatIconModule } from '@angular/material/icon';
import { MatButtonModule } from '@angular/material/button';
import { MatCardModule } from '@angular/material/card';
import { TableStyleConfig, DEFAULT_TABLE_STYLE } from '../shared/models/table-style.model';
import { TableStylePreviewComponent } from '../shared/components/table-style-preview/table-style-preview.component';
import { TableStyleStorageService } from '../shared/services/table-style-storage.service';

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
    TableStylePreviewComponent
  ],
  templateUrl: './multi-table-merge.component.html',
  styleUrl: './multi-table-merge.component.scss'
})
export class MultiTableMergeComponent {
  tableStyle = signal<TableStyleConfig>({ ...DEFAULT_TABLE_STYLE });
  previewExpanded = signal<boolean>(false); // 预览区域展开状态，默认折叠
  previewData = signal<any[][]>([]); // 预览数据，待后续实现多表合并功能时填充
  showResetButton = signal<boolean>(false); // 是否显示重置按钮

  constructor(private tableStyleStorage: TableStyleStorageService) {
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
}

