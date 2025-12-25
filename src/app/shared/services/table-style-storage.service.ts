import { Injectable } from '@angular/core';
import { TableStyleConfig, DEFAULT_TABLE_STYLE } from '../models/table-style.model';

@Injectable({
  providedIn: 'root'
})
export class TableStyleStorageService {
  // 统一的存储键名，所有模块共享
  private readonly STORAGE_KEY = 'xrexcel_table_style';

  /**
   * 保存表格风格配置到 localStorage
   */
  saveTableStyle(style: TableStyleConfig): void {
    try {
      localStorage.setItem(this.STORAGE_KEY, JSON.stringify(style));
    } catch (error) {
      console.error('保存表格风格配置失败:', error);
    }
  }

  /**
   * 从 localStorage 加载表格风格配置
   */
  loadTableStyle(): TableStyleConfig | null {
    try {
      const stored = localStorage.getItem(this.STORAGE_KEY);
      if (stored) {
        return JSON.parse(stored) as TableStyleConfig;
      }
    } catch (error) {
      console.error('加载表格风格配置失败:', error);
    }
    return null;
  }

  /**
   * 检查是否有保存的自定义配置
   */
  hasCustomStyle(): boolean {
    return this.loadTableStyle() !== null;
  }

  /**
   * 重置为默认配置（删除保存的配置）
   */
  resetToDefault(): void {
    try {
      localStorage.removeItem(this.STORAGE_KEY);
    } catch (error) {
      console.error('重置表格风格配置失败:', error);
    }
  }

  /**
   * 获取默认配置
   */
  getDefaultStyle(): TableStyleConfig {
    return { ...DEFAULT_TABLE_STYLE };
  }
}

