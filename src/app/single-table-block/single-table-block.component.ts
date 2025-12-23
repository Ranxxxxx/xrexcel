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
import { TableStyleConfig, DEFAULT_TABLE_STYLE } from '../shared/models/table-style.model';

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
    MatExpansionModule
  ],
  templateUrl: './single-table-block.component.html',
  styleUrl: './single-table-block.component.scss'
})
export class SingleTableBlockComponent {
  tableStyle = signal<TableStyleConfig>({ ...DEFAULT_TABLE_STYLE });

  styleOptions = ['商务风格', '简约风格', '经典风格', '现代风格'];

  fontOptions = ['微软雅黑', '宋体', '黑体', 'Arial', 'Times New Roman'];

  fontSizeOptions = [8, 9, 10, 11, 12, 14, 16, 18, 20];

  borderStyleOptions = [
    { value: 'thin', label: '细边框' },
    { value: 'medium', label: '中等边框' },
    { value: 'thick', label: '粗边框' }
  ];

  updateStyle(key: keyof TableStyleConfig, value: any) {
    this.tableStyle.update(config => ({ ...config, [key]: value }));
  }
}

