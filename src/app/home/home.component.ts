import { Component } from '@angular/core';
import { RouterLink } from '@angular/router';
import { CommonModule } from '@angular/common';
import { MatCardModule } from '@angular/material/card';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';

@Component({
  selector: 'app-home',
  standalone: true,
  imports: [
    CommonModule,
    RouterLink,
    MatCardModule,
    MatButtonModule,
    MatIconModule
  ],
  templateUrl: './home.component.html',
  styleUrl: './home.component.scss'
})
export class HomeComponent {
  features = [
    {
      id: 'summary-category',
      title: '汇总-分类',
      route: '/summary-category',
      icon: 'bar_chart',
      description: '对Excel数据进行汇总和分类处理'
    },
    {
      id: 'single-table-block',
      title: '单表格处理',
      route: '/single-table-block',
      icon: 'table_chart',
      description: '对Excel单个表格进行处理'
    },
    {
      id: 'multi-table-merge',
      title: '多表格数据合并',
      route: '/multi-table-merge',
      icon: 'merge_type',
      description: '将多个表格的数据进行合并处理'
    }
  ];
}

