import { Component, inject, OnInit, signal } from '@angular/core';
import { RouterLink } from '@angular/router';
import { CommonModule } from '@angular/common';
import { MatCardModule } from '@angular/material/card';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { PrivacyNoticeComponent } from '../shared/components/privacy-notice/privacy-notice.component';
import { VersionService } from '../shared/services/version.service';

@Component({
  selector: 'app-home',
  standalone: true,
  imports: [
    CommonModule,
    RouterLink,
    MatCardModule,
    MatButtonModule,
    MatIconModule,
    PrivacyNoticeComponent
  ],
  templateUrl: './home.component.html',
  styleUrl: './home.component.scss'
})
export class HomeComponent implements OnInit {
  private versionService = inject(VersionService);
  version = signal<string>(''); // 从 package.json 读取版本号

  async ngOnInit() {
    const version = await this.versionService.getVersion();
    this.version.set(version);
  }

  features = [
    {
      id: 'summary-category',
      title: '汇总-分类',
      route: '/summary-category',
      icon: 'bar_chart',
      description: '对Excel数据进行汇总和分类处理'
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

