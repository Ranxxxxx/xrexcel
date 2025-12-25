import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MatIconModule } from '@angular/material/icon';

@Component({
  selector: 'app-privacy-notice',
  standalone: true,
  imports: [
    CommonModule,
    MatIconModule
  ],
  templateUrl: './privacy-notice.component.html',
  styleUrl: './privacy-notice.component.scss'
})
export class PrivacyNoticeComponent {
}

