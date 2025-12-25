import { Component, input, output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MatCardModule } from '@angular/material/card';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatProgressBarModule } from '@angular/material/progress-bar';
import { FormsModule } from '@angular/forms';

@Component({
  selector: 'app-confirm-dialog',
  standalone: true,
  imports: [
    CommonModule,
    MatCardModule,
    MatButtonModule,
    MatIconModule,
    MatFormFieldModule,
    MatInputModule,
    MatProgressBarModule,
    FormsModule
  ],
  templateUrl: './confirm-dialog.component.html',
  styleUrl: './confirm-dialog.component.scss'
})
export class ConfirmDialogComponent {
  // 输入属性
  visible = input<boolean>(false);
  fileName = input<string>('');
  progress = input<number>(0);
  isProcessing = input<boolean>(false);

  // 输出事件
  fileNameChange = output<string>();
  close = output<void>();
  confirm = output<void>();

  onFileNameChange(value: string) {
    this.fileNameChange.emit(value);
  }

  onClose() {
    this.close.emit();
  }

  onConfirm() {
    this.confirm.emit();
  }

  onOverlayClick() {
    // 如果不在处理中，允许点击关闭
    if (!this.isProcessing() && this.progress() === 0) {
      this.onClose();
    }
  }
}

