import { Injectable, signal } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class VersionService {
  private versionSignal = signal<string>('0.0.0');
  private versionLoaded = false;

  constructor(private http: HttpClient) {}

  async getVersion(): Promise<string> {
    if (!this.versionLoaded) {
      try {
        const packageJson = await firstValueFrom(this.http.get<{ version: string }>('/package.json'));
        this.versionSignal.set(packageJson.version || '0.0.0');
        this.versionLoaded = true;
      } catch (error) {
        console.warn('无法读取 package.json，使用默认版本号');
        this.versionLoaded = true;
      }
    }
    return this.versionSignal();
  }

  getVersionSync(): string {
    return this.versionSignal();
  }
}

