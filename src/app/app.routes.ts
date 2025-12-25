import { Routes } from '@angular/router';
import { HomeComponent } from './home/home.component';
import { SummaryCategoryComponent } from './summary-category/summary-category.component';
import { MultiTableMergeComponent } from './multi-table-merge/multi-table-merge.component';

export const routes: Routes = [
  { path: '', component: HomeComponent },
  { path: 'summary-category', component: SummaryCategoryComponent },
  { path: 'multi-table-merge', component: MultiTableMergeComponent },
  { path: '**', redirectTo: '' }
];
