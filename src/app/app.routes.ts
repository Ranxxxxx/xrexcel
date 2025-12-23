import { Routes } from '@angular/router';
import { HomeComponent } from './home/home.component';
import { SummaryCategoryComponent } from './summary-category/summary-category.component';
import { SingleTableBlockComponent } from './single-table-block/single-table-block.component';
import { MultiTableMergeComponent } from './multi-table-merge/multi-table-merge.component';

export const routes: Routes = [
  { path: '', component: HomeComponent },
  { path: 'summary-category', component: SummaryCategoryComponent },
  { path: 'single-table-block', component: SingleTableBlockComponent },
  { path: 'multi-table-merge', component: MultiTableMergeComponent },
  { path: '**', redirectTo: '' }
];
