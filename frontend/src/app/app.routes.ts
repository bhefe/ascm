import { Routes } from '@angular/router';
import { UploadComponent } from './upload/upload.component';
import { ReportComponent } from './report/report.component';

export const routes: Routes = [
  { path: '', component: UploadComponent },
  { path: 'report', component: ReportComponent },
  { path: '**', redirectTo: '' },
];
