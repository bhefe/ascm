import { Component } from '@angular/core';
import { Router } from '@angular/router';
import { CommonModule } from '@angular/common';
import { ComplianceService, CheckResponse } from '../compliance.service';

@Component({
  selector: 'app-upload',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './upload.component.html',
  styleUrls: ['./upload.component.css'],
})
export class UploadComponent {
  selectedFile: File | null = null;
  loading = false;
  error = '';

  constructor(private svc: ComplianceService, private router: Router) {}

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    if (input.files?.length) {
      this.selectedFile = input.files[0];
      this.error = '';
    }
  }

  onDrop(event: DragEvent): void {
    event.preventDefault();
    const files = event.dataTransfer?.files;
    if (files?.length) {
      this.selectedFile = files[0];
      this.error = '';
    }
  }

  onDragOver(event: DragEvent): void {
    event.preventDefault();
  }

  submit(): void {
    if (!this.selectedFile) return;
    if (!this.selectedFile.name.endsWith('.csv')) {
      this.error = 'Please upload a valid .csv file.';
      return;
    }
    this.loading = true;
    this.error = '';
    this.svc.check(this.selectedFile).subscribe({
      next: (res: CheckResponse) => {
        this.loading = false;
        this.router.navigate(['/report'], { state: { data: res } });
      },
      error: (err) => {
        this.loading = false;
        this.error = err.error?.detail ?? 'An error occurred. Please try again.';
      },
    });
  }
}
