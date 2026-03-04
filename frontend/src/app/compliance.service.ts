import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';

export interface SoftwareResult {
  software: string;
  version: string;
  publisher: string;
  status: 'Allowed' | 'Not Allowed' | 'Not Found';
  matched: string;
  confidence: number;
  match_type: string;
  risk_level: string;
  risk_reasoning: string;
}

export interface CheckResponse {
  results: SoftwareResult[];
  counts: { Allowed: number; 'Not Allowed': number; 'Not Found': number };
  meta: Record<string, string>;
  total: number;
  filename: string;
  generated: string;
  csv_b64: string;
}

@Injectable({ providedIn: 'root' })
export class ComplianceService {
  private readonly api = 'http://localhost:8000/api';

  constructor(private http: HttpClient) {}

  check(file: File): Observable<CheckResponse> {
    const form = new FormData();
    form.append('csv_file', file);
    return this.http.post<CheckResponse>(`${this.api}/check`, form);
  }

  download(csvB64: string, filename: string): void {
    const form = new FormData();
    form.append('csv_b64', csvB64);
    form.append('filename', filename);
    this.http.post(`${this.api}/download`, form, { responseType: 'blob' }).subscribe(blob => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `compliance_${filename.replace('.csv', '')}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    });
  }
}
