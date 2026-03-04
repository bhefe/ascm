import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { CheckResponse, SoftwareResult, ComplianceService } from '../compliance.service';

@Component({
  selector: 'app-report',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './report.component.html',
  styleUrls: ['./report.component.css'],
})
export class ReportComponent implements OnInit {
  data!: CheckResponse;
  filtered: SoftwareResult[] = [];
  activeFilter = 'all';
  searchText = '';
  expandedRow: number | null = null;

  constructor(private router: Router, private svc: ComplianceService) {}

  ngOnInit(): void {
    const nav = this.router.getCurrentNavigation();
    const state = nav?.extras?.state ?? history.state;
    if (!state?.data) {
      this.router.navigate(['/']);
      return;
    }
    this.data = state.data;
    this.applyFilter();
  }

  setFilter(filter: string): void {
    this.activeFilter = filter;
    this.applyFilter();
  }

  applyFilter(): void {
    let rows = this.data.results;
    if (this.activeFilter !== 'all') {
      const map: Record<string, string> = {
        allowed: 'Allowed',
        'not-allowed': 'Not Allowed',
        'not-found': 'Not Found',
      };
      rows = rows.filter(r => r.status === map[this.activeFilter]);
    }
    if (this.searchText.trim()) {
      const q = this.searchText.toLowerCase();
      rows = rows.filter(r =>
        r.software.toLowerCase().includes(q) ||
        r.publisher.toLowerCase().includes(q) ||
        r.matched.toLowerCase().includes(q)
      );
    }
    this.filtered = rows;
  }

  badgeClass(status: string): string {
    return status === 'Allowed' ? 'badge allowed'
         : status === 'Not Allowed' ? 'badge not-allowed'
         : 'badge not-found';
  }

  riskClass(level: string): string {
    return level === 'Low' ? 'risk-badge low'
         : level === 'Medium' ? 'risk-badge medium'
         : level === 'High' ? 'risk-badge high'
         : '';
  }

  toggleRow(index: number): void {
    this.expandedRow = this.expandedRow === index ? null : index;
  }

  download(): void {
    this.svc.download(this.data.csv_b64, this.data.filename);
  }

  newCheck(): void {
    this.router.navigate(['/']);
  }
}
