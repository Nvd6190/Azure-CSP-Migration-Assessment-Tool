import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { AssessmentResponse } from '../models/assessment.model';

@Injectable({ providedIn: 'root' })
export class MigrationService {
  private apiUrl = '/api';

  constructor(private http: HttpClient) {}

  assessFile(file: File): Observable<AssessmentResponse> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<AssessmentResponse>(`${this.apiUrl}/assess-json`, formData);
  }

  assessRegionFile(file: File): Observable<AssessmentResponse> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<AssessmentResponse>(`${this.apiUrl}/assess-region-json`, formData);
  }

  assessJioFile(file: File): Observable<AssessmentResponse> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<AssessmentResponse>(`${this.apiUrl}/assess-jio-json`, formData);
  }

  assessAwsFile(file: File): Observable<AssessmentResponse> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<AssessmentResponse>(`${this.apiUrl}/assess-aws-json`, formData);
  }

  assessGcpFile(file: File): Observable<AssessmentResponse> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<AssessmentResponse>(`${this.apiUrl}/assess-gcp-json`, formData);
  }

  getDownloadUrl(downloadId: string): string {
    return `${this.apiUrl}/download/${downloadId}`;
  }

  refreshRules(): Observable<any> {
    return this.http.post<any>(`${this.apiUrl}/rules/refresh`, {});
  }

  getRules(): Observable<any> {
    return this.http.get<any>(`${this.apiUrl}/rules`);
  }

  refreshJioData(file: File): Observable<any> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<any>(`${this.apiUrl}/jio/refresh`, formData);
  }
}
