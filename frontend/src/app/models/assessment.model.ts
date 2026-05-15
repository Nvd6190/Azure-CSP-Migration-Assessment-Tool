export interface AssessmentResource {
  [key: string]: string | number | undefined;
  NAME?: string;
  TYPE?: string;
  'RESOURCE GROUP'?: string;
  LOCATION?: string;
  SUBSCRIPTION?: string;
  'SUBSCRIPTION MOVE SUPPORTED'?: string;
  'NORMALIZED TYPE'?: string;
  REMARKS?: string;
}

export interface AssessmentSummary {
  total: number;
  yes: number;
  no: number;
  review: number;
  conditional?: number;
}

export interface AssessmentResponse {
  summary: AssessmentSummary;
  resources: AssessmentResource[];
  downloadId: string;
}
