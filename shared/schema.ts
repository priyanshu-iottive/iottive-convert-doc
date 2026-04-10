// No database needed for this utility app
// Just type definitions

export interface ConversionRequest {
  file: File;
  clientName: string;
  contactName: string;
  contactEmail: string;
  projectTitle: string;
}

export interface ConversionResult {
  success: boolean;
  error?: string;
}
