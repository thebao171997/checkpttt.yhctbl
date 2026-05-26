export interface ProcessedProcedure {
  id: string;
  patientId: string;
  patientName: string;
  procedureName: string;
  doctorName: string;
  machineId: string;
  startTime: Date;
  endTime: Date;
  rawRow: any;
  rowNumber: number;
}

export type ConflictType = 
  | 'PATIENT_TIME_OVERLAP' 
  | 'DOCTOR_TIME_OVERLAP' 
  | 'MACHINE_TIME_OVERLAP'
  | 'MISSING_MACHINE_ID'
  | 'INVALID_TIME' 
  | 'MISSING_DATA'
  | 'CLS_TIME_OVERLAP';

export interface Conflict {
  id: string;
  type: ConflictType;
  message: string;
  procedures: ProcessedProcedure[];
}

export interface AnalysisResult {
  procedures: ProcessedProcedure[];
  conflicts: Conflict[];
  totalPatients: number;
  totalDoctors: number;
  errors: string[];
}
