declare module "*.css";

interface IltUnitSummary {
  id: number;
  name: string;
}

interface IltCourseDetails {
  name: string;
  talentLmsId: number;
  units: IltUnitSummary[];
}

interface CourseProgressUser {
  employeeNumber: string;
  firstName: string;
  lastName: string;
  talentLmsUserId: number | null;
  status: string;
  totalTime: number;
}

interface CourseProgressUnit {
  unitId: number;
  unitName: string;
  users: CourseProgressUser[];
}

interface CourseProgressDetails {
  courseId: number;
  courseName: string;
  units: CourseProgressUnit[];
}

interface ProgressStats {
  totalJobs: number;
  processedJobs: number;
  completed: number;
  failed: number;
  notAttempted: number;
  notEnrolled: number;
  notChecked: number;
}

interface ExcelSummary {
  fileName: string;
  sheetName: string;
  totalRows: number;
  totalColumns: number;
  columns: string[];
  courses: string[];
  matchedCourses: number;
  unmatchedCourses: string[];
  iltCourses: string[];
  iltCourseDetails: IltCourseDetails[];
  courseProgressDetails: CourseProgressDetails[];
  uniqueUsers: number;
  matchedUsers: number;
  unmatchedUsers: string[];
  progressStats: ProgressStats;
}

interface AnalysisProgress {
  stage:
    | "reading"
    | "processingExcel"
    | "fetchingCourses"
    | "matchingCourses"
    | "checkingIlt"
    | "fetchingIltUnits"
    | "checkingUsers"
    | "matchingUsers"
    | "checkingUnitProgress"
    | "writingExcel"
    | "savingExcel"
    | "removingStrikeThrough";

  current: number;
  total: number;
}

interface SelectExcelResult {
  cancelled: boolean;
  summary?: ExcelSummary;
  error?: string;
}

interface GetSettingsResult {
  success: boolean;
  baseUrl: string;
  hasApiKey: boolean;
  error?: string;
}

interface SaveSettingsInput {
  baseUrl: string;
  apiKey?: string;
}

interface SaveSettingsResult {
  success: boolean;
  baseUrl?: string;
  hasApiKey?: boolean;
  error?: string;
}

interface Window {
  electronAPI: {
    selectExcel: () => Promise<SelectExcelResult>;
    getSettings: () => Promise<GetSettingsResult>;
    saveSettings: (settings: SaveSettingsInput) => Promise<SaveSettingsResult>;
    onAnalysisProgress: (
      callback: (progress: AnalysisProgress) => void,
    ) => void;
  };
}
