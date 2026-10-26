import { app, BrowserWindow, dialog, ipcMain, safeStorage } from "electron";
import path from "node:path";
import fs from "node:fs/promises";
import started from "electron-squirrel-startup";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import {
  getAllCourses,
  isCourseIlt,
  getUserIdByEmployeeNumber,
  getCourseIltUnits,
  getUnitProgress,
} from "./services/talentlms";

declare const MAIN_WINDOW_WEBPACK_ENTRY: string;
declare const MAIN_WINDOW_PRELOAD_WEBPACK_ENTRY: string;

if (started) {
  app.quit();
}

const createWindow = (): void => {
  const mainWindow = new BrowserWindow({
    width: 960,
    height: 720,
    minWidth: 960,
    maxWidth: 960,
    minHeight: 720,
    maxHeight: 720,
    resizable: false,
    maximizable: false,
    fullscreenable: false,
    webPreferences: {
      preload: MAIN_WINDOW_PRELOAD_WEBPACK_ENTRY,
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadURL(MAIN_WINDOW_WEBPACK_ENTRY);

  // mainWindow.webContents.openDevTools();
};

function hasCompletionDate(value: string): boolean {
  const normalizedValue = value.trim();

  return (
    normalizedValue !== "" &&
    normalizedValue !== "-" &&
    normalizedValue !== "–" &&
    normalizedValue !== "—"
  );
}

function normalizeCourseName(name: string): string {
  return name
    .normalize("NFKC")
    .replace(/\s+/gu, " ")
    .trim()
    .toLocaleLowerCase("pt-PT");
}

function normalizeUnitName(name: string): string {
  return name
    .normalize("NFKC")
    .replace(/\s+/gu, " ")
    .trim()
    .toLocaleLowerCase("pt-PT");
}

const DEFAULT_TALENTLMS_BASE_URL = "https://homaacademy.talentlms.com";

function normalizeBaseUrl(value: string): string {
  const trimmed = value.trim();

  if (!trimmed) {
    throw new Error("A Base URL não pode estar vazia.");
  }

  let parsedUrl: URL;

  try {
    parsedUrl = new URL(trimmed);
  } catch {
    throw new Error("A Base URL introduzida não é válida.");
  }

  if (parsedUrl.protocol !== "https:" && parsedUrl.protocol !== "http:") {
    throw new Error("A Base URL tem de começar por http:// ou https://.");
  }
  return trimmed.replace(/\/+$/u, "");
}

async function removeStrikethroughFromWorkbook(
  filePath: string,
): Promise<void> {
  const fileBuffer = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(fileBuffer);
  const stylesFile = zip.file("xl/styles.xml");

  if (!stylesFile) {
    console.warn(
      "Não foi encontrado xl/styles.xml. O strikethrough não foi alterado.",
    );

    return;
  }

  let stylesXml = await stylesFile.async("string");
  stylesXml = stylesXml.replace(/<strike\b[^>]*\/>/gi, "");
  stylesXml = stylesXml.replace(/<strike\b[^>]*>[\s\S]*?<\/strike>/gi, "");
  zip.file("xl/styles.xml", stylesXml);

  const cleanedBuffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: {
      level: 6,
    },
  });

  await fs.writeFile(filePath, cleanedBuffer);
}

interface StoredSettings {
  talentLMSBaseUrl?: string;
  talentLMSApiKeyEncrypted?: string;
}

const SETTINGS_FILE_NAME = "settings.json";

function getSettingsPath(): string {
  return path.join(app.getPath("userData"), SETTINGS_FILE_NAME);
}

async function readSettings(): Promise<StoredSettings> {
  const settingsPath = getSettingsPath();

  try {
    const content = await fs.readFile(settingsPath, "utf8");

    return JSON.parse(content) as StoredSettings;
  } catch (error) {
    const fileError = error as NodeJS.ErrnoException;

    if (fileError.code === "ENOENT") {
      return {};
    }

    throw error;
  }
}

async function writeSettings(settings: StoredSettings): Promise<void> {
  const settingsPath = getSettingsPath();

  await fs.mkdir(path.dirname(settingsPath), {
    recursive: true,
  });

  await fs.writeFile(settingsPath, JSON.stringify(settings, null, 2), "utf8");
}

async function getSavedTalentLmsBaseUrl(): Promise<string> {
  const settings = await readSettings();

  const savedBaseUrl = settings.talentLMSBaseUrl;

  if (!savedBaseUrl) {
    return DEFAULT_TALENTLMS_BASE_URL;
  }

  try {
    return normalizeBaseUrl(savedBaseUrl);
  } catch {
    return DEFAULT_TALENTLMS_BASE_URL;
  }
}

async function getSavedTalentLmsApiKey(): Promise<string> {
  const settings = await readSettings();

  if (!settings.talentLMSApiKeyEncrypted) {
    return "";
  }

  if (!safeStorage.isEncryptionAvailable()) {
    throw new Error(
      "A encriptação segura não está disponível neste computador.",
    );
  }

  const encryptedBuffer = Buffer.from(
    settings.talentLMSApiKeyEncrypted,
    "base64",
  );

  return safeStorage.decryptString(encryptedBuffer);
}

async function saveTalentLmsSettings(
  baseUrl: string,
  apiKey?: string,
): Promise<{ baseUrl: string; hasApiKey: boolean }> {
  const normalizedBaseUrl = normalizeBaseUrl(baseUrl);

  const currentSettings = await readSettings();

  let encryptedApiKey = currentSettings.talentLMSApiKeyEncrypted;

  const normalizedApiKey = apiKey?.trim() ?? "";

  if (normalizedApiKey) {
    if (!safeStorage.isEncryptionAvailable()) {
      throw new Error(
        "A encriptação segura não está disponível neste computador.",
      );
    }

    const encrypted = safeStorage.encryptString(normalizedApiKey);
    encryptedApiKey = encrypted.toString("base64");
  }

  await writeSettings({
    ...currentSettings,
    talentLMSBaseUrl: normalizedBaseUrl,
    talentLMSApiKeyEncrypted: encryptedApiKey,
  });

  return {
    baseUrl: normalizedBaseUrl,
    hasApiKey: Boolean(encryptedApiKey),
  };
}

function registerSettingsHandlers(): void {
  ipcMain.removeHandler("settings:get");

  ipcMain.removeHandler("settings:save");

  ipcMain.handle(
    "settings:get",

    async () => {
      try {
        const [baseUrl, apiKey] = await Promise.all([
          getSavedTalentLmsBaseUrl(),
          getSavedTalentLmsApiKey(),
        ]);

        return {
          success: true,
          baseUrl,
          hasApiKey: apiKey.length > 0,
        };
      } catch (error) {
        console.error("Erro ao ler settings:", error);

        return {
          success: false,
          baseUrl: DEFAULT_TALENTLMS_BASE_URL,
          hasApiKey: false,
          error: "Não foi possível ler a configuração.",
        };
      }
    },
  );

  ipcMain.handle(
    "settings:save",

    async (_event, input: SaveSettingsInput) => {
      try {
        const result = await saveTalentLmsSettings(input.baseUrl, input.apiKey);

        return {
          success: true,
          baseUrl: result.baseUrl,
          hasApiKey: result.hasApiKey,
        };
      } catch (error) {
        console.error("Erro ao guardar settings:", error);

        return {
          success: false,
          error:
            error instanceof Error
              ? error.message
              : "Não foi possível guardar a configuração.",
        };
      }
    },
  );
}

const HEADER_ROW = 4;
const DATA_START_ROW = HEADER_ROW + 1;

const IS_DEV_TEST = false;

const DEV_PROGRESS_LIMIT = 100;

interface IltUnitData {
  id: number;
  name: string;
}

interface CourseAnalysisData {
  rows: number[];
  talentLmsId: number | null;
  isIlt: boolean | null;
  iltUnits: IltUnitData[];
}

interface UserIltProgressData {
  courseId: number;
  courseName: string;
  unitId: number;
  unitName: string;
  status: string;
  totalTime: number;
}

interface UserAnalysisData {
  rows: number[];
  firstName: string;
  lastName: string;
  talentLmsUserId: number | null;
  iltProgress: UserIltProgressData[];
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

let currentAnalysis: {
  filePath: string;
  workbook: ExcelJS.Workbook;
  worksheet: ExcelJS.Worksheet;
  courses: Map<string, CourseAnalysisData>;
  users: Map<string, UserAnalysisData>;
} | null = null;

ipcMain.handle("excel:select", async (event): Promise<SelectExcelResult> => {
  const result = await dialog.showOpenDialog({
    title: "Selecionar relatório TalentLMS",
    buttonLabel: "Selecionar relatório",
    properties: ["openFile"],
    filters: [
      {
        name: "Ficheiros Excel",
        extensions: ["xlsx"],
      },
    ],
  });

  if (result.canceled || result.filePaths.length === 0) {
    return {
      cancelled: true,
    };
  }

  let apiKey: string;
  let baseUrl: string;

  try {
    [apiKey, baseUrl] = await Promise.all([
      getSavedTalentLmsApiKey(),
      getSavedTalentLmsBaseUrl(),
    ]);
  } catch (error) {
    console.error("Erro ao carregar configuração TalentLMS:", error);

    return {
      cancelled: false,
      error: "Não foi possível carregar a configuração da TalentLMS.",
    };
  }

  if (!apiKey) {
    return {
      cancelled: false,
      error: "A API key da TalentLMS não está configurada.",
    };
  }

  const talentLmsConfig = {
    baseUrl,
    apiKey,
    apiVersion: "2025-01-01",
  };

  const filePath = result.filePaths[0];

  try {
    const workbook = new ExcelJS.Workbook();

    event.sender.send("analysis:progress", {
      stage: "reading",
      current: 0,
      total: 0,
    });

    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.worksheets[0];

    if (!worksheet) {
      return {
        cancelled: false,
        error: "O ficheiro Excel não contém nenhuma folha.",
      };
    }

    const columns: string[] = [];
    let courseColumnNumber: number | null = null;
    let employeeColumnNumber: number | null = null;
    let firstNameColumnNumber: number | null = null;
    let lastNameColumnNumber: number | null = null;
    let completionDateColumnNumber: number | null = null;
    let durationColumnNumber: number | null = null;

    worksheet.getRow(HEADER_ROW).eachCell(
      {
        includeEmpty: false,
      },

      (cell, columnNumber) => {
        const columnName = cell.text.trim();
        const normalizedColumnName = columnName.toLocaleLowerCase("pt-PT");

        if (columnName) {
          columns.push(columnName);
        }

        if (normalizedColumnName === "curso") {
          courseColumnNumber = columnNumber;
        }

        if (normalizedColumnName === "número de colaborador") {
          employeeColumnNumber = columnNumber;
        }

        if (normalizedColumnName === "nome") {
          firstNameColumnNumber = columnNumber;
        }

        if (normalizedColumnName === "sobrenome") {
          lastNameColumnNumber = columnNumber;
        }

        if (normalizedColumnName === "data de conclusão") {
          completionDateColumnNumber = columnNumber;
        }

        if (normalizedColumnName === "duração (horas)") {
          durationColumnNumber = columnNumber;
        }
      },
    );

    if (courseColumnNumber === null) {
      return {
        cancelled: false,
        error: 'Não foi encontrada a coluna "Curso".',
      };
    }

    if (employeeColumnNumber === null) {
      return {
        cancelled: false,
        error: 'Não foi encontrada a coluna "Número de Colaborador".',
      };
    }

    if (completionDateColumnNumber === null) {
      return {
        cancelled: false,
        error: 'Não foi encontrada a coluna "Data de conclusão".',
      };
    }

    if (durationColumnNumber === null) {
      return {
        cancelled: false,
        error: 'Não foi encontrada a coluna "Duração (horas)".',
      };
    }

    const courses = new Map<string, CourseAnalysisData>();
    const totalRows = worksheet.actualRowCount - DATA_START_ROW + 1;

    for (
      let rowNumber = DATA_START_ROW;
      rowNumber <= worksheet.actualRowCount;
      rowNumber++
    ) {
      const row = worksheet.getRow(rowNumber);
      const courseName = row.getCell(courseColumnNumber).text.trim();

      if (courseName) {
        const existingCourse = courses.get(courseName);

        if (existingCourse) {
          existingCourse.rows.push(rowNumber);
        } else {
          courses.set(courseName, {
            rows: [rowNumber],
            talentLmsId: null,
            isIlt: null,
            iltUnits: [],
          });
        }
      }

      const current = rowNumber - DATA_START_ROW + 1;

      if (current % 500 === 0 || current === totalRows) {
        event.sender.send("analysis:progress", {
          stage: "processingExcel",
          current,
          total: totalRows,
        });

        await new Promise<void>((resolve) => setImmediate(resolve));
      }
    }

    const uniqueCourses = Array.from(courses.keys());

    if (!uniqueCourses.length) {
      return {
        cancelled: false,
        error: "Não foram encontrados cursos no relatório.",
      };
    }

    event.sender.send("analysis:progress", {
      stage: "fetchingCourses",
      current: 0,
      total: 0,
    });

    const talentLmsCourses = await getAllCourses(talentLmsConfig);

    const talentLmsCoursesByName = new Map(
      talentLmsCourses.map((course) => [
        normalizeCourseName(course.name),
        course,
      ]),
    );

    const totalCoursesToMatch = courses.size;
    let currentCourse = 0;

    for (const [courseName, courseData] of courses) {
      const talentLmsCourse = talentLmsCoursesByName.get(
        normalizeCourseName(courseName),
      );

      if (talentLmsCourse) {
        courseData.talentLmsId = talentLmsCourse.id;
      }

      currentCourse++;

      event.sender.send("analysis:progress", {
        stage: "matchingCourses",
        current: currentCourse,
        total: totalCoursesToMatch,
      });

      await new Promise<void>((resolve) => setImmediate(resolve));
    }

    const coursesWithId = Array.from(courses.values()).filter(
      (courseData) => courseData.talentLmsId !== null,
    );

    let checkedCourses = 0;

    const COURSE_BATCH_SIZE = 2;

    for (let i = 0; i < coursesWithId.length; i += COURSE_BATCH_SIZE) {
      const batch = coursesWithId.slice(i, i + COURSE_BATCH_SIZE);
      const results = await Promise.all(
        batch.map(async (courseData) => {
          if (courseData.talentLmsId === null) {
            return {
              courseData,
              isIlt: false,
            };
          }

          const isIlt = await isCourseIlt(
            talentLmsConfig,
            courseData.talentLmsId,
          );

          return {
            courseData,
            isIlt,
          };
        }),
      );

      for (const item of results) {
        item.courseData.isIlt = item.isIlt;
        checkedCourses++;

        event.sender.send("analysis:progress", {
          stage: "checkingIlt",
          current: checkedCourses,
          total: coursesWithId.length,
        });
      }
    }

    const iltCoursesData = Array.from(courses.values()).filter(
      (courseData) => courseData.isIlt === true,
    );

    const iltCourses = Array.from(courses.entries())
      .filter(([, courseData]) => courseData.isIlt === true)
      .map(([courseName]) => courseName);

    let checkedIltUnits = 0;

    const ILT_UNIT_BATCH_SIZE = 2;

    for (let i = 0; i < iltCoursesData.length; i += ILT_UNIT_BATCH_SIZE) {
      const batch = iltCoursesData.slice(i, i + ILT_UNIT_BATCH_SIZE);
      const results = await Promise.all(
        batch.map(async (courseData) => {
          if (courseData.talentLmsId === null) {
            return {
              courseData,
              units: [] as IltUnitData[],
            };
          }

          const units = await getCourseIltUnits(
            talentLmsConfig,
            courseData.talentLmsId,
          );

          return {
            courseData,
            units,
          };
        }),
      );

      for (const item of results) {
        item.courseData.iltUnits = item.units;
        checkedIltUnits++;

        event.sender.send("analysis:progress", {
          stage: "fetchingIltUnits",
          current: checkedIltUnits,
          total: iltCoursesData.length,
        });
      }
    }

    const iltCourseDetails = Array.from(courses.entries())
      .filter(
        ([, courseData]) =>
          courseData.isIlt === true && courseData.talentLmsId !== null,
      )
      .map(([courseName, courseData]) => ({
        name: courseName,
        talentLmsId: courseData.talentLmsId as number,
        units: courseData.iltUnits.map((unit) => ({
          id: unit.id,
          name: unit.name,
        })),
      }));

    const users = new Map<string, UserAnalysisData>();

    const totalIltRows = iltCoursesData.reduce(
      (total, courseData) => total + courseData.rows.length,
      0,
    );

    let checkedRows = 0;

    for (const courseData of iltCoursesData) {
      for (const rowNumber of courseData.rows) {
        const row = worksheet.getRow(rowNumber);
        const employeeNumber = row.getCell(employeeColumnNumber).text.trim();

        const firstName =
          firstNameColumnNumber !== null
            ? row.getCell(firstNameColumnNumber).text.trim()
            : "";

        const lastName =
          lastNameColumnNumber !== null
            ? row.getCell(lastNameColumnNumber).text.trim()
            : "";

        if (employeeNumber) {
          const existingUser = users.get(employeeNumber);

          if (existingUser) {
            existingUser.rows.push(rowNumber);

            if (!existingUser.firstName && firstName) {
              existingUser.firstName = firstName;
            }

            if (!existingUser.lastName && lastName) {
              existingUser.lastName = lastName;
            }
          } else {
            users.set(employeeNumber, {
              rows: [rowNumber],
              firstName,
              lastName,
              talentLmsUserId: null,
              iltProgress: [],
            });
          }
        }

        checkedRows++;

        if (checkedRows % 100 === 0 || checkedRows === totalIltRows) {
          event.sender.send("analysis:progress", {
            stage: "checkingUsers",
            current: checkedRows,
            total: totalIltRows,
          });

          await new Promise<void>((resolve) => setImmediate(resolve));
        }
      }
    }

    let matchedUsers = 0;
    let checkedUsers = 0;

    const unmatchedUsers: string[] = [];
    const userEntries = Array.from(users.entries());
    const USER_BATCH_SIZE = 2;

    for (let i = 0; i < userEntries.length; i += USER_BATCH_SIZE) {
      const batch = userEntries.slice(i, i + USER_BATCH_SIZE);

      const results = await Promise.all(
        batch.map(async ([employeeNumber, userData]) => {
          const talentLmsUserId = await getUserIdByEmployeeNumber(
            talentLmsConfig,
            employeeNumber,
          );

          return {
            employeeNumber,
            userData,
            talentLmsUserId,
          };
        }),
      );

      for (const item of results) {
        item.userData.talentLmsUserId = item.talentLmsUserId;

        if (item.talentLmsUserId !== null) {
          matchedUsers++;
        } else {
          unmatchedUsers.push(item.employeeNumber);
        }

        checkedUsers++;

        event.sender.send("analysis:progress", {
          stage: "matchingUsers",
          current: checkedUsers,
          total: users.size,
        });
      }
    }

    const unitProgressJobs: {
      employeeNumber: string;
      userData: UserAnalysisData;
      userId: number;
      courseId: number;
      courseName: string;
      unitId: number;
      unitName: string;
    }[] = [];

    for (const [courseName, courseData] of courses) {
      if (courseData.isIlt !== true || courseData.talentLmsId === null) {
        continue;
      }

      const courseEmployeeNumbers = new Set<string>();

      for (const rowNumber of courseData.rows) {
        const row = worksheet.getRow(rowNumber);

        const employeeNumber = row.getCell(employeeColumnNumber).text.trim();

        if (employeeNumber) {
          courseEmployeeNumbers.add(employeeNumber);
        }
      }

      for (const employeeNumber of courseEmployeeNumbers) {
        const userData = users.get(employeeNumber);

        if (!userData || userData.talentLmsUserId === null) {
          continue;
        }

        for (const unit of courseData.iltUnits) {
          unitProgressJobs.push({
            employeeNumber,
            userData,
            userId: userData.talentLmsUserId,
            courseId: courseData.talentLmsId,
            courseName,
            unitId: unit.id,
            unitName: unit.name,
          });
        }
      }
    }

    console.log("TOTAL UNIT PROGRESS JOBS:", unitProgressJobs.length);

    const jobsToProcess = IS_DEV_TEST
      ? unitProgressJobs.slice(0, DEV_PROGRESS_LIMIT)
      : unitProgressJobs;

    console.log(
      IS_DEV_TEST
        ? `DEV MODE: ${jobsToProcess.length} de ${unitProgressJobs.length} jobs`
        : `PRODUCTION MODE: ${jobsToProcess.length} jobs`,
    );

    let checkedUnitProgress = 0;

    const UNIT_PROGRESS_BATCH_SIZE = 2;

    for (let i = 0; i < jobsToProcess.length; i += UNIT_PROGRESS_BATCH_SIZE) {
      const batch = jobsToProcess.slice(i, i + UNIT_PROGRESS_BATCH_SIZE);
      const results = await Promise.all(
        batch.map(async (job) => {
          const progress = await getUnitProgress(
            talentLmsConfig,
            job.userId,
            job.unitId,
          );

          return {
            job,
            progress,
          };
        }),
      );

      for (const item of results) {
        item.job.userData.iltProgress.push({
          courseId: item.job.courseId,
          courseName: item.job.courseName,
          unitId: item.job.unitId,
          unitName: item.job.unitName,
          status: item.progress?.status ?? "not_enrolled",
          totalTime: item.progress?.totalTime ?? 0,
        });

        checkedUnitProgress++;

        event.sender.send("analysis:progress", {
          stage: "checkingUnitProgress",
          current: checkedUnitProgress,
          total: jobsToProcess.length,
        });
      }

      await new Promise<void>((resolve) => setTimeout(resolve, 250));
    }

    let completedCount = 0;
    let failedCount = 0;
    let notAttemptedCount = 0;
    let notEnrolledCount = 0;

    for (const userData of users.values()) {
      for (const progress of userData.iltProgress) {
        if (progress.status === "completed") {
          completedCount++;
        } else if (progress.status === "failed") {
          failedCount++;
        } else if (progress.status === "not_attempted") {
          notAttemptedCount++;
        } else if (progress.status === "not_enrolled") {
          notEnrolledCount++;
        }
      }
    }

    const notCheckedCount = Math.max(
      unitProgressJobs.length - jobsToProcess.length,
      0,
    );

    const courseProgressDetails: CourseProgressDetails[] = [];

    for (const [courseName, courseData] of courses) {
      if (courseData.isIlt !== true || courseData.talentLmsId === null) {
        continue;
      }

      const courseEmployeeNumbers = new Set<string>();

      for (const rowNumber of courseData.rows) {
        const row = worksheet.getRow(rowNumber);
        const employeeNumber = row.getCell(employeeColumnNumber).text.trim();

        if (employeeNumber) {
          courseEmployeeNumbers.add(employeeNumber);
        }
      }

      const logicalUnitsMap = new Map<
        string,
        { name: string; physicalUnits: IltUnitData[] }
      >();

      for (const unit of courseData.iltUnits) {
        const logicalKey = normalizeUnitName(unit.name);

        const existing = logicalUnitsMap.get(logicalKey);

        if (existing) {
          existing.physicalUnits.push(unit);
        } else {
          logicalUnitsMap.set(logicalKey, {
            name: unit.name.trim(),
            physicalUnits: [unit],
          });
        }
      }

      console.log(
        `Curso ${courseName}: ${courseData.iltUnits.length} físicas -> ${logicalUnitsMap.size} lógicas`,
      );

      const units: CourseProgressUnit[] = [];

      for (const logicalUnit of logicalUnitsMap.values()) {
        const unitUsers: CourseProgressUser[] = [];
        const physicalUnitIds = new Set(
          logicalUnit.physicalUnits.map((unit) => unit.id),
        );

        for (const employeeNumber of courseEmployeeNumbers) {
          const userData = users.get(employeeNumber);

          if (!userData) {
            continue;
          }

          const matchingProgress = userData.iltProgress.filter(
            (progress) =>
              progress.courseId === courseData.talentLmsId &&
              physicalUnitIds.has(progress.unitId),
          );

          let status = "not_checked";

          let totalTime = 0;

          if (userData.talentLmsUserId === null) {
            status = "user_not_found";
          } else if (matchingProgress.length > 0) {
            const completed = matchingProgress.filter(
              (progress) => progress.status === "completed",
            );

            if (completed.length > 0) {
              status = "completed";

              totalTime = completed.reduce(
                (total, progress) => total + progress.totalTime,
                0,
              );
            } else {
              const pending = matchingProgress.find(
                (progress) => progress.status === "pending",
              );

              const failed = matchingProgress.find(
                (progress) => progress.status === "failed",
              );

              const notAttempted = matchingProgress.find(
                (progress) => progress.status === "not_attempted",
              );

              const notEnrolled = matchingProgress.find(
                (progress) => progress.status === "not_enrolled",
              );

              if (pending) {
                status = "pending";

                totalTime = pending.totalTime;
              } else if (failed) {
                status = "failed";

                totalTime = failed.totalTime;
              } else if (notAttempted) {
                status = "not_attempted";

                totalTime = 0;
              } else if (notEnrolled) {
                status = "not_enrolled";

                totalTime = 0;
              }
            }
          }

          unitUsers.push({
            employeeNumber,
            firstName: userData.firstName,
            lastName: userData.lastName,
            talentLmsUserId: userData.talentLmsUserId,
            status,
            totalTime,
          });
        }

        const representativeUnit = logicalUnit.physicalUnits[0];

        units.push({
          unitId: representativeUnit.id,
          unitName: logicalUnit.name,
          users: unitUsers,
        });
      }

      courseProgressDetails.push({
        courseId: courseData.talentLmsId,
        courseName,

        units,
      });
    }

    const iltHoursByEmployeeAndCourse = new Map<string, number>();
    const iltStatusesByEmployeeAndCourse = new Map<string, string[]>();

    for (const course of courseProgressDetails) {
      const employeeTotals = new Map<string, number>();
      const employeeStatuses = new Map<string, string[]>();

      for (const unit of course.units) {
        for (const user of unit.users) {
          const currentTotal = employeeTotals.get(user.employeeNumber) ?? 0;

          employeeTotals.set(
            user.employeeNumber,
            currentTotal + user.totalTime,
          );

          const statuses = employeeStatuses.get(user.employeeNumber) ?? [];

          statuses.push(user.status);
          employeeStatuses.set(user.employeeNumber, statuses);
        }
      }

      for (const [employeeNumber, totalSeconds] of employeeTotals) {
        const key = `${employeeNumber}::${normalizeCourseName(
          course.courseName,
        )}`;

        iltHoursByEmployeeAndCourse.set(key, totalSeconds / 3600);

        iltStatusesByEmployeeAndCourse.set(
          key,
          employeeStatuses.get(employeeNumber) ?? [],
        );
      }
    }

    const iltColumnNumber = durationColumnNumber + 1;
    const headerRow = worksheet.getRow(HEADER_ROW);
    const durationHeaderCell = headerRow.getCell(durationColumnNumber);
    const iltHeaderCell = headerRow.getCell(iltColumnNumber);

    iltHeaderCell.style = {
      ...durationHeaderCell.style,
    };

    iltHeaderCell.alignment = {
      ...durationHeaderCell.alignment,
    };

    iltHeaderCell.value = "ILT (horas)";

    worksheet.getColumn(iltColumnNumber).width = 14;

    const ILT_COLORS = {
      green: "FFC6EFCE",
      yellow: "FFFFEB9C",
      orange: "FFF4B183",
      red: "FFFFC7CE",
    };

    const totalExcelRows = worksheet.actualRowCount - DATA_START_ROW + 1;

    let writtenExcelRows = 0;

    event.sender.send("analysis:progress", {
      stage: "writingExcel",
      current: 0,
      total: totalExcelRows,
    });

    const updateExcelProgress = async (): Promise<void> => {
      writtenExcelRows++;

      if (writtenExcelRows % 100 === 0 || writtenExcelRows === totalExcelRows) {
        event.sender.send("analysis:progress", {
          stage: "writingExcel",
          current: writtenExcelRows,
          total: totalExcelRows,
        });

        await new Promise<void>((resolve) => setImmediate(resolve));
      }
    };

    for (
      let rowNumber = DATA_START_ROW;
      rowNumber <= worksheet.actualRowCount;
      rowNumber++
    ) {
      const row = worksheet.getRow(rowNumber);

      const employeeNumber = row.getCell(employeeColumnNumber).text.trim();

      const courseName = row.getCell(courseColumnNumber).text.trim();

      if (!employeeNumber || !courseName) {
        await updateExcelProgress();
        continue;
      }

      const key = `${employeeNumber}::${normalizeCourseName(courseName)}`;

      const iltHours = iltHoursByEmployeeAndCourse.get(key);

      const iltCell = row.getCell(iltColumnNumber);

      if (iltHours === undefined) {
        iltCell.value = null;
        await updateExcelProgress();
        continue;
      }

      const hours = Math.floor(iltHours);
      const fractionalPart = iltHours - hours;

      const minutes = Math.round(fractionalPart * 60);

      const hourAndMinuteString = `${hours}h ${minutes}m`;

      iltCell.value = hourAndMinuteString;
      iltCell.numFmt = "0.00";

      const completionDate = row
        .getCell(completionDateColumnNumber)
        .text.trim();

      const courseCompleted = hasCompletionDate(completionDate);

      const statuses = iltStatusesByEmployeeAndCourse.get(key) ?? [];

      const hasPending = statuses.includes("pending");

      const attendedAnyIlt =
        statuses.includes("completed") || statuses.includes("failed");

      let fillColor: string;

      if (hasPending) {
        fillColor = ILT_COLORS.yellow;
      } else if (courseCompleted) {
        fillColor = ILT_COLORS.green;
      } else if (attendedAnyIlt) {
        fillColor = ILT_COLORS.orange;
      } else {
        fillColor = ILT_COLORS.red;
      }

      iltCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: fillColor,
        },
      };

      await updateExcelProgress();
    }

    const parsedPath = path.parse(filePath);
    const timestamp = Date.now();
    const randomNumber = Math.floor(1000 + Math.random() * 9000);

    const outputPath = path.join(
      parsedPath.dir,
      `${parsedPath.name}_ILT_${timestamp}_${randomNumber}${parsedPath.ext}`,
    );

    event.sender.send("analysis:progress", {
      stage: "savingExcel",
      current: 0,
      total: 0,
    });

    await workbook.xlsx.writeFile(outputPath);

    event.sender.send("analysis:progress", {
      stage: "removingStrikeThrough",
      current: 0,
      total: 0,
    });

    console.log("A remover strikethrough de xl/styles.xml...");

    await removeStrikethroughFromWorkbook(outputPath);

    console.log("STRIKETHROUGH REMOVIDO.");
    console.log("RELATÓRIO ILT CRIADO:", outputPath);

    const unmatchedCourses: string[] = [];

    let matchedCourses = 0;

    for (const [courseName, courseData] of courses) {
      if (courseData.talentLmsId !== null) {
        matchedCourses++;
      } else {
        unmatchedCourses.push(courseName);
      }
    }

    currentAnalysis = {
      filePath,
      workbook,
      worksheet,
      courses,
      users,
    };

    return {
      cancelled: false,
      summary: {
        fileName: path.basename(filePath),
        sheetName: worksheet.name,
        totalRows: Math.max(worksheet.actualRowCount - DATA_START_ROW + 1, 0),
        totalColumns: columns.length + 1,
        columns: [...columns, "ILT (horas)"],
        courses: uniqueCourses,
        matchedCourses,
        unmatchedCourses,
        iltCourses,
        iltCourseDetails,
        courseProgressDetails,
        uniqueUsers: users.size,
        matchedUsers,
        unmatchedUsers,
        progressStats: {
          totalJobs: unitProgressJobs.length,
          processedJobs: jobsToProcess.length,
          completed: completedCount,
          failed: failedCount,
          notAttempted: notAttemptedCount,
          notEnrolled: notEnrolledCount,
          notChecked: notCheckedCount,
        },
      },
    };
  } catch (error) {
    console.error("Erro durante a análise:", error);

    return {
      cancelled: false,
      error:
        error instanceof Error
          ? `Não foi possível concluir a análise: ${error.message}`
          : "Não foi possível concluir a análise.",
    };
  }
});

//Coisas de Electron

app.whenReady().then(() => {
  registerSettingsHandlers();
  createWindow();
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
