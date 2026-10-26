interface TalentLMSConfig {
  baseUrl: string;
  apiKey: string;
  apiVersion: string;
}

export interface TalentLMSCourse {
  id: number;
  name: string;
}

interface CoursesResponse {
  _data: TalentLMSCourse[];
  _meta?: {
    pagination?: {
      page: number;
      page_size: number;
      total_pages: number;
      total_results: number;
    };
  };
}

const MAX_RETRIES = 6;
const BASE_RETRY_DELAY_MS = 2000;

let rateLimitBlockedUntil = 0;

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getRetryDelay(response: Response, attempt: number): number {
  const retryAfter = response.headers.get("retry-after");

  if (retryAfter) {
    const seconds = Number(retryAfter);

    if (Number.isFinite(seconds)) {
      return Math.max(seconds * 1000, 1000);
    }

    const retryDate = Date.parse(retryAfter);

    if (!Number.isNaN(retryDate)) {
      return Math.max(retryDate - Date.now(), 1000);
    }
  }

  const exponentialDelay = BASE_RETRY_DELAY_MS * 2 ** attempt;

  const jitter = Math.floor(Math.random() * 500);

  return Math.min(exponentialDelay + jitter, 30000);
}

async function talentLmsFetch(
  config: TalentLMSConfig,
  url: URL,
): Promise<Response> {
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    const remainingCooldown = rateLimitBlockedUntil - Date.now();

    if (remainingCooldown > 0) {
      await sleep(remainingCooldown);
    }

    const response = await fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        "X-API-Key": config.apiKey,
        "X-API-Version": config.apiVersion,
      },
    });

    if (response.status !== 429) {
      return response;
    }

    if (attempt === MAX_RETRIES) {
      const errorBody = await response.text();

      throw new Error(
        `TalentLMS rate limit excedido após várias tentativas: ${errorBody}`,
      );
    }

    const waitTime = getRetryDelay(response, attempt);

    rateLimitBlockedUntil = Math.max(
      rateLimitBlockedUntil,
      Date.now() + waitTime,
    );

    console.log(
      `TalentLMS rate limit. A aguardar ${Math.round(
        waitTime / 1000,
      )}s antes de repetir...`,
    );

    await sleep(waitTime);
  }

  throw new Error("Não foi possível concluir o pedido à TalentLMS.");
}

export async function getAllCourses(
  config: TalentLMSConfig,
): Promise<TalentLMSCourse[]> {
  const allCourses: TalentLMSCourse[] = [];

  let page = 1;

  // eslint-disable-next-line no-constant-condition
  while (true) {
    const baseUrl = config.baseUrl.replace(/\/+$/, "");

    const url = new URL(`${baseUrl}/api/v2/courses`);

    url.searchParams.set("page[number]", String(page));

    url.searchParams.set("page[size]", "100");

    const response = await talentLmsFetch(config, url);

    if (!response.ok) {
      const errorBody = await response.text();

      throw new Error(
        `TalentLMS respondeu com ${response.status}: ${errorBody}`,
      );
    }

    const data = (await response.json()) as CoursesResponse;

    allCourses.push(...data._data);

    const totalPages = data._meta?.pagination?.total_pages ?? 1;

    if (page >= totalPages) {
      break;
    }

    page++;
  }

  return allCourses;
}

export async function isCourseIlt(
  config: TalentLMSConfig,
  courseId: number,
): Promise<boolean> {
  const baseUrl = config.baseUrl.replace(/\/+$/, "");

  const url = new URL(`${baseUrl}/api/v2/courses/${courseId}`);

  const response = await talentLmsFetch(config, url);

  if (!response.ok) {
    const errorBody = await response.text();

    throw new Error(`TalentLMS respondeu com ${response.status}: ${errorBody}`);
  }

  const data = await response.json();

  const iltField = data._data.custom_fields?.find(
    (field: { name: string; value: string }) => field.name === "ILT",
  );

  return iltField?.value === "on";
}

export async function getUserIdByEmployeeNumber(
  config: TalentLMSConfig,
  employeeNumber: string,
): Promise<number | null> {
  const baseUrl = config.baseUrl.replace(/\/+$/, "");

  const url = new URL(`${baseUrl}/api/v2/users`);

  url.searchParams.set("filter[custom_field_value][eq]", employeeNumber);

  const response = await talentLmsFetch(config, url);

  if (!response.ok) {
    const errorBody = await response.text();

    throw new Error(`TalentLMS respondeu com ${response.status}: ${errorBody}`);
  }

  const data = await response.json();

  if (!data._data || data._data.length === 0) {
    return null;
  }

  return data._data[0].id;
}

export interface IltUnitData {
  id: number;
  name: string;
}

interface UnitsResponse {
  _data: {
    id: number;
    name: string;
    type: string;
  }[];

  _meta?: {
    pagination?: {
      page: number;
      page_size: number;
      total_pages: number;
      total_results: number;
    };
  };
}

export async function getCourseIltUnits(
  config: TalentLMSConfig,
  courseId: number,
): Promise<IltUnitData[]> {
  const allIltUnits: IltUnitData[] = [];

  const baseUrl = config.baseUrl.replace(/\/+$/, "");

  let page = 1;

  const PAGE_SIZE = 50;

  // eslint-disable-next-line no-constant-condition
  while (true) {
    const url = new URL(`${baseUrl}/api/v2/courses/${courseId}/units`);

    url.searchParams.set("page[number]", String(page));

    url.searchParams.set("page[size]", String(PAGE_SIZE));

    const response = await talentLmsFetch(config, url);

    if (!response.ok) {
      const errorBody = await response.text();

      throw new Error(
        `TalentLMS respondeu com ${response.status}: ${errorBody}`,
      );
    }

    const data = (await response.json()) as UnitsResponse;

    const iltUnits = data._data
      .filter((unit) => unit.type === "Instructor-led training")
      .map((unit) => ({
        id: unit.id,
        name: unit.name,
      }));

    allIltUnits.push(...iltUnits);

    const totalPages = data._meta?.pagination?.total_pages ?? 1;

    console.log(
      `Curso ${courseId}: página ${page}/${totalPages} - ${data._data.length} units recebidas, ${iltUnits.length} ILT`,
    );

    if (page >= totalPages) {
      break;
    }

    page++;
  }

  console.log(`Curso ${courseId}: ${allIltUnits.length} ILT units no total`);

  return allIltUnits;
}

interface UnitProgress {
  status: string;
  totalTime: number;
}

export async function getUnitProgress(
  config: TalentLMSConfig,
  userId: number,
  unitId: number,
): Promise<UnitProgress | null> {
  const baseUrl = config.baseUrl.replace(/\/+$/, "");

  const url = new URL(`${baseUrl}/api/v2/unit-progress`);

  url.searchParams.set("user_id", String(userId));

  url.searchParams.set("unit_id", String(unitId));

  const response = await talentLmsFetch(config, url);

  // Aqui 404 é um resultado normal:
  // o utilizador não está inscrito nessa unit.
  if (response.status === 404) {
    return null;
  }

  if (!response.ok) {
    const errorBody = await response.text();

    throw new Error(`TalentLMS respondeu com ${response.status}: ${errorBody}`);
  }

  const data = await response.json();

  return {
    status: data._data.status,
    totalTime: data._data.total_time,
  };
}
