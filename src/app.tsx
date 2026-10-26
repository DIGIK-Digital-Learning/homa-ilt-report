import React, { useEffect, useMemo, useState } from "react";

import { createRoot } from "react-dom/client";

type ActiveView = "summary" | "courses";

const DEFAULT_BASE_URL = "https://homaacademy.talentlms.com";

function App() {
  const [summary, setSummary] = useState<ExcelSummary | null>(null);
  const [analysisProgress, setAnalysisProgress] =
    useState<AnalysisProgress | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [activeView, setActiveView] = useState<ActiveView>("summary");
  const [openCourses, setOpenCourses] = useState<Set<number>>(new Set());
  const [openUnits, setOpenUnits] = useState<Set<string>>(new Set());
  const [baseUrlInput, setBaseUrlInput] = useState(DEFAULT_BASE_URL);
  const [savedBaseUrl, setSavedBaseUrl] = useState(DEFAULT_BASE_URL);
  const [apiKeyInput, setApiKeyInput] = useState("");
  const [showApiKey, setShowApiKey] = useState(false);
  const [hasSavedApiKey, setHasSavedApiKey] = useState(false);
  const [isEditingSettings, setIsEditingSettings] = useState(false);
  const [isSavingSettings, setIsSavingSettings] = useState(false);
  const [settingsMessage, setSettingsMessage] = useState<string | null>(null);

  const progressPercentage =
    analysisProgress && analysisProgress.total > 0
      ? Math.round((analysisProgress.current / analysisProgress.total) * 100)
      : 0;

  useEffect(() => {
    window.electronAPI.onAnalysisProgress((progress) => {
      setAnalysisProgress(progress);
    });

    const loadSettings = async (): Promise<void> => {
      try {
        const result = await window.electronAPI.getSettings();

        if (!result.success) {
          setSettingsMessage(
            result.error ?? "Não foi possível carregar a configuração.",
          );

          return;
        }

        setBaseUrlInput(result.baseUrl);
        setSavedBaseUrl(result.baseUrl);
        setHasSavedApiKey(result.hasApiKey);

        setIsEditingSettings(!result.hasApiKey);
      } catch (loadError) {
        console.error(loadError);

        setSettingsMessage("Não foi possível carregar a configuração.");
        setIsEditingSettings(true);
      }
    };

    void loadSettings();
  }, []);

  const handleSaveSettings = async (): Promise<void> => {
    const trimmedBaseUrl = baseUrlInput.trim();
    const trimmedApiKey = apiKeyInput.trim();

    if (!trimmedBaseUrl) {
      setSettingsMessage("Introduz a Base URL.");

      return;
    }

    if (!hasSavedApiKey && !trimmedApiKey) {
      setSettingsMessage("Introduz a API Key.");
      return;
    }

    setIsSavingSettings(true);
    setSettingsMessage(null);

    try {
      const result = await window.electronAPI.saveSettings({
        baseUrl: trimmedBaseUrl,
        apiKey: trimmedApiKey || undefined,
      });

      if (!result.success) {
        setSettingsMessage(
          result.error ?? "Não foi possível guardar a configuração.",
        );

        return;
      }

      const newBaseUrl = result.baseUrl ?? trimmedBaseUrl;

      setBaseUrlInput(newBaseUrl);
      setSavedBaseUrl(newBaseUrl);
      setHasSavedApiKey(Boolean(result.hasApiKey));
      setApiKeyInput("");
      setShowApiKey(false);
      setSettingsMessage(null);

      setIsEditingSettings(false);
    } catch (saveError) {
      console.error(saveError);

      setSettingsMessage("Ocorreu um erro ao guardar a configuração.");
    } finally {
      setIsSavingSettings(false);
    }
  };

  const handleCancelSettings = (): void => {
    setBaseUrlInput(savedBaseUrl);
    setApiKeyInput("");
    setShowApiKey(false);
    setSettingsMessage(null);
    setIsEditingSettings(false);
  };

  const handleSelectExcel = async (): Promise<void> => {
    if (!hasSavedApiKey) {
      setError("Configura primeiro a TalentLMS.");

      return;
    }

    setIsLoading(true);
    setError(null);
    setSummary(null);
    setActiveView("summary");
    setAnalysisProgress(null);
    setOpenCourses(new Set());
    setOpenUnits(new Set());

    try {
      const result = await window.electronAPI.selectExcel();

      if (result.cancelled) {
        return;
      }

      if (result.error) {
        setError(result.error);

        return;
      }

      if (result.summary) {
        setSummary(result.summary);
      }
    } catch (selectionError) {
      console.error(selectionError);

      setError("Ocorreu um erro inesperado ao selecionar o relatório.");
    } finally {
      setIsLoading(false);
    }
  };

  const getProgressDetail = (): string | null => {
    if (!analysisProgress || analysisProgress.total === 0) {
      return null;
    }

    switch (analysisProgress.stage) {
      case "processingExcel":
        return `${analysisProgress.current.toLocaleString(
          "pt-PT",
        )} de ${analysisProgress.total.toLocaleString("pt-PT")} linhas`;

      case "matchingCourses":
        return `${analysisProgress.current.toLocaleString(
          "pt-PT",
        )} de ${analysisProgress.total.toLocaleString("pt-PT")} cursos`;

      case "checkingIlt":
        return `${analysisProgress.current.toLocaleString(
          "pt-PT",
        )} de ${analysisProgress.total.toLocaleString("pt-PT")} cursos`;

      case "fetchingIltUnits":
        return `${analysisProgress.current.toLocaleString(
          "pt-PT",
        )} de ${analysisProgress.total.toLocaleString("pt-PT")} cursos ILT`;

      case "checkingUsers":
        return `${analysisProgress.current.toLocaleString(
          "pt-PT",
        )} de ${analysisProgress.total.toLocaleString("pt-PT")} registos`;

      case "matchingUsers":
        return `${analysisProgress.current.toLocaleString(
          "pt-PT",
        )} de ${analysisProgress.total.toLocaleString("pt-PT")} colaboradores`;

      case "checkingUnitProgress":
        return `${analysisProgress.current.toLocaleString(
          "pt-PT",
        )} de ${analysisProgress.total.toLocaleString("pt-PT")}`;

      case "writingExcel":
        return `${analysisProgress.current.toLocaleString(
          "pt-PT",
        )} de ${analysisProgress.total.toLocaleString("pt-PT")} linhas`;

      default:
        return null;
    }
  };

  const getProgressText = (): string => {
    switch (analysisProgress?.stage) {
      case "reading":
        return "A abrir o Excel...";
      case "processingExcel":
        return "A analisar o relatório...";
      case "fetchingCourses":
        return "A obter cursos...";
      case "matchingCourses":
        return "A associar cursos...";
      case "checkingIlt":
        return "A identificar cursos ILT...";
      case "fetchingIltUnits":
        return "A obter unidades ILT...";
      case "checkingUsers":
        return "A identificar colaboradores...";
      case "matchingUsers":
        return "A associar colaboradores...";
      case "checkingUnitProgress":
        return "A verificar progresso ILT...";
      case "writingExcel":
        return "A preencher o relatório...";
      case "savingExcel":
        return "A guardar o Excel...";
      case "removingStrikeThrough":
        return "A finalizar o Excel...";
      default:
        return "A preparar análise...";
    }
  };

  const getStatusLabel = (status: string): string => {
    switch (status) {
      case "completed":
        return "Concluído";
      case "failed":
        return "Não concluído";
      case "pending":
        return "Pendente";
      case "not_attempted":
        return "Não iniciado";
      case "not_enrolled":
        return "Sem inscrição";
      case "not_checked":
        return "Não verificado";
      case "user_not_found":
        return "Não encontrado";
      default:
        return status;
    }
  };

  const formatTime = (seconds: number): string => {
    if (!seconds || seconds <= 0) {
      return "—";
    }

    const hours = Math.floor(seconds / 3600);
    const minutes = Math.round((seconds % 3600) / 60);

    if (hours > 0 && minutes > 0) {
      return `${hours}h ${minutes}m`;
    }

    if (hours > 0) {
      return `${hours}h`;
    }

    return `${minutes}m`;
  };

  const getFullName = (user: CourseProgressUser): string => {
    const fullName = `${user.firstName} ${user.lastName}`.trim();

    return fullName || `Colaborador ${user.employeeNumber}`;
  };

  const getStatusCounts = (users: CourseProgressUser[]) => {
    const counts = {
      completed: 0,
      failed: 0,
      pending: 0,
      notAttempted: 0,
      notEnrolled: 0,
      notChecked: 0,
      userNotFound: 0,
    };

    for (const user of users) {
      switch (user.status) {
        case "completed":
          counts.completed++;
          break;
        case "failed":
          counts.failed++;
          break;
        case "pending":
          counts.pending++;
          break;
        case "not_attempted":
          counts.notAttempted++;
          break;
        case "not_enrolled":
          counts.notEnrolled++;
          break;
        case "user_not_found":
          counts.userNotFound++;
          break;
        default:
          counts.notChecked++;
          break;
      }
    }

    return counts;
  };

  const toggleCourse = (courseId: number) => {
    setOpenCourses((current) => {
      const next = new Set(current);

      if (next.has(courseId)) {
        next.delete(courseId);
      } else {
        next.add(courseId);
      }

      return next;
    });
  };

  const toggleUnit = (courseId: number, unitId: number) => {
    const key = `${courseId}-${unitId}`;

    setOpenUnits((current) => {
      const next = new Set(current);

      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }

      return next;
    });
  };

  const totalIltUnits = useMemo(() => {
    if (!summary) {
      return 0;
    }

    return summary.courseProgressDetails.reduce(
      (total, course) => total + course.units.length,
      0,
    );
  }, [summary]);

  return (
    <main className="app">
      <header className="header">
        <span className="product-label">TalentLMS</span>

        <h1>Análise de formação ILT</h1>
      </header>
      <section className="settings-panel">
        {!isEditingSettings && hasSavedApiKey ? (
          <div className="settings-compact">
            <div className="settings-compact-info">
              <div>
                <strong>TalentLMS</strong>
                <span>{savedBaseUrl.replace(/^https?:\/\//u, "")}</span>
              </div>
              <span className="api-status saved">Configurado</span>
            </div>
            <button
              type="button"
              className="text-button"
              onClick={() => {
                setSettingsMessage(null);
                setIsEditingSettings(true);
              }}
            >
              Alterar
            </button>
          </div>
        ) : (
          <>
            <div className="settings-title-row">
              <strong>Configuração TalentLMS</strong>
              {hasSavedApiKey && (
                <button
                  type="button"
                  className="text-button"
                  onClick={handleCancelSettings}
                >
                  Cancelar
                </button>
              )}
            </div>
            <div className="settings-form">
              <div className="setting-field">
                <label htmlFor="talentlms-base-url">Base URL</label>
                <input
                  id="talentlms-base-url"
                  type="url"
                  value={baseUrlInput}
                  onChange={(event) => {
                    setBaseUrlInput(event.target.value);
                    setSettingsMessage(null);
                  }}
                  placeholder="https://empresa.talentlms.com"
                  spellCheck={false}
                />
              </div>
              <div className="setting-field">
                <label htmlFor="talentlms-api-key">API Key</label>
                <div className="api-key-input-wrap">
                  <input
                    id="talentlms-api-key"
                    type={showApiKey ? "text" : "password"}
                    value={apiKeyInput}
                    onChange={(event) => {
                      setApiKeyInput(event.target.value);
                      setSettingsMessage(null);
                    }}
                    placeholder={
                      hasSavedApiKey
                        ? "Deixar vazio para manter"
                        : "Introduzir API Key"
                    }
                    autoComplete="off"
                    spellCheck={false}
                  />
                  <button
                    type="button"
                    className="api-key-toggle"
                    onClick={() => setShowApiKey((current) => !current)}
                    disabled={apiKeyInput.length === 0}
                  >
                    {showApiKey ? "Ocultar" : "Mostrar"}
                  </button>
                </div>
              </div>
              <button
                type="button"
                className="primary-button settings-save-button"
                onClick={handleSaveSettings}
                disabled={
                  isSavingSettings ||
                  !baseUrlInput.trim() ||
                  (!hasSavedApiKey && !apiKeyInput.trim())
                }
              >
                {isSavingSettings ? "A guardar..." : "Guardar"}
              </button>
            </div>
            {settingsMessage && (
              <span className="settings-error">{settingsMessage}</span>
            )}
          </>
        )}
      </section>
      <section className="import-panel">
        <div className="excel-icon">XLSX</div>

        <div className="import-copy">
          <strong>
            {summary ? "Relatório analisado" : "Importar relatório"}
          </strong>
          {summary && <span>{summary.fileName}</span>}
        </div>
        <button
          type="button"
          className="primary-button"
          onClick={handleSelectExcel}
          disabled={isLoading || !hasSavedApiKey || isEditingSettings}
        >
          {!hasSavedApiKey
            ? "Configura primeiro"
            : isLoading
              ? "A analisar..."
              : summary
                ? "Novo relatório"
                : "Selecionar Excel"}
        </button>
      </section>
      {isLoading && analysisProgress && (
        <section className="progress-panel">
          <div className="progress-copy">
            <strong>{getProgressText()}</strong>
            {getProgressDetail() && <span>{getProgressDetail()}</span>}
          </div>
          {analysisProgress.total > 0 && (
            <div className="progress-row">
              <progress
                value={analysisProgress.current}
                max={analysisProgress.total}
              />
              <strong>{progressPercentage}%</strong>
            </div>
          )}
        </section>
      )}
      {error && (
        <section className="message error-message">
          <strong>Não foi possível concluir a análise.</strong>
          <span>{error}</span>
        </section>
      )}
      {summary && (
        <section className="summary-panel">
          <div className="summary-file">
            <h2>{summary.fileName}</h2>
          </div>
          <nav className="view-tabs" aria-label="Vistas da análise">
            <button
              type="button"
              className={activeView === "summary" ? "active" : ""}
              onClick={() => setActiveView("summary")}
            >
              Resumo
            </button>
            <button
              type="button"
              className={activeView === "courses" ? "active" : ""}
              onClick={() => setActiveView("courses")}
            >
              Cursos ILT
            </button>
          </nav>
          {activeView === "summary" && (
            <div className="view-content">
              <div className="stats four-columns">
                <div className="stat">
                  <span>Registos</span>
                  <strong>{summary.totalRows.toLocaleString("pt-PT")}</strong>
                </div>
                <div className="stat">
                  <span>Cursos</span>
                  <strong>{summary.courses.length}</strong>
                </div>
                <div className="stat">
                  <span>Cursos ILT</span>
                  <strong>{summary.courseProgressDetails.length}</strong>
                </div>
                <div className="stat">
                  <span>Unidades ILT</span>
                  <strong>{totalIltUnits}</strong>
                </div>
              </div>
              <div className="stats three-columns compact-stats">
                <div className="stat">
                  <span>Colaboradores ILT</span>
                  <strong>{summary.uniqueUsers.toLocaleString("pt-PT")}</strong>
                </div>
                <div className="stat">
                  <span>Associados</span>
                  <strong>
                    {summary.matchedUsers.toLocaleString("pt-PT")}
                  </strong>
                </div>
                <div className="stat">
                  <span>Não encontrados</span>
                  <strong>{summary.unmatchedUsers.length}</strong>
                </div>
              </div>
              {summary.progressStats.processedJobs <
                summary.progressStats.totalJobs && (
                <div className="message info-message">
                  <strong>Modo de teste ativo</strong>
                  <span>
                    Foram verificadas {summary.progressStats.processedJobs} de{" "}
                    {summary.progressStats.totalJobs.toLocaleString("pt-PT")}{" "}
                    verificações.
                  </span>
                </div>
              )}
              {summary.unmatchedCourses.length > 0 && (
                <div className="message warning-message">
                  <strong>
                    {summary.unmatchedCourses.length === 1
                      ? "1 curso não foi encontrado."
                      : `${summary.unmatchedCourses.length} cursos não foram encontrados.`}
                  </strong>
                  <ul>
                    {summary.unmatchedCourses.map((course) => (
                      <li key={course}>{course}</li>
                    ))}
                  </ul>
                </div>
              )}
            </div>
          )}
          {activeView === "courses" && (
            <div className="view-content">
              <div className="view-heading">
                <h3>Cursos ILT</h3>
                <span className="count-badge">
                  {summary.courseProgressDetails.length} cursos
                </span>
              </div>
              {summary.progressStats.notChecked > 0 && (
                <div className="message info-message slim-message">
                  <span>Existem verificações ainda não processadas.</span>
                </div>
              )}
              <div className="course-list">
                {summary.courseProgressDetails.map((course) => {
                  const isCourseOpen = openCourses.has(course.courseId);
                  const participantNumbers = new Set(
                    course.units.flatMap((unit) =>
                      unit.users.map((user) => user.employeeNumber),
                    ),
                  );
                  const allCourseUsers = course.units.flatMap(
                    (unit) => unit.users,
                  );
                  const courseCounts = getStatusCounts(allCourseUsers);
                  return (
                    <article className="course-card" key={course.courseId}>
                      <button
                        type="button"
                        className="course-toggle"
                        onClick={() => toggleCourse(course.courseId)}
                        aria-expanded={isCourseOpen}
                      >
                        <div className="course-main-info">
                          <span className="course-tag">ILT</span>
                          <div>
                            <h4>{course.courseName}</h4>
                            <span className="course-meta">
                              {course.units.length}{" "}
                              {course.units.length === 1
                                ? "unidade"
                                : "unidades"}
                              <span aria-hidden="true">•</span>
                              {participantNumbers.size.toLocaleString(
                                "pt-PT",
                              )}{" "}
                              colaboradores
                            </span>
                          </div>
                        </div>
                        <div className="course-summary-side">
                          <div className="mini-statuses">
                            {courseCounts.completed > 0 && (
                              <span className="mini-status completed">
                                {courseCounts.completed} concluídos
                              </span>
                            )}
                            {courseCounts.notAttempted > 0 && (
                              <span className="mini-status pending">
                                {courseCounts.notAttempted} não iniciados
                              </span>
                            )}
                          </div>
                          <span
                            className={`chevron ${isCourseOpen ? "open" : ""}`}
                          >
                            ›
                          </span>
                        </div>
                      </button>
                      {isCourseOpen && (
                        <div className="course-body">
                          {course.units.length === 0 ? (
                            <div className="empty-state">
                              Não foram encontradas unidades ILT.
                            </div>
                          ) : (
                            course.units.map((unit) => {
                              const unitKey = `${course.courseId}-${unit.unitId}`;
                              const isUnitOpen = openUnits.has(unitKey);
                              const counts = getStatusCounts(unit.users);
                              return (
                                <section
                                  className="unit-card"
                                  key={unit.unitId}
                                >
                                  <button
                                    type="button"
                                    className="unit-toggle"
                                    onClick={() =>
                                      toggleUnit(course.courseId, unit.unitId)
                                    }
                                    aria-expanded={isUnitOpen}
                                  >
                                    <div className="unit-title-wrap">
                                      <span className="unit-marker" />
                                      <div>
                                        <strong>{unit.unitName}</strong>
                                        <span>
                                          {unit.users.length.toLocaleString(
                                            "pt-PT",
                                          )}{" "}
                                          colaboradores
                                        </span>
                                      </div>
                                    </div>
                                    <div className="unit-summary">
                                      <span className="unit-completed">
                                        {counts.completed} concluídos
                                      </span>
                                      <span
                                        className={`chevron small ${
                                          isUnitOpen ? "open" : ""
                                        }`}
                                      >
                                        ›
                                      </span>
                                    </div>
                                  </button>
                                  {isUnitOpen && (
                                    <div className="participant-table-wrap">
                                      <div className="participant-table-header">
                                        <span>Colaborador</span>
                                        <span>Estado</span>
                                        <span>Tempo</span>
                                      </div>
                                      <div className="participant-list">
                                        {unit.users.map((user) => (
                                          <div
                                            className="participant-row"
                                            key={`${unit.unitId}-${user.employeeNumber}`}
                                          >
                                            <div className="participant-person">
                                              <span className="person-avatar">
                                                {(
                                                  user.firstName ||
                                                  user.employeeNumber
                                                )
                                                  .slice(0, 2)
                                                  .toUpperCase()}
                                              </span>
                                              <div>
                                                <strong>
                                                  {getFullName(user)}
                                                </strong>
                                                <span>
                                                  #{user.employeeNumber}
                                                  {user.talentLmsUserId !==
                                                    null &&
                                                    ` · #${user.talentLmsUserId}`}
                                                </span>
                                              </div>
                                            </div>
                                            <div>
                                              <span
                                                className={`status-badge status-${user.status}`}
                                              >
                                                {getStatusLabel(user.status)}
                                              </span>
                                            </div>
                                            <span className="time-value">
                                              {formatTime(user.totalTime)}
                                            </span>
                                          </div>
                                        ))}
                                      </div>
                                    </div>
                                  )}
                                </section>
                              );
                            })
                          )}
                        </div>
                      )}
                    </article>
                  );
                })}
              </div>
            </div>
          )}
        </section>
      )}
    </main>
  );
}

const rootElement = document.getElementById("root");

if (!rootElement) {
  throw new Error("Não foi encontrado o elemento #root");
}

createRoot(rootElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
