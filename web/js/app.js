// 皮肤病理报告系统 - Frontend JavaScript
function startApp() {
  // Global state
  const state = {
    initialized: false,
    config: {},
    todayPatients: [],
    selectedPatient: null,
    isEditMode: false,
    imageCount: 1,
    checkNoTimer: null
  };

  // Elements
  const tabs = document.querySelectorAll(".nav-tab-btn");

  // Refresh all active workspace patient lists after sync or rename/rearchive
  function refreshAllPatientListsQuietly() {
    try { if (typeof loadTodayPatients === "function") loadTodayPatients(); } catch (_) {}
    try { if (typeof loadMakingPatients === "function" && makingState && makingState.patients) loadMakingPatients(false); } catch (_) {}
    try { if (typeof loadReviewPatients === "function" && reviewState && reviewState.patients) loadReviewPatients(false); } catch (_) {}
    try { if (typeof loadIhcPatients === "function" && ihcState && ihcState.patients) loadIhcPatients(false); } catch (_) {}
    try {
      const searchTab = document.getElementById("tab-search");
      if (searchTab && searchTab.classList.contains("active") && typeof doSearch === "function") doSearch();
    } catch (_) {}
    try {
      const outTab = document.getElementById("tab-output");
      if (outTab && outTab.classList.contains("active")) {
        if (typeof loadOutputRegularPatients === "function") loadOutputRegularPatients();
        if (typeof loadOutputSpecialPatients === "function") loadOutputSpecialPatients();
        if (typeof loadOutputIhcPatients === "function") loadOutputIhcPatients();
      }
    } catch (_) {}
  }

  // Handle sequential normal / IHC re-archive confirmations after edit or external Word sync
  async function promptRearchiveAfterUpdate(data, initialPatient) {
    let savedPatient = initialPatient;
    if (!data || !savedPatient) return savedPatient;

    if (data.pending_blocked_archived) {
      alert(
        "⚠️ 【提示：待审核/重切患者已标注需重新归档】\n\n" +
        "患者基本信息及草稿文件已同步更新，但由于该报告当前处于【待审核 / 重切】状态，\n" +
        "请在【报告审核】工作区完成专家审核后，再执行重新归档！"
      );
    }

    if (data.ask_rearchive) {
      const wantRearchive = confirm(
        "患者信息及关联文件（除已归档报告外）已同步更新！\n" +
        "是否立即重新归档【普通/原病理报告】？\n\n" +
        "点击【确定】立即自动归档（旧版归档文件将自动移入回收站备份）；\n" +
        "点击【取消】标注为【需要重新归档】。"
      );
      if (wantRearchive) {
        try {
          const archData = await fetchJsonWithSafetyConfirm("/api/reports/archive", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              pathology_no: savedPatient.pathology_no,
              source: data.archive_source || "maker"
            })
          });
          if (archData.success) {
            showToast(archData.message || "报告已自动重新归档！", "success");
            if (archData.duplicate_clinical_photos) {
              alert(archData.duplicate_prompt || "【提示】该患者在临床照片母目录中匹配到多个重名文件夹，请手动核对归档临床照片！");
            }
            if (archData.patient) savedPatient = archData.patient;
          } else if (!archData.cancelled) {
            showToast("重新归档失败: " + (archData.error || "未知错误"), "error");
          }
        } catch (archErr) {
          showToast("重新归档请求异常: " + archErr.message, "error");
        }
      } else {
        showToast("已将原报告标注为【需要重新归档】", "info");
      }
    }

    if (data.ask_ihc_rearchive) {
      const wantIhcRearchive = confirm(
        "检测到该患者同时存在【已归档免疫组化报告】，基本公共信息已同步更新！\n" +
        "是否立即重新归档【免疫组化报告】？\n\n" +
        "点击【确定】立即自动归档免疫组化报告（旧版归档将自动移入回收站备份）；\n" +
        "点击【取消】标注免疫组化为【需要重新归档】。"
      );
      if (wantIhcRearchive) {
        try {
          const ihcArchData = await fetchJsonWithSafetyConfirm("/api/ihc/archive", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ pathology_no: savedPatient.pathology_no })
          });
          if (ihcArchData.success) {
            showToast(ihcArchData.message || "免疫组化报告已自动重新归档！", "success");
            if (ihcArchData.patient) {
              savedPatient = Object.assign({}, savedPatient, { ihc_report: ihcArchData.ihc_report || ihcArchData.patient.ihc_report });
            }
          } else if (!ihcArchData.cancelled) {
            showToast("免疫组化重新归档失败: " + (ihcArchData.error || "未知错误"), "error");
          }
        } catch (ihcErr) {
          showToast("免疫组化重新归档请求异常: " + ihcErr.message, "error");
        }
      } else {
        showToast("已将免疫组化报告标注为【需要重新归档】", "info");
      }
    }

    return savedPatient;
  }

  // Native window.confirm prompt to resolve external Word modifications (forward import or reverse overwrite)
  let _externalSyncPromptActive = false;
  async function resolveExternalSyncForPatient(sourceObj, scopeHint = "all", isWriteConflict = false) {
    if (!sourceObj || _externalSyncPromptActive) return null;
    const pNo = sourceObj.pathology_no || "";
    if (!pNo) return null;

    let files = [];
    if (Array.isArray(sourceObj.modified_files) && sourceObj.modified_files.length) {
      files = sourceObj.modified_files;
    } else {
      const regFiles = Array.isArray(sourceObj.external_modified_files) ? sourceObj.external_modified_files : [];
      const ihcFiles = Array.isArray(sourceObj.ihc_external_modified_files) ? sourceObj.ihc_external_modified_files : [];
      if (scopeHint === "ihc") {
        files = ihcFiles.length ? ihcFiles : regFiles;
      } else if (scopeHint === "normal") {
        files = regFiles;
      } else {
        files = regFiles.concat(ihcFiles);
      }
    }
    if (!files.length) return null;

    _externalSyncPromptActive = true;
    try {
      const primary = sourceObj.primary_source || files[0];
      const fileLines = files.map(f => `• ${f.label || f.role}: ${f.filename || f.path}`).join("\n");
      const nameTxt = sourceObj.name ? `${sourceObj.name} (${pNo})` : pNo;
      const prefixReason = isWriteConflict
        ? "系统准备写入/保存报告前检测到以下 Word 文件已被外部修改，已自动拦截直接覆盖：\n"
        : "检测到以下本地 Word 文档已在系统外修改并保存：\n";

      const wantForward = confirm(
        `⚠️ 【检测到外部 Word 已修改】\n\n` +
        `患者：${nameTxt}\n` +
        `${prefixReason}${fileLines}\n\n` +
        `是否将外部 Word 的最新修改【导入并同步到系统】？\n` +
        `• 点击【确定】导入外部 Word 修改到系统数据库及同级文档\n` +
        `• 点击【取消】暂不导入（可选择反写覆盖或保持现状）`
      );

      if (wantForward) {
        const res = await fetch(`/api/patients/${encodeURIComponent(pNo)}/sync_from_docx`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            action: "forward",
            source_role: primary.role,
            source_path: primary.path
          })
        });
        const syncData = await res.json();
        if (!syncData.success) {
          alert("⚠️ 外部 Word 导入失败：\n\n" + (syncData.error || "未知错误"));
          return { success: false, blocked: true, error: syncData.error };
        }
        showToast(syncData.message || "已导入外部 Word 修改！", "success");
        let syncedP = syncData.patient;
        syncedP = await promptRearchiveAfterUpdate(syncData, syncedP);
        refreshAllPatientListsQuietly();
        return { success: true, action: "forward", patient: syncedP, syncData };
      }

      const wantReverse = confirm(
        `⚠️ 【是否按系统数据覆盖外部 Word？】\n\n` +
        `您刚才取消了“导入外部修改”。\n` +
        `是否将上述外部修改过的文件移入 Windows 回收站备份，并按系统当前记录重新覆盖生成？\n\n` +
        `• 点击【确定】移入回收站并反写覆盖外部 Word\n` +
        `• 点击【取消】${isWriteConflict ? "终止本次写入操作（保留外部 Word 修改）" : "暂不处理，保留当前外部 Word 状态"}`
      );
      if (wantReverse) {
        const revData = await fetchJsonWithSafetyConfirm(`/api/patients/${encodeURIComponent(pNo)}/sync_from_docx`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            action: "reverse",
            source_role: primary.role,
            source_path: primary.path
          })
        });
        if (revData && revData.success) {
          showToast(revData.message || "已按系统数据覆盖外部 Word（旧文件已移入回收站）", "success");
          refreshAllPatientListsQuietly();
          return { success: true, action: "reverse", patient: revData.patient };
        }
        return { success: false, cancelled: true, error: (revData && revData.error) || "反写覆盖失败" };
      }

      return { success: false, cancelled: true };
    } finally {
      _externalSyncPromptActive = false;
    }
  }

  // Unified JSON API fetch with interactive confirmation for permanent-delete fallback, IHC draft overwrite, and external Word conflict
  async function fetchJsonWithSafetyConfirm(url, options = {}) {
    let curUrl = url;
    let curOptions = Object.assign({}, options);
    for (let attempt = 0; attempt < 4; attempt++) {
      const res = await fetch(curUrl, curOptions);
      const data = await res.json();
      if (data && data.success) {
        return data;
      }
      if (data && data.confirm_sync) {
        const syncOutcome = await resolveExternalSyncForPatient(data, "all", true);
        if (!syncOutcome || !syncOutcome.success) {
          return {
            success: false,
            cancelled: true,
            error: (syncOutcome && syncOutcome.error) || "已取消写入，保留外部 Word 修改内容"
          };
        }
        if (syncOutcome.action === "forward") {
          // For archive calls, user may proceed cleanly with newly imported draft; for text-editing forms, do not overwrite imported fields with stale inputs!
          const isPureArchive = curUrl.includes("/archive");
          if (!isPureArchive) {
            return {
              success: true,
              synced_forward: true,
              patient: syncOutcome.patient,
              ihc_report: syncOutcome.patient ? syncOutcome.patient.ihc_report : null,
              message: "已改为导入外部 Word 的最新修改并完成同步（未用旧表单盲目覆盖）！"
            };
          }
        }
        // Reverse overwrite succeeded (or forward before archive): retry original request once
        continue;
      }
      if (data && data.confirm_permanent) {
        const okPerm = confirm(
          "⚠️ 【回收站移入失败 — 防永久删除拦截】\n\n" +
          (data.error || "无法将旧文件或目标路径移入 Windows 回收站。") +
          "\n\n点击【确定】将执行永久删除/直接覆盖（不可恢复）；\n点击【取消】立即终止本次操作以保护现有文件。"
        );
        if (!okPerm) {
          return { success: false, cancelled: true, error: "已取消操作（已阻止静默永久删除）" };
        }
        const sep = curUrl.includes("?") ? "&" : "?";
        if (!curUrl.includes("force_permanent=")) {
          curUrl = curUrl + sep + "force_permanent=1";
        }
        if (curOptions.body && typeof curOptions.body === "string") {
          try {
            const parsed = JSON.parse(curOptions.body);
            parsed.force_permanent = true;
            curOptions = Object.assign({}, curOptions, { body: JSON.stringify(parsed) });
          } catch (_) {}
        }
        continue;
      }
      if (data && data.confirm_required) {
        const okOver = confirm(
          "⚠️ 【覆盖已有草稿确认】\n\n" +
          (data.error || "目标文件已存在，是否确定覆盖？")
        );
        if (!okOver) {
          return { success: false, cancelled: true, error: "已取消覆盖操作" };
        }
        if (curOptions.body && typeof curOptions.body === "string") {
          try {
            const parsed = JSON.parse(curOptions.body);
            parsed.force_overwrite = true;
            curOptions = Object.assign({}, curOptions, { body: JSON.stringify(parsed) });
          } catch (_) {}
        }
        continue;
      }
      return data;
    }
    return { success: false, error: "操作被中止" };
  }
  const panels = document.querySelectorAll(".workspace-panel");
  const setupWizardModal = document.getElementById("setup-wizard-modal");
  const wizardForm = document.getElementById("wizard-form");
  const currentWorkdirDisplay = document.getElementById("current-workdir-display");
  const currentDoctorsDisplay = document.getElementById("current-doctors-display");

  // Form Elements
  const patientForm = document.getElementById("patient-form");
  const editOldPathologyNoInput = document.getElementById("edit-old-pathology-no");
  const formModeTitle = document.getElementById("form-mode-title");
  const editCancelBtnBox = document.getElementById("edit-cancel-btn-box");
  const btnCancelEdit = document.getElementById("btn-cancel-edit");
  const btnSubmitText = document.getElementById("btn-submit-text");

  const pathologyPrefix = document.getElementById("pathology-prefix");
  const pathologyDate = document.getElementById("pathology-date");
  const pathologyNoInput = document.getElementById("pathology-no");
  const btnAutoNextNo = document.getElementById("btn-auto-next-no");
  const pathologyParsedInfo = document.getElementById("pathology-parsed-info");

  const patientNameInput = document.getElementById("patient-name");
  const patientGenderInput = document.getElementById("patient-gender");
  const patientAgeInput = document.getElementById("patient-age");
  const patientAgeUnitSelect = document.getElementById("patient-age-unit");
  const patientRefDocInput = document.getElementById("patient-ref-doc");
  const patientInpatientNoInput = document.getElementById("patient-inpatient-no");
  const patientOutpatientNoInput = document.getElementById("patient-outpatient-no");
  const patientSendUnitInput = document.getElementById("patient-send-unit");
  const patientDepartmentInput = document.getElementById("patient-department");
  const patientDiagInput = document.getElementById("patient-diag");
  const patientImageCountInput = document.getElementById("patient-image-count");
  const dynamicSitesWrapper = document.getElementById("dynamic-sites-wrapper");
  const patientHistoryInput = document.getElementById("patient-history");
  const btnResetForm = document.getElementById("btn-reset-form");

  // Image count button group
  const imageCountBtns = document.querySelectorAll("#image-count-btn-group .count-btn");

  // Today's list elements
  const todayCountBadge = document.getElementById("today-count-badge");
  const todayPatientList = document.getElementById("today-patient-list");
  const btnRefreshToday = document.getElementById("btn-refresh-today");
  const entryFilterReportDate = document.getElementById("entry-filter-report-date");
  const btnEntryReportToday = document.getElementById("btn-entry-report-today");
  const btnEntryReportAll = document.getElementById("btn-entry-report-all");
  const entryFilterSubDate = document.getElementById("entry-filter-sub-date");
  const btnEntrySubAll = document.getElementById("btn-entry-sub-all");
  const entryFilterKeyword = document.getElementById("entry-filter-keyword");
  const btnEntrySearch = document.getElementById("btn-entry-search");
  const selectedPatientBox = document.getElementById("selected-patient-box");
  const selectedPatientTitle = document.getElementById("selected-patient-title");
  const selectedPatientDiagBadge = document.getElementById("selected-patient-diag-badge");
  const selectedWordPath = document.getElementById("selected-word-path");
  const selectedPhotoPath = document.getElementById("selected-photo-path");
  const btnCopyWordPath = document.getElementById("btn-copy-word-path");
  const btnCopyPhotoPath = document.getElementById("btn-copy-photo-path");
  const btnPreviewSelected = document.getElementById("btn-preview-selected");
  const btnOpenWordFile = document.getElementById("btn-open-word-file");
  const btnOpenPhotoDir = document.getElementById("btn-open-photo-dir");
  const btnEntryFindClinical = document.getElementById("btn-entry-find-clinical");
  const btnEditSelected = document.getElementById("btn-edit-selected");

  // Search Elements
  const searchForm = document.getElementById("search-form");
  const searchReportStartDateInput = document.getElementById("search-report-start-date");
  const searchReportEndDateInput = document.getElementById("search-report-end-date");
  const searchStartDateInput = document.getElementById("search-start-date");
  const searchEndDateInput = document.getElementById("search-end-date");
  const searchPrefixSelect = document.getElementById("search-prefix");
  const searchArchivedSelect = document.getElementById("search-archived");
  const searchReviewStatusSelect = document.getElementById("search-review-status");
  const searchNameInput = document.getElementById("search-name");
  const searchPathologyNoInput = document.getElementById("search-pathology-no");
  const searchInpatientNoInput = document.getElementById("search-inpatient-no");
  const searchOutpatientNoInput = document.getElementById("search-outpatient-no");
  const searchKeywordInput = document.getElementById("search-keyword");
  const btnResetSearch = document.getElementById("btn-reset-search");
  const searchCountBadge = document.getElementById("search-count-badge");
  const searchResultsTbody = document.getElementById("search-results-tbody");

  const btnSearchReportToday = document.getElementById("btn-search-report-today");
  const btnSearchReportWeek = document.getElementById("btn-search-report-week");
  const btnSearchReportMonth = document.getElementById("btn-search-report-month");
  const btnSearchReportAll = document.getElementById("btn-search-report-all");
  const btnSearchDateToday = document.getElementById("btn-search-date-today");
  const btnSearchDateWeek = document.getElementById("btn-search-date-week");
  const btnSearchDateMonth = document.getElementById("btn-search-date-month");
  const btnSearchDateAll = document.getElementById("btn-search-date-all");

  // Settings Elements
  const settingsForm = document.getElementById("settings-form");
  const settingWorkDir = document.getElementById("setting-work-dir");
  const settingReportDoc = document.getElementById("setting-report-doc");
  const settingReviewDoc = document.getElementById("setting-review-doc");

  // Preview Modal Elements
  const reportPreviewModal = document.getElementById("report-preview-modal");
  const reportPreviewBody = document.getElementById("report-preview-body");
  const modalPreviewFilename = document.getElementById("modal-preview-filename");
  const btnModalOpenWord = document.getElementById("btn-modal-open-word");
  const btnClosePreviewModal = document.getElementById("btn-close-preview-modal");

  // Toast Container
  const toastContainer = document.getElementById("toast-container");

  // ================= Shared Helpers =================
  function formatLocalIsoDate(d = new Date()) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  function getTodayDateStr() {
    return (state && state.config && state.config.today) ? state.config.today : formatLocalIsoDate();
  }

  function bindSidebarListFilter({
    reportDateEl, btnReportTodayEl, btnReportAllEl,
    subDateEl, btnSubAllEl, keywordEl, btnSearchEl, onReload
  }) {
    if (reportDateEl) reportDateEl.addEventListener("change", onReload);
    if (subDateEl) subDateEl.addEventListener("change", onReload);
    if (btnReportTodayEl && reportDateEl) {
      btnReportTodayEl.addEventListener("click", () => {
        reportDateEl.value = getTodayDateStr();
        onReload();
      });
    }
    if (btnReportAllEl && reportDateEl) {
      btnReportAllEl.addEventListener("click", () => {
        reportDateEl.value = "";
        onReload();
      });
    }
    if (btnSubAllEl && subDateEl) {
      btnSubAllEl.addEventListener("click", () => {
        subDateEl.value = "";
        onReload();
      });
    }
    if (btnSearchEl) {
      btnSearchEl.addEventListener("click", onReload);
    }
    if (keywordEl) {
      keywordEl.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          onReload();
        }
      });
    }
  }

  function getPrefixGroupParam(allChkEl, itemChkNodes) {
    if (allChkEl && allChkEl.checked) return "ALL";
    const sel = Array.from(itemChkNodes || []).filter(c => c.checked).map(c => c.value);
    return sel.length ? sel.join(",") : "__EMPTY__";
  }

  function bindPrefixCheckboxGroup(allChkEl, itemChkNodes, onChange) {
    if (allChkEl) {
      allChkEl.addEventListener("change", () => {
        Array.from(itemChkNodes || []).forEach(chk => chk.checked = allChkEl.checked);
        if (onChange) onChange();
      });
    }
    Array.from(itemChkNodes || []).forEach(chk => {
      chk.addEventListener("change", () => {
        if (allChkEl) {
          allChkEl.checked = Array.from(itemChkNodes || []).every(c => c.checked);
        }
        if (onChange) onChange();
      });
    });
  }

  function bindRangeQuickDateButtons({ startEl, endEl, btnTodayEl, btnWeekEl, btnMonthEl, btnAllEl, onReload, triggerOnInput = true }) {
    if (triggerOnInput) {
      if (startEl) startEl.addEventListener("change", onReload);
      if (endEl) endEl.addEventListener("change", onReload);
    }
    if (btnTodayEl) {
      btnTodayEl.addEventListener("click", () => {
        const today = getTodayDateStr();
        if (startEl) startEl.value = today;
        if (endEl) endEl.value = today;
        onReload();
      });
    }
    if (btnWeekEl) {
      btnWeekEl.addEventListener("click", () => {
        const now = new Date();
        const day = now.getDay() || 7;
        const monday = new Date(now);
        monday.setDate(now.getDate() - day + 1);
        if (startEl) startEl.value = formatLocalIsoDate(monday);
        if (endEl) endEl.value = formatLocalIsoDate(now);
        onReload();
      });
    }
    if (btnMonthEl) {
      btnMonthEl.addEventListener("click", () => {
        const now = new Date();
        const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
        if (startEl) startEl.value = formatLocalIsoDate(firstDay);
        if (endEl) endEl.value = formatLocalIsoDate(now);
        onReload();
      });
    }
    if (btnAllEl) {
      btnAllEl.addEventListener("click", () => {
        if (startEl) startEl.value = "";
        if (endEl) endEl.value = "";
        onReload();
      });
    }
  }

  function formatAgeDisplay(age) {
    const raw = String(age || "").trim();
    if (!raw) return "-";
    return /[岁月底天日]$/.test(raw) ? raw : `${raw}岁`;
  }

  function composeAgeValue(raw, unit) {
    const val = String(raw || "").trim();
    if (!val) return "";
    if (unit === "自由输入") return val;
    if (/[岁月底天日]$/.test(val)) return val;
    return `${val}${unit || "岁"}`;
  }

  function populateAgeInputs(ageStr) {
    const raw = String(ageStr || "").trim();
    if (!raw) {
      patientAgeInput.value = "";
      if (patientAgeUnitSelect) patientAgeUnitSelect.value = "岁";
      return;
    }
    const m = raw.match(/^(\d+(?:\.\d+)?)\s*(岁|月|日)$/);
    if (m) {
      patientAgeInput.value = m[1];
      if (patientAgeUnitSelect) patientAgeUnitSelect.value = m[2];
    } else if (/^\d+(?:\.\d+)?$/.test(raw)) {
      patientAgeInput.value = raw;
      if (patientAgeUnitSelect) patientAgeUnitSelect.value = "岁";
    } else {
      patientAgeInput.value = raw;
      if (patientAgeUnitSelect) patientAgeUnitSelect.value = "自由输入";
    }
  }

  function syncUnitDeptEditable(prefixVal, keepValues = false) {
    if (!patientSendUnitInput || !patientDepartmentInput) return;
    const pfx = String(prefixVal || "").trim().toUpperCase();
    const canEdit = (pfx === "H" || pfx === "T");
    patientSendUnitInput.readOnly = !canEdit;
    patientDepartmentInput.readOnly = !canEdit;
    if (!canEdit && !keepValues) {
      patientSendUnitInput.value = "本院";
      patientDepartmentInput.value = "皮肤科";
    }
  }

  function getEffectivePrefixForUnitDept() {
    const noVal = (pathologyNoInput.value || "").trim().toUpperCase();
    const m = noVal.match(/^([BLHT])/);
    if (m) return m[1];
    return (pathologyPrefix.value || "").trim().toUpperCase();
  }

  function toIsoDateString(val) {
    const s = String(val || "").trim();
    if (!s) return "";
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const m = s.match(/^(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})日?$/);
    if (m) {
      return `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(m[3]).padStart(2, "0")}`;
    }
    return "";
  }

  function renderExternalModBadge(p, ihcMode = false) {
    if (!p) return "";
    const flag = ihcMode ? !!p.ihc_external_modified : !!p.external_modified;
    if (!flag) return "";
    return `<span class="badge-status badge-external-mod" title="检测到本地 Word 文档已在外部被修改，点击可选择导入或覆盖">⚠️ Word已外改</span>`;
  }

  function renderReviewStatusBadges(p) {
    const rev = String(p.review_status || "");
    if (rev === "完成审核") {
      let html = `<span class="badge-status badge-approved">完成审核</span>`;
      if (p.is_needs_ihc === 1) {
        html += `<span class="badge-status badge-ihc">免疫组化</span>`;
      }
      return html;
    }
    let html = "";
    if (p.is_pending_review === 1 || rev.includes("待审核")) {
      html += `<span class="badge-status badge-pending">待审核</span>`;
    }
    if (p.is_needs_ihc === 1 || rev.includes("免疫组化")) {
      html += `<span class="badge-status badge-ihc">免疫组化</span>`;
    }
    if (p.is_needs_recut === 1 || rev.includes("重切")) {
      html += `<span class="badge-status badge-recut">重切</span>`;
    }
    return html;
  }

  function getMakingCardStatusClass(p) {
    if (p.archived === 2) return "status-fill-rearchive";
    if (p.archived === 1) return "status-fill-archived";
    const rev = String(p.review_status || "");
    if (rev === "完成审核") return "status-fill-approved";
    if (p.is_pending_review === 1 || rev.includes("待审核")) return "status-fill-pending";
    if (p.is_needs_ihc === 1 || rev.includes("免疫组化")) return "status-fill-ihc";
    if (p.is_needs_recut === 1 || rev.includes("重切")) return "status-fill-recut";
    if (p.final_report_path) return "status-fill-saved";
    return "";
  }

  async function findClinicalPhotosForPatient(patient) {
    if (!patient) return;
    try {
      const res = await fetch("/api/clinical_photos/find", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: patient.name, pathology_no: patient.pathology_no })
      });
      const data = await res.json();
      if (data.multiple) {
        alert(data.message);
        showToast(data.message, "warning");
      } else if (data.found) {
        showToast(data.message, "success");
      } else {
        showToast(data.message, "error");
      }
    } catch (e) {
      showToast("查询临床照片失败: " + e.message, "error");
    }
  }

  // ================= Toast Helper =================
  function showToast(message, type = "success") {
    const toast = document.createElement("div");
    toast.className = `toast toast-${type}`;
    const icon = type === "success" ? "✓" : (type === "error" ? "✕" : "ℹ");
    toast.innerHTML = `<span>${icon}</span> <span>${message}</span>`;
    toastContainer.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = "0";
      toast.style.transition = "opacity 0.3s";
      setTimeout(() => toast.remove(), 300);
    }, 3200);
  }

  function findMatchingPhoto(photosList, pPath) {
    if (!pPath || !Array.isArray(photosList) || !photosList.length) return null;
    let hit = photosList.find(ph => ph.path === pPath);
    if (!hit) {
      const baseName = String(pPath).split(/[/\\]/).pop().toLowerCase();
      if (baseName) {
        hit = photosList.find(ph => (ph.filename || "").toLowerCase() === baseName);
      }
    }
    return hit || null;
  }

  // ================= Clipboard Helper =================
  async function copyToClipboard(text, successMsg = "路径已成功复制到剪贴板！") {
    if (!text || text === "-") return;
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
      } else {
        const textArea = document.createElement("textarea");
        textArea.value = text;
        textArea.style.position = "fixed";
        textArea.style.left = "-999999px";
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        document.execCommand("copy");
        textArea.remove();
      }
      showToast(successMsg, "success");
    } catch (err) {
      showToast("复制失败，请手动复制: " + text, "error");
    }
  }

  // ================= Navigation Tabs =================
  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      tabs.forEach(t => t.classList.remove("active"));
      panels.forEach(p => p.classList.remove("active"));

      tab.classList.add("active");
      const targetId = tab.dataset.tab;
      const targetPanel = document.getElementById(targetId);
      if (targetPanel) targetPanel.classList.add("active");

      if (targetId === "tab-entry") {
        loadTodayPatients();
      } else if (targetId === "tab-search") {
        // 默认报告日期与送检日期均为全部时间，便于直接输入条件从所有患者中查询
        doSearch();
      } else if (targetId === "tab-making") {
        initMakingDefaultDates();
        loadMakingPatients(true);
      } else if (targetId === "tab-review") {
        loadReviewPatients(true);
      } else if (targetId === "tab-ihc") {
        loadIhcPatients();
      } else if (targetId === "tab-output") {
        loadOutputRegularPatients();
        loadOutputSpecialPatients();
        loadOutputIhcPatients();
      } else if (targetId === "tab-signing") {
        initSigningDateRange();
        loadSigningCandidates();
      }
    });
  });

  // ================= Init System Status =================
  async function checkSystemStatus() {
    try {
      const res = await fetch("/api/status");
      const data = await res.json();
      state.initialized = data.initialized;
      state.config = data;

      if (!data.initialized) {
        setupWizardModal.style.display = "flex";
        const wizWorkDir = document.getElementById("wizard-work-dir");
        const wizRepDoc = document.getElementById("wizard-report-doc");
        const wizRevDoc = document.getElementById("wizard-review-doc");
        if (wizWorkDir && data.work_dir) wizWorkDir.value = data.work_dir;
        if (wizRepDoc && data.report_doctor) wizRepDoc.value = data.report_doctor;
        if (wizRevDoc && data.review_doctor) wizRevDoc.value = data.review_doctor;
      } else {
        setupWizardModal.style.display = "none";
        currentWorkdirDisplay.textContent = data.work_dir ? data.work_dir.split(/[\\/]/).pop() : "已连接工作目录";
        currentWorkdirDisplay.parentElement.title = data.work_dir;
        currentDoctorsDisplay.textContent = `${data.report_doctor} / ${data.review_doctor}`;

        settingWorkDir.value = data.work_dir || "";
        settingReportDoc.value = data.report_doctor || "";
        settingReviewDoc.value = data.review_doctor || "";

        // Default date to today
        const todayStr = data.today || formatLocalIsoDate();
        pathologyDate.value = todayStr;
        if (entryFilterReportDate && !state.entryDefaultsInitialized) {
          entryFilterReportDate.value = todayStr;
          state.entryDefaultsInitialized = true;
        }
        initMakingDefaultDates();

        // Auto generate first sequence number if pathology no is empty
        if (!pathologyNoInput.value) {
          fetchNextPathologyNo();
        }

        loadTodayPatients();
      }
    } catch (err) {
      showToast("无法连接系统后台: " + err.message, "error");
    }
  }

  // ================= Setup Wizard Form =================
  wizardForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const workDir = document.getElementById("wizard-work-dir").value.trim();
    const repDoc = document.getElementById("wizard-report-doc").value.trim();
    const revDoc = document.getElementById("wizard-review-doc").value.trim();

    try {
      const res = await fetch("/api/init", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          work_dir: workDir,
          report_doctor: repDoc,
          review_doctor: revDoc
        })
      });
      const data = await res.json();
      if (data.success) {
        showToast("工作目录初始化完成！", "success");
        setupWizardModal.style.display = "none";
        checkSystemStatus();
      } else {
        showToast(data.error || "初始化失败", "error");
      }
    } catch (err) {
      showToast("初始化请求异常: " + err.message, "error");
    }
  });

  // ================= Settings Form =================
  settingsForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const updates = {
      work_dir: settingWorkDir.value.trim(),
      report_doctor: settingReportDoc.value.trim(),
      review_doctor: settingReviewDoc.value.trim()
    };
    try {
      const res = await fetch("/api/config", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(updates)
      });
      const data = await res.json();
      if (data.success) {
        showToast("系统设置已更新！", "success");
        checkSystemStatus();
      } else {
        showToast(data.error || "保存失败", "error");
      }
    } catch (err) {
      showToast("请求失败: " + err.message, "error");
    }
  });

  // ================= Pathology No & Real-time Validation =================
  async function fetchNextPathologyNo() {
    const prefix = pathologyPrefix.value;
    const dateVal = pathologyDate.value;
    try {
      const res = await fetch(`/api/next_no?prefix=${encodeURIComponent(prefix)}&date=${encodeURIComponent(dateVal)}`);
      const data = await res.json();
      if (data.success && data.next_no) {
        pathologyNoInput.value = data.next_no;
        validatePathologyNo();
      }
    } catch (err) {
      console.error("fetchNextPathologyNo error:", err);
    }
  }

  btnAutoNextNo.addEventListener("click", fetchNextPathologyNo);

  pathologyPrefix.addEventListener("change", () => {
    syncUnitDeptEditable(pathologyPrefix.value);
    if (!state.isEditMode) {
      fetchNextPathologyNo();
    } else {
      validatePathologyNo();
    }
  });

  if (patientAgeUnitSelect) {
    patientAgeUnitSelect.addEventListener("change", () => {
      patientAgeInput.placeholder = patientAgeUnitSelect.value === "自由输入" ? "如 1岁3月" : "如 45";
    });
  }

  pathologyDate.addEventListener("change", () => {
    if (!state.isEditMode) {
      fetchNextPathologyNo();
    } else {
      validatePathologyNo();
    }
  });

  pathologyNoInput.addEventListener("input", () => {
    syncUnitDeptEditable(getEffectivePrefixForUnitDept(), true);
    if (state.checkNoTimer) clearTimeout(state.checkNoTimer);
    state.checkNoTimer = setTimeout(validatePathologyNo, 300);
  });

  async function validatePathologyNo() {
    const val = pathologyNoInput.value.trim().toUpperCase();
    if (!val) {
      pathologyParsedInfo.innerHTML = "";
      return;
    }

    // Parse format: [Prefix]YYYYMMDDNN
    const match = val.match(/^([BLHT]?)(\d{4})(\d{2})(\d{2})(\d+)$/);
    if (!match) {
      pathologyParsedInfo.innerHTML = `<span class="feedback-dup">⚠️ 病理号格式不符合规范 (示例: 2026101501 或 L2026101501)</span>`;
      return;
    }

    const [_, prefix, y, m, d, seq] = match;
    const prefixNameMap = { "": "门诊/普通", "B": "病房", "L": "分院区", "H": "会诊", "T": "特殊" };
    const dateDesc = `${y}年${parseInt(m, 10)}月${parseInt(d, 10)}日`;
    const prefixDesc = prefixNameMap[prefix] || prefix;

    // Check duplicate in database
    try {
      const res = await fetch(`/api/check_no?no=${encodeURIComponent(val)}`);
      const data = await res.json();
      if (data.exists) {
        const isSelf = state.isEditMode && editOldPathologyNoInput.value === val;
        if (!isSelf) {
          pathologyParsedInfo.innerHTML = `
            <span class="feedback-dup">
              ⚠️ 病理号 [${val}] 已存在！对应患者：${data.patient.name} (送检日期: ${data.patient.submission_date})
            </span>
          `;
          return;
        }
      }

      pathologyParsedInfo.innerHTML = `
        <span class="feedback-ok">
          ✓ 解析成功：${prefixDesc} | 送检日期：${dateDesc} | 第 ${parseInt(seq, 10)} 份
        </span>
      `;
    } catch (e) {
      pathologyParsedInfo.innerHTML = `<span class="feedback-ok">✓ 格式有效：送检日期 ${dateDesc}</span>`;
    }
  }

  // ================= Image Count & Dynamic Biopsy Sites =================
  function setImageCount(count, existingSites = []) {
    state.imageCount = count;
    patientImageCountInput.value = count;

    imageCountBtns.forEach(btn => {
      btn.classList.toggle("active", parseInt(btn.dataset.count, 10) === count);
    });

    renderDynamicBiopsySites(count, existingSites);
  }

  imageCountBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      const currentValues = getBiopsySiteValues();
      setImageCount(parseInt(btn.dataset.count, 10), currentValues);
    });
  });

  function renderDynamicBiopsySites(count, values = []) {
    dynamicSitesWrapper.innerHTML = "";

    if (count === 1) {
      const val = values[0] || "";
      dynamicSitesWrapper.innerHTML = `
        <div class="site-input-group">
          <input type="text" class="form-control biopsy-site-input" id="site-input-1" value="${val}" placeholder="如：左额部红斑伴脱屑">
        </div>
      `;
    } else {
      const container = document.createElement("div");
      container.className = "dynamic-sites-container";

      for (let i = 1; i <= count; i++) {
        const val = values[i - 1] || "";
        const group = document.createElement("div");
        group.className = "site-input-group";
        group.innerHTML = `
          <span class="site-badge">${i}</span>
          <input type="text" class="form-control biopsy-site-input" id="site-input-${i}" value="${val}" placeholder="部位 ${i} 名称与皮损特点">
        `;
        container.appendChild(group);
      }
      dynamicSitesWrapper.appendChild(container);
    }
  }

  function getBiopsySiteValues() {
    const inputs = dynamicSitesWrapper.querySelectorAll(".biopsy-site-input");
    const arr = [];
    inputs.forEach(input => {
      arr.push(input.value.trim());
    });
    return arr;
  }

  // Initial render with count = 1
  setImageCount(1);

  // ================= Patient Form Submit (Create / Edit) =================
  patientForm.addEventListener("submit", async (e) => {
    e.preventDefault();

    const pathologyNo = pathologyNoInput.value.trim().toUpperCase();
    const name = patientNameInput.value.trim();
    const gender = patientGenderInput.value;
    const ageUnit = patientAgeUnitSelect ? patientAgeUnitSelect.value : "岁";
    const age = composeAgeValue(patientAgeInput.value, ageUnit);
    const refDoc = patientRefDocInput.value.trim();
    const inpatientNo = patientInpatientNoInput.value.trim();
    const outpatientNo = patientOutpatientNoInput.value.trim();
    const sendUnit = patientSendUnitInput ? (patientSendUnitInput.value.trim() || "本院") : "本院";
    const department = patientDepartmentInput ? (patientDepartmentInput.value.trim() || "皮肤科") : "皮肤科";
    const diag = patientDiagInput.value.trim();
    const history = patientHistoryInput.value.trim();
    const sites = getBiopsySiteValues();

    if (!pathologyNo) {
      showToast("请输入病理号", "error");
      pathologyNoInput.focus();
      return;
    }
    if (!name) {
      showToast("请输入患者姓名", "error");
      patientNameInput.focus();
      return;
    }

    const payload = {
      pathology_no: pathologyNo,
      name: name,
      gender: gender,
      age: age,
      referring_doctor: refDoc,
      inpatient_no: inpatientNo,
      outpatient_no: outpatientNo,
      send_unit: sendUnit,
      department: department,
      clinical_diagnosis: diag,
      image_count: state.imageCount,
      biopsy_sites: sites,
      medical_history: history
    };

    try {
      let data;
      if (state.isEditMode) {
        const oldNo = editOldPathologyNoInput.value.trim();
        data = await fetchJsonWithSafetyConfirm(`/api/patients/${encodeURIComponent(oldNo)}`, {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });
      } else {
        const res = await fetch("/api/patients", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });
        data = await res.json();
      }

      if (data.success) {
        showToast(data.message || "操作成功！", "success");
        let savedPatient = data.patient;

        if (data.synced_forward) {
          exitEditMode();
          resetPatientForm();
          await loadTodayPatients();
          if (savedPatient) selectPatient(savedPatient, false);
          return;
        }

        if (state.isEditMode && savedPatient) {
          savedPatient = await promptRearchiveAfterUpdate(data, savedPatient);
        }

        const wasEditMode = state.isEditMode;
        const inheritPrefix = (!wasEditMode && savedPatient && savedPatient.prefix !== undefined)
          ? (savedPatient.prefix || "")
          : (!wasEditMode ? (pathologyPrefix.value || "") : undefined);
        const inheritSubDate = (!wasEditMode && savedPatient && savedPatient.submission_date_raw)
          ? savedPatient.submission_date_raw
          : (!wasEditMode ? (pathologyDate.value || "") : undefined);

        exitEditMode();
        if (!wasEditMode) {
          resetPatientForm({ inheritPrefix, inheritSubDate });
        } else {
          resetPatientForm();
        }
        await loadTodayPatients();
        if (makingState && makingState.patients) {
          loadMakingPatients();
        }

        if (savedPatient) {
          selectPatient(savedPatient);
        }
      } else {
        showToast(data.error || "保存失败", "error");
      }
    } catch (err) {
      showToast("提交异常: " + err.message, "error");
    }
  });

  // ================= Reset & Exit Edit =================
  function resetPatientForm(options = {}) {
    const opts = options && typeof options === "object" ? options : {};
    const targetPrefix = opts.inheritPrefix !== undefined ? String(opts.inheritPrefix || "") : "";
    const targetDate = opts.inheritSubDate || getTodayDateStr();

    patientForm.reset();
    editOldPathologyNoInput.value = "";
    pathologyPrefix.value = targetPrefix;
    if (patientAgeUnitSelect) patientAgeUnitSelect.value = "岁";
    patientAgeInput.placeholder = "如 45";
    syncUnitDeptEditable(targetPrefix);
    pathologyDate.value = targetDate;
    setImageCount(1);
    fetchNextPathologyNo();
  }

  btnResetForm.addEventListener("click", () => resetPatientForm());

  function enterEditMode(patient) {
    state.isEditMode = true;
    formModeTitle.textContent = `✏️ 修改患者信息 (${patient.pathology_no} - ${patient.name})`;
    btnSubmitText.textContent = "保存修改并同步Word文件";
    editCancelBtnBox.style.display = "block";
    editOldPathologyNoInput.value = patient.pathology_no;

    // Fill form
    pathologyNoInput.value = patient.pathology_no;
    pathologyPrefix.value = patient.prefix || "";
    if (patient.submission_date_raw) pathologyDate.value = patient.submission_date_raw;

    patientNameInput.value = patient.name || "";
    patientGenderInput.value = patient.gender || "女";
    populateAgeInputs(patient.age);
    if (patientAgeUnitSelect) {
      patientAgeInput.placeholder = patientAgeUnitSelect.value === "自由输入" ? "如 1岁3月" : "如 45";
    }
    patientRefDocInput.value = patient.referring_doctor || "";
    patientInpatientNoInput.value = patient.inpatient_no || "";
    patientOutpatientNoInput.value = patient.outpatient_no || "";
    if (patientSendUnitInput) patientSendUnitInput.value = patient.send_unit || "本院";
    if (patientDepartmentInput) patientDepartmentInput.value = patient.department || "皮肤科";
    syncUnitDeptEditable(patient.prefix || getEffectivePrefixForUnitDept(), true);
    patientDiagInput.value = patient.clinical_diagnosis || "";
    patientHistoryInput.value = patient.medical_history || "";

    const count = parseInt(patient.image_count, 10) || 1;
    setImageCount(count, patient.biopsy_sites || []);

    validatePathologyNo();
    patientForm.scrollIntoView({ behavior: "smooth" });
  }

  function exitEditMode() {
    state.isEditMode = false;
    formModeTitle.textContent = "📝 患者信息录入";
    btnSubmitText.textContent = "保存并生成病理报告";
    editCancelBtnBox.style.display = "none";
    editOldPathologyNoInput.value = "";
  }

  btnCancelEdit.addEventListener("click", () => {
    exitEditMode();
    resetPatientForm();
  });

  // ================= Load Entry Patient List =================
  async function loadTodayPatients() {
    try {
      const params = new URLSearchParams();
      params.set("filter_mode", "1");
      params.set("report_date", entryFilterReportDate ? entryFilterReportDate.value : getTodayDateStr());
      if (entryFilterSubDate && entryFilterSubDate.value) {
        params.set("submission_date", entryFilterSubDate.value);
      }
      if (entryFilterKeyword && entryFilterKeyword.value.trim()) {
        params.set("keyword", entryFilterKeyword.value.trim());
      }
      const res = await fetch(`/api/patients/today?${params.toString()}`);
      const data = await res.json();
      if (data.success) {
        state.todayPatients = data.patients || [];
        renderTodayPatientList(state.todayPatients);
      }
    } catch (err) {
      console.error("loadTodayPatients error:", err);
    }
  }

  btnRefreshToday.addEventListener("click", loadTodayPatients);
  bindSidebarListFilter({
    reportDateEl: entryFilterReportDate,
    btnReportTodayEl: btnEntryReportToday,
    btnReportAllEl: btnEntryReportAll,
    subDateEl: entryFilterSubDate,
    btnSubAllEl: btnEntrySubAll,
    keywordEl: entryFilterKeyword,
    btnSearchEl: btnEntrySearch,
    onReload: loadTodayPatients
  });

  function renderTodayPatientList(patients) {
    todayCountBadge.textContent = `${patients.length} 人`;
    todayPatientList.innerHTML = "";

    if (patients.length === 0) {
      todayPatientList.innerHTML = `<div class="patient-list-empty">暂无符合条件的患者记录</div>`;
      selectedPatientBox.style.display = "none";
      return;
    }

    patients.forEach(p => {
      const item = document.createElement("div");
      item.className = `patient-item ${p.external_modified ? "external-modified-card" : ""}`.trim();
      if (state.selectedPatient && state.selectedPatient.pathology_no === p.pathology_no) {
        item.classList.add("selected");
      }

      const extBadge = renderExternalModBadge(p, false);
      item.innerHTML = `
        <div class="patient-item-header">
          <span class="patient-name-box">${p.name}</span>
          <span style="display: inline-flex; gap: 4px; align-items: center;">
            ${extBadge}
            <span class="patient-pathology-badge">${p.pathology_no}</span>
          </span>
        </div>
        <div class="patient-meta">
          <span>${p.gender || "-"} / ${formatAgeDisplay(p.age)}</span>
          <span>送检医生: ${p.referring_doctor || "未填"}</span>
          <span>图: ${p.image_count || 1}张</span>
        </div>
        <div class="patient-diag">
          <strong>诊断:</strong> ${p.clinical_diagnosis || "暂无临床诊断"}
        </div>
      `;

      item.addEventListener("click", () => {
        selectPatient(p, true);
      });

      todayPatientList.appendChild(item);
    });
  }

  // ================= Select Patient & Detail Actions =================
  async function selectPatient(patient, fromUserClick = false) {
    let effPatient = patient;
    if (fromUserClick && effPatient && effPatient.external_modified) {
      const outcome = await resolveExternalSyncForPatient(effPatient, "normal", false);
      if (outcome && outcome.success && outcome.patient) {
        effPatient = outcome.patient;
        if (state.isEditMode && editOldPathologyNoInput.value.trim() === effPatient.pathology_no) {
          enterEditMode(effPatient);
        }
      }
    }
    state.selectedPatient = effPatient;

    // Highlight item in today's list
    const items = todayPatientList.querySelectorAll(".patient-item");
    items.forEach(el => {
      const badge = el.querySelector(".patient-pathology-badge");
      if (badge && badge.textContent === effPatient.pathology_no) {
        el.classList.add("selected");
      } else {
        el.classList.remove("selected");
      }
    });

    selectedPatientBox.style.display = "block";
    selectedPatientTitle.textContent = `${effPatient.name} (${effPatient.pathology_no})`;
    selectedPatientDiagBadge.textContent = effPatient.clinical_diagnosis || "无诊断";

    selectedWordPath.textContent = `Word报告: ${effPatient.word_path || "-"}`;
    selectedWordPath.title = effPatient.word_path || "";

    selectedPhotoPath.textContent = `照片目录: ${effPatient.photo_dir || "-"}`;
    selectedPhotoPath.title = effPatient.photo_dir || "";
  }

  // Copy Path buttons
  btnCopyWordPath.addEventListener("click", () => {
    if (state.selectedPatient && state.selectedPatient.word_path) {
      copyToClipboard(state.selectedPatient.word_path, "已复制Word报告文件完整路径！");
    }
  });

  btnCopyPhotoPath.addEventListener("click", () => {
    if (state.selectedPatient && state.selectedPatient.photo_dir) {
      copyToClipboard(state.selectedPatient.photo_dir, "已复制病理照片文件夹完整路径！");
    }
  });

  // Open Word File
  btnOpenWordFile.addEventListener("click", async () => {
    if (!state.selectedPatient || !state.selectedPatient.word_path) return;
    try {
      const res = await fetch("/api/open_file", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ path: state.selectedPatient.word_path })
      });
      const data = await res.json();
      if (data.success) {
        showToast("已在本地 Word 中打开文件", "success");
      } else {
        showToast(data.error || "无法打开文件", "error");
      }
    } catch (e) {
      showToast("打开异常: " + e.message, "error");
    }
  });

  // Open Photo Dir in Explorer
  btnOpenPhotoDir.addEventListener("click", async () => {
    if (!state.selectedPatient || !state.selectedPatient.photo_dir) return;
    try {
      const res = await fetch("/api/open_dir", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ path: state.selectedPatient.photo_dir })
      });
      const data = await res.json();
      if (data.success) {
        showToast("已在资源管理器中打开文件夹", "success");
      } else {
        showToast(data.error || "无法打开文件夹", "error");
      }
    } catch (e) {
      showToast("打开异常: " + e.message, "error");
    }
  });

  // Clinical photo lookup button in Entry
  if (btnEntryFindClinical) {
    btnEntryFindClinical.addEventListener("click", () => {
      findClinicalPhotosForPatient(state.selectedPatient);
    });
  }

  // Edit Selected Patient
  btnEditSelected.addEventListener("click", async () => {
    if (state.selectedPatient) {
      let targetP = state.selectedPatient;
      if (targetP.external_modified) {
        const outcome = await resolveExternalSyncForPatient(targetP, "normal", false);
        if (outcome && outcome.success && outcome.patient) targetP = outcome.patient;
      }
      enterEditMode(targetP);
    }
  });

  // Preview Selected Patient
  btnPreviewSelected.addEventListener("click", () => {
    if (state.selectedPatient) {
      openReportPreviewModal(state.selectedPatient.pathology_no);
    }
  });

  // ================= A4 Word Report Preview Modal =================
  async function openReportPreviewModal(pathologyNo, kind = "normal") {
    try {
      reportPreviewBody.innerHTML = `<div style="text-align: center; padding: 60px; color: #666;">正在加载报告预览...</div>`;
      reportPreviewModal.style.display = "flex";

      const url = kind === "ihc"
        ? `/api/ihc/${encodeURIComponent(pathologyNo)}/preview`
        : `/api/patients/${encodeURIComponent(pathologyNo)}/preview`;
      let res = await fetch(url);
      let data = await res.json();
      if (data.success && data.patient) {
        const hasExt = (kind === "ihc") ? !!data.patient.ihc_external_modified : !!data.patient.external_modified;
        if (hasExt) {
          const outcome = await resolveExternalSyncForPatient(data.patient, kind === "ihc" ? "ihc" : "normal", false);
          if (outcome && outcome.success) {
            res = await fetch(url);
            data = await res.json();
          }
        }
      }
      if (data.success) {
        const reportPath = data.doc_path || data.patient.final_report_path || data.patient.word_path;
        modalPreviewFilename.textContent = reportPath ? reportPath.split(/[\\/]/).pop() : `${data.patient.pathology_no}_${data.patient.name}.docx`;
        reportPreviewBody.innerHTML = data.html;

        btnModalOpenWord.onclick = async () => {
          if (reportPath) {
            await fetch("/api/open_file", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ path: reportPath })
            });
          }
        };
      } else {
        reportPreviewBody.innerHTML = `<div style="color: red; padding: 40px;">预览失败: ${data.error}</div>`;
      }
    } catch (e) {
      reportPreviewBody.innerHTML = `<div style="color: red; padding: 40px;">加载预览异常: ${e.message}</div>`;
    }
  }

  btnClosePreviewModal.addEventListener("click", () => {
    reportPreviewModal.style.display = "none";
  });

  reportPreviewModal.addEventListener("click", (e) => {
    if (e.target === reportPreviewModal) {
      reportPreviewModal.style.display = "none";
    }
  });

  // ================= Search Workspace Logic (Requirement 2) =================
  async function doSearch() {
    const params = new URLSearchParams();
    if (searchReportStartDateInput && searchReportStartDateInput.value) params.append("report_start_date", searchReportStartDateInput.value);
    if (searchReportEndDateInput && searchReportEndDateInput.value) params.append("report_end_date", searchReportEndDateInput.value);
    if (searchStartDateInput && searchStartDateInput.value) params.append("start_date", searchStartDateInput.value);
    if (searchEndDateInput && searchEndDateInput.value) params.append("end_date", searchEndDateInput.value);
    if (searchPrefixSelect && searchPrefixSelect.value && searchPrefixSelect.value !== "ALL") params.append("prefix", searchPrefixSelect.value);
    if (searchArchivedSelect && searchArchivedSelect.value && searchArchivedSelect.value !== "ALL") params.append("archived", searchArchivedSelect.value);
    if (searchReviewStatusSelect && searchReviewStatusSelect.value && searchReviewStatusSelect.value !== "ALL") params.append("review_status", searchReviewStatusSelect.value);
    if (searchNameInput.value.trim()) params.append("name", searchNameInput.value.trim());
    if (searchPathologyNoInput.value.trim()) params.append("pathology_no", searchPathologyNoInput.value.trim());
    if (searchInpatientNoInput.value.trim()) params.append("inpatient_no", searchInpatientNoInput.value.trim());
    if (searchOutpatientNoInput.value.trim()) params.append("outpatient_no", searchOutpatientNoInput.value.trim());
    if (searchKeywordInput && searchKeywordInput.value.trim()) params.append("keyword", searchKeywordInput.value.trim());

    try {
      const res = await fetch(`/api/patients/search?${params.toString()}`);
      const data = await res.json();
      if (data.success) {
        renderSearchResults(data.patients || []);
      }
    } catch (e) {
      showToast("查询异常: " + e.message, "error");
    }
  }

  // Quick date range buttons for search (report date & submission date)
  bindRangeQuickDateButtons({
    startEl: searchReportStartDateInput,
    endEl: searchReportEndDateInput,
    btnTodayEl: btnSearchReportToday,
    btnWeekEl: btnSearchReportWeek,
    btnMonthEl: btnSearchReportMonth,
    btnAllEl: btnSearchReportAll,
    onReload: doSearch,
    triggerOnInput: false
  });
  bindRangeQuickDateButtons({
    startEl: searchStartDateInput,
    endEl: searchEndDateInput,
    btnTodayEl: btnSearchDateToday,
    btnWeekEl: btnSearchDateWeek,
    btnMonthEl: btnSearchDateMonth,
    btnAllEl: btnSearchDateAll,
    onReload: doSearch,
    triggerOnInput: false
  });

  searchForm.addEventListener("submit", (e) => {
    e.preventDefault();
    doSearch();
  });

  // One-click cancel all filters (Requirement 2)
  btnResetSearch.addEventListener("click", () => {
    if (searchReportStartDateInput) searchReportStartDateInput.value = "";
    if (searchReportEndDateInput) searchReportEndDateInput.value = "";
    if (searchStartDateInput) searchStartDateInput.value = "";
    if (searchEndDateInput) searchEndDateInput.value = "";
    if (searchPrefixSelect) searchPrefixSelect.value = "ALL";
    if (searchArchivedSelect) searchArchivedSelect.value = "ALL";
    if (searchReviewStatusSelect) searchReviewStatusSelect.value = "ALL";
    searchNameInput.value = "";
    searchPathologyNoInput.value = "";
    searchInpatientNoInput.value = "";
    searchOutpatientNoInput.value = "";
    if (searchKeywordInput) searchKeywordInput.value = "";
    doSearch();
    showToast("已清空所有筛选条件，显示全量结果！", "info");
  });

  function renderSearchResults(patients) {
    searchCountBadge.textContent = `${patients.length} 条记录`;
    searchResultsTbody.innerHTML = "";

    if (patients.length === 0) {
      searchResultsTbody.innerHTML = `
        <tr>
          <td colspan="10" style="text-align: center; color: var(--text-muted); padding: 40px;">
            未找到匹配的患者记录
          </td>
        </tr>
      `;
      return;
    }

    patients.forEach(p => {
      const tr = document.createElement("tr");

      const sitesStr = Array.isArray(p.biopsy_sites) ? p.biopsy_sites.join("；") : (p.biopsy_sites || "-");

      // Status badges
      let statusHtml = renderReviewStatusBadges(p);
      if (statusHtml) statusHtml = statusHtml.replace(/class="badge-status/g, 'style="margin-right: 4px;" class="badge-status');

      if (p.archived === 1) {
        statusHtml += `<span class="badge-status badge-archived">已归档</span>`;
      } else if (p.archived === 2) {
        statusHtml += `<span class="badge-status badge-rearchive">需重新归档</span>`;
      } else if (!statusHtml) {
        statusHtml = `<span style="color: var(--text-muted); font-size: 11px;">未归档</span>`;
      }
      if (p.external_modified || p.ihc_external_modified) {
        statusHtml += ` <span class="badge-status badge-external-mod" title="检测到磁盘 Word 已在外部修改">⚠️ Word已外改</span>`;
      }

      tr.innerHTML = `
        <td style="font-weight: 600; font-family: monospace; color: var(--primary);">${p.pathology_no}</td>
        <td style="font-weight: 600;">${p.name}</td>
        <td>${p.gender || "-"}</td>
        <td>${formatAgeDisplay(p.age)}</td>
        <td>${p.clinical_diagnosis || "-"}</td>
        <td style="max-width: 140px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${sitesStr}">${sitesStr}</td>
        <td>${p.referring_doctor || "-"}</td>
        <td>${p.submission_date || "-"}</td>
        <td>${statusHtml}</td>
        <td>
          <div class="search-action-btn-group">
            <button class="btn btn-secondary btn-tbl-edit" title="修改患者基本信息">✏️ 信息</button>
            <button class="btn btn-secondary btn-tbl-making" title="进入制作工作区编辑报告" style="color: #0284c7; border-color: #bae6fd; background: #f0f9ff;">🔬 制作</button>
            <button class="btn btn-secondary btn-tbl-preview" title="预览Word报告">👁️ 预览</button>
            <button class="btn btn-secondary btn-tbl-paths" title="查看并定位报告涉及的所有文件夹与文件">📂 目录</button>
            <button class="btn btn-secondary btn-tbl-status" title="修改审核/归档等各种状态并同步文件夹">🏷️ 状态</button>
            <button class="btn btn-secondary btn-tbl-delete" title="彻底删除报告及所有相关目录" style="color: #dc2626; border-color: #fecaca; background: #fef2f2;">🗑️ 删除</button>
          </div>
        </td>
      `;

      // 1. Edit basic patient info
      tr.querySelector(".btn-tbl-edit").addEventListener("click", async () => {
        tabs.forEach(t => t.classList.toggle("active", t.dataset.tab === "tab-entry"));
        panels.forEach(pan => pan.classList.toggle("active", pan.id === "tab-entry"));
        await selectPatient(p, true);
        enterEditMode(state.selectedPatient || p);
      });

      // 2. Open in Report Making
      tr.querySelector(".btn-tbl-making").addEventListener("click", () => {
        tabs.forEach(t => t.classList.toggle("active", t.dataset.tab === "tab-making"));
        panels.forEach(pan => pan.classList.toggle("active", pan.id === "tab-making"));
        selectMakingPatient(p, true, true);
      });

      // 3. Preview Word
      tr.querySelector(".btn-tbl-preview").addEventListener("click", () => {
        openReportPreviewModal(p.pathology_no);
      });

      // 4. Open associated paths modal
      tr.querySelector(".btn-tbl-paths").addEventListener("click", () => {
        openPatientPathsModal(p.pathology_no, p.name);
      });

      // 5. Manage status modal
      tr.querySelector(".btn-tbl-status").addEventListener("click", () => {
        openStatusChangeModal(p);
      });

      // 6. Delete patient modal
      tr.querySelector(".btn-tbl-delete").addEventListener("click", () => {
        openDeletePatientModal(p);
      });

      searchResultsTbody.appendChild(tr);
    });
  }

  // ================= Patient Paths Modal Logic =================
  const patientPathsModal = document.getElementById("patient-paths-modal");
  const pathsModalPatientInfo = document.getElementById("paths-modal-patient-info");
  const pathsModalList = document.getElementById("paths-modal-list");
  const btnClosePathsModal = document.getElementById("btn-close-paths-modal");

  async function openPatientPathsModal(pathologyNo, patientName) {
    pathsModalPatientInfo.textContent = `${patientName} (${pathologyNo})`;
    pathsModalList.innerHTML = `<div style="text-align: center; padding: 30px; color: var(--text-muted);">正在检索关联路径...</div>`;
    patientPathsModal.style.display = "flex";

    try {
      const res = await fetch(`/api/patients/${encodeURIComponent(pathologyNo)}/paths`);
      const data = await res.json();
      if (data.success) {
        renderPatientPathsList(data.paths);
      } else {
        pathsModalList.innerHTML = `<div style="color: red; padding: 20px;">检索失败: ${data.error}</div>`;
      }
    } catch (e) {
      pathsModalList.innerHTML = `<div style="color: red; padding: 20px;">请求异常: ${e.message}</div>`;
    }
  }

  function renderPatientPathsList(paths) {
    pathsModalList.innerHTML = "";
    const list = Array.isArray(paths) ? paths : [];

    if (!list.length) {
      pathsModalList.innerHTML = `<div style="text-align: center; padding: 20px; color: var(--text-muted);">暂无关联目录信息</div>`;
      return;
    }

    list.forEach(it => {
      const card = document.createElement("div");
      card.className = "path-item-card";
      const exists = !!it.exists;
      const pathVal = it.path || "-";
      const badge = exists
        ? `<span class="path-badge-exists">✓ 存在</span>`
        : `<span class="path-badge-missing">✕ 未生成</span>`;

      card.innerHTML = `
        <div class="path-item-info">
          <div class="path-item-header">
            <span class="path-item-title">${it.name}</span>
            ${badge}
          </div>
          <div class="path-item-val" title="${pathVal}">${pathVal}</div>
        </div>
        <div class="path-item-actions">
          <button type="button" class="btn btn-secondary btn-sm btn-open-explorer" ${!exists ? "disabled" : ""}>
            📂 定位打开
          </button>
          <button type="button" class="btn btn-secondary btn-sm btn-copy-path" ${pathVal === "-" ? "disabled" : ""}>
            📋 复制路径
          </button>
        </div>
      `;

      const btnOpen = card.querySelector(".btn-open-explorer");
      if (btnOpen && exists) {
        btnOpen.addEventListener("click", async () => {
          await fetch("/api/open_explorer_select", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ path: pathVal })
          });
        });
      }

      const btnCopy = card.querySelector(".btn-copy-path");
      if (btnCopy && pathVal !== "-") {
        btnCopy.addEventListener("click", () => {
          copyToClipboard(pathVal, `已复制路径: ${it.title}`);
        });
      }

      pathsModalList.appendChild(card);
    });
  }

  btnClosePathsModal.addEventListener("click", () => {
    patientPathsModal.style.display = "none";
  });
  patientPathsModal.addEventListener("click", (e) => {
    if (e.target === patientPathsModal) patientPathsModal.style.display = "none";
  });

  // ================= Status Management Modal Logic =================
  const statusChangeModal = document.getElementById("status-change-modal");
  const statusModalPatientInfo = document.getElementById("status-modal-patient-info");
  const statusModalReview = document.getElementById("status-modal-review");
  const statusModalArchived = document.getElementById("status-modal-archived");
  const btnCloseStatusModal = document.getElementById("btn-close-status-modal");
  const statusChangeForm = document.getElementById("status-change-form");
  const btnSubmitStatusModal = document.getElementById("btn-submit-status-modal");
  let activeStatusPatient = null;

  function openStatusChangeModal(patient) {
    activeStatusPatient = patient;
    statusModalPatientInfo.textContent = `患者：${patient.name} | 病理号：${patient.pathology_no}`;
    statusModalReview.value = patient.review_status || "无";
    statusModalArchived.value = String(patient.archived !== undefined && patient.archived !== null ? patient.archived : 0);
    statusChangeModal.style.display = "flex";
  }

  statusChangeForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (!activeStatusPatient) return;
    btnSubmitStatusModal.disabled = true;
    btnSubmitStatusModal.innerHTML = `<span>⏳ 保存并同步中...</span>`;

    try {
      const payload = {
        review_status: statusModalReview.value,
        archived: parseInt(statusModalArchived.value) || 0
      };
      const data = await fetchJsonWithSafetyConfirm(`/api/patients/${encodeURIComponent(activeStatusPatient.pathology_no)}/status`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      if (data.success) {
        showToast(data.message, "success");
        statusChangeModal.style.display = "none";
        doSearch();
        loadMakingPatients();
      } else {
        showToast("修改状态失败: " + data.error, "error");
      }
    } catch (err) {
      showToast("请求异常: " + err.message, "error");
    } finally {
      btnSubmitStatusModal.disabled = false;
      btnSubmitStatusModal.innerHTML = `<span>💾 保存修改并同步文件</span>`;
    }
  });

  btnCloseStatusModal.addEventListener("click", () => {
    statusChangeModal.style.display = "none";
  });
  statusChangeModal.addEventListener("click", (e) => {
    if (e.target === statusChangeModal) statusChangeModal.style.display = "none";
  });

  // ================= Delete Patient Modal Logic =================
  const deletePatientModal = document.getElementById("delete-patient-modal");
  const deleteModalPatientInfo = document.getElementById("delete-modal-patient-info");
  const btnCancelDeleteModal = document.getElementById("btn-cancel-delete-modal");
  const btnConfirmDeleteModal = document.getElementById("btn-confirm-delete-modal");
  let activeDeletePatient = null;

  function openDeletePatientModal(patient) {
    activeDeletePatient = patient;
    deleteModalPatientInfo.textContent = `患者：${patient.name} | 病理号：${patient.pathology_no}`;
    deletePatientModal.style.display = "flex";
  }

  btnConfirmDeleteModal.addEventListener("click", async () => {
    if (!activeDeletePatient) return;
    btnConfirmDeleteModal.disabled = true;
    btnConfirmDeleteModal.innerHTML = `<span>⏳ 正在彻底删除...</span>`;

    try {
      const data = await fetchJsonWithSafetyConfirm(`/api/patients/${encodeURIComponent(activeDeletePatient.pathology_no)}`, {
        method: "DELETE"
      });
      if (data.success) {
        showToast(data.message, "success");
        deletePatientModal.style.display = "none";
        doSearch();
        loadTodayPatients();
        loadMakingPatients();
      } else {
        showToast("删除失败: " + data.error, "error");
      }
    } catch (err) {
      showToast("删除请求异常: " + err.message, "error");
    } finally {
      btnConfirmDeleteModal.disabled = false;
      btnConfirmDeleteModal.innerHTML = `<span>🗑️ 确认彻底删除</span>`;
    }
  });

  btnCancelDeleteModal.addEventListener("click", () => {
    deletePatientModal.style.display = "none";
  });
  deletePatientModal.addEventListener("click", (e) => {
    if (e.target === deletePatientModal) deletePatientModal.style.display = "none";
  });

  // ==========================================================
  // Phase 2: Report Making Workspace (2. 报告制作)
  // ==========================================================
  const makingState = {
    patients: [],
    selectedPatient: null,
    photos: [],
    slots: [], // [{slot: 1, site: "...", photo: photoObj|null}]
    imageCount: 1,
    recommendedTemplates: [],
    defaultsInitialized: false
  };

  const makingFilterDate = document.getElementById("making-filter-date");
  const btnMakingToday = document.getElementById("btn-making-today");
  const btnMakingAllDates = document.getElementById("btn-making-all-dates");
  const makingFilterSubDate = document.getElementById("making-filter-sub-date");
  const btnMakingSubAll = document.getElementById("btn-making-sub-all");
  const makingFilterKeyword = document.getElementById("making-filter-keyword");
  const btnMakingSearch = document.getElementById("btn-making-search");
  const btnRefreshMaking = document.getElementById("btn-refresh-making");
  const makingCountBadge = document.getElementById("making-count-badge");
  const btnMakingBatchArchiveToday = document.getElementById("btn-making-batch-archive-today");
  const makingPatientList = document.getElementById("making-patient-list");

  const makingEmptyHint = document.getElementById("making-empty-hint");
  const makingEditorPanel = document.getElementById("making-editor-panel");

  const mkBadgeName = document.getElementById("mk-badge-name");
  const mkBadgeNo = document.getElementById("mk-badge-no");
  const mkBadgeInfo = document.getElementById("mk-badge-info");
  const mkBadgeDiag = document.getElementById("mk-badge-diag");

  const btnMkFindClinical = document.getElementById("btn-mk-find-clinical");
  const btnMkOpenPhotoDir = document.getElementById("btn-mk-open-photo-dir");
  const btnMkCopyPhotoPath = document.getElementById("btn-mk-copy-photo-path");
  const btnMkOpenWordFile = document.getElementById("btn-mk-open-word-file");

  const btnMkRefreshPhotos = document.getElementById("btn-mk-refresh-photos");
  const btnMkAddSlot = document.getElementById("btn-mk-add-slot");
  const btnMkAutoLayout = document.getElementById("btn-mk-auto-layout");
  const mkAvailPhotosCount = document.getElementById("mk-avail-photos-count");
  const mkAvailablePhotosList = document.getElementById("mk-available-photos-list");
  const mkSlotsStatusDesc = document.getElementById("mk-slots-status-desc");
  const mkSlotsGrid = document.getElementById("mk-slots-grid");

  const btnOpenTemplateModal = document.getElementById("btn-open-template-modal");
  const mkRecommendContainer = document.getElementById("mk-recommend-container");
  const mkRecommendChips = document.getElementById("mk-recommend-chips");
  const mkMicroscopicDesc = document.getElementById("mk-microscopic-desc");

  const btnMkRegenDiag = document.getElementById("btn-mk-regen-diag");
  const mkPathologyDiag = document.getElementById("mk-pathology-diag");

  const mkIsPendingReview = document.getElementById("mk-is-pending-review");
  const mkIsNeedsIhc = document.getElementById("mk-is-needs-ihc");
  const mkIsNeedsRecut = document.getElementById("mk-is-needs-recut");
  const mkSaveStatus = document.getElementById("mk-save-status");

  const btnMkPreview = document.getElementById("btn-mk-preview");
  const btnMkOpenFinalWord = document.getElementById("btn-mk-open-final-word");
  const btnMkOpenExplorer = document.getElementById("btn-mk-open-explorer");
  const btnMkPrintReport = document.getElementById("btn-mk-print-report");
  const btnMkSaveReport = document.getElementById("btn-mk-save-report");
  const btnMkArchiveReport = document.getElementById("btn-mk-archive-report");

  function initMakingDefaultDates() {
    if (!makingState.defaultsInitialized && makingFilterDate) {
      makingFilterDate.value = getTodayDateStr();
      makingState.defaultsInitialized = true;
    }
  }

  // Load patients for Report Making
  async function loadMakingPatients(reloadSelectedPhotos = false) {
    try {
      initMakingDefaultDates();
      const repDateVal = makingFilterDate ? (makingFilterDate.value || "") : "";
      const subDateVal = makingFilterSubDate ? (makingFilterSubDate.value || "") : "";
      const kwVal = makingFilterKeyword.value.trim();
      const params = new URLSearchParams();
      if (repDateVal) params.set("report_date", repDateVal);
      if (subDateVal) params.set("submission_date", subDateVal);
      if (kwVal) params.set("keyword", kwVal);

      makingPatientList.innerHTML = `<div class="patient-list-empty">正在加载报告列表...</div>`;
      const res = await fetch(`/api/patients/report_making?${params.toString()}`);
      const data = await res.json();

      if (data.success) {
        makingState.patients = data.patients;
        makingCountBadge.textContent = `${data.count} 人`;
        renderMakingPatientList();

        // If previously selected patient exists in list, re-select
        if (makingState.selectedPatient) {
          const found = data.patients.find(p => p.pathology_no === makingState.selectedPatient.pathology_no);
          if (found) {
            selectMakingPatient(found, reloadSelectedPhotos);
          }
        }
      } else {
        makingPatientList.innerHTML = `<div class="patient-list-empty" style="color: var(--danger);">${data.error}</div>`;
      }
    } catch (err) {
      makingPatientList.innerHTML = `<div class="patient-list-empty" style="color: var(--danger);">加载失败: ${err.message}</div>`;
    }
  }

  function renderMakingPatientList() {
    if (!makingState.patients.length) {
      makingPatientList.innerHTML = `<div class="patient-list-empty">暂无符合条件的报告记录</div>`;
      return;
    }

    makingPatientList.innerHTML = "";
    makingState.patients.forEach(p => {
      const card = document.createElement("div");
      const isSelected = makingState.selectedPatient && makingState.selectedPatient.pathology_no === p.pathology_no;
      const statusFillClass = getMakingCardStatusClass(p);
      card.className = `patient-list-item ${statusFillClass} ${p.external_modified ? "external-modified-card" : ""} ${isSelected ? "selected" : ""}`.trim();

      const isApprovedArchived = (p.review_status === "完成审核" && p.archived === 1);
      const isApprovedUnarchived = (p.review_status === "完成审核" && p.archived !== 1);

      // 完成审核已归档采用普通已归档报告形式（保留免疫组化标签）；完成审核未归档给出“需重新归档”提示
      const statusBadge = isApprovedArchived
        ? (p.is_needs_ihc === 1 ? `<span class="badge-status badge-ihc">免疫组化</span>` : "")
        : renderReviewStatusBadges(p);

      let archBadge = "";
      if (p.archived === 1) {
        archBadge = `<span class="badge-status badge-archived">已归档</span>`;
      } else if (p.archived === 2 || isApprovedUnarchived) {
        archBadge = `<span class="badge-status badge-rearchive">需重新归档</span>`;
      } else if (p.final_report_path && !statusBadge) {
        archBadge = `<span class="badge-status badge-saved">已保存</span>`;
      }
      const extModBadge = renderExternalModBadge(p, false);

      const diagText = p.pathology_diag || p.clinical_diagnosis || "未诊断";
      card.innerHTML = `
        <div class="patient-item-header">
          <span class="patient-item-name">${p.name}</span>
          <span class="patient-item-no">${p.pathology_no}</span>
        </div>
        <div class="patient-item-meta">
          <span>${p.gender || "-"} / ${formatAgeDisplay(p.age)}</span>
          <span>图: ${p.image_count || 1}张</span>
        </div>
        <div class="patient-item-body">
          <div class="patient-item-diag" title="${diagText}">${diagText}</div>
          ${(statusBadge || archBadge || extModBadge) ? `<div class="patient-item-badges">${extModBadge}${statusBadge}${archBadge}</div>` : ""}
        </div>
      `;

      card.addEventListener("click", () => {
        selectMakingPatient(p, true, true);
      });

      makingPatientList.appendChild(card);
    });
  }

  // Select patient to edit in Report Making
  async function selectMakingPatient(patient, reloadPhotos = true, fromUserClick = false) {
    let effPatient = patient;
    if (fromUserClick && effPatient && effPatient.external_modified) {
      const outcome = await resolveExternalSyncForPatient(effPatient, "normal", false);
      if (outcome && outcome.success && outcome.patient) {
        effPatient = outcome.patient;
      }
    }
    makingState.selectedPatient = effPatient;
    makingEmptyHint.style.display = "none";
    makingEditorPanel.style.display = "flex";

    // Update list selection highlight
    document.querySelectorAll("#making-patient-list .patient-list-item").forEach(el => {
      const noEl = el.querySelector(".patient-item-no");
      if (noEl && noEl.textContent.trim() === effPatient.pathology_no) {
        el.classList.add("selected");
      } else {
        el.classList.remove("selected");
      }
    });

    // Populate banner
    mkBadgeName.textContent = effPatient.name;
    mkBadgeNo.textContent = effPatient.pathology_no;
    const ageStr = effPatient.age ? formatAgeDisplay(effPatient.age) : "";
    mkBadgeInfo.textContent = `${effPatient.gender || ""} · ${ageStr}`;
    mkBadgeDiag.textContent = `临床诊断: ${effPatient.clinical_diagnosis || "无"}`;

    // Status checkboxes
    const revStr = String(effPatient.review_status || "");
    mkIsPendingReview.checked = effPatient.is_pending_review === 1 || revStr.includes("待审核");
    mkIsNeedsIhc.checked = effPatient.is_needs_ihc === 1 || revStr.includes("免疫组化");
    if (mkIsNeedsRecut) {
      mkIsNeedsRecut.checked = effPatient.is_needs_recut === 1 || revStr.includes("重切");
    }

    // Text fields
    mkMicroscopicDesc.value = effPatient.microscopic_desc || "";
    mkPathologyDiag.value = effPatient.pathology_diag || "";
    mkSaveStatus.textContent = "就绪";
    mkSaveStatus.style.color = "var(--text-muted)";

    if (reloadPhotos) {
      await fetchPatientPhotosAndSetupSlots(effPatient);
    }
  }

  // Fetch photos from patient's directory and setup slots
  async function fetchPatientPhotosAndSetupSlots(patient) {
    try {
      mkAvailablePhotosList.innerHTML = `<div class="photo-strip-empty">正在扫描照片...</div>`;
      const res = await fetch(`/api/patients/${encodeURIComponent(patient.pathology_no)}/photos`);
      const data = await res.json();

      if (data.success) {
        makingState.photos = data.photos || [];
        mkAvailPhotosCount.textContent = makingState.photos.length;
        makingState.recommendedTemplates = data.recommended_templates || [];

        renderAvailablePhotosStrip();
        renderRecommendChips(data.recommended_templates);

        // Auto-fill diagnosis if field is currently blank
        if (!mkPathologyDiag.value && data.auto_diagnosis) {
          mkPathologyDiag.value = data.auto_diagnosis;
        }

        // Parse existing selected_images or initialize slots
        initSlotsFromPatientData(patient);
        renderMakingSlots();
      } else {
        mkAvailablePhotosList.innerHTML = `<div class="photo-strip-empty" style="color: var(--danger);">${data.error}</div>`;
      }
    } catch (e) {
      mkAvailablePhotosList.innerHTML = `<div class="photo-strip-empty" style="color: var(--danger);">读取照片失败: ${e.message}</div>`;
    }
  }

  // Render horizontal photos strip
  function renderAvailablePhotosStrip() {
    if (!makingState.photos.length) {
      mkAvailablePhotosList.innerHTML = `<div class="photo-strip-empty">该患者文件夹下暂无显微镜图片（可在病理照片文件夹中存入“序号+诊断.jpg”）</div>`;
      return;
    }

    mkAvailablePhotosList.innerHTML = "";
    makingState.photos.forEach(ph => {
      const card = document.createElement("div");
      card.className = "photo-strip-card";
      const idxText = ph.indices && ph.indices.length ? ph.indices.join(".") : "";
      const tagHtml = idxText ? `<span class="photo-strip-tag">${idxText}</span>` : "";

      card.innerHTML = `
        <img src="${ph.url}" class="photo-strip-img" alt="${ph.filename}">
        <div class="photo-strip-name" title="${ph.filename}">${tagHtml} ${ph.diagnosis || ph.filename}</div>
      `;

      card.addEventListener("click", () => {
        openPhotoZoom(ph.url, `${ph.filename} (${ph.diagnosis || ""})`);
      });

      mkAvailablePhotosList.appendChild(card);
    });
  }

  // Render recommended template chips
  function renderRecommendChips(recs) {
    if (!recs || !recs.length) {
      mkRecommendContainer.style.display = "none";
      mkRecommendChips.innerHTML = "";
      return;
    }

    mkRecommendContainer.style.display = "flex";
    mkRecommendChips.innerHTML = "";

    recs.forEach(t => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "recommend-chip";
      chip.textContent = `+ ${t.name}`;
      chip.title = `点击将【${t.name}】镜下描述填入文本框`;

      chip.addEventListener("click", () => {
        if (mkMicroscopicDesc.value.trim()) {
          mkMicroscopicDesc.value += "\n" + t.desc;
        } else {
          mkMicroscopicDesc.value = t.desc;
        }
        showToast(`已追加推荐模板【${t.name}】至镜下描述`, "success");
      });

      mkRecommendChips.appendChild(chip);
    });
  }

  // Initialize slots
  function initSlotsFromPatientData(patient) {
    const totalSlots = Math.max(parseInt(patient.image_count) || 1, 1);
    makingState.imageCount = totalSlots;

    let biopsySites = patient.biopsy_sites || [];
    if (typeof biopsySites === "string") {
      biopsySites = biopsySites.split(";").map(s => s.trim()).filter(Boolean);
    }

    let existingSelected = [];
    if (patient.selected_images) {
      if (typeof patient.selected_images === "string") {
        try { existingSelected = JSON.parse(patient.selected_images); } catch (e) { existingSelected = []; }
      } else if (Array.isArray(patient.selected_images)) {
        existingSelected = patient.selected_images;
      }
    }

    makingState.slots = [];
    for (let i = 1; i <= totalSlots; i++) {
      const siteName = (biopsySites.length >= i) ? biopsySites[i - 1] : "";
      let matchedPhoto = null;

      // Check if slot had previously saved photo
      if (existingSelected.length) {
        const item = existingSelected.find((it, idx) => {
          if (typeof it === "object" && it !== null) {
            return it.slot === i;
          }
          return idx === i - 1;
        });

        if (item) {
          const pPath = typeof item === "object" ? item.path : item;
          matchedPhoto = findMatchingPhoto(makingState.photos, pPath);
        }
      }

      // If no saved photo, auto-match by photo indices (e.g. 1 -> photo with index 1)
      if (!matchedPhoto && makingState.photos.length) {
        matchedPhoto = makingState.photos.find(ph => {
          if (ph.indices && ph.indices.includes(String(i))) return true;
          return ph.index_num === i;
        });
      }

      makingState.slots.push({
        slot: i,
        site: siteName,
        photo: matchedPhoto || null
      });
    }
  }

  // Render slots grid in Report Making
  function renderMakingSlots() {
    mkSlotsStatusDesc.textContent = `当前共 ${makingState.slots.length} 个排版插槽 (对应Word第1页排版规则)`;
    mkSlotsGrid.innerHTML = "";

    makingState.slots.forEach((s, idx) => {
      const card = document.createElement("div");
      card.className = "slot-card";

      const hasPhoto = s.photo !== null;
      let previewHtml = "";
      if (hasPhoto) {
        previewHtml = `
          <div class="slot-preview-box has-photo" title="点击放大图片">
            <img src="${s.photo.url}" class="slot-preview-img" alt="插槽 ${s.slot}">
          </div>
        `;
      } else {
        previewHtml = `
          <div class="slot-preview-box">
            <div class="slot-empty-text">
              <span>🖼️</span>
              <span>待选择图片</span>
            </div>
          </div>
        `;
      }

      // Build options for select dropdown
      let selectOptions = `<option value="">-- 未选择图片 --</option>`;
      makingState.photos.forEach(ph => {
        const isSel = s.photo && s.photo.path === ph.path;
        const idxT = ph.indices && ph.indices.length ? `[${ph.indices.join(".")}] ` : "";
        selectOptions += `<option value="${ph.path}" ${isSel ? "selected" : ""}>${idxT}${ph.diagnosis || ph.filename}</option>`;
      });

      card.innerHTML = `
        <div class="slot-card-header">
          <span>第 ${s.slot} 张图片</span>
          ${makingState.slots.length > 1 ? `<button type="button" class="btn btn-secondary btn-sm btn-remove-slot" style="padding: 1px 6px; font-size: 11px;" title="删除此预留位">✕</button>` : ""}
        </div>
        ${previewHtml}
        <div class="form-group" style="margin-bottom: 4px;">
          <input type="text" class="slot-site-input" placeholder="部位名称 (如: 面部)" value="${s.site || ""}">
        </div>
        <div class="form-group" style="margin-bottom: 2px;">
          <select class="slot-photo-select">
            ${selectOptions}
          </select>
        </div>
        <div class="slot-actions">
          ${hasPhoto ? `<button type="button" class="btn btn-secondary btn-sm btn-zoom-slot" style="padding: 2px 6px; font-size: 11px;">🔍 放大</button>` : `<span></span>`}
          ${hasPhoto ? `<button type="button" class="btn btn-secondary btn-sm btn-clear-slot-photo" style="padding: 2px 6px; font-size: 11px; color: var(--danger);">移除图片</button>` : `<span></span>`}
        </div>
      `;

      // Event listeners
      const siteInput = card.querySelector(".slot-site-input");
      siteInput.addEventListener("input", (e) => {
        s.site = e.target.value;
      });

      const photoSelect = card.querySelector(".slot-photo-select");
      photoSelect.addEventListener("change", (e) => {
        const chosenPath = e.target.value;
        s.photo = makingState.photos.find(p => p.path === chosenPath) || null;
        renderMakingSlots();
      });

      if (hasPhoto) {
        const previewBox = card.querySelector(".slot-preview-box");
        previewBox.addEventListener("click", () => {
          openPhotoZoom(s.photo.url, `第 ${s.slot} 张: ${s.photo.filename}`);
        });

        const zoomBtn = card.querySelector(".btn-zoom-slot");
        if (zoomBtn) {
          zoomBtn.addEventListener("click", () => {
            openPhotoZoom(s.photo.url, `第 ${s.slot} 张: ${s.photo.filename}`);
          });
        }

        const clearBtn = card.querySelector(".btn-clear-slot-photo");
        if (clearBtn) {
          clearBtn.addEventListener("click", () => {
            s.photo = null;
            renderMakingSlots();
          });
        }
      }

      const removeBtn = card.querySelector(".btn-remove-slot");
      if (removeBtn) {
        removeBtn.addEventListener("click", () => {
          removeMakingSlot(idx);
        });
      }

      mkSlotsGrid.appendChild(card);
    });
  }

  // Add slot button
  btnMkAddSlot.addEventListener("click", () => {
    const nextSlotNo = makingState.slots.length + 1;
    if (nextSlotNo > 9) {
      showToast("报告模板最多支持 9 张图片预留位", "error");
      return;
    }
    makingState.slots.push({
      slot: nextSlotNo,
      site: "",
      photo: null
    });
    makingState.imageCount = makingState.slots.length;
    renderMakingSlots();
    showToast(`已增加第 ${nextSlotNo} 个图片预留位`, "success");
  });

  // Remove slot
  function removeMakingSlot(slotIndex) {
    if (makingState.slots.length <= 1) return;
    makingState.slots.splice(slotIndex, 1);
    // Renumber remaining slots
    makingState.slots.forEach((s, i) => s.slot = i + 1);
    makingState.imageCount = makingState.slots.length;
    renderMakingSlots();
  }

  // Auto layout button: re-maps photos by index
  btnMkAutoLayout.addEventListener("click", () => {
    makingState.slots.forEach(s => {
      const matched = makingState.photos.find(ph => {
        if (ph.indices && ph.indices.includes(String(s.slot))) return true;
        return ph.index_num === s.slot;
      });
      if (matched) {
        s.photo = matched;
      }
    });
    renderMakingSlots();
    showToast("已按照图片编号完成一键自动排版！", "success");
  });

  // Re-generate diagnosis button
  btnMkRegenDiag.addEventListener("click", async () => {
    if (!makingState.selectedPatient) return;
    try {
      const slotCnt = makingState.slots.length || makingState.imageCount || 1;
      const res = await fetch(`/api/patients/${encodeURIComponent(makingState.selectedPatient.pathology_no)}/photos?image_count=${slotCnt}`);
      const data = await res.json();
      if (data.success && data.auto_diagnosis) {
        mkPathologyDiag.value = data.auto_diagnosis;
        showToast("已根据图片命名重新合成病理诊断！", "success");
      } else {
        showToast("未能从图片命名中提取出有效诊断", "error");
      }
    } catch (e) {
      showToast("合成诊断失败: " + e.message, "error");
    }
  });

  // Refresh photos button
  btnMkRefreshPhotos.addEventListener("click", async () => {
    if (!makingState.selectedPatient) return;
    await fetchPatientPhotosAndSetupSlots(makingState.selectedPatient);
    showToast("已重新扫描病理照片目录！", "success");
  });

  // Save Report (2.7)
  btnMkSaveReport.addEventListener("click", async () => {
    if (!makingState.selectedPatient) return;
    const p = makingState.selectedPatient;

    const selectedImgs = makingState.slots.map(s => ({
      slot: s.slot,
      path: s.photo ? s.photo.path : null,
      site: (s.site || "").trim()
    }));

    const biopsySites = makingState.slots.map(s => (s.site || "").trim());

    const payload = {
      pathology_no: p.pathology_no,
      microscopic_desc: mkMicroscopicDesc.value.trim(),
      pathology_diag: mkPathologyDiag.value.trim(),
      selected_images: selectedImgs,
      image_count: makingState.slots.length,
      biopsy_sites: biopsySites,
      is_pending_review: mkIsPendingReview.checked,
      is_needs_ihc: mkIsNeedsIhc.checked,
      is_needs_recut: mkIsNeedsRecut ? mkIsNeedsRecut.checked : false
    };

    btnMkSaveReport.disabled = true;
    btnMkSaveReport.innerHTML = `<span>⏳ 保存中...</span>`;

    try {
      const data = await fetchJsonWithSafetyConfirm("/api/reports/save", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      if (data.success) {
        showToast(data.message, "success");
        if (data.duplicate_clinical_photos) {
          alert(data.duplicate_prompt || "【提示】该患者在临床照片母目录中匹配到多个重名文件夹，无法自动归档临床照片！\n已在资源管理器中同时打开【临床照片母目录】和【特殊归档目标目录】，请操作者手动筛选并将正确的照片文件夹归档复制。");
        }
        makingState.selectedPatient = data.patient;
        if (data.synced_forward && data.patient) {
          await selectMakingPatient(data.patient, true, false);
        }
        mkSaveStatus.textContent = data.synced_forward
          ? `已导入外部Word修改 (${new Date().toLocaleTimeString()})`
          : `已保存并输出 (${new Date().toLocaleTimeString()})`;
        mkSaveStatus.style.color = "var(--success)";
        loadMakingPatients();
      } else {
        showToast(`保存失败: ${data.error}`, "error");
      }
    } catch (e) {
      showToast(`网络或系统异常: ${e.message}`, "error");
    } finally {
      btnMkSaveReport.disabled = false;
      btnMkSaveReport.innerHTML = `<span>💾 保存报告并输出 (2.7)</span>`;
    }
  });

  // Archive Report (2.8 & Requirement 3)
  btnMkArchiveReport.addEventListener("click", async () => {
    if (!makingState.selectedPatient) return;
    const p = makingState.selectedPatient;

    // Check prohibition for 待审核 and 重切 (unless already 完成审核; IHC-only follows regular archiving)
    const isApproved = (p.review_status === "完成审核") && !mkIsPendingReview.checked && !(mkIsNeedsRecut && mkIsNeedsRecut.checked);
    const isPending = mkIsPendingReview.checked || (!isApproved && (p.is_pending_review == 1 || (p.review_status && p.review_status.includes("待审核"))));
    const isRecut = (!isApproved && ((mkIsNeedsRecut && mkIsNeedsRecut.checked) || p.is_needs_recut == 1 || (p.review_status && p.review_status.includes("重切"))));

    if (isPending || isRecut) {
      alert("⚠️ 【禁止直接归档】\n\n该报告已被标记为【待审核 / 重切】！\n在点击保存时，系统已自动将其放入【特殊患者归档】待审队列。\n\n操作指引：请移步【报告审核】工作区由上级医生完成专家审核后，方可进行归档。");
      return;
    }

    btnMkArchiveReport.disabled = true;
    btnMkArchiveReport.innerHTML = `<span>⏳ 归档中...</span>`;

    try {
      const data = await fetchJsonWithSafetyConfirm("/api/reports/archive", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pathology_no: p.pathology_no, source: isApproved ? "review" : "maker" })
      });

      if (data.success) {
        showToast(`患者 [${p.name}] 报告归档成功！已保存至【已归档报告】。`, "success");
        if (data.duplicate_clinical_photos) {
          alert(data.duplicate_prompt || "【提示】该患者在临床照片母目录中匹配到多个重名文件夹，无法自动归档临床照片！\n已在资源管理器中同时打开【临床照片母目录】和【特殊归档目标目录】，请操作者手动筛选并将正确的照片文件夹归档复制。");
        }
        makingState.selectedPatient = data.patient;
        loadMakingPatients();
      } else {
        alert("归档失败: " + data.error);
        showToast(`归档失败: ${data.error}`, "error");
      }
    } catch (e) {
      showToast(`归档异常: ${e.message}`, "error");
    } finally {
      btnMkArchiveReport.disabled = false;
      btnMkArchiveReport.innerHTML = `<span>📦 一键归档 (2.8)</span>`;
    }
  });

  // Preview button in Making
  btnMkPreview.addEventListener("click", () => {
    if (makingState.selectedPatient) {
      openReportPreviewModal(makingState.selectedPatient.pathology_no);
    }
  });

  // Open Word button in Making
  btnMkOpenFinalWord.addEventListener("click", async () => {
    if (!makingState.selectedPatient) return;
    const reportPath = makingState.selectedPatient.final_report_path || makingState.selectedPatient.word_path;
    if (reportPath) {
      await fetch("/api/open_file", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ path: reportPath })
      });
    } else {
      showToast("当前报告文件尚未生成", "error");
    }
  });

  // Open in Windows Explorer button in Making (Requirement 2)
  btnMkOpenExplorer.addEventListener("click", async () => {
    if (!makingState.selectedPatient) return;
    const reportPath = makingState.selectedPatient.final_report_path || makingState.selectedPatient.word_path;
    if (reportPath) {
      try {
        await fetch("/api/open_explorer_select", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ path: reportPath })
        });
      } catch (err) {
        showToast("打开资源管理器异常: " + err.message, "error");
      }
    } else {
      showToast("当前报告文件尚未生成", "error");
    }
  });

  // Print Report button in Making (Requirement 2: auto-saves and prints Page 1 only)
  btnMkPrintReport.addEventListener("click", async () => {
    if (!makingState.selectedPatient) return;
    const p = makingState.selectedPatient;

    const selectedImgs = makingState.slots.map(s => ({
      slot: s.slot,
      path: s.photo ? s.photo.path : null,
      site: (s.site || "").trim()
    }));

    const payload = {
      pathology_no: p.pathology_no,
      microscopic_desc: mkMicroscopicDesc.value.trim(),
      pathology_diag: mkPathologyDiag.value.trim(),
      selected_images: selectedImgs,
      image_count: makingState.slots.length,
      biopsy_sites: makingState.slots.map(s => (s.site || "").trim()),
      is_pending_review: mkIsPendingReview.checked,
      is_needs_ihc: mkIsNeedsIhc.checked,
      is_needs_recut: mkIsNeedsRecut ? mkIsNeedsRecut.checked : false
    };

    btnMkPrintReport.disabled = true;
    btnMkPrintReport.innerHTML = `<span>⏳ 正在保存并准备打印...</span>`;

    try {
      const data = await fetchJsonWithSafetyConfirm("/api/reports/print", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      if (data.success) {
        makingState.selectedPatient = data.patient;
        if (data.synced_forward && data.patient) {
          await selectMakingPatient(data.patient, true, false);
          mkSaveStatus.textContent = `已导入外部Word修改 (${new Date().toLocaleTimeString()})`;
          mkSaveStatus.style.color = "var(--success)";
          showToast(data.message, "success");
          loadMakingPatients();
          return;
        }
        mkSaveStatus.textContent = `已保存并打印 (${new Date().toLocaleTimeString()})`;
        mkSaveStatus.style.color = "var(--success)";
        loadMakingPatients();

        if (data.direct_printed) {
          showToast("报告已保存，已调用Word发送第1页至默认打印机！", "success");
        } else {
          showToast(data.message || "调用Word打印失败，请检查Word及默认打印机", "error");
        }
      } else {
        showToast("打印失败: " + data.error, "error");
      }
    } catch (e) {
      showToast("打印请求异常: " + e.message, "error");
    } finally {
      btnMkPrintReport.disabled = false;
      btnMkPrintReport.innerHTML = `<span>🖨️ 一键打印第一页</span>`;
    }
  });

  // Find clinical photo button (2.2)
  btnMkFindClinical.addEventListener("click", () => {
    findClinicalPhotosForPatient(makingState.selectedPatient);
  });

  // Open photo dir button in Making
  btnMkOpenPhotoDir.addEventListener("click", async () => {
    if (!makingState.selectedPatient || !makingState.selectedPatient.photo_dir) return;
    await fetch("/api/open_dir", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: makingState.selectedPatient.photo_dir })
    });
  });

  // Copy photo dir path button in Making
  if (btnMkCopyPhotoPath) {
    btnMkCopyPhotoPath.addEventListener("click", () => {
      if (!makingState.selectedPatient || !makingState.selectedPatient.photo_dir) return;
      copyToClipboard(makingState.selectedPatient.photo_dir, "已复制病理照片文件夹完整路径！");
    });
  }

  // Batch archive today's saved-unarchived reports (normal & IHC-only, excluding 待审核/重切)
  if (btnMakingBatchArchiveToday) {
    btnMakingBatchArchiveToday.addEventListener("click", async () => {
      btnMakingBatchArchiveToday.disabled = true;
      btnMakingBatchArchiveToday.innerHTML = `<span>⏳ 正在检索今日可归档报告...</span>`;
      try {
        const preRes = await fetch("/api/reports/batch_archive_today", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ dry_run: true })
        });
        const preData = await preRes.json();
        if (!preData.success) {
          showToast("检索可归档报告失败: " + (preData.error || "未知错误"), "error");
          return;
        }
        const cands = preData.candidates || [];
        if (!cands.length) {
          showToast("今日（报告日期为今日）暂无已保存未归档的可归档报告（普通及仅免疫组化）", "info");
          return;
        }
        const previewLines = cands.slice(0, 8).map(c => `• ${c.pathology_no}  ${c.name}`).join("\n");
        const extraHint = cands.length > 8 ? `\n... 等共 ${cands.length} 份` : "";
        const ok = confirm(
          `检测到今日（${preData.today}）共有 ${cands.length} 份已保存未归档报告（含普通报告及仅免疫组化报告，不含待审核/重切）：\n\n` +
          `${previewLines}${extraHint}\n\n确定要一键全部归档吗？`
        );
        if (!ok) return;

        btnMakingBatchArchiveToday.innerHTML = `<span>⏳ 正在批量归档 ${cands.length} 份...</span>`;
        const data = await fetchJsonWithSafetyConfirm("/api/reports/batch_archive_today", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({})
        });
        if (data.success) {
          showToast(data.message, data.failed_count ? "info" : "success");
        } else {
          showToast(data.message || data.error || "批量归档失败", "error");
        }
        if (data.failed && data.failed.length) {
          const failMsg = data.failed.map(f => `${f.pathology_no} (${f.name}): ${f.error}`).join("\n");
          alert(`以下报告归档失败，请检查是否被Word占用：\n\n${failMsg}`);
        }
        if (makingState.selectedPatient && data.archived_nos && data.archived_nos.includes(makingState.selectedPatient.pathology_no)) {
          makingState.selectedPatient.archived = 1;
        }
        await loadMakingPatients();
      } catch (err) {
        showToast("批量归档请求异常: " + err.message, "error");
      } finally {
        btnMakingBatchArchiveToday.disabled = false;
        btnMakingBatchArchiveToday.innerHTML = `<span>📦 一键归档今日已保存未归档报告</span>`;
      }
    });
  }

  // Open draft word file button in Making
  btnMkOpenWordFile.addEventListener("click", async () => {
    if (!makingState.selectedPatient || !makingState.selectedPatient.word_path) return;
    await fetch("/api/open_file", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: makingState.selectedPatient.word_path })
    });
  });

  // Date and filter buttons
  bindSidebarListFilter({
    reportDateEl: makingFilterDate,
    btnReportTodayEl: btnMakingToday,
    btnReportAllEl: btnMakingAllDates,
    subDateEl: makingFilterSubDate,
    btnSubAllEl: btnMakingSubAll,
    keywordEl: makingFilterKeyword,
    btnSearchEl: btnMakingSearch,
    onReload: loadMakingPatients
  });

  btnRefreshMaking.addEventListener("click", () => {
    loadMakingPatients(true);
  });

  // ==========================================================
  // Phase 2: Report Review Workspace (2.9 独立的报告审核页面)
  // ==========================================================
  const reviewState = {
    patients: [],
    selectedPatient: null,
    photos: [],
    slots: [],
    reviewPhotoDir: ""
  };

  const reviewCountBadge = document.getElementById("review-count-badge");
  const btnRefreshReview = document.getElementById("btn-refresh-review");
  const reviewShowApprovedToggle = document.getElementById("review-show-approved-toggle");
  const reviewFilterReportDate = document.getElementById("review-filter-report-date");
  const btnReviewReportToday = document.getElementById("btn-review-report-today");
  const btnReviewReportAll = document.getElementById("btn-review-report-all");
  const reviewFilterSubDate = document.getElementById("review-filter-sub-date");
  const btnReviewSubAll = document.getElementById("btn-review-sub-all");
  const reviewFilterKeyword = document.getElementById("review-filter-keyword");
  const btnReviewSearch = document.getElementById("btn-review-search");
  const reviewPatientList = document.getElementById("review-patient-list");

  const reviewEmptyHint = document.getElementById("review-empty-hint");
  const reviewEditorPanel = document.getElementById("review-editor-panel");

  const revBadgeName = document.getElementById("rev-badge-name");
  const revBadgeNo = document.getElementById("rev-badge-no");
  const revBadgeStatus = document.getElementById("rev-badge-status");
  const revBadgeDiag = document.getElementById("rev-badge-diag");

  const btnRevFindClinical = document.getElementById("btn-rev-find-clinical");
  const btnRevOpenPhotoDir = document.getElementById("btn-rev-open-photo-dir");
  const btnRevCopyPhotoPath = document.getElementById("btn-rev-copy-photo-path");
  const revPhotoPathBox = document.getElementById("rev-photo-path-box");
  const revPhotoPathText = document.getElementById("rev-photo-path-text");

  const revReportDoc = document.getElementById("rev-report-doc");
  const revReviewDoc = document.getElementById("rev-review-doc");
  const revReportDate = document.getElementById("rev-report-date");

  const btnRevRefreshPhotos = document.getElementById("btn-rev-refresh-photos");
  const btnRevAddSlot = document.getElementById("btn-rev-add-slot");
  const btnRevAutoLayout = document.getElementById("btn-rev-auto-layout");
  const revSlotsGrid = document.getElementById("rev-slots-grid");

  const btnRevOpenTemplateModal = document.getElementById("btn-rev-open-template-modal");
  const revMicroscopicDesc = document.getElementById("rev-microscopic-desc");
  const revPathologyDiag = document.getElementById("rev-pathology-diag");
  const revSaveStatus = document.getElementById("rev-save-status");

  const btnRevPreview = document.getElementById("btn-rev-preview");
  const btnRevOpenWord = document.getElementById("btn-rev-open-word");
  const btnRevPrintReport = document.getElementById("btn-rev-print-report");
  const btnRevSaveOverwrite = document.getElementById("btn-rev-save-overwrite");
  const btnRevApprove = document.getElementById("btn-rev-approve");

  // Load review patients
  async function loadReviewPatients(reloadSelectedPhotos = false) {
    try {
      const includeApproved = reviewShowApprovedToggle.checked ? "1" : "0";
      const params = new URLSearchParams();
      params.set("include_approved", includeApproved);
      if (reviewFilterReportDate && reviewFilterReportDate.value) {
        params.set("report_date", reviewFilterReportDate.value);
      }
      if (reviewFilterSubDate && reviewFilterSubDate.value) {
        params.set("submission_date", reviewFilterSubDate.value);
      }
      if (reviewFilterKeyword && reviewFilterKeyword.value.trim()) {
        params.set("keyword", reviewFilterKeyword.value.trim());
      }
      reviewPatientList.innerHTML = `<div class="patient-list-empty">正在加载待审核数据...</div>`;

      const res = await fetch(`/api/reviews/list?${params.toString()}`);
      const data = await res.json();

      if (data.success) {
        const list = data.patients || [];
        reviewState.patients = list;
        reviewCountBadge.textContent = `${list.length} 人`;
        renderReviewPatientList();

        if (reviewState.selectedPatient) {
          const found = list.find(p => p.pathology_no === reviewState.selectedPatient.pathology_no);
          if (found) {
            selectReviewPatient(found, reloadSelectedPhotos);
          }
        }
      } else {
        reviewPatientList.innerHTML = `<div class="patient-list-empty" style="color: var(--danger);">${data.error}</div>`;
      }
    } catch (e) {
      reviewPatientList.innerHTML = `<div class="patient-list-empty" style="color: var(--danger);">加载失败: ${e.message}</div>`;
    }
  }

  function renderReviewPatientList() {
    if (!reviewState.patients.length) {
      reviewPatientList.innerHTML = `<div class="patient-list-empty">暂无待审核或重切报告</div>`;
      return;
    }

    reviewPatientList.innerHTML = "";
    reviewState.patients.forEach(p => {
      const card = document.createElement("div");
      const isSelected = reviewState.selectedPatient && reviewState.selectedPatient.pathology_no === p.pathology_no;
      const statusFillClass = getMakingCardStatusClass(p);
      card.className = `patient-list-item ${statusFillClass} ${p.external_modified ? "external-modified-card" : ""} ${isSelected ? "selected" : ""}`.trim();

      const badgeHtml = `${renderExternalModBadge(p, false)}${renderReviewStatusBadges(p)}`;

      card.innerHTML = `
        <div class="patient-item-header">
          <span class="patient-item-name">${p.name}</span>
          <span class="patient-item-no">${p.pathology_no}</span>
        </div>
        <div class="patient-item-body">
          <div class="patient-item-diag">${p.pathology_diag || p.clinical_diagnosis || "未诊断"}</div>
          <div class="patient-item-badges">${badgeHtml}</div>
        </div>
      `;

      card.addEventListener("click", () => {
        selectReviewPatient(p, true, true);
      });

      reviewPatientList.appendChild(card);
    });
  }

  // Select patient in Review Workspace
  async function selectReviewPatient(patient, reloadPhotos = true, fromUserClick = false) {
    let effPatient = patient;
    if (fromUserClick && effPatient && effPatient.external_modified) {
      const outcome = await resolveExternalSyncForPatient(effPatient, "normal", false);
      if (outcome && outcome.success && outcome.patient) {
        effPatient = outcome.patient;
      }
    }
    reviewState.selectedPatient = effPatient;
    reviewEmptyHint.style.display = "none";
    reviewEditorPanel.style.display = "flex";

    document.querySelectorAll("#review-patient-list .patient-list-item").forEach(el => {
      const noEl = el.querySelector(".patient-item-no");
      if (noEl && noEl.textContent.trim() === effPatient.pathology_no) {
        el.classList.add("selected");
      } else {
        el.classList.remove("selected");
      }
    });

    revBadgeName.textContent = effPatient.name;
    revBadgeNo.textContent = effPatient.pathology_no;
    revBadgeStatus.textContent = effPatient.review_status || "待审核";
    revBadgeDiag.textContent = `诊断: ${effPatient.pathology_diag || effPatient.clinical_diagnosis || "未诊断"}`;

    // Doctor & date overrides (Requirement 2.9 & 3.1)
    revReportDoc.value = effPatient.report_doctor || state.config.report_doctor || "张医生";
    revReviewDoc.value = effPatient.review_doctor || state.config.review_doctor || "李主任";
    revReportDate.value = toIsoDateString(effPatient.report_date_raw || effPatient.report_date);

    revMicroscopicDesc.value = effPatient.microscopic_desc || "";
    revPathologyDiag.value = effPatient.pathology_diag || "";
    if (effPatient.archived === 2) {
      revSaveStatus.textContent = "已同步草稿，请至【报告制作】重新归档";
      revSaveStatus.style.color = "var(--warning)";
    } else {
      revSaveStatus.textContent = "就绪";
      revSaveStatus.style.color = "var(--text-muted)";
    }

    const copiedPhotoDir = effPatient.review_photo_dir || effPatient.photo_dir || "";
    reviewState.reviewPhotoDir = copiedPhotoDir;
    if (revPhotoPathBox && revPhotoPathText) {
      if (copiedPhotoDir) {
        revPhotoPathBox.style.display = "flex";
        revPhotoPathText.textContent = copiedPhotoDir;
        revPhotoPathText.title = copiedPhotoDir;
      } else {
        revPhotoPathBox.style.display = "none";
      }
    }

    if (reloadPhotos) {
      await fetchReviewPhotosAndSlots(effPatient);
    }
  }

  async function fetchReviewPhotosAndSlots(patient) {
    try {
      const res = await fetch(`/api/reviews/${encodeURIComponent(patient.pathology_no)}/photos`);
      const data = await res.json();
      if (data.success) {
        reviewState.photos = data.photos || [];
        reviewState.reviewPhotoDir = data.photo_dir || patient.review_photo_dir || patient.photo_dir || "";
        if (revPhotoPathBox && revPhotoPathText) {
          if (reviewState.reviewPhotoDir) {
            revPhotoPathBox.style.display = "flex";
            revPhotoPathText.textContent = reviewState.reviewPhotoDir;
            revPhotoPathText.title = reviewState.reviewPhotoDir;
          } else {
            revPhotoPathBox.style.display = "none";
          }
        }

        // Setup slots
        const totalSlots = Math.max(parseInt(patient.image_count) || 1, 1);
        let biopsySites = patient.biopsy_sites || [];
        if (typeof biopsySites === "string") {
          biopsySites = biopsySites.split(";").map(s => s.trim()).filter(Boolean);
        }

        let existingSelected = [];
        if (patient.selected_images) {
          if (typeof patient.selected_images === "string") {
            try { existingSelected = JSON.parse(patient.selected_images); } catch (e) {}
          } else if (Array.isArray(patient.selected_images)) {
            existingSelected = patient.selected_images;
          }
        }

        reviewState.slots = [];
        for (let i = 1; i <= totalSlots; i++) {
          const siteName = (biopsySites.length >= i) ? biopsySites[i - 1] : "";
          let matchedPhoto = null;

          if (existingSelected.length) {
            const item = existingSelected.find((it, idx) => {
              if (typeof it === "object" && it !== null) return it.slot === i;
              return idx === i - 1;
            });
            if (item) {
              const pPath = typeof item === "object" ? item.path : item;
              matchedPhoto = findMatchingPhoto(reviewState.photos, pPath);
            }
          }

          if (!matchedPhoto && reviewState.photos.length) {
            matchedPhoto = reviewState.photos.find(ph => (ph.indices && ph.indices.includes(String(i))) || ph.index_num === i);
          }

          reviewState.slots.push({
            slot: i,
            site: siteName,
            photo: matchedPhoto || null
          });
        }

        renderReviewSlots();
      }
    } catch (e) {
      showToast("读取审核图片失败: " + e.message, "error");
    }
  }

  function renderReviewSlots() {
    revSlotsGrid.innerHTML = "";
    reviewState.slots.forEach((s, idx) => {
      const card = document.createElement("div");
      card.className = "slot-card";

      const hasPhoto = s.photo !== null;
      let previewHtml = hasPhoto
        ? `<div class="slot-preview-box has-photo" title="点击放大图片"><img src="${s.photo.url}" class="slot-preview-img" alt="插槽 ${s.slot}"></div>`
        : `<div class="slot-preview-box"><div class="slot-empty-text"><span>🖼️</span><span>待选择图片</span></div></div>`;

      let selectOptions = `<option value="">-- 未选择图片 --</option>`;
      reviewState.photos.forEach(ph => {
        const isSel = s.photo && s.photo.path === ph.path;
        const idxT = ph.indices && ph.indices.length ? `[${ph.indices.join(".")}] ` : "";
        selectOptions += `<option value="${ph.path}" ${isSel ? "selected" : ""}>${idxT}${ph.diagnosis || ph.filename}</option>`;
      });

      card.innerHTML = `
        <div class="slot-card-header">
          <span>第 ${s.slot} 张图片</span>
          ${reviewState.slots.length > 1 ? `<button type="button" class="btn btn-secondary btn-sm btn-remove-rev-slot" style="padding: 1px 6px; font-size: 11px;">✕</button>` : ""}
        </div>
        ${previewHtml}
        <div class="form-group" style="margin-bottom: 4px;">
          <input type="text" class="slot-site-input" placeholder="部位名称" value="${s.site || ""}">
        </div>
        <div class="form-group" style="margin-bottom: 2px;">
          <select class="slot-photo-select">${selectOptions}</select>
        </div>
      `;

      card.querySelector(".slot-site-input").addEventListener("input", e => s.site = e.target.value);
      card.querySelector(".slot-photo-select").addEventListener("change", e => {
        s.photo = reviewState.photos.find(p => p.path === e.target.value) || null;
        renderReviewSlots();
      });

      if (hasPhoto) {
        card.querySelector(".slot-preview-box").addEventListener("click", () => {
          openPhotoZoom(s.photo.url, `第 ${s.slot} 张: ${s.photo.filename}`);
        });
      }

      const removeBtn = card.querySelector(".btn-remove-rev-slot");
      if (removeBtn) {
        removeBtn.addEventListener("click", () => {
          if (reviewState.slots.length > 1) {
            reviewState.slots.splice(idx, 1);
            reviewState.slots.forEach((item, i) => item.slot = i + 1);
            renderReviewSlots();
          }
        });
      }

      revSlotsGrid.appendChild(card);
    });
  }

  btnRevAddSlot.addEventListener("click", () => {
    const next = reviewState.slots.length + 1;
    if (next > 9) return showToast("最多支持 9 张插槽", "error");
    reviewState.slots.push({ slot: next, site: "", photo: null });
    renderReviewSlots();
  });

  btnRevAutoLayout.addEventListener("click", () => {
    reviewState.slots.forEach(s => {
      const matched = reviewState.photos.find(ph => (ph.indices && ph.indices.includes(String(s.slot))) || ph.index_num === s.slot);
      if (matched) s.photo = matched;
    });
    renderReviewSlots();
    showToast("已按照图片编号完成一键自动排版！", "success");
  });

  btnRevRefreshPhotos.addEventListener("click", () => {
    if (reviewState.selectedPatient) {
      fetchReviewPhotosAndSlots(reviewState.selectedPatient);
    }
  });

  function buildReviewSavePayload(p) {
    const selectedImgs = reviewState.slots.map(s => ({
      slot: s.slot,
      path: s.photo ? s.photo.path : null,
      site: (s.site || "").trim()
    }));

    return {
      pathology_no: p.pathology_no,
      report_doctor: revReportDoc.value.trim(),
      review_doctor: revReviewDoc.value.trim(),
      report_date: revReportDate.value.trim(),
      microscopic_desc: revMicroscopicDesc.value.trim(),
      pathology_diag: revPathologyDiag.value.trim(),
      selected_images: selectedImgs,
      image_count: reviewState.slots.length,
      biopsy_sites: reviewState.slots.map(s => (s.site || "").trim()),
      is_pending_review: p.is_pending_review,
      is_needs_ihc: p.is_needs_ihc,
      is_needs_recut: p.is_needs_recut
    };
  }

  function applyReviewSaveResponseToUi(data, statusPrefix = "已同步草稿") {
    reviewState.selectedPatient = data.patient;
    if (data.patient && data.patient.review_photo_dir) {
      reviewState.reviewPhotoDir = data.patient.review_photo_dir;
      if (revPhotoPathBox && revPhotoPathText) {
        revPhotoPathBox.style.display = "flex";
        revPhotoPathText.textContent = reviewState.reviewPhotoDir;
        revPhotoPathText.title = reviewState.reviewPhotoDir;
      }
    }
    if (data.patient && data.patient.archived === 2) {
      revSaveStatus.textContent = `${statusPrefix}，请至【报告制作】重新归档 (${new Date().toLocaleTimeString()})`;
      revSaveStatus.style.color = "var(--warning)";
    } else {
      revSaveStatus.textContent = `${statusPrefix} (${new Date().toLocaleTimeString()})`;
      revSaveStatus.style.color = "var(--success)";
    }
    loadReviewPatients(false);
  }

  // Save and overwrite across output, patient archive, and special archives (2.9)
  btnRevSaveOverwrite.addEventListener("click", async () => {
    if (!reviewState.selectedPatient) return;
    const payload = buildReviewSavePayload(reviewState.selectedPatient);

    btnRevSaveOverwrite.disabled = true;
    btnRevSaveOverwrite.innerHTML = `<span>⏳ 保存并同步覆盖中...</span>`;

    try {
      const data = await fetchJsonWithSafetyConfirm("/api/reviews/save", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      if (data.success) {
        showToast(data.message || "审核内容已保存，并已同步两份草稿文件！", "success");
        if (data.synced_forward && data.patient) {
          await selectReviewPatient(data.patient, true, false);
        }
        applyReviewSaveResponseToUi(data, data.synced_forward ? "已导入外部修改" : "已同步草稿");
      } else {
        showToast("保存失败: " + data.error, "error");
      }
    } catch (e) {
      showToast("保存异常: " + e.message, "error");
    } finally {
      btnRevSaveOverwrite.disabled = false;
      btnRevSaveOverwrite.innerHTML = `<span>💾 保存修改并同步草稿</span>`;
    }
  });

  // Save and print Page 1 only in Review Workspace (same workflow as Report Making)
  if (btnRevPrintReport) {
    btnRevPrintReport.addEventListener("click", async () => {
      if (!reviewState.selectedPatient) return;
      const payload = buildReviewSavePayload(reviewState.selectedPatient);

      btnRevPrintReport.disabled = true;
      btnRevPrintReport.innerHTML = `<span>⏳ 正在保存并准备打印...</span>`;

      try {
        const data = await fetchJsonWithSafetyConfirm("/api/reviews/print", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });
        if (data.success) {
          if (data.synced_forward && data.patient) {
            await selectReviewPatient(data.patient, true, false);
            applyReviewSaveResponseToUi(data, "已导入外部修改");
            showToast(data.message, "success");
            return;
          }
          applyReviewSaveResponseToUi(data, "已保存并打印");
          if (data.direct_printed) {
            showToast("审核报告已保存，已调用Word发送第1页至默认打印机！", "success");
          } else {
            showToast(data.message || "调用Word打印失败，请检查Word及默认打印机", "error");
          }
        } else {
          showToast("打印失败: " + data.error, "error");
        }
      } catch (e) {
        showToast("打印请求异常: " + e.message, "error");
      } finally {
        btnRevPrintReport.disabled = false;
        btnRevPrintReport.innerHTML = `<span>🖨️ 一键打印第一页</span>`;
      }
    });
  }

  // Approve patient review (2.9)
  btnRevApprove.addEventListener("click", async () => {
    if (!reviewState.selectedPatient) return;
    const p = reviewState.selectedPatient;

    btnRevApprove.disabled = true;
    btnRevApprove.innerHTML = `<span>⏳ 提交审核状态...</span>`;

    try {
      const res = await fetch("/api/reviews/approve", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pathology_no: p.pathology_no })
      });
      const data = await res.json();

      if (data.success) {
        showToast(data.message, "success");
        reviewState.selectedPatient = data.patient;
        loadReviewPatients();
      } else {
        showToast("审核失败: " + data.error, "error");
      }
    } catch (e) {
      showToast("审核异常: " + e.message, "error");
    } finally {
      btnRevApprove.disabled = false;
      btnRevApprove.innerHTML = `<span>✅ 完成审核 (2.9)</span>`;
    }
  });

  btnRevPreview.addEventListener("click", () => {
    if (reviewState.selectedPatient) {
      openReportPreviewModal(reviewState.selectedPatient.pathology_no);
    }
  });

  btnRevOpenWord.addEventListener("click", async () => {
    if (!reviewState.selectedPatient) return;
    const reportPath = reviewState.selectedPatient.final_report_path || reviewState.selectedPatient.word_path;
    if (reportPath) {
      await fetch("/api/open_file", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ path: reportPath })
      });
    }
  });

  function getActiveReviewPhotoDir() {
    if (reviewState.reviewPhotoDir) return reviewState.reviewPhotoDir;
    if (reviewState.selectedPatient) {
      return reviewState.selectedPatient.review_photo_dir || reviewState.selectedPatient.photo_dir || "";
    }
    return "";
  }

  async function openReviewPhotoDir() {
    const dirPath = getActiveReviewPhotoDir();
    if (!dirPath) return;
    await fetch("/api/open_dir", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: dirPath })
    });
  }

  function copyReviewPhotoPath() {
    const dirPath = getActiveReviewPhotoDir();
    if (!dirPath) return;
    copyToClipboard(dirPath, "已复制复制后的病理照片目录路径！");
  }

  btnRevFindClinical.addEventListener("click", () => {
    findClinicalPhotosForPatient(reviewState.selectedPatient);
  });

  btnRevOpenPhotoDir.addEventListener("click", openReviewPhotoDir);
  if (btnRevCopyPhotoPath) btnRevCopyPhotoPath.addEventListener("click", copyReviewPhotoPath);

  reviewShowApprovedToggle.addEventListener("change", () => {
    loadReviewPatients(true);
  });

  bindSidebarListFilter({
    reportDateEl: reviewFilterReportDate,
    btnReportTodayEl: btnReviewReportToday,
    btnReportAllEl: btnReviewReportAll,
    subDateEl: reviewFilterSubDate,
    btnSubAllEl: btnReviewSubAll,
    keywordEl: reviewFilterKeyword,
    btnSearchEl: btnReviewSearch,
    onReload: () => loadReviewPatients(true)
  });

  btnRefreshReview.addEventListener("click", () => {
    loadReviewPatients(true);
  });

  // ==========================================================
  // Phase 2.5: IHC Workspace (独立免疫组化模块)
  // ==========================================================
  const ihcState = {
    patients: [],
    selectedPatient: null,
    photos: [],
    slots: [],
    hasUnsavedChanges: false
  };

  const ihcCountBadge = document.getElementById("ihc-count-badge");
  const btnRefreshIhc = document.getElementById("btn-refresh-ihc");
  const ihcFilterReportDate = document.getElementById("ihc-filter-report-date");
  const btnIhcReportToday = document.getElementById("btn-ihc-report-today");
  const btnIhcReportAll = document.getElementById("btn-ihc-report-all");
  const ihcFilterSubDate = document.getElementById("ihc-filter-sub-date");
  const btnIhcSubAll = document.getElementById("btn-ihc-sub-all");
  const ihcFilterKeyword = document.getElementById("ihc-filter-keyword");
  const btnIhcSearch = document.getElementById("btn-ihc-search");
  const ihcPatientList = document.getElementById("ihc-patient-list");

  const ihcEmptyHint = document.getElementById("ihc-empty-hint");
  const ihcEditorPanel = document.getElementById("ihc-editor-panel");

  const ihcBadgeName = document.getElementById("ihc-badge-name");
  const ihcBadgeNo = document.getElementById("ihc-badge-no");
  const ihcBadgeStatus = document.getElementById("ihc-badge-status");
  const ihcBadgeDiag = document.getElementById("ihc-badge-diag");

  const btnIhcDelete = document.getElementById("btn-ihc-delete");
  const btnIhcCreate = document.getElementById("btn-ihc-create");
  const btnIhcCreateInline = document.getElementById("btn-ihc-create-inline");
  const btnIhcOpenPhotoDir = document.getElementById("btn-ihc-open-photo-dir");
  const btnIhcCopyPhotoPath = document.getElementById("btn-ihc-copy-photo-path");
  const btnIhcFindClinical = document.getElementById("btn-ihc-find-clinical");

  const ihcNotCreatedBanner = document.getElementById("ihc-not-created-banner");
  const ihcNotCreatedTitle = document.getElementById("ihc-not-created-title");
  const ihcNotCreatedDesc = document.getElementById("ihc-not-created-desc");
  const ihcPhotoPathBox = document.getElementById("ihc-photo-path-box");
  const ihcPhotoPathText = document.getElementById("ihc-photo-path-text");

  const ihcReportDoc = document.getElementById("ihc-report-doc");
  const ihcReviewDoc = document.getElementById("ihc-review-doc");
  const ihcReportDate = document.getElementById("ihc-report-date");

  const btnIhcRefreshPhotos = document.getElementById("btn-ihc-refresh-photos");
  const btnIhcAddSlot = document.getElementById("btn-ihc-add-slot");
  const btnIhcAutoLayout = document.getElementById("btn-ihc-auto-layout");
  const ihcAvailPhotosCount = document.getElementById("ihc-avail-photos-count");
  const ihcAvailablePhotosList = document.getElementById("ihc-available-photos-list");
  const ihcSlotsStatusDesc = document.getElementById("ihc-slots-status-desc");
  const ihcSlotsGrid = document.getElementById("ihc-slots-grid");

  const btnIhcOpenTemplateModal = document.getElementById("btn-ihc-open-template-modal");
  const ihcRecommendContainer = document.getElementById("ihc-recommend-container");
  const ihcRecommendChips = document.getElementById("ihc-recommend-chips");
  const ihcMicroscopicDesc = document.getElementById("ihc-microscopic-desc");

  const btnIhcRegenDiag = document.getElementById("btn-ihc-regen-diag");
  const ihcPathologyDiag = document.getElementById("ihc-pathology-diag");
  const ihcSaveStatus = document.getElementById("ihc-save-status");

  const btnIhcPreview = document.getElementById("btn-ihc-preview");
  const btnIhcOpenWord = document.getElementById("btn-ihc-open-word");
  const btnIhcOpenExplorer = document.getElementById("btn-ihc-open-explorer");
  const btnIhcSave = document.getElementById("btn-ihc-save");
  const btnIhcArchive = document.getElementById("btn-ihc-archive");

  function markIhcDirty() {
    if (!ihcState.selectedPatient || !(ihcState.selectedPatient.ihc_report && ihcState.selectedPatient.ihc_report.ihc_created)) return;
    ihcState.hasUnsavedChanges = true;
    ihcSaveStatus.textContent = "有未保存修改";
    ihcSaveStatus.style.color = "var(--warning)";
  }

  [ihcReportDoc, ihcReviewDoc, ihcReportDate, ihcMicroscopicDesc, ihcPathologyDiag].forEach(el => {
    if (el) {
      el.addEventListener("input", markIhcDirty);
      el.addEventListener("change", markIhcDirty);
    }
  });

  function setIhcEditorEnabled(enabled) {
    [ihcReportDoc, ihcReviewDoc, ihcReportDate, ihcMicroscopicDesc, ihcPathologyDiag].forEach(el => {
      if (el) el.disabled = !enabled;
    });
    [
      btnIhcOpenPhotoDir, btnIhcCopyPhotoPath,
      btnIhcRefreshPhotos, btnIhcAddSlot, btnIhcAutoLayout, btnIhcOpenTemplateModal,
      btnIhcRegenDiag, btnIhcPreview, btnIhcOpenWord, btnIhcOpenExplorer,
      btnIhcSave, btnIhcArchive
    ].forEach(btn => {
      if (btn) btn.disabled = !enabled;
    });
  }

  async function loadIhcPatients() {
    try {
      const params = new URLSearchParams();
      if (ihcFilterReportDate && ihcFilterReportDate.value) {
        params.set("report_date", ihcFilterReportDate.value);
      }
      if (ihcFilterSubDate && ihcFilterSubDate.value) {
        params.set("submission_date", ihcFilterSubDate.value);
      }
      if (ihcFilterKeyword && ihcFilterKeyword.value.trim()) {
        params.set("keyword", ihcFilterKeyword.value.trim());
      }
      ihcPatientList.innerHTML = `<div class="patient-list-empty">正在加载免疫组化数据...</div>`;

      const res = await fetch(`/api/ihc/list?${params.toString()}`);
      const data = await res.json();

      if (data.success) {
        const list = data.patients || [];
        ihcState.patients = list;
        ihcCountBadge.textContent = `${list.length} 人`;
        renderIhcPatientList();

        if (ihcState.selectedPatient) {
          const found = list.find(p => p.pathology_no === ihcState.selectedPatient.pathology_no);
          if (found) {
            selectIhcPatient(found, false);
          }
        }
      } else {
        ihcPatientList.innerHTML = `<div class="patient-list-empty" style="color: var(--danger);">${data.error}</div>`;
      }
    } catch (e) {
      ihcPatientList.innerHTML = `<div class="patient-list-empty" style="color: var(--danger);">加载失败: ${e.message}</div>`;
    }
  }

  function renderIhcPatientList() {
    if (!ihcState.patients.length) {
      ihcPatientList.innerHTML = `<div class="patient-list-empty">暂无标记免疫组化的患者</div>`;
      return;
    }

    ihcPatientList.innerHTML = "";
    ihcState.patients.forEach(p => {
      const ihc = p.ihc_report || {};
      const card = document.createElement("div");
      const isSelected = ihcState.selectedPatient && ihcState.selectedPatient.pathology_no === p.pathology_no;
      let fillClass = "status-fill-ihc";
      let statusBadge = `<span class="badge-status badge-pending">未创建组化报告</span>`;
      if (ihc.archived === 1) {
        fillClass = "status-fill-archived";
        statusBadge = `<span class="badge-status badge-archived">组化已归档</span>`;
      } else if (ihc.archived === 2) {
        fillClass = "status-fill-rearchive";
        statusBadge = `<span class="badge-status badge-rearchive">组化需重归</span>`;
      } else if (ihc.ihc_created) {
        fillClass = "status-fill-saved";
        statusBadge = `<span class="badge-status badge-ihc">已创建草稿</span>`;
      }

      let origArchBadge = p.archived === 0
        ? `<span class="badge-status badge-recut">原报告未归档</span>`
        : "";
      const ihcExtBadge = renderExternalModBadge(p, true);

      card.className = `patient-list-item ${fillClass} ${p.ihc_external_modified ? "external-modified-card" : ""} ${isSelected ? "selected" : ""}`.trim();
      const diagText = (ihc.ihc_created && ihc.pathology_diag) ? ihc.pathology_diag : (p.pathology_diag || p.clinical_diagnosis || "未诊断");

      card.innerHTML = `
        <div class="patient-item-header">
          <span class="patient-item-name">${p.name}</span>
          <span class="patient-item-no">${p.pathology_no}</span>
        </div>
        <div class="patient-item-meta">
          <span>${p.gender || "-"} / ${formatAgeDisplay(p.age)}</span>
          <span>送检: ${p.submission_date || "-"}</span>
        </div>
        <div class="patient-item-body">
          <div class="patient-item-diag" title="${diagText}">${diagText}</div>
          <div class="patient-item-badges">${ihcExtBadge}${statusBadge}${origArchBadge}</div>
        </div>
      `;

      card.addEventListener("click", () => {
        selectIhcPatient(p, true, true);
      });

      ihcPatientList.appendChild(card);
    });
  }

  async function selectIhcPatient(patient, reloadPhotos = true, fromUserClick = false) {
    let effPatient = patient;
    if (fromUserClick && effPatient && effPatient.ihc_external_modified) {
      const outcome = await resolveExternalSyncForPatient(effPatient, "ihc", false);
      if (outcome && outcome.success && outcome.patient) {
        effPatient = outcome.patient;
      }
    }
    ihcState.selectedPatient = effPatient;
    ihcState.hasUnsavedChanges = false;
    ihcEmptyHint.style.display = "none";
    ihcEditorPanel.style.display = "flex";

    document.querySelectorAll("#ihc-patient-list .patient-list-item").forEach(el => {
      const noEl = el.querySelector(".patient-item-no");
      if (noEl && noEl.textContent.trim() === effPatient.pathology_no) {
        el.classList.add("selected");
      } else {
        el.classList.remove("selected");
      }
    });

    const ihc = effPatient.ihc_report || {};
    const created = !!ihc.ihc_created;

    ihcBadgeName.textContent = effPatient.name;
    ihcBadgeNo.textContent = effPatient.pathology_no;

    if (!created) {
      ihcBadgeStatus.textContent = "未创建免疫组化报告";
      ihcBadgeDiag.textContent = `原诊断: ${effPatient.pathology_diag || effPatient.clinical_diagnosis || "未诊断"}`;
      btnIhcCreate.innerHTML = `<span>🧪 创建免疫组化报告</span>`;
      ihcNotCreatedBanner.style.display = "flex";
      ihcPhotoPathBox.style.display = "none";
      ihcPhotoPathText.textContent = "-";
      ihcPhotoPathText.title = "";
      if (effPatient.archived === 0) {
        ihcNotCreatedTitle.textContent = "⚠️ 原报告尚未归档，无法创建免疫组化报告";
        ihcNotCreatedDesc.textContent = "请先在【报告制作】（或【报告审核】）完成原报告归档至“已归档报告”后，再点击【创建免疫组化报告】。";
      } else {
        ihcNotCreatedTitle.textContent = "尚未创建该患者的免疫组化报告副本";
        ihcNotCreatedDesc.textContent = "请点击【创建免疫组化报告】，系统将复制该患者“已归档报告”及病理照片目录至“特殊患者归档\\免疫组化”中（不复制临床照片）。";
      }

      ihcReportDoc.value = effPatient.report_doctor || state.config.report_doctor || "张医生";
      ihcReviewDoc.value = effPatient.review_doctor || state.config.review_doctor || "李主任";
      ihcReportDate.value = toIsoDateString(effPatient.report_date_raw || effPatient.report_date);
      ihcMicroscopicDesc.value = "";
      ihcPathologyDiag.value = "";
      ihcState.photos = [];
      ihcState.slots = [];
      ihcAvailPhotosCount.textContent = "0";
      ihcAvailablePhotosList.innerHTML = `<div class="photo-strip-empty">请先点击“创建免疫组化报告”以建立独立免疫组化照片目录</div>`;
      ihcSlotsGrid.innerHTML = "";
      ihcSlotsStatusDesc.textContent = "尚未创建免疫组化报告";
      ihcRecommendContainer.style.display = "none";
      ihcRecommendChips.innerHTML = "";
      ihcSaveStatus.textContent = "未创建免疫组化报告";
      ihcSaveStatus.style.color = "var(--text-muted)";
      setIhcEditorEnabled(false);
      return;
    }

    ihcBadgeStatus.textContent = ihc.archived === 1 ? "免疫组化已归档" : (ihc.archived === 2 ? "免疫组化需重归" : "免疫组化草稿");
    ihcBadgeDiag.textContent = `诊断: ${ihc.pathology_diag || effPatient.pathology_diag || "未诊断"}`;
    btnIhcCreate.innerHTML = `<span>🔄 重建免疫组化报告</span>`;
    ihcNotCreatedBanner.style.display = "none";
    ihcPhotoPathBox.style.display = "flex";
    ihcPhotoPathText.textContent = ihc.ihc_photo_dir || "-";
    ihcPhotoPathText.title = ihc.ihc_photo_dir || "";

    ihcReportDoc.value = ihc.report_doctor || effPatient.report_doctor || state.config.report_doctor || "张医生";
    ihcReviewDoc.value = ihc.review_doctor || effPatient.review_doctor || state.config.review_doctor || "李主任";
    ihcReportDate.value = toIsoDateString(ihc.report_date_raw || ihc.report_date || effPatient.report_date_raw || effPatient.report_date);
    ihcMicroscopicDesc.value = ihc.microscopic_desc || "";
    ihcPathologyDiag.value = ihc.pathology_diag || "";

    if (ihc.archived === 1) {
      ihcSaveStatus.textContent = "免疫组化报告已归档";
      ihcSaveStatus.style.color = "var(--success)";
    } else if (ihc.archived === 2) {
      ihcSaveStatus.textContent = "已保存修改，需重新归档";
      ihcSaveStatus.style.color = "var(--warning)";
    } else {
      ihcSaveStatus.textContent = "已创建草稿，就绪";
      ihcSaveStatus.style.color = "var(--text-muted)";
    }

    setIhcEditorEnabled(true);

    if (reloadPhotos) {
      await fetchIhcPhotosAndSlots(patient);
    }
  }

  async function fetchIhcPhotosAndSlots(patient) {
    const ihc = patient.ihc_report || {};
    if (!ihc.ihc_created) return;
    try {
      ihcAvailablePhotosList.innerHTML = `<div class="photo-strip-empty">正在扫描免疫组化照片目录...</div>`;
      const res = await fetch(`/api/ihc/${encodeURIComponent(patient.pathology_no)}/photos`);
      const data = await res.json();
      if (data.success) {
        ihcState.photos = data.photos || [];
        ihcAvailPhotosCount.textContent = ihcState.photos.length;
        renderIhcAvailablePhotosStrip();
        renderIhcRecommendChips(data.recommended_templates || []);

        const totalSlots = Math.max(parseInt(ihc.image_count || patient.image_count) || 1, 1);
        let biopsySites = ihc.biopsy_sites && ihc.biopsy_sites.length ? ihc.biopsy_sites : (patient.biopsy_sites || []);
        if (typeof biopsySites === "string") {
          biopsySites = biopsySites.split(";").map(s => s.trim()).filter(Boolean);
        }
        let existingSelected = ihc.selected_images || [];
        if (typeof existingSelected === "string") {
          try { existingSelected = JSON.parse(existingSelected); } catch (e) { existingSelected = []; }
        }

        ihcState.slots = [];
        for (let i = 1; i <= totalSlots; i++) {
          const siteName = (biopsySites.length >= i) ? biopsySites[i - 1] : "";
          let matchedPhoto = null;
          if (Array.isArray(existingSelected) && existingSelected.length) {
            const item = existingSelected.find((it, idx) => {
              if (typeof it === "object" && it !== null) return it.slot === i;
              return idx === i - 1;
            });
            if (item) {
              const pPath = typeof item === "object" ? item.path : item;
              if (pPath) {
                matchedPhoto = ihcState.photos.find(ph => ph.path === pPath);
              }
            }
          }
          if (!matchedPhoto && ihcState.photos.length) {
            matchedPhoto = ihcState.photos.find(ph => (ph.indices && ph.indices.includes(String(i))) || ph.index_num === i);
          }
          ihcState.slots.push({ slot: i, site: siteName, photo: matchedPhoto || null });
        }
        renderIhcSlots();
      }
    } catch (e) {
      showToast("读取免疫组化照片失败: " + e.message, "error");
    }
  }

  function renderIhcAvailablePhotosStrip() {
    if (!ihcState.photos.length) {
      ihcAvailablePhotosList.innerHTML = `<div class="photo-strip-empty">复制后的免疫组化照片目录中暂无图片（可打开目录录入照片后点击“重新读取照片”）</div>`;
      return;
    }
    ihcAvailablePhotosList.innerHTML = "";
    ihcState.photos.forEach(ph => {
      const card = document.createElement("div");
      card.className = "photo-strip-card";
      const idxText = ph.indices && ph.indices.length ? ph.indices.join(".") : "";
      const tagHtml = idxText ? `<span class="photo-strip-tag">${idxText}</span>` : "";
      card.innerHTML = `
        <img src="${ph.url}" class="photo-strip-img" alt="${ph.filename}">
        <div class="photo-strip-name" title="${ph.filename}">${tagHtml} ${ph.diagnosis || ph.filename}</div>
      `;
      card.addEventListener("click", () => {
        openPhotoZoom(ph.url, `${ph.filename} (${ph.diagnosis || ""})`);
      });
      ihcAvailablePhotosList.appendChild(card);
    });
  }

  function renderIhcRecommendChips(recs) {
    if (!recs || !recs.length) {
      ihcRecommendContainer.style.display = "none";
      ihcRecommendChips.innerHTML = "";
      return;
    }
    ihcRecommendContainer.style.display = "flex";
    ihcRecommendChips.innerHTML = "";
    recs.forEach(t => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "recommend-chip";
      chip.textContent = `+ ${t.name}`;
      chip.title = `点击将【${t.name}】镜下描述追加到输入框`;
      chip.addEventListener("click", () => {
        if (ihcMicroscopicDesc.value.trim()) {
          ihcMicroscopicDesc.value += "\n" + t.desc;
        } else {
          ihcMicroscopicDesc.value = t.desc;
        }
        markIhcDirty();
        showToast(`已追加推荐模板【${t.name}】至镜下描述`, "success");
      });
      ihcRecommendChips.appendChild(chip);
    });
  }

  function renderIhcSlots() {
    ihcSlotsStatusDesc.textContent = `当前共 ${ihcState.slots.length} 个排版插槽`;
    ihcSlotsGrid.innerHTML = "";
    ihcState.slots.forEach((s, idx) => {
      const card = document.createElement("div");
      card.className = "slot-card";
      const hasPhoto = s.photo !== null;
      const previewHtml = hasPhoto
        ? `<div class="slot-preview-box has-photo" title="点击放大图片"><img src="${s.photo.url}" class="slot-preview-img" alt="插槽 ${s.slot}"></div>`
        : `<div class="slot-preview-box"><div class="slot-empty-text"><span>🖼️</span><span>待选择图片</span></div></div>`;

      let selectOptions = `<option value="">-- 未选择图片 --</option>`;
      ihcState.photos.forEach(ph => {
        const isSel = s.photo && s.photo.path === ph.path;
        const idxT = ph.indices && ph.indices.length ? `[${ph.indices.join(".")}] ` : "";
        selectOptions += `<option value="${ph.path}" ${isSel ? "selected" : ""}>${idxT}${ph.diagnosis || ph.filename}</option>`;
      });

      card.innerHTML = `
        <div class="slot-card-header">
          <span>第 ${s.slot} 张图片</span>
          ${ihcState.slots.length > 1 ? `<button type="button" class="btn btn-secondary btn-sm btn-remove-ihc-slot" style="padding: 1px 6px; font-size: 11px;">✕</button>` : ""}
        </div>
        ${previewHtml}
        <div class="form-group" style="margin-bottom: 4px;">
          <input type="text" class="slot-site-input" placeholder="部位名称" value="${s.site || ""}">
        </div>
        <div class="form-group" style="margin-bottom: 2px;">
          <select class="slot-photo-select">${selectOptions}</select>
        </div>
        <div class="slot-actions">
          ${hasPhoto ? `<button type="button" class="btn btn-secondary btn-sm btn-zoom-ihc-slot" style="padding: 2px 6px; font-size: 11px;">🔍 放大</button>` : `<span></span>`}
          ${hasPhoto ? `<button type="button" class="btn btn-secondary btn-sm btn-clear-ihc-slot" style="padding: 2px 6px; font-size: 11px; color: var(--danger);">移除图片</button>` : `<span></span>`}
        </div>
      `;

      card.querySelector(".slot-site-input").addEventListener("input", e => {
        s.site = e.target.value;
        markIhcDirty();
      });
      card.querySelector(".slot-photo-select").addEventListener("change", e => {
        s.photo = ihcState.photos.find(p => p.path === e.target.value) || null;
        markIhcDirty();
        renderIhcSlots();
      });
      if (hasPhoto) {
        card.querySelector(".slot-preview-box").addEventListener("click", () => {
          openPhotoZoom(s.photo.url, `第 ${s.slot} 张: ${s.photo.filename}`);
        });
        const zoomBtn = card.querySelector(".btn-zoom-ihc-slot");
        if (zoomBtn) {
          zoomBtn.addEventListener("click", () => {
            openPhotoZoom(s.photo.url, `第 ${s.slot} 张: ${s.photo.filename}`);
          });
        }
        const clearBtn = card.querySelector(".btn-clear-ihc-slot");
        if (clearBtn) {
          clearBtn.addEventListener("click", () => {
            s.photo = null;
            markIhcDirty();
            renderIhcSlots();
          });
        }
      }
      const removeBtn = card.querySelector(".btn-remove-ihc-slot");
      if (removeBtn) {
        removeBtn.addEventListener("click", () => {
          if (ihcState.slots.length > 1) {
            ihcState.slots.splice(idx, 1);
            ihcState.slots.forEach((item, i) => item.slot = i + 1);
            markIhcDirty();
            renderIhcSlots();
          }
        });
      }
      ihcSlotsGrid.appendChild(card);
    });
  }

  async function handleCreateIhcReport() {
    if (!ihcState.selectedPatient) return;
    const p = ihcState.selectedPatient;
    const ihc = p.ihc_report || {};
    let forceOverwrite = false;
    if (ihc.ihc_created) {
      if (!confirm("该患者已存在免疫组化报告草稿，重新创建将用【已归档报告】重新覆盖免疫组化目录中的报告文档（旧草稿将提前备份至回收站，已录入的照片文件将保留），确定继续吗？")) {
        return;
      }
      forceOverwrite = true;
    }

    btnIhcCreate.disabled = true;
    try {
      const data = await fetchJsonWithSafetyConfirm("/api/ihc/create", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pathology_no: p.pathology_no, force_overwrite: forceOverwrite })
      });
      if (data.success) {
        showToast(data.message, "success");
        ihcState.selectedPatient = data.patient;
        await selectIhcPatient(data.patient, true);
        loadIhcPatients();
      } else {
        alert(data.error || "创建免疫组化报告失败");
        showToast(data.error || "创建失败", "error");
      }
    } catch (e) {
      showToast("创建异常: " + e.message, "error");
    } finally {
      btnIhcCreate.disabled = false;
    }
  }

  async function handleSaveIhcReport() {
    if (!ihcState.selectedPatient) return false;
    const p = ihcState.selectedPatient;
    const selectedImgs = ihcState.slots.map(s => ({
      slot: s.slot,
      path: s.photo ? s.photo.path : null,
      site: (s.site || "").trim()
    }));
    const payload = {
      pathology_no: p.pathology_no,
      report_doctor: ihcReportDoc.value.trim(),
      review_doctor: ihcReviewDoc.value.trim(),
      report_date: ihcReportDate.value.trim(),
      microscopic_desc: ihcMicroscopicDesc.value.trim(),
      pathology_diag: ihcPathologyDiag.value.trim(),
      selected_images: selectedImgs,
      image_count: ihcState.slots.length || 1,
      biopsy_sites: ihcState.slots.map(s => (s.site || "").trim())
    };

    btnIhcSave.disabled = true;
    btnIhcSave.innerHTML = `<span>⏳ 保存中...</span>`;
    try {
      const data = await fetchJsonWithSafetyConfirm("/api/ihc/save", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      if (data.success) {
        showToast(data.message, "success");
        ihcState.selectedPatient = data.patient;
        if (data.synced_forward && data.patient) {
          await selectIhcPatient(data.patient, true, false);
        }
        ihcState.hasUnsavedChanges = false;
        ihcSaveStatus.textContent = data.synced_forward
          ? `已导入外部Word修改 (${new Date().toLocaleTimeString()})`
          : `已独立保存 (${new Date().toLocaleTimeString()})`;
        ihcSaveStatus.style.color = "var(--success)";
        loadIhcPatients();
        return true;
      } else {
        showToast("保存失败: " + data.error, "error");
        return false;
      }
    } catch (e) {
      showToast("保存异常: " + e.message, "error");
      return false;
    } finally {
      btnIhcSave.disabled = false;
      btnIhcSave.innerHTML = `<span>💾 保存免疫组化报告</span>`;
    }
  }

  async function handleDeleteIhcReport() {
    if (!ihcState.selectedPatient) return;
    const p = ihcState.selectedPatient;
    const ihc = p.ihc_report || {};
    const confirmMsg = ihc.ihc_created
      ? `⚠️ 【危险操作警告】确定要删除患者 [${p.name}] (${p.pathology_no}) 已生成的免疫组化报告并撤销其免疫组化状态吗？\n\n执行后：\n1. 该患者“特殊患者归档\\免疫组化”目录及已归档的《*_免疫组化报告.docx》（如有）将被移入回收站；\n2. 患者免疫组化标记将被完全撤销，变为普通报告（原有报告与普通归档不受影响）；\n3. 该患者将立即从免疫组化列表中移除。\n\n此操作请谨慎确认，是否继续？`
      : `⚠️ 【危险操作警告】确定要撤销患者 [${p.name}] (${p.pathology_no}) 的免疫组化状态吗？\n\n撤销后该患者将完全变为普通报告并从免疫组化列表中移除，是否继续？`;
    if (!confirm(confirmMsg)) {
      return;
    }

    if (btnIhcDelete) btnIhcDelete.disabled = true;
    try {
      const data = await fetchJsonWithSafetyConfirm("/api/ihc/delete", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pathology_no: p.pathology_no })
      });
      if (data.success) {
        showToast(data.message || "已删除免疫组化报告并恢复为普通报告", "success");
        ihcState.selectedPatient = null;
        ihcState.hasUnsavedChanges = false;
        ihcEditorPanel.style.display = "none";
        ihcEmptyHint.style.display = "flex";
        await loadIhcPatients();
        if (makingState && makingState.patients) loadMakingPatients();
        loadOutputIhcPatients();
        loadOutputRegularPatients();
      } else {
        alert(data.error || "删除失败");
        showToast(data.error || "删除失败", "error");
      }
    } catch (e) {
      showToast("删除免疫组化报告异常: " + e.message, "error");
    } finally {
      if (btnIhcDelete) btnIhcDelete.disabled = false;
    }
  }

  if (btnIhcDelete) btnIhcDelete.addEventListener("click", handleDeleteIhcReport);
  btnIhcCreate.addEventListener("click", handleCreateIhcReport);
  if (btnIhcCreateInline) btnIhcCreateInline.addEventListener("click", handleCreateIhcReport);

  async function openCopiedIhcPhotoDir() {
    const ihc = ihcState.selectedPatient && ihcState.selectedPatient.ihc_report;
    if (!ihc || !ihc.ihc_photo_dir) return;
    await fetch("/api/open_dir", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: ihc.ihc_photo_dir })
    });
  }
  function copyCopiedIhcPhotoPath() {
    const ihc = ihcState.selectedPatient && ihcState.selectedPatient.ihc_report;
    if (!ihc || !ihc.ihc_photo_dir) return;
    copyToClipboard(ihc.ihc_photo_dir, "已复制复制后的免疫组化病理照片目录路径！");
  }

  btnIhcOpenPhotoDir.addEventListener("click", openCopiedIhcPhotoDir);
  btnIhcCopyPhotoPath.addEventListener("click", copyCopiedIhcPhotoPath);
  btnIhcFindClinical.addEventListener("click", () => {
    findClinicalPhotosForPatient(ihcState.selectedPatient);
  });

  btnIhcRefreshPhotos.addEventListener("click", async () => {
    if (ihcState.selectedPatient) {
      await fetchIhcPhotosAndSlots(ihcState.selectedPatient);
      showToast("已重新扫描复制后的免疫组化照片目录！", "success");
    }
  });

  btnIhcAddSlot.addEventListener("click", () => {
    const next = ihcState.slots.length + 1;
    if (next > 9) return showToast("最多支持 9 张插槽", "error");
    ihcState.slots.push({ slot: next, site: "", photo: null });
    markIhcDirty();
    renderIhcSlots();
  });

  btnIhcAutoLayout.addEventListener("click", () => {
    ihcState.slots.forEach(s => {
      const matched = ihcState.photos.find(ph => (ph.indices && ph.indices.includes(String(s.slot))) || ph.index_num === s.slot);
      if (matched) s.photo = matched;
    });
    markIhcDirty();
    renderIhcSlots();
    showToast("已按照图片编号完成一键自动排版！", "success");
  });

  btnIhcRegenDiag.addEventListener("click", async () => {
    if (!ihcState.selectedPatient) return;
    try {
      const slotCnt = ihcState.slots.length || 1;
      const res = await fetch(`/api/ihc/${encodeURIComponent(ihcState.selectedPatient.pathology_no)}/photos?image_count=${slotCnt}`);
      const data = await res.json();
      if (data.success && data.auto_diagnosis) {
        ihcPathologyDiag.value = data.auto_diagnosis;
        markIhcDirty();
        showToast("已根据免疫组化照片命名重新合成病理诊断！", "success");
      } else {
        showToast("未能从免疫组化照片命名中提取出有效诊断", "error");
      }
    } catch (e) {
      showToast("合成诊断失败: " + e.message, "error");
    }
  });

  btnIhcOpenTemplateModal.addEventListener("click", () => {
    openTemplateModal(ihcMicroscopicDesc);
  });

  btnIhcSave.addEventListener("click", handleSaveIhcReport);

  btnIhcArchive.addEventListener("click", async () => {
    if (!ihcState.selectedPatient) return;
    if (ihcState.hasUnsavedChanges) {
      if (!confirm("当前免疫组化报告有尚未保存的修改，是否先保存后再归档？\n点击【确定】将先保存当前修改并归档；点击【取消】取消操作。")) {
        return;
      }
      const savedOk = await handleSaveIhcReport();
      if (!savedOk) return;
    }

    const p = ihcState.selectedPatient;
    btnIhcArchive.disabled = true;
    btnIhcArchive.innerHTML = `<span>⏳ 归档中...</span>`;
    try {
      const data = await fetchJsonWithSafetyConfirm("/api/ihc/archive", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pathology_no: p.pathology_no })
      });
      if (data.success) {
        showToast(data.message, "success");
        ihcState.selectedPatient = data.patient;
        ihcSaveStatus.textContent = "免疫组化报告已归档";
        ihcSaveStatus.style.color = "var(--success)";
        loadIhcPatients();
      } else {
        alert("归档失败: " + data.error);
        showToast("归档失败: " + data.error, "error");
      }
    } catch (e) {
      showToast("归档异常: " + e.message, "error");
    } finally {
      btnIhcArchive.disabled = false;
      btnIhcArchive.innerHTML = `<span>📦 归档免疫组化报告</span>`;
    }
  });

  btnIhcPreview.addEventListener("click", () => {
    if (ihcState.selectedPatient) {
      openReportPreviewModal(ihcState.selectedPatient.pathology_no, "ihc");
    }
  });

  btnIhcOpenWord.addEventListener("click", async () => {
    const ihc = ihcState.selectedPatient && ihcState.selectedPatient.ihc_report;
    if (ihc && ihc.ihc_doc_path) {
      await fetch("/api/open_file", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ path: ihc.ihc_doc_path })
      });
    }
  });

  btnIhcOpenExplorer.addEventListener("click", async () => {
    const ihc = ihcState.selectedPatient && ihcState.selectedPatient.ihc_report;
    const target = (ihc && (ihc.ihc_doc_path || ihc.ihc_dir)) || "";
    if (target) {
      await fetch("/api/open_explorer_select", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ path: target })
      });
    }
  });

  bindSidebarListFilter({
    reportDateEl: ihcFilterReportDate,
    btnReportTodayEl: btnIhcReportToday,
    btnReportAllEl: btnIhcReportAll,
    subDateEl: ihcFilterSubDate,
    btnSubAllEl: btnIhcSubAll,
    keywordEl: ihcFilterKeyword,
    btnSearchEl: btnIhcSearch,
    onReload: loadIhcPatients
  });

  btnRefreshIhc.addEventListener("click", () => {
    loadIhcPatients();
  });

  // ==========================================================
  // Text Template Library Modal (普通模板 + 免疫组化模板)
  // ==========================================================
  let activeTemplateTargetTextarea = null;
  let allLoadedTemplates = [];
  let currentTmplCategory = "全部";

  const templateLibraryModal = document.getElementById("template-library-modal");
  const btnCloseTemplateModal = document.getElementById("btn-close-template-modal");
  const tmplCatBtns = document.querySelectorAll(".tmpl-cat-btn");
  const tmplSearchInput = document.getElementById("tmpl-search-input");
  const btnClearTmplSearch = document.getElementById("btn-clear-tmpl-search");
  const tmplItemsList = document.getElementById("tmpl-items-list");
  const catCountAll = document.getElementById("cat-count-all");
  const catCountNormal = document.getElementById("cat-count-normal");
  const catCountIhc = document.getElementById("cat-count-ihc");

  async function openTemplateModal(targetTextarea) {
    activeTemplateTargetTextarea = targetTextarea;
    templateLibraryModal.style.display = "flex";
    tmplSearchInput.value = "";
    if (!allLoadedTemplates.length) {
      await loadAllTemplates();
    }
    renderFilteredTemplates();
  }

  async function loadAllTemplates() {
    try {
      tmplItemsList.innerHTML = `<div style="text-align: center; padding: 40px; color: var(--text-muted);">正在加载模板库...</div>`;
      const res = await fetch("/api/templates");
      const data = await res.json();
      if (data.success) {
        allLoadedTemplates = data.templates || [];
        const normCount = allLoadedTemplates.filter(t => t.category === "普通模板").length;
        const ihcCount = allLoadedTemplates.filter(t => t.category === "免疫组化模板").length;
        catCountAll.textContent = allLoadedTemplates.length;
        catCountNormal.textContent = normCount;
        catCountIhc.textContent = ihcCount;
      }
    } catch (e) {
      tmplItemsList.innerHTML = `<div style="color: red; padding: 40px;">加载模板异常: ${e.message}</div>`;
    }
  }

  function renderFilteredTemplates() {
    const kw = tmplSearchInput.value.trim().toLowerCase();
    let filtered = allLoadedTemplates;

    if (currentTmplCategory !== "全部") {
      filtered = filtered.filter(t => t.category === currentTmplCategory);
    }

    if (kw) {
      filtered = filtered.filter(t => t.name.toLowerCase().includes(kw) || t.desc.toLowerCase().includes(kw));
    }

    if (!filtered.length) {
      tmplItemsList.innerHTML = `<div style="text-align: center; padding: 40px; color: var(--text-muted);">未搜索到匹配的文字模板</div>`;
      return;
    }

    tmplItemsList.innerHTML = "";
    filtered.forEach(t => {
      const card = document.createElement("div");
      card.className = "tmpl-card";

      card.innerHTML = `
        <div class="tmpl-card-header">
          <span class="tmpl-title">${t.name}</span>
          <span class="tmpl-category-pill">${t.category}</span>
        </div>
        <div class="tmpl-desc-text">${t.desc}</div>
        <div class="tmpl-actions">
          <button type="button" class="btn btn-secondary btn-sm btn-append-tmpl">追加到输入框</button>
          <button type="button" class="btn btn-primary btn-sm btn-insert-tmpl">覆盖填入</button>
        </div>
      `;

      card.querySelector(".btn-insert-tmpl").addEventListener("click", () => {
        if (activeTemplateTargetTextarea) {
          activeTemplateTargetTextarea.value = t.desc;
          activeTemplateTargetTextarea.dispatchEvent(new Event("input", { bubbles: true }));
          showToast(`已将【${t.name}】填入镜下描述`, "success");
          templateLibraryModal.style.display = "none";
        }
      });

      card.querySelector(".btn-append-tmpl").addEventListener("click", () => {
        if (activeTemplateTargetTextarea) {
          if (activeTemplateTargetTextarea.value.trim()) {
            activeTemplateTargetTextarea.value += "\n" + t.desc;
          } else {
            activeTemplateTargetTextarea.value = t.desc;
          }
          activeTemplateTargetTextarea.dispatchEvent(new Event("input", { bubbles: true }));
          showToast(`已将【${t.name}】追加到镜下描述`, "success");
          templateLibraryModal.style.display = "none";
        }
      });

      tmplItemsList.appendChild(card);
    });
  }

  tmplCatBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      tmplCatBtns.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      currentTmplCategory = btn.dataset.cat;
      renderFilteredTemplates();
    });
  });

  tmplSearchInput.addEventListener("input", () => {
    renderFilteredTemplates();
  });

  btnClearTmplSearch.addEventListener("click", () => {
    tmplSearchInput.value = "";
    renderFilteredTemplates();
  });

  btnCloseTemplateModal.addEventListener("click", () => {
    templateLibraryModal.style.display = "none";
  });

  btnOpenTemplateModal.addEventListener("click", () => {
    openTemplateModal(mkMicroscopicDesc);
  });

  btnRevOpenTemplateModal.addEventListener("click", () => {
    openTemplateModal(revMicroscopicDesc);
  });

  // ==========================================================
  // Photo Zoom Modal
  // ==========================================================
  const photoZoomModal = document.getElementById("photo-zoom-modal");
  const zoomPhotoImg = document.getElementById("zoom-photo-img");
  const zoomPhotoTitle = document.getElementById("zoom-photo-title");
  const btnCloseZoomModal = document.getElementById("btn-close-zoom-modal");

  function openPhotoZoom(url, title = "照片预览") {
    zoomPhotoImg.src = url;
    zoomPhotoTitle.textContent = title;
    photoZoomModal.style.display = "flex";
  }

  btnCloseZoomModal.addEventListener("click", () => {
    photoZoomModal.style.display = "none";
  });

  photoZoomModal.addEventListener("click", (e) => {
    if (e.target === photoZoomModal) {
      photoZoomModal.style.display = "none";
    }
  });

  // ==========================================================
  // Phase 3.1: Report Output Workspace (3. 报告输出与集中打印)
  // ==========================================================
  const outputState = {
    regPatients: [],
    specPatients: [],
    ihcPatients: []
  };

  const outputRegCountBadge = document.getElementById("output-reg-count-badge");
  const btnRefreshOutputReg = document.getElementById("btn-refresh-output-reg");
  const btnOutputRegQuery = document.getElementById("btn-output-reg-query");
  const btnOutputRegPrintToday = document.getElementById("btn-output-reg-print-today");
  const outputRegPatientList = document.getElementById("output-reg-patient-list");
  const outputRegReportStart = document.getElementById("output-reg-report-start");
  const outputRegReportEnd = document.getElementById("output-reg-report-end");
  const btnOutputRegRepToday = document.getElementById("btn-output-reg-rep-today");
  const btnOutputRegRepWeek = document.getElementById("btn-output-reg-rep-week");
  const btnOutputRegRepMonth = document.getElementById("btn-output-reg-rep-month");
  const btnOutputRegRepAll = document.getElementById("btn-output-reg-rep-all");
  const outputRegSubStart = document.getElementById("output-reg-sub-start");
  const outputRegSubEnd = document.getElementById("output-reg-sub-end");
  const btnOutputRegSubAll = document.getElementById("btn-output-reg-sub-all");
  const outputRegPfxAll = document.getElementById("output-reg-pfx-all");
  const outputRegPfxChks = document.querySelectorAll(".output-reg-pfx-chk");

  const outputSpecCountBadge = document.getElementById("output-spec-count-badge");
  const btnRefreshOutputSpec = document.getElementById("btn-refresh-output-spec");
  const btnOutputSpecQuery = document.getElementById("btn-output-spec-query");
  const btnOutputSpecPrintToday = document.getElementById("btn-output-spec-print-today");
  const outputSpecPatientList = document.getElementById("output-spec-patient-list");
  const outputSpecReportStart = document.getElementById("output-spec-report-start");
  const outputSpecReportEnd = document.getElementById("output-spec-report-end");
  const btnOutputSpecRepToday = document.getElementById("btn-output-spec-rep-today");
  const btnOutputSpecRepWeek = document.getElementById("btn-output-spec-rep-week");
  const btnOutputSpecRepMonth = document.getElementById("btn-output-spec-rep-month");
  const btnOutputSpecRepAll = document.getElementById("btn-output-spec-rep-all");
  const outputSpecSubStart = document.getElementById("output-spec-sub-start");
  const outputSpecSubEnd = document.getElementById("output-spec-sub-end");
  const btnOutputSpecSubAll = document.getElementById("btn-output-spec-sub-all");
  const outputSpecPfxAll = document.getElementById("output-spec-pfx-all");
  const outputSpecPfxChks = document.querySelectorAll(".output-spec-pfx-chk");

  const outputIhcCountBadge = document.getElementById("output-ihc-count-badge");
  const btnRefreshOutputIhc = document.getElementById("btn-refresh-output-ihc");
  const btnOutputIhcQuery = document.getElementById("btn-output-ihc-query");
  const btnOutputIhcPrintToday = document.getElementById("btn-output-ihc-print-today");
  const outputIhcPatientList = document.getElementById("output-ihc-patient-list");
  const outputIhcReportStart = document.getElementById("output-ihc-report-start");
  const outputIhcReportEnd = document.getElementById("output-ihc-report-end");
  const btnOutputIhcRepToday = document.getElementById("btn-output-ihc-rep-today");
  const btnOutputIhcRepWeek = document.getElementById("btn-output-ihc-rep-week");
  const btnOutputIhcRepMonth = document.getElementById("btn-output-ihc-rep-month");
  const btnOutputIhcRepAll = document.getElementById("btn-output-ihc-rep-all");
  const outputIhcSubStart = document.getElementById("output-ihc-sub-start");
  const outputIhcSubEnd = document.getElementById("output-ihc-sub-end");
  const btnOutputIhcSubAll = document.getElementById("btn-output-ihc-sub-all");
  const outputIhcPfxAll = document.getElementById("output-ihc-pfx-all");
  const outputIhcPfxChks = document.querySelectorAll(".output-ihc-pfx-chk");

  async function loadOutputPatientsGroup(cfg) {
    try {
      cfg.listEl.innerHTML = `<div class="patient-list-empty">${cfg.loadingText}</div>`;
      const params = new URLSearchParams();
      params.set("type", cfg.type);
      params.set("prefixes", getPrefixGroupParam(cfg.pfxAllEl, cfg.pfxChks));
      if (cfg.repStartEl && cfg.repStartEl.value) params.set("report_start_date", cfg.repStartEl.value);
      if (cfg.repEndEl && cfg.repEndEl.value) params.set("report_end_date", cfg.repEndEl.value);
      if (cfg.subStartEl && cfg.subStartEl.value) params.set("sub_start_date", cfg.subStartEl.value);
      if (cfg.subEndEl && cfg.subEndEl.value) params.set("sub_end_date", cfg.subEndEl.value);

      const res = await fetch(`/api/output/patients?${params.toString()}`);
      const data = await res.json();
      if (data.success) {
        const list = data.patients || [];
        if (cfg.type === "special") {
          outputState.specPatients = list;
        } else if (cfg.type === "ihc") {
          outputState.ihcPatients = list;
        } else {
          outputState.regPatients = list;
        }
        cfg.badgeEl.textContent = `${list.length} 人`;
        renderOutputList(cfg.listEl, list, cfg.type);
      } else {
        cfg.listEl.innerHTML = `<div class="patient-list-empty" style="color: var(--danger);">${data.error}</div>`;
      }
    } catch (e) {
      cfg.listEl.innerHTML = `<div class="patient-list-empty" style="color: var(--danger);">加载异常: ${e.message}</div>`;
    }
  }

  function loadOutputRegularPatients() {
    return loadOutputPatientsGroup({
      type: "regular",
      listEl: outputRegPatientList,
      badgeEl: outputRegCountBadge,
      pfxAllEl: outputRegPfxAll,
      pfxChks: outputRegPfxChks,
      repStartEl: outputRegReportStart,
      repEndEl: outputRegReportEnd,
      subStartEl: outputRegSubStart,
      subEndEl: outputRegSubEnd,
      loadingText: "正在加载常规已归档报告..."
    });
  }

  function loadOutputSpecialPatients() {
    return loadOutputPatientsGroup({
      type: "special",
      listEl: outputSpecPatientList,
      badgeEl: outputSpecCountBadge,
      pfxAllEl: outputSpecPfxAll,
      pfxChks: outputSpecPfxChks,
      repStartEl: outputSpecReportStart,
      repEndEl: outputSpecReportEnd,
      subStartEl: outputSpecSubStart,
      subEndEl: outputSpecSubEnd,
      loadingText: "正在加载特殊审核完成报告..."
    });
  }

  function loadOutputIhcPatients() {
    return loadOutputPatientsGroup({
      type: "ihc",
      listEl: outputIhcPatientList,
      badgeEl: outputIhcCountBadge,
      pfxAllEl: outputIhcPfxAll,
      pfxChks: outputIhcPfxChks,
      repStartEl: outputIhcReportStart,
      repEndEl: outputIhcReportEnd,
      subStartEl: outputIhcSubStart,
      subEndEl: outputIhcSubEnd,
      loadingText: "正在加载免疫组化已归档报告..."
    });
  }

  // Render items in Output list
  function renderOutputList(container, patients, reportType = "regular") {
    const isSpecial = (reportType === "special" || reportType === true);
    const isIhc = (reportType === "ihc");
    if (!patients.length) {
      container.innerHTML = `<div class="patient-list-empty">暂无符合条件的报告单</div>`;
      return;
    }

    container.innerHTML = "";
    patients.forEach(p => {
      const card = document.createElement("div");
      const hasOutExtMod = isIhc ? !!p.ihc_external_modified : !!p.external_modified;
      card.className = `output-item-card ${hasOutExtMod ? "external-modified-card" : ""}`.trim();

      const ageStr = formatAgeDisplay(p.age);
      const diagStr = p.pathology_diag || p.clinical_diagnosis || "未诊断";
      const repDate = p.report_date || p.submission_date || "-";

      let statusBadge = "";
      if (isSpecial) {
        statusBadge = `<span class="badge-status badge-approved" style="font-size: 11px;">✓ 完成审核</span>`;
      } else if (isIhc) {
        statusBadge = p.archived === 2
          ? `<span class="badge-status badge-rearchive" style="font-size: 11px;">需重新归档</span>`
          : `<span class="badge-status badge-ihc" style="font-size: 11px;">🧪 免疫组化报告</span>`;
      } else {
        statusBadge = `<span class="badge-status badge-archived" style="font-size: 11px;">已归档</span>`;
        if (p.is_needs_ihc === 1) {
          statusBadge += ` <span class="badge-status badge-ihc" style="font-size: 11px;">免疫组化</span>`;
        }
      }
      if (hasOutExtMod) {
        statusBadge += ` ${renderExternalModBadge(p, isIhc)}`;
      }

      card.innerHTML = `
        <div class="output-item-header">
          <div>
            <span class="output-item-name">${p.name}</span>
            <span class="output-item-no">${p.pathology_no}</span>
          </div>
          <div>${statusBadge}</div>
        </div>
        <div class="output-item-meta">
          <span>${p.gender || "-"} · ${ageStr}</span>
          <span>日期: ${repDate}</span>
          ${p.review_doctor ? `<span>审核: ${p.review_doctor}</span>` : ""}
        </div>
        <div class="output-item-diag" title="${diagStr}">
          <strong>诊断:</strong> ${diagStr}
        </div>
        <div class="output-item-actions">
          <button type="button" class="btn btn-secondary btn-sm btn-out-preview" title="直接预览当前报告">👁️ 预览</button>
          <button type="button" class="btn btn-secondary btn-sm btn-out-word" title="在外部Word中打开报告">📄 Word</button>
          <button type="button" class="btn btn-secondary btn-sm btn-out-folder" title="在资源管理器中定位归档报告（已归档报告）">📂 归档位置</button>
          <button type="button" class="btn btn-secondary btn-sm btn-out-print" title="单独打印当前报告第一页" style="color: #0369a1; border-color: #bae6fd;">🖨️ 打印第1页</button>
          <button type="button" class="btn btn-secondary btn-sm btn-out-revoke-arch" title="撤销归档：修改数据库并从已归档报告中移除对应文件" style="color: #c2410c; border-color: #fed7aa;">↩️ 撤销归档</button>
          <button type="button" class="btn btn-secondary btn-sm btn-out-revoke-rev" ${!isSpecial ? "disabled" : ""} title="${isSpecial ? '撤销审核：返回待审核列表并恢复对应特殊归档' : (isIhc ? '免疫组化报告无需审核' : '常规报告无需审核')}" style="${isSpecial ? 'color: #9333ea; border-color: #e9d5ff;' : 'opacity: 0.5; cursor: not-allowed;'}">↩️ 撤销审核</button>
        </div>
      `;

      // 1. Preview
      card.querySelector(".btn-out-preview").addEventListener("click", () => {
        openReportPreviewModal(p.pathology_no, isIhc ? "ihc" : "normal");
      });

      // 2. Open Word from output / IHC draft directory
      card.querySelector(".btn-out-word").addEventListener("click", async () => {
        const reportPath = p.final_report_path || p.word_path;
        if (reportPath) {
          await fetch("/api/open_file", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ path: reportPath })
          });
        } else {
          showToast("未检测到报告文件路径", "error");
        }
      });

      // 3. Open archive dir in Explorer (Requirement 3.1)
      card.querySelector(".btn-out-folder").addEventListener("click", async () => {
        try {
          const res = await fetch("/api/output/open_archive_dir", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ pathology_no: p.pathology_no, report_type: reportType })
          });
          const data = await res.json();
          if (data.success) {
            showToast(data.message, "success");
          } else {
            showToast("打开目录失败: " + data.error, "error");
          }
        } catch (err) {
          showToast("打开异常: " + err.message, "error");
        }
      });

      // 4. Print Page 1
      card.querySelector(".btn-out-print").addEventListener("click", async () => {
        if (hasOutExtMod) {
          await resolveExternalSyncForPatient(p, isIhc ? "ihc" : "normal", false);
        }
        const btnPrint = card.querySelector(".btn-out-print");
        btnPrint.disabled = true;
        btnPrint.innerHTML = `<span>⏳ 打印中...</span>`;
        try {
          const res = await fetch("/api/output/batch_print", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ pathology_nos: [p.pathology_no], report_type: reportType })
          });
          const data = await res.json();
          if (data.success) {
            if (data.direct_printed) {
              showToast(`患者 [${p.name}] 报告第1页已调用Word发送至默认打印机！`, "success");
            } else {
              showToast(data.message || "Word打印失败，请检查Word及默认打印机", "error");
            }
          } else {
            showToast("打印失败: " + data.error, "error");
          }
        } catch (err) {
          showToast("打印请求异常: " + err.message, "error");
        } finally {
          btnPrint.disabled = false;
          btnPrint.innerHTML = `<span>🖨️ 打印第1页</span>`;
        }
      });

      // 5. Revoke Archive
      card.querySelector(".btn-out-revoke-arch").addEventListener("click", async () => {
        const confirmMsg = isIhc
          ? `确定要撤销患者 [${p.name}] (病理号: ${p.pathology_no}) 的免疫组化报告归档吗？\n撤销后将自动从“已归档报告”中移除该免疫组化报告文件（不会影响原报告和“特殊患者归档\\免疫组化”目录中的草稿）。`
          : `确定要撤销患者 [${p.name}] (病理号: ${p.pathology_no}) 的归档吗？\n撤销后将自动从“已归档报告”中移除该报告文件，并将其从本列表移出。`;
        if (!confirm(confirmMsg)) {
          return;
        }
        try {
          const data = await fetchJsonWithSafetyConfirm("/api/output/revoke_archive", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ pathology_no: p.pathology_no, report_type: reportType })
          });
          if (data.success) {
            showToast(data.message, "success");
            if (isSpecial) loadOutputSpecialPatients();
            else if (isIhc) loadOutputIhcPatients();
            else loadOutputRegularPatients();
          } else {
            showToast("撤销归档失败: " + data.error, "error");
          }
        } catch (err) {
          showToast("操作异常: " + err.message, "error");
        }
      });

      // 6. Revoke Approval (Only for Special Approved list)
      const btnRevokeRev = card.querySelector(".btn-out-revoke-rev");
      if (isSpecial) {
        btnRevokeRev.addEventListener("click", async () => {
          if (!confirm(`确定要撤销患者 [${p.name}] (病理号: ${p.pathology_no}) 的“完成审核”状态吗？\n撤销后患者将返回报告审核工作区，并同步恢复特殊归档文件夹。`)) {
            return;
          }
          try {
            const data = await fetchJsonWithSafetyConfirm("/api/output/revoke_approval", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ pathology_no: p.pathology_no })
            });
            if (data.success) {
              showToast(data.message, "success");
              loadOutputSpecialPatients();
            } else {
              showToast("撤销审核失败: " + data.error, "error");
            }
          } catch (err) {
            showToast("操作异常: " + err.message, "error");
          }
        });
      }

      container.appendChild(card);
    });
  }

  // Filter Prefix, Date Ranges & Refresh bindings for Output
  bindPrefixCheckboxGroup(outputRegPfxAll, outputRegPfxChks, loadOutputRegularPatients);
  bindPrefixCheckboxGroup(outputSpecPfxAll, outputSpecPfxChks, loadOutputSpecialPatients);
  bindPrefixCheckboxGroup(outputIhcPfxAll, outputIhcPfxChks, loadOutputIhcPatients);

  bindRangeQuickDateButtons({
    startEl: outputRegReportStart,
    endEl: outputRegReportEnd,
    btnTodayEl: btnOutputRegRepToday,
    btnWeekEl: btnOutputRegRepWeek,
    btnMonthEl: btnOutputRegRepMonth,
    btnAllEl: btnOutputRegRepAll,
    onReload: loadOutputRegularPatients
  });
  bindRangeQuickDateButtons({
    startEl: outputRegSubStart,
    endEl: outputRegSubEnd,
    btnAllEl: btnOutputRegSubAll,
    onReload: loadOutputRegularPatients
  });

  bindRangeQuickDateButtons({
    startEl: outputSpecReportStart,
    endEl: outputSpecReportEnd,
    btnTodayEl: btnOutputSpecRepToday,
    btnWeekEl: btnOutputSpecRepWeek,
    btnMonthEl: btnOutputSpecRepMonth,
    btnAllEl: btnOutputSpecRepAll,
    onReload: loadOutputSpecialPatients
  });
  bindRangeQuickDateButtons({
    startEl: outputSpecSubStart,
    endEl: outputSpecSubEnd,
    btnAllEl: btnOutputSpecSubAll,
    onReload: loadOutputSpecialPatients
  });

  bindRangeQuickDateButtons({
    startEl: outputIhcReportStart,
    endEl: outputIhcReportEnd,
    btnTodayEl: btnOutputIhcRepToday,
    btnWeekEl: btnOutputIhcRepWeek,
    btnMonthEl: btnOutputIhcRepMonth,
    btnAllEl: btnOutputIhcRepAll,
    onReload: loadOutputIhcPatients
  });
  bindRangeQuickDateButtons({
    startEl: outputIhcSubStart,
    endEl: outputIhcSubEnd,
    btnAllEl: btnOutputIhcSubAll,
    onReload: loadOutputIhcPatients
  });

  if (btnOutputRegQuery) btnOutputRegQuery.addEventListener("click", loadOutputRegularPatients);
  if (btnOutputSpecQuery) btnOutputSpecQuery.addEventListener("click", loadOutputSpecialPatients);
  if (btnOutputIhcQuery) btnOutputIhcQuery.addEventListener("click", loadOutputIhcPatients);
  btnRefreshOutputReg.addEventListener("click", loadOutputRegularPatients);
  btnRefreshOutputSpec.addEventListener("click", loadOutputSpecialPatients);
  if (btnRefreshOutputIhc) btnRefreshOutputIhc.addEventListener("click", loadOutputIhcPatients);

  async function runBatchPrintToday(reportType, pfxAllEl, pfxChks, btnEl, emptyMsg, confirmLabel) {
    btnEl.disabled = true;
    btnEl.innerHTML = `<span>⏳ 正在检索今日报告...</span>`;
    try {
      const pfx = getPrefixGroupParam(pfxAllEl, pfxChks);
      const res = await fetch(`/api/output/patients?type=${encodeURIComponent(reportType)}&prefixes=${encodeURIComponent(pfx)}&today_only=1`);
      const data = await res.json();
      const todayList = data.patients || [];

      if (!todayList.length) {
        showToast(emptyMsg, "error");
        return;
      }
      if (!confirm(`检测到今日共有 ${todayList.length} 份${confirmLabel}，确定要全部调用Word打印（仅打印第1页）吗？`)) {
        return;
      }

      btnEl.innerHTML = `<span>⏳ 正在批量打印 ${todayList.length} 份...</span>`;
      const printRes = await fetch("/api/output/batch_print", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pathology_nos: todayList.map(p => p.pathology_no), report_type: reportType })
      });
      const printData = await printRes.json();
      if (printData.success && printData.direct_printed) {
        showToast(printData.message, "success");
      } else {
        showToast(printData.message || ("批量打印异常: " + (printData.error || "部分或全部打印失败")), "error");
      }
    } catch (err) {
      showToast("请求失败: " + err.message, "error");
    } finally {
      btnEl.disabled = false;
      btnEl.innerHTML = `<span>🖨️ 一键打印今日已归档报告 (第1页)</span>`;
    }
  }

  // Batch Print Today Regular
  btnOutputRegPrintToday.addEventListener("click", () => {
    runBatchPrintToday("regular", outputRegPfxAll, outputRegPfxChks, btnOutputRegPrintToday, "未找到今日已归档的常规报告单", "已归档常规报告");
  });

  // Batch Print Today Special
  btnOutputSpecPrintToday.addEventListener("click", () => {
    runBatchPrintToday("special", outputSpecPfxAll, outputSpecPfxChks, btnOutputSpecPrintToday, "未找到今日已归档的特殊审核完成报告", "特殊审核完成报告");
  });

  // Batch Print Today IHC
  if (btnOutputIhcPrintToday) {
    btnOutputIhcPrintToday.addEventListener("click", () => {
      runBatchPrintToday("ihc", outputIhcPfxAll, outputIhcPfxChks, btnOutputIhcPrintToday, "未找到今日已归档的免疫组化报告", "已归档免疫组化报告");
    });
  }

  // ==========================================================
  // Phase 3.2: Signing Sheet Workspace (3.2 报告领取签名单)
  // ==========================================================
  const signingState = {
    candidates: [],
    generatedFilePath: null,
    generatedFilename: null,
    generatedWordPath: null,
    generatedWordFilename: null,
    defaultsInitialized: false
  };

  const signingReportStartDate = document.getElementById("signing-report-start-date");
  const signingReportEndDate = document.getElementById("signing-report-end-date");
  const signingStartDate = document.getElementById("signing-start-date");
  const signingEndDate = document.getElementById("signing-end-date");
  const btnSigningDateToday = document.getElementById("btn-signing-date-today");
  const btnSigningDateWeek = document.getElementById("btn-signing-date-week");
  const btnSigningDateMonth = document.getElementById("btn-signing-date-month");
  const btnSigningDateAll = document.getElementById("btn-signing-date-all");
  const btnSigningSubAll = document.getElementById("btn-signing-sub-all");

  const pfxChkAll = document.getElementById("pfx-chk-all");
  const pfxItemChks = document.querySelectorAll(".pfx-item-chk");

  const btnSigningQuery = document.getElementById("btn-signing-query");
  const btnSigningGenerate = document.getElementById("btn-signing-generate");
  const signingCandidatesBadge = document.getElementById("signing-candidates-badge");
  const signingSelectAllToggle = document.getElementById("signing-select-all-toggle");
  const signingCandidatesTbody = document.getElementById("signing-candidates-tbody");

  const signingResultCard = document.getElementById("signing-result-card");
  const signingGeneratedFilename = document.getElementById("signing-generated-filename");
  const btnSigningOpenExcel = document.getElementById("btn-signing-open-excel");
  const btnSigningOpenWord = document.getElementById("btn-signing-open-word");
  const btnSigningOpenDir = document.getElementById("btn-signing-open-dir");
  const btnSigningPrintSheet = document.getElementById("btn-signing-print-sheet");
  const signingPreviewContainer = document.getElementById("signing-preview-container");

  function initSigningDateRange() {
    if (!signingState.defaultsInitialized && signingReportStartDate && signingReportEndDate) {
      const today = getTodayDateStr();
      signingReportStartDate.value = today;
      signingReportEndDate.value = today;
      signingState.defaultsInitialized = true;
    }
  }

  // Quick date range buttons (Report date: 今日/本周/本月/全部; Submission date: 全部)
  bindRangeQuickDateButtons({
    startEl: signingReportStartDate,
    endEl: signingReportEndDate,
    btnTodayEl: btnSigningDateToday,
    btnWeekEl: btnSigningDateWeek,
    btnMonthEl: btnSigningDateMonth,
    btnAllEl: btnSigningDateAll,
    onReload: loadSigningCandidates
  });
  bindRangeQuickDateButtons({
    startEl: signingStartDate,
    endEl: signingEndDate,
    btnAllEl: btnSigningSubAll,
    onReload: loadSigningCandidates
  });

  // Prefix checkbox synchronization
  bindPrefixCheckboxGroup(pfxChkAll, pfxItemChks);

  function getSelectedPrefixesParam() {
    return getPrefixGroupParam(pfxChkAll, pfxItemChks);
  }

  // Load Candidates
  async function loadSigningCandidates() {
    try {
      signingCandidatesTbody.innerHTML = `
        <tr><td colspan="8" style="text-align: center; color: var(--text-muted); padding: 30px;">正在检索数据库患者记录...</td></tr>
      `;
      const repS = signingReportStartDate ? signingReportStartDate.value : "";
      const repE = signingReportEndDate ? signingReportEndDate.value : "";
      const sDate = signingStartDate.value;
      const eDate = signingEndDate.value;
      const pfxParam = getSelectedPrefixesParam();

      const params = new URLSearchParams();
      if (repS) params.set("report_start_date", repS);
      if (repE) params.set("report_end_date", repE);
      if (sDate) params.set("start_date", sDate);
      if (eDate) params.set("end_date", eDate);
      if (pfxParam) params.set("prefixes", pfxParam);

      const res = await fetch(`/api/signing/patients?${params.toString()}`);
      const data = await res.json();

      if (data.success) {
        signingState.candidates = data.patients || [];
        signingCandidatesBadge.textContent = `${signingState.candidates.length} 人`;
        renderSigningCandidatesTable();
      } else {
        signingCandidatesTbody.innerHTML = `<tr><td colspan="8" style="color: red; text-align: center; padding: 20px;">${data.error}</td></tr>`;
      }
    } catch (err) {
      signingCandidatesTbody.innerHTML = `<tr><td colspan="8" style="color: red; text-align: center; padding: 20px;">加载失败: ${err.message}</td></tr>`;
    }
  }

  function renderSigningCandidatesTable() {
    if (!signingState.candidates.length) {
      signingCandidatesTbody.innerHTML = `
        <tr><td colspan="8" style="text-align: center; color: var(--text-muted); padding: 30px;">所选时间与前缀范围内未找到任何患者</td></tr>
      `;
      return;
    }

    signingCandidatesTbody.innerHTML = "";
    signingState.candidates.forEach((p, idx) => {
      const tr = document.createElement("tr");
      const ageStr = formatAgeDisplay(p.age);

      tr.innerHTML = `
        <td style="text-align: center;">
          <input type="checkbox" class="signing-patient-select-chk" data-no="${p.pathology_no}" checked>
        </td>
        <td style="font-family: monospace; font-weight: 600; color: var(--primary);">${p.pathology_no}</td>
        <td style="font-weight: 600;">${p.name}</td>
        <td>${p.gender || "-"}</td>
        <td>${ageStr}</td>
        <td>${p.referring_doctor || "-"}</td>
        <td>${p.submission_date || "-"}</td>
        <td>${p.clinical_diagnosis || "-"}</td>
      `;

      signingCandidatesTbody.appendChild(tr);
    });
  }

  signingSelectAllToggle.addEventListener("change", () => {
    const chks = signingCandidatesTbody.querySelectorAll(".signing-patient-select-chk");
    chks.forEach(c => c.checked = signingSelectAllToggle.checked);
  });

  btnSigningQuery.addEventListener("click", loadSigningCandidates);

  // Generate Excel Signing Sheet (Requirement 3.2)
  btnSigningGenerate.addEventListener("click", async () => {
    const selectedChks = signingCandidatesTbody.querySelectorAll(".signing-patient-select-chk:checked");
    if (!selectedChks.length) {
      showToast("请至少勾选一位患者生成签名单", "error");
      return;
    }

    const selectedNos = Array.from(selectedChks).map(c => c.dataset.no);

    btnSigningGenerate.disabled = true;
    btnSigningGenerate.innerHTML = `<span>⏳ 正在生成签名单 (Excel + Word)...</span>`;

    try {
      const data = await fetchJsonWithSafetyConfirm("/api/signing/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          pathology_nos: selectedNos,
          start_date: signingStartDate.value,
          end_date: signingEndDate.value,
          report_start_date: signingReportStartDate ? signingReportStartDate.value : "",
          report_end_date: signingReportEndDate ? signingReportEndDate.value : ""
        })
      });

      if (data.success) {
        showToast(data.message, "success");
        signingState.generatedFilePath = data.file_path;
        signingState.generatedFilename = data.filename;
        signingState.generatedWordPath = data.word_path || null;
        signingState.generatedWordFilename = data.word_filename || null;

        signingGeneratedFilename.textContent = data.word_filename || data.filename;
        signingGeneratedFilename.title = data.word_path || data.file_path;
        signingPreviewContainer.innerHTML = data.html;
        signingResultCard.style.display = "flex";

        // Scroll into view on mobile/tablet
        signingResultCard.scrollIntoView({ behavior: "smooth" });
      } else {
        showToast("生成失败: " + data.error, "error");
      }
    } catch (err) {
      showToast("生成签名单请求异常: " + err.message, "error");
    } finally {
      btnSigningGenerate.disabled = false;
      btnSigningGenerate.innerHTML = `<span>🚀 生成签名单 (Excel + Word)</span>`;
    }
  });

  // Open generated Excel
  btnSigningOpenExcel.addEventListener("click", async () => {
    if (!signingState.generatedFilePath) return;
    await fetch("/api/open_file", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: signingState.generatedFilePath })
    });
  });

  // Open generated Word
  if (btnSigningOpenWord) {
    btnSigningOpenWord.addEventListener("click", async () => {
      if (!signingState.generatedWordPath) return;
      await fetch("/api/open_file", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ path: signingState.generatedWordPath })
      });
    });
  }

  // Open 领取表 folder in Explorer
  btnSigningOpenDir.addEventListener("click", async () => {
    const targetPath = signingState.generatedWordPath || signingState.generatedFilePath;
    if (!targetPath) return;
    await fetch("/api/open_explorer_select", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: targetPath })
    });
  });

  // Print Signing Sheet (Word)
  btnSigningPrintSheet.addEventListener("click", async () => {
    const targetPath = signingState.generatedWordPath || signingState.generatedFilePath;
    if (!targetPath) return;
    try {
      const res = await fetch("/api/signing/print", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          file_path: targetPath,
          word_path: signingState.generatedWordPath
        })
      });
      const data = await res.json();
      if (data.success) {
        if (data.direct_printed) {
          showToast("Word签名单已直接发送至默认打印机！", "success");
        } else {
          showToast(data.message || "调用Word打印签名单失败，请检查Word及默认打印机", "error");
        }
      } else {
        showToast("打印失败: " + data.error, "error");
      }
    } catch (err) {
      showToast("打印异常: " + err.message, "error");
    }
  });

  // Start initialization check
  checkSystemStatus();

  // Disclaimer modal controls (显示免责声明)
  const disclaimerModal = document.getElementById("disclaimer-modal");
  const btnDisclaimerAgree = document.getElementById("btn-disclaimer-agree");
  const btnDisclaimerExit = document.getElementById("btn-disclaimer-exit");

  if (btnDisclaimerAgree && disclaimerModal) {
    btnDisclaimerAgree.addEventListener("click", () => {
      disclaimerModal.style.display = "none";
    });
  }

  if (btnDisclaimerExit && disclaimerModal) {
    btnDisclaimerExit.addEventListener("click", async () => {
      try {
        await fetch("/api/system/exit", { method: "POST" });
      } catch (e) {}
      document.body.innerHTML = `
        <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; font-family: sans-serif; background: #f8fafc; color: #334155; text-align: center; padding: 20px;">
          <h2 style="color: #64748b; margin-bottom: 12px; font-size: 24px;">⚖️ 已安全退出系统</h2>
          <p style="color: #94a3b8; font-size: 15px; max-width: 480px; line-height: 1.6;">您已拒绝免责声明，系统后台服务已停止。您可以直接关闭当前浏览器窗口或标签页。</p>
        </div>
      `;
      try {
        window.close();
      } catch (e) {}
    });
  }
}

// Safely execute startApp whether DOMContentLoaded has already fired or not
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", startApp);
} else {
  startApp();
}
