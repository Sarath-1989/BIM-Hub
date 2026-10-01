(() => {
  "use strict";

  const CONFIG = window.ASTERWIX_PORTAL_CONFIG;
  const GRAPH_SCOPES = ["User.Read", "Sites.ReadWrite.All"];
  const FOLDERS = { employees: "employees", projects: "projects", tasks: "tasks", workspaces: "employee-workspaces", issues: "bim-issues", registers: "bim-registers" };
  const CORE_FOLDERS = ["employees", "projects", "tasks"];
  const TASK_STATUSES = ["Not started", "In progress", "Ready for QA/QC", "QA/QC review", "Revise & resubmit", "Ready to submit", "Submitted", "Client review", "Approved", "Blocked", "Completed"];
  const STAFF_TASK_STATUSES = ["Not started", "In progress", "Ready for QA/QC", "Revise & resubmit", "Ready to submit", "Blocked"];
  const state = { account: null, profile: null, role: "Staff", projects: [], tasks: [], workLogs: [], employees: [], issues: [], registers: [], missingFolders: [], bimStorage: { issues: false, registers: false }, inactive: false };
  const byId = (id) => document.getElementById(id);
  let editingEmployeeEmail = "";
  let editingProjectId = "";
  let editingTaskId = "";
  let editingWorkLogId = "";
  const isManager = () => state.role === "Admin" || state.role === "Team Lead";
  const isAdmin = () => state.role === "Admin";
  const isCoordinator = () => state.role === "Team Lead";
  const dubaiDate = () => new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Dubai", year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
  const esc = (value = "") => String(value).replace(/[&<>'"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[c]));
  const emailKey = (email) => `employee-${btoa(unescape(encodeURIComponent(email.toLowerCase()))).replace(/[+/=]/g, "-")}`;
  const recordId = (prefix) => `${prefix}-${crypto.randomUUID()}`;
  let msalInstance;

  function toast(message, kind = "") {
    const item = document.createElement("div");
    item.className = `toast ${kind}`;
    item.textContent = message;
    byId("toast-region").append(item);
    setTimeout(() => item.remove(), 4500);
  }

  function setSync(text, error = false) {
    const el = byId("sync-status");
    el.textContent = text;
    el.style.color = error ? "#c64040" : "";
  }

  function msalConfig() {
    return {
      auth: { clientId: CONFIG.clientId, authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`, redirectUri: window.location.origin + window.location.pathname, navigateToLoginRequestUrl: false },
      cache: { cacheLocation: "sessionStorage" }
    };
  }

  async function initialiseAuth() {
    if (!window.msal || !CONFIG?.clientId || !CONFIG?.tenantId || !CONFIG?.driveId) throw new Error("Microsoft login or SharePoint storage configuration is missing.");
    msalInstance = new msal.PublicClientApplication(msalConfig());
    await msalInstance.initialize();
    const response = await msalInstance.handleRedirectPromise();
    state.account = response?.account || msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
    if (state.account) msalInstance.setActiveAccount(state.account);
  }

  async function signIn() {
    try {
      byId("sign-in-status").textContent = "Opening Asterwix Microsoft sign-in…";
      const response = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES, prompt: "select_account" });
      msalInstance.setActiveAccount(response.account);
      state.account = response.account;
      await openPortal();
    } catch (error) {
      console.error(error);
      byId("sign-in-status").textContent = "Microsoft sign-in could not complete.";
      toast("Microsoft sign-in could not complete. Please try again.", "error");
    }
  }

  async function accessToken() {
    if (!state.account) throw new Error("Sign in is required.");
    try { return (await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: state.account })).accessToken; }
    catch { return (await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES, account: state.account })).accessToken; }
  }

  async function graph(path, options = {}) {
    const token = await accessToken();
    const graphUrl = path.startsWith("https://graph.microsoft.com/v1.0/") ? path : `https://graph.microsoft.com/v1.0${path}`;
    const response = await fetch(graphUrl, { ...options, headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(options.headers || {}) } });
    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      const error = new Error(body?.error?.message || `SharePoint request failed (${response.status}).`);
      error.status = response.status;
      throw error;
    }
    if (response.status === 204) return null;
    return (response.headers.get("content-type") || "").includes("json") ? response.json() : response.text();
  }

  async function listGraphCollection(path) {
    const values = [];
    let next = path;
    while (next) {
      const response = await graph(next);
      values.push(...(response.value || []));
      next = response["@odata.nextLink"] || "";
    }
    return values;
  }

  async function mapWithConcurrency(items, mapper, limit = 8) {
    const results = new Array(items.length);
    let nextIndex = 0;
    const worker = async () => {
      while (nextIndex < items.length) {
        const index = nextIndex++;
        results[index] = await mapper(items[index]);
      }
    };
    await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
    return results;
  }
  function folderPath(key) { return `${CONFIG.storageFolder}/${FOLDERS[key]}`; }
  function filePath(key, id) { return `${folderPath(key)}/${id}.json`; }
  function workspacePath(email = accountEmail()) { return `${folderPath("workspaces")}/${emailKey(String(email).toLowerCase())}`; }
  function workspaceLogsPath(email = accountEmail()) { return `${workspacePath(email)}/work-logs`; }
  function drivePath(path, suffix = "") { return `/drives/${CONFIG.driveId}/root:/${path}:${suffix}`; }

  async function checkPaths(paths, append = false) {
    const checks = await Promise.all(paths.map(async ({ label, path }) => {
      try { await graph(drivePath(path)); return null; }
      catch (error) { if (error.status === 404) return label; throw error; }
    }));
    const missing = checks.filter(Boolean);
    state.missingFolders = append ? [...state.missingFolders, ...missing] : missing;
  }

  async function checkFolders() {
    await checkPaths(CORE_FOLDERS.map((key) => ({ label: key, path: folderPath(key) })));
  }

  async function getFolder(path) { return graph(drivePath(path)); }

  async function ensureFolder(parentPath, name) {
    const path = `${parentPath}/${name}`;
    try { return await getFolder(path); }
    catch (error) {
      if (error.status !== 404) throw error;
      try {
        return await graph(drivePath(parentPath, "/children"), { method: "POST", body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }) });
      } catch (createError) {
        if (createError.status === 409) return getFolder(path);
        throw createError;
      }
    }
  }

  async function ensureWorkspace(email = accountEmail()) {
    await ensureFolder(CONFIG.storageFolder, FOLDERS.workspaces);
    const workspace = await ensureFolder(folderPath("workspaces"), emailKey(String(email).toLowerCase()));
    await ensureFolder(`${folderPath("workspaces")}/${workspace.name}`, "work-logs");
    return workspace;
  }

  async function inviteToFolder(path, email, role) {
    const item = await getFolder(path);
    await graph(`/drives/${CONFIG.driveId}/items/${item.id}/invite`, {
      method: "POST",
      body: JSON.stringify({ recipients: [{ email }], roles: [role], requireSignIn: true, sendInvitation: false, retainInheritedPermissions: true })
    });
  }

  async function inviteToItem(path, email, role) {
    const item = await graph(drivePath(path));
    await graph(`/drives/${CONFIG.driveId}/items/${item.id}/invite`, {
      method: "POST",
      body: JSON.stringify({ recipients: [{ email }], roles: [role], requireSignIn: true, sendInvitation: false, retainInheritedPermissions: true })
    });
  }

  async function maybeInviteToFolder(key, email, role) {
    try { await inviteToFolder(folderPath(key), email, role); }
    catch (error) { if (error.status !== 404) throw error; }
  }

  function permissionEmails(permission) {
    const identities = [permission.grantedToV2, permission.grantedTo, ...(permission.grantedToIdentitiesV2 || []), ...(permission.grantedToIdentities || [])];
    return identities.flatMap((identity) => [identity?.user?.email, identity?.user?.userPrincipalName, identity?.siteUser?.email, identity?.siteUser?.loginName]).filter(Boolean).map((email) => String(email).split("|").pop().toLowerCase());
  }

  async function revokeDirectAccess(path, email) {
    let item;
    try { item = await getFolder(path); }
    catch (error) { if (error.status === 404) return; throw error; }
    const permissions = await listGraphCollection(`/drives/${CONFIG.driveId}/items/${item.id}/permissions`);
    const matching = permissions.filter((permission) => permissionEmails(permission).includes(String(email).toLowerCase()));
    for (const permission of matching) {
      try { await graph(`/drives/${CONFIG.driveId}/items/${item.id}/permissions/${permission.id}`, { method: "DELETE" }); }
      catch (error) { if (error.status !== 404) throw error; }
    }
  }
  async function revokeEmployeePortalAccess(employee) {
    const email = String(employee.Email || "").toLowerCase();
    if (!email || employee.Role === "Admin") return;
    const paths = [folderPath("employees"), folderPath("projects"), folderPath("tasks"), folderPath("workspaces"), workspacePath(email), folderPath("issues"), folderPath("registers"), ...state.tasks.filter((task) => (task.AssigneeEmail || "").toLowerCase() === email && task.id).map((task) => filePath("tasks", task.id))];
    for (const path of [...new Set(paths)]) await revokeDirectAccess(path, email);
  }

  async function provisionEmployeeWorkspace(employee) {
    const email = String(employee.Email || "").toLowerCase();
    if (!email || employee.Active === "No") return;
    await ensureWorkspace(email);
    if (employee.Role === "Admin") return;
    await inviteToFolder(folderPath("employees"), email, "read");
    await inviteToFolder(folderPath("projects"), email, employee.Role === "Team Lead" ? "write" : "read");
    await inviteToFolder(folderPath("tasks"), email, employee.Role === "Team Lead" ? "write" : "read");
    await inviteToFolder(workspacePath(email), email, "write");
    await Promise.all(state.tasks.filter((task) => !isTaskInRecycleBin(task) && (task.AssigneeEmail || "").toLowerCase() === email && task.id).map((task) => inviteToItem(filePath("tasks", task.id), email, "write")));
    if (employee.Role === "Team Lead") {
      await inviteToFolder(folderPath("workspaces"), email, "read");
      await maybeInviteToFolder("issues", email, "write");
      await maybeInviteToFolder("registers", email, "write");
    }
  }

  async function readRecord(item) {
    try { return await graph(`/drives/${CONFIG.driveId}/items/${item.id}/content`); }
    catch (error) { console.warn("Skipped unreadable record", item.name, error); return null; }
  }

  async function listRecordsAt(path) {
    const items = await listGraphCollection(`${drivePath(path, "/children")}?$top=999`);
    const records = await mapWithConcurrency(items.filter((item) => item.file && item.name.endsWith(".json")), readRecord);
    return records.filter(Boolean);
  }
  async function listRecords(key) { return listRecordsAt(folderPath(key)); }

  async function recordItemsAt(path) {
    try {
      const items = await listGraphCollection(`${drivePath(path, "/children")}?$top=999`);
      return items.filter((item) => item.file && item.name.endsWith(".json"));
    } catch (error) {
      if (error.status === 404) return [];
      throw error;
    }
  }

  async function workspaceLogRecordItems() {
    let folders;
    try {
      folders = (await listGraphCollection(`${drivePath(folderPath("workspaces"), "/children")}?$top=999`)).filter((item) => item.folder);
    } catch (error) {
      if (error.status === 404) return [];
      throw error;
    }
    const groups = await mapWithConcurrency(folders, async (folder) => recordItemsAt(`${folderPath("workspaces")}/${folder.name}/work-logs`), 6);
    return groups.flat();
  }

  async function portalResetRecordGroups() {
    const [tasks, workLogs] = await Promise.all([
      recordItemsAt(folderPath("tasks")),
      workspaceLogRecordItems()
    ]);
    return [
      { key: "tasks", label: "task assignment", items: tasks },
      { key: "workLogs", label: "work-hour log", items: workLogs }
    ];
  }

  async function deleteDriveItems(items) {
    const results = await mapWithConcurrency(items, async (item) => {
      try {
        await graph(`/drives/${CONFIG.driveId}/items/${item.id}`, { method: "DELETE" });
        return { item, deleted: true };
      } catch (error) {
        return { item, error };
      }
    }, 6);
    const failures = results.filter((result) => result?.error);
    return { deleted: items.length - failures.length, failures };
  }

  async function listOptionalRecords(key) {
    try {
      const records = await listRecords(key);
      state.bimStorage[key] = true;
      return records;
    } catch (error) {
      if (error.status === 404 || error.status === 403) {
        state.bimStorage[key] = false;
        return [];
      }
      throw error;
    }
  }

  async function saveRecordAt(path, id, fields) {
    const record = { ...fields, id, updatedAt: new Date().toISOString() };
    await graph(drivePath(`${path}/${id}.json`, "/content"), { method: "PUT", body: JSON.stringify(record) });
    return record;
  }

  async function saveRecord(key, id, fields) {
    return saveRecordAt(folderPath(key), id, fields);
  }

  async function listAllWorkspaceLogs() {
    const folders = (await listGraphCollection(`${drivePath(folderPath("workspaces"), "/children")}?$top=999`)).filter((item) => item.folder);
    const logs = await mapWithConcurrency(folders, async (folder) => {
      try { return await listRecordsAt(`${folderPath("workspaces")}/${folder.name}/work-logs`); }
      catch (error) { console.warn("Skipped unreadable employee workspace", folder.name, error); return []; }
    }, 6);
    return logs.flat();
  }
  async function loadData() {
    state.issues = []; state.registers = []; state.bimStorage = { issues: false, registers: false }; state.inactive = false;
    await checkFolders();
    if (state.missingFolders.length) return;
    [state.projects, state.tasks, state.employees] = await Promise.all([listRecords("projects"), listRecords("tasks"), listRecords("employees")]);
    setRole();
    if (state.inactive) return;
    await checkPaths([{ label: isManager() ? FOLDERS.workspaces : "your personal work folder", path: isManager() ? folderPath("workspaces") : workspaceLogsPath() }], true);
    if (state.missingFolders.length) return;
    const loadedWorkLogs = isManager() ? await listAllWorkspaceLogs() : await listRecordsAt(workspaceLogsPath());
    const email = accountEmail().toLowerCase();
    state.workLogs = isAdmin() ? loadedWorkLogs : isCoordinator() ? loadedWorkLogs.filter((entry) => canManageProject(projectByCode(entry.ProjectCode)) || (entry.EmployeeEmail || "").toLowerCase() === email) : loadedWorkLogs;
    if (isManager()) [state.issues, state.registers] = await Promise.all([listOptionalRecords("issues"), listOptionalRecords("registers")]);
  }
  function accountEmail() { return state.profile?.mail || state.profile?.userPrincipalName || state.account?.username || ""; }

  function setRole() {
    const email = accountEmail().toLowerCase();
    const isBootstrapAdmin = email === CONFIG.bootstrapAdminEmail.toLowerCase();
    const matchingEmployee = state.employees.find((employee) => (employee.Email || "").toLowerCase() === email);
    state.inactive = !isBootstrapAdmin && matchingEmployee?.Active === "No";
    const entry = state.inactive ? null : matchingEmployee;
    state.role = isBootstrapAdmin ? "Admin" : entry?.Role || (state.inactive ? "Inactive" : "Staff");
  }

  function setProfileUI() {
    const name = state.profile.displayName || state.account.name || state.account.username;
    const email = accountEmail();
    byId("profile-name").textContent = name;
    byId("profile-role").textContent = `${state.role} · ${email}`;
    byId("profile-initials").textContent = name.split(/\s+/).slice(0, 2).map((part) => part[0]).join("").toUpperCase();
    byId("role-badge").textContent = state.role;
    document.querySelectorAll(".manager-only:not(.page-view)").forEach((element) => element.classList.toggle("hidden", !isManager()));
    document.querySelectorAll(".admin-only:not(.page-view)").forEach((element) => element.classList.toggle("hidden", !isAdmin()));
    document.querySelectorAll(".page-view.manager-only").forEach((element) => element.classList.toggle("role-restricted", !isManager()));
    document.querySelectorAll(".page-view.admin-only").forEach((element) => element.classList.toggle("role-restricted", !isAdmin()));
  }

  function coordinatorEmailFor(project) {
    return String(project.CoordinatorEmail || project.AssignedCoordinatorEmail || "").trim().toLowerCase();
  }

  function managedProjects() {
    const email = accountEmail().toLowerCase();
    return isAdmin() ? state.projects : state.projects.filter((project) => coordinatorEmailFor(project) === email);
  }

  function canManageProject(project) {
    return Boolean(isAdmin() || (isCoordinator() && project && coordinatorEmailFor(project) === accountEmail().toLowerCase()));
  }

  function scopedManagerRecords(records) {
    return isAdmin() ? records : records.filter((record) => canManageProject(projectByCode(record.ProjectCode)));
  }
  function isTaskInRecycleBin(task) {
    return Boolean(task?.DeletedAt || task?.InRecycleBin === "Yes");
  }

  function canRecycleTask(task) {
    return Boolean(task?.id && canManageProject(projectByCode(task.ProjectCode)));
  }

  function taskRecordState() {
    const control = byId("task-filter-record-state");
    return isManager() && control?.value === "recycle" ? "recycle" : "active";
  }

  function visibleTasks({ includeDeleted = false } = {}) {
    const email = accountEmail().toLowerCase();
    let tasks;
    if (isAdmin()) tasks = state.tasks;
    else if (isCoordinator()) {
      const codes = new Set(managedProjects().map((project) => project.ProjectCode));
      tasks = state.tasks.filter((task) => codes.has(task.ProjectCode) || (task.AssigneeEmail || "").toLowerCase() === email);
    } else {
      tasks = state.tasks.filter((task) => (task.AssigneeEmail || "").toLowerCase() === email);
    }
    return includeDeleted ? tasks : tasks.filter((task) => !isTaskInRecycleBin(task));
  }

  function canUpdateTask(task) {
    const isAssignee = (task.AssigneeEmail || "").toLowerCase() === accountEmail().toLowerCase();
    return isAdmin() || canManageProject(projectByCode(task.ProjectCode)) || isAssignee;
  }

  function projectByCode(code) { return state.projects.find((project) => project.ProjectCode === code); }

  function projectCodeKey(value) { return String(value || "").trim().toUpperCase(); }

  function projectCodeReferenceGroups(oldCode) {
    const key = projectCodeKey(oldCode);
    const matches = (record) => projectCodeKey(record.ProjectCode) === key;
    const groups = [
      ...state.tasks.filter(matches).map((record) => ({ key: "tasks", label: "task", record })),
      ...state.workLogs.filter(matches).map((record) => ({ key: "workLogs", label: "work log", record })),
      ...state.issues.filter(matches).map((record) => ({ key: "issues", label: "BIM issue", record })),
      ...state.registers.filter(matches).map((record) => ({ key: "registers", label: "model / sheet register", record }))
    ];
    const invalid = groups.find((reference) => !reference.record?.id || (reference.key === "workLogs" && !String(reference.record.EmployeeEmail || "").trim()));
    if (invalid) throw new Error(`A linked ${invalid.label} is missing its record identity. The project code was not changed, so no information was lost.`);
    return groups;
  }

  function projectCodeReferenceSummary(references) {
    const labels = { tasks: "task", workLogs: "work log", issues: "BIM issue", registers: "model / sheet register" };
    return Object.entries(labels).map(([key, label]) => {
      const count = references.filter((reference) => reference.key === key).length;
      return count ? `${count} ${label}${count === 1 ? "" : "s"}` : "";
    }).filter(Boolean).join(", ");
  }

  function projectCodeMigrationFields(reference, projectCode, audit) {
    const action = `Project code changed from ${audit.from} to ${audit.to}`;
    const history = Array.isArray(reference.record.UpdateHistory) ? reference.record.UpdateHistory : [];
    return {
      ...reference.record,
      ProjectCode: projectCode,
      UpdatedBy: audit.by,
      LastUpdateAction: action,
      UpdateHistory: [...history, { at: audit.at, by: audit.by, action }]
    };
  }

  async function writeProjectCodeReference(reference, fields) {
    if (reference.key === "workLogs") return saveRecordAt(workspaceLogsPath(reference.record.EmployeeEmail), reference.record.id, fields);
    return saveRecord(reference.key, reference.record.id, fields);
  }

  async function rollbackProjectCodeReferences(references, writer = writeProjectCodeReference) {
    let failures = 0;
    for (let index = 0; index < references.length; index += 6) {
      const batch = references.slice(index, index + 6);
      const results = await Promise.allSettled(batch.map((reference) => writer(reference, reference.record)));
      failures += results.filter((result) => result.status === "rejected").length;
    }
    return failures;
  }

  async function migrateProjectCodeReferences(references, projectCode, audit, writer = writeProjectCodeReference) {
    const applied = [];
    for (let index = 0; index < references.length; index += 6) {
      const batch = references.slice(index, index + 6);
      const results = await Promise.allSettled(batch.map(async (reference) => {
        await writer(reference, projectCodeMigrationFields(reference, projectCode, audit));
        return reference;
      }));
      results.forEach((result) => { if (result.status === "fulfilled") applied.push(result.value); });
      const failed = results.find((result) => result.status === "rejected");
      if (failed) {
        const rollbackFailures = await rollbackProjectCodeReferences([...applied].reverse(), writer);
        const reason = failed.reason?.message || "a linked record could not be updated";
        const rollbackNote = rollbackFailures ? " Some linked records could not be restored automatically; do not retry until an Admin checks SharePoint." : " Linked records already changed were restored.";
        throw new Error(`Project code was not changed because ${reason}.${rollbackNote}`);
      }
    }
    return applied;
  }

  function coordinatorEmployees() { return state.employees.filter((employee) => employee.Active !== "No" && employee.Role === "Team Lead"); }

  function modellerEmployees() { return state.employees.filter((employee) => employee.Active !== "No" && employee.Role !== "Admin" && /Modell?er|Technician/i.test(employee.Designation || "")); }

  function toUtcDate(value) {
    const text = String(value || "").slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
    const date = new Date(`${text}T00:00:00Z`);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function dateKey(date) { return date.toISOString().slice(0, 10); }

  function normalDayMinutes(date) {
    const day = date.getUTCDay();
    return day === 0 ? 0 : day === 6 ? 240 : 480;
  }

  function normalTimeBetween(startValue, endValue) {
    const start = toUtcDate(startValue);
    const end = toUtcDate(endValue);
    if (!start || !end || end < start) return { calendarDays: 0, normalMinutes: 0, normalWorkDays: 0 };
    let calendarDays = 0;
    let normalMinutes = 0;
    const cursor = new Date(start);
    while (cursor <= end) {
      calendarDays += 1;
      normalMinutes += normalDayMinutes(cursor);
      cursor.setUTCDate(cursor.getUTCDate() + 1);
    }
    return { calendarDays, normalMinutes, normalWorkDays: normalMinutes / 480 };
  }

  function projectStartInfo(project) {
    const explicitStart = String(project.StartDate || project.ProjectStartDate || "").slice(0, 10);
    const recordedStart = String(project.createdAt || "").slice(0, 10);
    return { value: explicitStart || recordedStart, usesRecordedStart: !explicitStart && Boolean(recordedStart) };
  }

  function formatMinutes(minutes) {
    const total = Math.max(0, Math.round(Number(minutes) || 0));
    const hours = Math.floor(total / 60);
    const remainder = total % 60;
    return remainder ? `${hours}h ${remainder}m` : `${hours}h`;
  }

  function formatManDays(minutes) {
    const days = Math.max(0, Number(minutes) || 0) / 480;
    const value = Number.isInteger(days) ? days : days.toFixed(1);
    return `${value} ${Number(days) === 1 ? "man-day" : "man-days"}`;
  }

  function formatNormalTime(summary) {
    return summary.normalMinutes ? `${summary.normalWorkDays % 1 ? summary.normalWorkDays.toFixed(1) : summary.normalWorkDays} workdays · ${formatMinutes(summary.normalMinutes)}` : "Set project start and target dates";
  }

  function workLogMinutes(entry) {
    const saved = Number(entry.DurationMinutes);
    return Number.isFinite(saved) && saved > 0 ? saved : durationMinutes(entry.StartTime, entry.EndTime);
  }

  function isVoidedWorkLog(entry) {
    return Boolean(entry?.VoidedAt || entry?.ExcludedFromReporting === "Yes");
  }

  function workLogSignature(entry) {
    const employee = String(entry?.EmployeeEmail || "").trim().toLowerCase();
    const workDate = String(entry?.WorkDate || "").slice(0, 10);
    const project = String(entry?.ProjectCode || "").trim().toUpperCase();
    const task = String(entry?.TaskId || "").trim();
    const start = String(entry?.StartTime || "");
    const end = String(entry?.EndTime || "");
    return employee && workDate && start && end ? [employee, workDate, project, task, start, end].join("|") : "";
  }

  function uniqueWorkLogs(entries) {
    const seen = new Set();
    return entries.filter((entry) => !isVoidedWorkLog(entry)).filter((entry) => {
      const signature = workLogSignature(entry);
      if (!signature || !seen.has(signature)) {
        if (signature) seen.add(signature);
        return true;
      }
      return false;
    });
  }

  function workTimeRange(start, end) {
    const duration = durationMinutes(start, end);
    if (!duration) return null;
    const [hours, minutes] = String(start).split(":").map(Number);
    return { start: hours * 60 + minutes, end: hours * 60 + minutes + duration };
  }

  function workLogInterval(entry) {
    return workTimeRange(entry?.StartTime, entry?.EndTime);
  }

  function workLogTotalMinutes(entries) {
    const groups = new Map();
    let fallbackMinutes = 0;
    uniqueWorkLogs(entries).forEach((entry) => {
      const employee = String(entry?.EmployeeEmail || "").trim().toLowerCase();
      const workDate = String(entry?.WorkDate || "").slice(0, 10);
      const interval = workLogInterval(entry);
      if (!employee || !workDate || !interval) {
        fallbackMinutes += workLogMinutes(entry);
        return;
      }
      const key = `${employee}|${workDate}`;
      const group = groups.get(key) || { workDate, intervals: [] };
      group.intervals.push(interval);
      groups.set(key, group);
    });
    return fallbackMinutes + [...groups.values()].reduce((total, group) => {
      let lastEnd = -1;
      const rawMinutes = group.intervals.sort((a, b) => a.start - b.start || a.end - b.end).reduce((minutes, interval) => {
        const segmentStart = Math.max(interval.start, lastEnd);
        lastEnd = Math.max(lastEnd, interval.end);
        return interval.end > segmentStart ? minutes + interval.end - segmentStart : minutes;
      }, 0);
      const normalMinutes = normalDayMinutes(toUtcDate(group.workDate));
      return total + (normalMinutes ? Math.min(rawMinutes, normalMinutes) : rawMinutes);
    }, 0);
  }

  function workTimesOverlap(start, end, entry) {
    const candidate = workTimeRange(start, end);
    const existing = workLogInterval(entry);
    return Boolean(candidate && existing && candidate.start < existing.end && existing.start < candidate.end);
  }

  function monitoringDataForProject(project, today = dubaiDate()) {
    const startInfo = projectStartInfo(project);
    const start = startInfo.value;
    const target = String(project.TargetDate || "").slice(0, 10);
    const todayDate = toUtcDate(today);
    const startDate = toUtcDate(start);
    const targetDate = toUtcDate(target);
    const totalSchedule = normalTimeBetween(start, target);
    let elapsedSchedule = { calendarDays: 0, normalMinutes: 0, normalWorkDays: 0 };
    let remainingSchedule = { calendarDays: 0, normalMinutes: 0, normalWorkDays: 0 };
    if (startDate && targetDate && todayDate) {
      if (todayDate >= startDate) elapsedSchedule = normalTimeBetween(start, dateKey(todayDate > targetDate ? targetDate : todayDate));
      if (todayDate <= targetDate) remainingSchedule = normalTimeBetween(dateKey(todayDate > startDate ? todayDate : startDate), target);
    }
    const tasks = state.tasks.filter((task) => !isTaskInRecycleBin(task) && task.ProjectCode === project.ProjectCode);
    const logs = state.workLogs.filter((entry) => entry.ProjectCode === project.ProjectCode);
    const actualMinutes = workLogTotalMinutes(logs);
    const completedTasks = tasks.filter((task) => task.Status === "Completed").length;
    const overdueTasks = tasks.filter((task) => task.EndDate && task.EndDate < today && task.Status !== "Completed");
    const blockedTasks = tasks.filter((task) => task.Status === "Blocked");
    const members = new Set(tasks.map((task) => String(task.AssigneeEmail || "").toLowerCase()).filter(Boolean));
    const scheduleProgress = totalSchedule.normalMinutes ? Math.min(100, Math.round((elapsedSchedule.normalMinutes / totalSchedule.normalMinutes) * 100)) : null;
    return { project, start, target, startInfo, totalSchedule, elapsedSchedule, remainingSchedule, tasks, logs, actualMinutes, completedTasks, overdueTasks, blockedTasks, members, scheduleProgress };
  }

  function mondayOfWeek(dateValue) {
    const date = toUtcDate(dateValue);
    if (!date) return "";
    const offset = (date.getUTCDay() + 6) % 7;
    date.setUTCDate(date.getUTCDate() - offset);
    return dateKey(date);
  }

  function monthKey(dateValue) {
    return String(dateValue || "").slice(0, 7);
  }

  function monthStartKey(dateValue) {
    const month = monthKey(dateValue);
    return /^\d{4}-\d{2}$/.test(month) ? `${month}-01` : "";
  }

  function monthEndKey(month) {
    if (!/^\d{4}-\d{2}$/.test(month)) return "";
    const [year, value] = month.split("-").map(Number);
    return dateKey(new Date(Date.UTC(year, value, 0)));
  }

  function monthLabel(month) {
    const date = toUtcDate(`${month}-01`);
    return date ? new Intl.DateTimeFormat("en-GB", { timeZone: "UTC", month: "short", year: "numeric" }).format(date) : month || "—";
  }

  function recentMonthKeys(today = dubaiDate(), count = 6) {
    const first = toUtcDate(monthStartKey(today));
    if (!first) return [];
    const months = [];
    const cursor = new Date(first);
    for (let index = 0; index < count; index += 1) {
      months.push(dateKey(cursor).slice(0, 7));
      cursor.setUTCMonth(cursor.getUTCMonth() - 1);
    }
    return months;
  }

  function monitoringTeamRows(projectCode, today = dubaiDate()) {
    const weekStart = mondayOfWeek(today);
    const monthStart = monthStartKey(today);
    const relevantTasks = state.tasks.filter((task) => !isTaskInRecycleBin(task) && (!projectCode || task.ProjectCode === projectCode));
    const relevantLogs = state.workLogs.filter((entry) => !projectCode || entry.ProjectCode === projectCode);
    return state.employees.filter((employee) => employee.Active !== "No" && employee.Role !== "Admin").map((employee) => {
      const email = String(employee.Email || "").toLowerCase();
      const tasks = relevantTasks.filter((task) => String(task.AssigneeEmail || "").toLowerCase() === email);
      const logs = relevantLogs.filter((entry) => String(entry.EmployeeEmail || "").toLowerCase() === email);
      const todayMinutes = workLogTotalMinutes(logs.filter((entry) => entry.WorkDate === today));
      const weekMinutes = workLogTotalMinutes(logs.filter((entry) => entry.WorkDate >= weekStart && entry.WorkDate <= today));
      const monthMinutes = workLogTotalMinutes(logs.filter((entry) => entry.WorkDate >= monthStart && entry.WorkDate <= today));
      const latest = [...logs].sort((a, b) => `${b.WorkDate || ""}${b.StartTime || ""}`.localeCompare(`${a.WorkDate || ""}${a.StartTime || ""}`))[0];
      const openTasks = tasks.filter((task) => task.Status !== "Completed");
      return { employee, tasks, openTasks, todayMinutes, weekMinutes, monthMinutes, latest, noDailyLog: normalDayMinutes(toUtcDate(today)) > 0 && openTasks.length > 0 && todayMinutes === 0 };
    }).filter((row) => !projectCode || row.tasks.length || row.todayMinutes || row.weekMinutes || row.monthMinutes).sort((a, b) => b.monthMinutes - a.monthMinutes || b.todayMinutes - a.todayMinutes || b.openTasks.length - a.openTasks.length);
  }

  function monthlyWorkmanshipRows(projectCode, today = dubaiDate(), count = 6) {
    const currentMonth = monthKey(today);
    const visibleLogs = uniqueWorkLogs(state.workLogs.filter((entry) => !projectCode || entry.ProjectCode === projectCode));
    return recentMonthKeys(today, count).map((month) => {
      const periodEnd = month === currentMonth ? today : monthEndKey(month);
      const normalSchedule = normalTimeBetween(monthStartKey(month), periodEnd);
      const logs = visibleLogs.filter((entry) => monthKey(entry.WorkDate) === month);
      const actualMinutes = workLogTotalMinutes(logs);
      const contributors = new Set(logs.map((entry) => String(entry.EmployeeEmail || "").toLowerCase()).filter(Boolean));
      return { month, label: monthLabel(month), monthToDate: month === currentMonth, normalSchedule, actualMinutes, contributors, entryCount: uniqueWorkLogs(logs).length };
    });
  }

  function renderMetrics() {
    const today = dubaiDate();
    if (isAdmin()) {
      const activeProjects = state.projects.filter((project) => project.Status === "Active").length;
      const activeTeam = state.employees.filter((employee) => employee.Active !== "No" && employee.Role !== "Admin").length;
      const todayMinutes = workLogTotalMinutes(state.workLogs.filter((entry) => entry.WorkDate === today));
      const overdueTasks = state.tasks.filter((task) => !isTaskInRecycleBin(task) && task.EndDate && task.EndDate < today && task.Status !== "Completed").length;
      byId("metrics").innerHTML = [[String(activeProjects), "Active BIM projects"], [String(activeTeam), "Active team members"], [formatMinutes(todayMinutes), "Team workmanship today"], [String(overdueTasks), "Overdue BIM tasks"]].map(([value, label]) => `<div class="metric"><div class="metric-value">${esc(value)}</div><div class="metric-label">${esc(label)}</div></div>`).join("");
      return;
    }
    const tasks = visibleTasks();
    const accessibleProjects = isCoordinator() ? managedProjects() : state.projects.filter((project) => tasks.some((task) => task.ProjectCode === project.ProjectCode));
    const activeProjects = accessibleProjects.filter((project) => project.Status === "Active").length;
    const openTasks = tasks.filter((task) => task.Status !== "Completed");
    const dueTasks = openTasks.filter((task) => task.EndDate && task.EndDate <= today).length;
    const modelOrSheetDeliveries = isManager() ? scopedManagerRecords(state.registers).filter((item) => !["Approved", "Superseded"].includes(item.Status || "")).length : openTasks.filter((task) => /model|drawing/i.test(task.Deliverable || "")).length;
    const coordinationIssues = isManager() ? scopedManagerRecords(state.issues).filter((issue) => issue.Status !== "Closed").length : openTasks.filter((task) => task.Status === "Blocked" || ["Clash Coordination", "RFI"].includes(task.Deliverable)).length;
    byId("metrics").innerHTML = [[String(activeProjects), "Active BIM projects"], [String(dueTasks), "Tasks due / overdue"], [String(modelOrSheetDeliveries), "Model / sheet deliveries"], [String(coordinationIssues), "Open coordination issues"]].map(([value, label]) => `<div class="metric"><div class="metric-value">${esc(value)}</div><div class="metric-label">${esc(label)}</div></div>`).join("");
  }
  function renderWorkForm() {
    const select = byId("work-project");
    const selectedProject = select.value;
    const availableProjects = (isAdmin() ? state.projects : isCoordinator() ? managedProjects() : state.projects.filter((project) => visibleTasks().some((task) => task.ProjectCode === project.ProjectCode))).filter((project) => !["Completed", "Archived"].includes(project.Status));
    select.innerHTML = `<option value="">Select a project</option>${availableProjects.map((project) => `<option value="${esc(project.ProjectCode)}">${esc(project.ProjectCode)} · ${esc(project.Title)}</option>`).join("")}`;
    if ([...select.options].some((option) => option.value === selectedProject)) select.value = selectedProject;
    renderWorkTaskOptions();
    if (!byId("work-date").value) byId("work-date").value = dubaiDate();
    syncWorkEntryForSelection();
  }

  function workLogTasks() {
    const projectCode = byId("work-project").value;
    const email = accountEmail().toLowerCase();
    return visibleTasks().filter((task) => task.ProjectCode === projectCode && task.Status !== "Completed" && (task.AssigneeEmail || "").toLowerCase() === email);
  }
  function renderWorkTaskOptions() {
    const select = byId("work-task");
    const selectedTaskId = select.value;
    const tasks = workLogTasks();
    select.innerHTML = `<option value="">${byId("work-project").value ? "Select an active BIM task" : "Select a project first"}</option>${tasks.map((task) => `<option value="${esc(task.id || "")}">${esc(task.Title)} · ${esc(task.Discipline || "BIM")} · ${esc(task.Deliverable || "Task")}</option>`).join("")}`;
    if ([...select.options].some((option) => option.value === selectedTaskId)) select.value = selectedTaskId;
  }

  function ownWorkLogForSelection() {
    const employeeEmail = accountEmail().toLowerCase();
    const taskId = byId("work-task").value;
    const workDate = byId("work-date").value;
    if (!taskId || !workDate) return null;
    return state.workLogs.filter((entry) => !isVoidedWorkLog(entry) && String(entry.EmployeeEmail || "").toLowerCase() === employeeEmail && entry.TaskId === taskId && entry.WorkDate === workDate).sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")))[0] || null;
  }

  function setWorkEntryMode(entry = null) {
    const heading = byId("work-entry-heading");
    const help = byId("work-entry-help");
    const button = byId("work-save-button");
    if (entry) {
      heading.textContent = "Update today’s BIM work";
      help.textContent = "This task already has a work entry for the selected date. Update the same record—do not create a duplicate.";
      button.textContent = "Update work entry";
      return;
    }
    heading.textContent = "Log today’s BIM work";
    help.textContent = "Save one entry per task per day. Select the same task and date to update it.";
    button.textContent = "Save work entry";
  }

  function syncWorkEntryForSelection() {
    const entry = ownWorkLogForSelection();
    editingWorkLogId = entry?.id || "";
    if (entry) {
      byId("work-start").value = entry.StartTime || "";
      byId("work-end").value = entry.EndTime || "";
      byId("work-note").value = entry.WorkNote || "";
    } else {
      byId("work-start").value = "";
      byId("work-end").value = "";
      byId("work-note").value = "";
    }
    setWorkEntryMode(entry);
    updateDuration();
  }

  function renderRecentWork() {
    const email = accountEmail().toLowerCase();
    const entries = uniqueWorkLogs(state.workLogs.filter((entry) => isAdmin() || canManageProject(projectByCode(entry.ProjectCode)) || (entry.EmployeeEmail || "").toLowerCase() === email)).sort((a, b) => `${b.WorkDate || ""}${b.StartTime || ""}`.localeCompare(`${a.WorkDate || ""}${a.StartTime || ""}`)).slice(0, 7);
    byId("recent-work").innerHTML = entries.length ? entries.map((entry) => {
      const ownEntry = String(entry.EmployeeEmail || "").toLowerCase() === email;
      const taskStillAssigned = visibleTasks().some((task) => task.id === entry.TaskId && !isTaskInRecycleBin(task));
      const canEdit = Boolean(entry.id && ownEntry && taskStillAssigned);
      const canCorrect = Boolean(entry.id && isAdmin());
      return `<div class="activity-row"><strong>${esc(entry.TaskTitle || "Work entry")}</strong><span>${esc(entry.ProjectCode || "—")} · ${esc(entry.WorkDate || "")} · ${esc(entry.StartTime || "")}–${esc(entry.EndTime || "")} · ${esc(entry.EmployeeName || "")}</span>${canEdit ? `<button class="button button-quiet activity-action" type="button" data-edit-work="${esc(entry.id)}">Edit work entry</button>` : ""}${canCorrect ? `<button class="button button-quiet activity-action" type="button" data-void-work="${esc(entry.id)}">Mark incorrect</button>` : ""}</div>`;
    }).join("") : `<p class="muted">No work entries have been logged yet.</p>`;
  }

  function projectMonthlyWorkmanshipRows(projects, projectCode, today = dubaiDate()) {
    const month = monthKey(today);
    return projects.filter((project) => !projectCode || project.ProjectCode === projectCode).map((project) => {
      const logs = state.workLogs.filter((entry) => entry.ProjectCode === project.ProjectCode && monthKey(entry.WorkDate) === month);
      const tasks = state.tasks.filter((task) => !isTaskInRecycleBin(task) && task.ProjectCode === project.ProjectCode);
      const actualMinutes = workLogTotalMinutes(logs);
      const contributors = new Set(logs.map((entry) => String(entry.EmployeeEmail || "").toLowerCase()).filter(Boolean));
      const assignedMembers = new Set(tasks.map((task) => String(task.AssigneeEmail || "").toLowerCase()).filter(Boolean));
      const openTasks = tasks.filter((task) => task.Status !== "Completed");
      return { project, actualMinutes, contributors, assignedMembers, entryCount: uniqueWorkLogs(logs).length, openTasks, taskCount: tasks.length };
    }).sort((a, b) => b.actualMinutes - a.actualMinutes || String(a.project.ProjectCode || "").localeCompare(String(b.project.ProjectCode || "")));
  }

  function renderAdminMonitor() {
    const staffDashboard = byId("staff-dashboard-content");
    const monitor = byId("admin-monitor");
    if (!staffDashboard || !monitor) return;
    staffDashboard.classList.toggle("hidden", isAdmin());
    monitor.classList.toggle("hidden", !isAdmin());
    if (!isAdmin()) return;
    const today = dubaiDate();
    const projectSelect = byId("admin-monitor-project");
    const currentProjects = state.projects.filter((project) => project.Status !== "Archived");
    const selectedProjectCode = projectSelect.value;
    projectSelect.innerHTML = `<option value="">All current projects</option>${currentProjects.map((project) => `<option value="${esc(project.ProjectCode)}">${esc(project.ProjectCode)} · ${esc(project.Title)}</option>`).join("")}`;
    if ([...projectSelect.options].some((option) => option.value === selectedProjectCode)) projectSelect.value = selectedProjectCode;
    const projectCode = projectSelect.value;
    const projectData = currentProjects.filter((project) => !projectCode || project.ProjectCode === projectCode).map((project) => monitoringDataForProject(project, today));
    const teamRows = monitoringTeamRows(projectCode, today);
    const monthlyRows = monthlyWorkmanshipRows(projectCode, today);
    const projectMonthlyRows = projectMonthlyWorkmanshipRows(currentProjects, projectCode, today);
    const visibleLogs = uniqueWorkLogs(state.workLogs.filter((entry) => !projectCode || entry.ProjectCode === projectCode));
    const todayMinutes = workLogTotalMinutes(visibleLogs.filter((entry) => entry.WorkDate === today));
    const weekStart = mondayOfWeek(today);
    const weekMinutes = workLogTotalMinutes(visibleLogs.filter((entry) => entry.WorkDate >= weekStart && entry.WorkDate <= today));
    const currentMonth = monthlyRows[0] || { actualMinutes: 0, normalSchedule: { normalMinutes: 0 }, entryCount: 0, contributors: new Set() };
    const overdueTasks = projectData.flatMap((item) => item.overdueTasks);
    const blockedTasks = projectData.flatMap((item) => item.blockedTasks);
    const missingCoordinator = projectData.filter((item) => !coordinatorEmailFor(item.project));
    const noDailyLogRows = teamRows.filter((row) => row.noDailyLog);
    byId("admin-monitor-summary").innerHTML = [[formatMinutes(todayMinutes), "Workmanship logged today"], [formatMinutes(weekMinutes), "Workmanship logged this week"], [formatMinutes(currentMonth.actualMinutes), "Workmanship logged this month"], [formatMinutes(normalDayMinutes(toUtcDate(today))), "Normal member workday"], [String(overdueTasks.length), "Overdue BIM tasks"]].map(([value, label]) => `<div class="monitor-stat"><strong>${esc(value)}</strong><span>${esc(label)}</span></div>`).join("");
    const alerts = [
      overdueTasks.length ? { kind: "risk", title: `${overdueTasks.length} overdue task${overdueTasks.length === 1 ? "" : "s"}`, text: "Check task dates, assignees, and recovery action." } : null,
      blockedTasks.length ? { kind: "risk", title: `${blockedTasks.length} blocked task${blockedTasks.length === 1 ? "" : "s"}`, text: "Review the blocker, issue owner, and next coordination action." } : null,
      missingCoordinator.length ? { kind: "risk", title: `${missingCoordinator.length} project${missingCoordinator.length === 1 ? "" : "s"} without a Coordinator`, text: "Assign a Coordinator before task delegation." } : null,
      noDailyLogRows.length ? { kind: "notice", title: `${noDailyLogRows.length} team member${noDailyLogRows.length === 1 ? "" : "s"} with no log today`, text: "They have open tasks but no daily work entry yet." } : null
    ].filter(Boolean);
    byId("admin-monitor-alerts").innerHTML = alerts.length ? alerts.map((alert) => `<article class="monitor-alert ${alert.kind}"><strong>${esc(alert.title)}</strong><span>${esc(alert.text)}</span></article>`).join("") : `<article class="monitor-alert clear"><strong>No critical monitoring alerts</strong><span>Current project, task, and daily-log records do not show an immediate issue.</span></article>`;
    byId("admin-project-monitor").innerHTML = projectData.length ? `<p class="table-scroll-hint">Swipe left or right to see all columns</p><table class="data-table monitor-table"><thead><tr><th>Project / coordinator</th><th>Programme duration</th><th>Normal working time</th><th>Actual workmanship</th><th>Tasks / schedule</th></tr></thead><tbody>${projectData.map((item) => {
      const programme = item.start && item.target ? `${item.start} → ${item.target}${item.startInfo.usesRecordedStart ? " · recorded start" : ""}` : "Set project start and target dates";
      const taskText = item.tasks.length ? `${item.completedTasks}/${item.tasks.length} completed · ${item.members.size} assigned` : "No tasks assigned";
      const progress = item.scheduleProgress === null ? "Programme dates pending" : `${item.scheduleProgress}% programme elapsed`;
      const width = item.scheduleProgress === null ? 0 : item.scheduleProgress;
      return `<tr><td><strong>${esc(item.project.ProjectCode || "—")}</strong><br><span class="muted">${esc(item.project.Title || "Untitled project")}</span><br><span class="muted">${esc(item.project.CoordinatorName || item.project.CoordinatorEmail || "Coordinator required")}</span></td><td>${esc(programme)}<br><span class="muted">${esc(item.totalSchedule.calendarDays ? `${item.totalSchedule.calendarDays} calendar days` : "Programme dates pending")}</span></td><td>${esc(formatNormalTime(item.totalSchedule))}<br><span class="muted">${esc(item.remainingSchedule.normalMinutes ? `${formatMinutes(item.remainingSchedule.normalMinutes)} remaining` : item.target && item.target < today ? "Target date passed" : "—")}</span></td><td><strong>${esc(formatMinutes(item.actualMinutes))}</strong><br><span class="muted">${esc(formatManDays(item.actualMinutes))}</span></td><td>${esc(taskText)}<div class="monitor-progress" aria-label="${esc(progress)}"><span style="width:${width}%"></span></div><span class="muted">${esc(progress)}</span></td></tr>`;
    }).join("")}</tbody></table>` : `<p class="muted">No current projects match this monitoring filter.</p>`;
    byId("admin-team-monitor").innerHTML = teamRows.length ? `<p class="table-scroll-hint">Swipe left or right to see all columns</p><table class="data-table monitor-table"><thead><tr><th>Team member</th><th>Assigned work</th><th>Today</th><th>This week</th><th>This month</th><th>Latest activity</th></tr></thead><tbody>${teamRows.map((row) => {
      const name = row.employee.DisplayName || row.employee.Title || row.employee.Email || "Team member";
      const latest = row.latest ? `${row.latest.WorkDate || ""} · ${row.latest.TaskTitle || "Work entry"}` : "No work entry yet";
      const taskText = row.openTasks.length ? `${row.openTasks.length} open task${row.openTasks.length === 1 ? "" : "s"} · ${[...new Set(row.tasks.map((task) => task.ProjectCode).filter(Boolean))].join(", ")}` : "No open task";
      return `<tr><td><strong>${esc(name)}</strong><br><span class="muted">${esc(row.employee.Designation || row.employee.Role || "BIM team member")}</span></td><td>${esc(taskText)}</td><td><strong>${esc(formatMinutes(row.todayMinutes))}</strong><br><span class="muted">${row.noDailyLog ? "No log yet" : "Logged"}</span></td><td>${esc(formatMinutes(row.weekMinutes))}</td><td><strong>${esc(formatMinutes(row.monthMinutes))}</strong><br><span class="muted">${esc(formatManDays(row.monthMinutes))}</span></td><td>${esc(latest)}</td></tr>`;
    }).join("")}</tbody></table>` : `<p class="muted">No active team activity matches this monitoring filter.</p>`;
    byId("admin-monthly-workmanship").innerHTML = monthlyRows.length ? `<p class="table-scroll-hint">Swipe left or right to see all columns</p><table class="data-table monitor-table monthly-workmanship-table"><thead><tr><th>Month</th><th>Normal work time / member</th><th>Logged workmanship</th><th>Man-days</th><th>Contributors</th><th>Work entries</th></tr></thead><tbody>${monthlyRows.map((row) => `<tr><td><strong>${esc(row.label)}</strong><br><span class="muted">${row.monthToDate ? "Month to date" : "Full month"}</span></td><td>${esc(formatNormalTime(row.normalSchedule))}</td><td><strong>${esc(formatMinutes(row.actualMinutes))}</strong></td><td>${esc(formatManDays(row.actualMinutes))}</td><td>${esc(String(row.contributors.size))}</td><td>${esc(String(row.entryCount))}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">No monthly workmanship data is available for this monitoring filter.</p>`;

    byId("admin-project-monthly-workmanship").innerHTML = projectMonthlyRows.length ? `<p class="table-scroll-hint">Swipe left or right to see all columns</p><table class="data-table monitor-table project-monthly-workmanship-table"><thead><tr><th>Project / coordinator</th><th>Logged this month</th><th>Man-days</th><th>Contributors</th><th>Work entries</th><th>Active BIM tasks</th></tr></thead><tbody>${projectMonthlyRows.map((row) => {
      const coordinator = row.project.CoordinatorName || row.project.CoordinatorEmail || "Coordinator required";
      const teamText = `${row.contributors.size} logged · ${row.assignedMembers.size} assigned`;
      const taskText = row.taskCount ? `${row.openTasks.length} open / ${row.taskCount} total` : "No tasks assigned";
      return `<tr><td><strong>${esc(row.project.ProjectCode || "—")}</strong><br><span class="muted">${esc(row.project.Title || "Untitled project")}</span><br><span class="muted">${esc(coordinator)}</span></td><td><strong>${esc(formatMinutes(row.actualMinutes))}</strong><br><span class="muted">Month to date</span></td><td>${esc(formatManDays(row.actualMinutes))}</td><td>${esc(teamText)}</td><td>${esc(String(row.entryCount))}</td><td>${esc(taskText)}</td></tr>`;
    }).join("")}</tbody></table>` : `<p class="muted">No current projects match this monitoring filter.</p>`;
    const recent = [...visibleLogs].sort((a, b) => `${b.WorkDate || ""}${b.StartTime || ""}`.localeCompare(`${a.WorkDate || ""}${a.StartTime || ""}`)).slice(0, 10);
    byId("admin-activity-timeline").innerHTML = recent.length ? recent.map((entry) => `<div class="activity-row"><strong>${esc(entry.EmployeeName || entry.EmployeeEmail || "Team member")} · ${esc(entry.TaskTitle || "Work entry")}</strong><span>${esc(entry.ProjectCode || "—")} · ${esc(entry.WorkDate || "")} · ${esc(formatMinutes(workLogMinutes(entry)))} · ${esc(entry.WorkNote || "No note")}</span><button class="button button-quiet activity-action" type="button" data-void-work="${esc(entry.id || "")}">Mark incorrect</button></div>`).join("") : `<p class="muted">No work entries have been logged for this monitoring filter.</p>`;
  }

  function renderTaskFilters() {
    const setOptions = (id, options, placeholder) => {
      const select = byId(id);
      const previous = select.value;
      select.innerHTML = `<option value="">${esc(placeholder)}</option>${options.map(([value, label]) => `<option value="${esc(value)}">${esc(label)}</option>`).join("")}`;
      if ([...select.options].some((option) => option.value === previous)) select.value = previous;
    };
    const showingRecycleBin = taskRecordState() === "recycle";
    const tasks = visibleTasks({ includeDeleted: showingRecycleBin }).filter((task) => showingRecycleBin ? isTaskInRecycleBin(task) : !isTaskInRecycleBin(task));
    const projectCodes = new Set(tasks.map((task) => task.ProjectCode));
    managedProjects().forEach((project) => projectCodes.add(project.ProjectCode));
    const visibleProjects = state.projects.filter((project) => projectCodes.has(project.ProjectCode));
    const assigneeEmails = new Set(tasks.map((task) => (task.AssigneeEmail || "").toLowerCase()).filter(Boolean));
    setOptions("task-filter-project", visibleProjects.map((project) => [project.ProjectCode, `${project.ProjectCode} · ${project.Title}`]), "All projects");
    setOptions("task-filter-discipline", [...new Set(tasks.map((task) => task.Discipline).filter(Boolean))].sort().map((discipline) => [discipline, discipline]), "All disciplines");
    setOptions("task-filter-status", TASK_STATUSES.map((status) => [status, status]), "All statuses");
    setOptions("task-filter-assignee", state.employees.filter((employee) => employee.Active !== "No" && assigneeEmails.has((employee.Email || "").toLowerCase())).map((employee) => [employee.Email, `${employee.DisplayName || employee.Email} · ${employee.Designation || "BIM team member"}`]), "All assignees");
  }

  function filteredTasks(tasks = visibleTasks({ includeDeleted: taskRecordState() === "recycle" })) {
    const showingRecycleBin = taskRecordState() === "recycle";
    const project = byId("task-filter-project").value;
    const discipline = byId("task-filter-discipline").value;
    const status = byId("task-filter-status").value;
    const assignee = byId("task-filter-assignee").value.toLowerCase();
    const due = byId("task-filter-due").value;
    const today = dubaiDate();
    const nextWeek = new Date(`${today}T00:00:00Z`); nextWeek.setUTCDate(nextWeek.getUTCDate() + 7);
    const nextWeekDate = nextWeek.toISOString().slice(0, 10);
    return tasks.filter((task) => {
      if (showingRecycleBin ? !isTaskInRecycleBin(task) : isTaskInRecycleBin(task)) return false;
      if (project && task.ProjectCode !== project) return false;
      if (discipline && task.Discipline !== discipline) return false;
      if (status && (task.Status || "Not started") !== status) return false;
      if (assignee && (task.AssigneeEmail || "").toLowerCase() !== assignee) return false;
      if (!showingRecycleBin && due === "overdue" && !(task.EndDate && task.EndDate < today && task.Status !== "Completed")) return false;
      if (!showingRecycleBin && due === "due" && !(task.EndDate && task.EndDate <= today && task.Status !== "Completed")) return false;
      if (!showingRecycleBin && due === "next-7" && !(task.EndDate && task.EndDate >= today && task.EndDate <= nextWeekDate && task.Status !== "Completed")) return false;
      return true;
    });
  }

    function allowedTaskStatuses(task) {
    const currentStatus = task.Status || "Not started";
    if (isManager()) return TASK_STATUSES;
    return STAFF_TASK_STATUSES.includes(currentStatus) ? STAFF_TASK_STATUSES : [];
  }

  function renderTasks() {
    const tasks = filteredTasks();
    byId("tasks-list").innerHTML = tasks.length ? tasks.map((task) => {
      const project = projectByCode(task.ProjectCode);
      const status = task.Status || "Not started";
      const choices = allowedTaskStatuses(task);
      const recycled = isTaskInRecycleBin(task);
      const updateControls = !recycled && canUpdateTask(task) && task.id && choices.length ? `<div class="task-actions"><label class="task-status-control">Update status<select class="task-status-select">${choices.map((option) => `<option${option === status ? " selected" : ""}>${esc(option)}</option>`).join("")}</select></label><button class="button button-primary" type="button" data-task-id="${esc(task.id)}">Save status</button></div>` : "";
      const editControls = !recycled && canRecycleTask(task) && task.id ? `<div class="task-recycle-actions"><button class="button button-quiet" type="button" data-edit-task="${esc(task.id)}">Edit / reassign task</button></div>` : "";
      const recycleControls = canRecycleTask(task) && task.id ? (recycled
        ? `<div class="task-recycle-actions"><p class="task-recycle-note">${esc(task.DeletedAt ? `Moved to Recycle Bin by ${task.DeletedBy || "Asterwix team"} · ${new Date(task.DeletedAt).toLocaleDateString("en-GB")}` : "This task is in the Recycle Bin.")}</p><button class="button button-primary" type="button" data-restore-task="${esc(task.id)}">Restore task</button></div>`
        : `<div class="task-recycle-actions"><button class="button button-danger" type="button" data-delete-task="${esc(task.id)}">Delete task</button></div>`) : "";
      const deliveryDetails = [task.Discipline, task.Deliverable].filter(Boolean).join(" · ");
      const referenceDetails = [task.BIMStage, task.ModelDrawingNo, task.Revision ? `Rev ${task.Revision}` : ""].filter(Boolean).join(" · ");
      const qualityDetails = task.StatusUpdatedAt ? `Last updated by ${task.StatusUpdatedBy || "Asterwix team"} · ${new Date(task.StatusUpdatedAt).toLocaleDateString("en-GB")}` : "";
      return `<article class="task-card${recycled ? " task-card-recycled" : ""}"><p class="eyebrow">${esc(task.ProjectCode || "NO PROJECT")}</p><h2>${esc(task.Title)}</h2><p>${esc(deliveryDetails || "BIM delivery details not set")}</p><p>${esc(referenceDetails || project?.Client || "Asterwix project")}</p><p>${esc(task.AssigneeEmail || "")}</p><div class="task-meta"><span class="badge">${esc(recycled ? "Recycle Bin" : status)}</span><span>Due: ${esc(task.EndDate || "—")}</span></div>${qualityDetails ? `<p class="task-audit">${esc(qualityDetails)}</p>` : ""}${updateControls}${editControls}${recycleControls}</article>`;
    }).join("") : `<section class="card"><p class="muted">${taskRecordState() === "recycle" ? "No tasks are in the Recycle Bin for the current filters." : "No BIM tasks match the current filters."}</p></section>`;
  }

    function renderManagers() {
    if (!isManager()) return;
    const availableProjects = managedProjects().filter((project) => !["Completed", "Archived"].includes(project.Status));
    const coordinators = coordinatorEmployees();
    const taskAssignees = isAdmin() ? state.employees.filter((employee) => employee.Active !== "No") : modellerEmployees();
    byId("project-coordinator").innerHTML = `<option value="">Select BIM Coordinator / Team Lead</option>${coordinators.map((employee) => `<option value="${esc(employee.Email)}">${esc(employee.DisplayName || employee.Email)} · ${esc(employee.Designation || "BIM Coordinator")}</option>`).join("")}`;
    byId("task-project").innerHTML = `<option value="">Select project</option>${availableProjects.map((project) => `<option value="${esc(project.ProjectCode)}">${esc(project.ProjectCode)} · ${esc(project.Title)}</option>`).join("")}`;
    const activeEmployees = state.employees.filter((employee) => employee.Active !== "No").sort((a, b) => String(a.DisplayName || a.Email).localeCompare(String(b.DisplayName || b.Email)));
    byId("task-assignee").innerHTML = `<option value="">${isAdmin() ? "Select an active employee" : "Select a BIM modeller"}</option>${taskAssignees.map((employee) => `<option value="${esc(employee.Email)}">${esc(employee.DisplayName || employee.Email)} · ${esc(employee.Designation || "BIM team member")} · ${esc(employee.Discipline || "—")}</option>`).join("")}`;
    byId("task-form-heading").textContent = isAdmin() ? "Assign BIM task" : "Delegate task to BIM modeller";
    byId("task-assignment-note").textContent = isAdmin() ? "Assign the project coordinator's package or a direct BIM task." : "You can assign tasks only within projects where you are the assigned Coordinator.";
    byId("issue-project").innerHTML = `<option value="">Select project</option>${availableProjects.map((project) => `<option value="${esc(project.ProjectCode)}">${esc(project.ProjectCode)} · ${esc(project.Title)}</option>`).join("")}`;
    byId("issue-owner").innerHTML = `<option value="">Select responsible person</option>${activeEmployees.map((employee) => `<option value="${esc(employee.Email)}">${esc(employee.DisplayName || employee.Email)} · ${esc(employee.Designation || "BIM team member")}</option>`).join("")}`;
    byId("register-project").innerHTML = `<option value="">Select project</option>${availableProjects.map((project) => `<option value="${esc(project.ProjectCode)}">${esc(project.ProjectCode)} · ${esc(project.Title)}</option>`).join("")}`;
    const projectRows = isAdmin() ? state.projects : managedProjects();
    const projectActions = isAdmin() ? "<th>Action</th>" : "";
    const noProjectMessage = isAdmin()
      ? "No projects created yet. Archived projects and their records are retained."
      : "No projects are assigned to your account yet. Existing project records are retained; ask an Admin to assign or update the Coordinator."; 
    byId("projects-list").innerHTML = projectRows.length ? `<p class="table-scroll-hint">Swipe left or right to see all columns</p><table class="data-table"><thead><tr><th>Code</th><th>Project</th><th>Coordinator</th><th>Client</th><th>Status</th><th>Target</th>${projectActions}</tr></thead><tbody>${projectRows.map((project) => `<tr><td>${esc(project.ProjectCode)}</td><td>${esc(project.Title)}</td><td>${esc(project.CoordinatorName || project.CoordinatorEmail || project.AssignedCoordinatorEmail || "Unassigned — Admin action required")}</td><td>${esc(project.Client || "—")}</td><td>${esc(project.Status || "—")}</td><td>${esc(project.TargetDate || "—")}</td>${isAdmin() ? `<td><div class="table-actions"><button class="button button-quiet" type="button" data-edit-project="${esc(project.id || "")}">Edit</button>${project.Status !== "Archived" ? `<button class="button button-quiet" type="button" data-archive-project="${esc(project.id || "")}">Archive</button>` : ""}</div></td>` : ""}</tr>`).join("")}</tbody></table>` : `<p class="muted">${noProjectMessage}</p>`;
    const teamActionHeader = isAdmin() ? "<th>Action</th>" : "";
    const teamTableRows = state.employees.map((employee) => `<tr><td>${esc(employee.DisplayName || "—")}</td><td>${esc(employee.Designation || "—")}</td><td>${esc(employee.Discipline || "—")}</td><td>${esc(employee.Email || "—")}</td><td>${esc(employee.Role || "Staff")}</td><td>${esc(employee.Active || "Yes")}</td>${isAdmin() ? `<td><button class="button button-quiet" type="button" data-edit-employee="${esc(employee.Email || "")}">Edit</button></td>` : ""}</tr>`).join("");
    const teamMobileCards = state.employees.map((employee) => {
      const displayName = employee.DisplayName || employee.Email || "—";
      const initials = displayName.split(/\s+/).slice(0, 2).map((part) => part[0]).join("").toUpperCase();
      const active = employee.Active || "Yes";
      return `<article class="mobile-team-card"><div class="mobile-team-head"><div class="mobile-team-avatar">${esc(initials || "AX")}</div><div class="mobile-team-name"><strong>${esc(displayName)}</strong><span>${esc(employee.Designation || "BIM team member")}</span></div><span class="team-active-badge ${active === "No" ? "inactive" : ""}">${active === "No" ? "Inactive" : "Active"}</span></div><div class="mobile-team-details"><div><span>Discipline</span><strong>${esc(employee.Discipline || "—")}</strong></div><div><span>Portal role</span><strong>${esc(employee.Role || "Staff")}</strong></div><div class="full"><span>Email</span><strong>${esc(employee.Email || "—")}</strong></div></div>${isAdmin() ? `<div class="mobile-team-actions"><button class="button button-quiet" type="button" data-edit-employee="${esc(employee.Email || "")}">Edit team member</button></div>` : ""}</article>`;
    }).join("");
    byId("team-list").innerHTML = state.employees.length ? `<div class="mobile-team-list">${teamMobileCards}</div><div class="desktop-team-table"><p class="table-scroll-hint">Swipe left or right to see all columns</p><table class="data-table"><thead><tr><th>Name</th><th>Designation</th><th>Discipline</th><th>Email</th><th>Portal role</th><th>Active</th>${teamActionHeader}</tr></thead><tbody>${teamTableRows}</tbody></table></div>` : `<p class="muted">Add BIM team members after portal storage is ready.</p>`;
    renderIssueRegister();
    renderModelSheetRegister();
  }

  function renderIssueRegister() {
    if (!isManager()) return;
    const storageNote = state.bimStorage.issues ? "" : `<p class="muted register-note">Issue register storage is not ready yet. An Admin can prepare it from SharePoint setup.</p>`;
    const issues = [...scopedManagerRecords(state.issues)].sort((a, b) => `${a.Status === "Closed" ? 1 : 0}${a.DueDate || "9999"}`.localeCompare(`${b.Status === "Closed" ? 1 : 0}${b.DueDate || "9999"}`));
    byId("issues-list").innerHTML = storageNote || (issues.length ? `<p class="table-scroll-hint">Swipe left or right to see all columns</p><table class="data-table"><thead><tr><th>Issue</th><th>Project</th><th>Discipline</th><th>Reference</th><th>Responsible</th><th>Due</th><th>Status</th></tr></thead><tbody>${issues.map((issue) => `<tr><td><strong>${esc(issue.Title)}</strong><br><span class="muted">${esc(issue.IssueType || "Issue")} · ${esc(issue.Priority || "Medium")}</span></td><td>${esc(issue.ProjectCode || "—")}</td><td>${esc(issue.Discipline || "—")}</td><td>${esc(issue.Reference || "—")}</td><td>${esc(issue.OwnerName || issue.OwnerEmail || "—")}</td><td>${esc(issue.DueDate || "—")}</td><td>${esc(issue.Status || "Open")}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">No BIM issues have been logged yet.</p>`);
  }
  function renderModelSheetRegister() {
    if (!isManager()) return;
    const storageNote = state.bimStorage.registers ? "" : `<p class="muted register-note">Model and sheet register storage is not ready yet. An Admin can prepare it from SharePoint setup.</p>`;
    const records = [...scopedManagerRecords(state.registers)].sort((a, b) => `${a.ProjectCode || ""}${a.Number || ""}`.localeCompare(`${b.ProjectCode || ""}${b.Number || ""}`));
    byId("register-list").innerHTML = storageNote || (records.length ? `<p class="table-scroll-hint">Swipe left or right to see all columns</p><table class="data-table"><thead><tr><th>Type</th><th>Project</th><th>Discipline</th><th>Number / title</th><th>Revision</th><th>Stage</th><th>Status</th></tr></thead><tbody>${records.map((record) => `<tr><td>${esc(record.RecordType || "—")}</td><td>${esc(record.ProjectCode || "—")}</td><td>${esc(record.Discipline || "—")}</td><td><strong>${esc(record.Number || "—")}</strong><br><span class="muted">${esc(record.Title || "—")}</span></td><td>${esc(record.Revision || "—")}</td><td>${esc(record.BIMStage || "—")}</td><td>${esc(record.Status || "WIP")}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">No model or sheet delivery records have been added yet.</p>`);
  }
  function resetEmployeeForm() {
    editingEmployeeEmail = "";
    byId("employee-form").reset();
    byId("employee-email").readOnly = false;
    byId("employee-form-heading").textContent = "Add team member";
    byId("employee-cancel-button").classList.add("hidden");
    updateEmployeeSaveLabel();
  }

  function updateEmployeeSaveLabel() {
    byId("employee-save-button").textContent = editingEmployeeEmail ? "Save changes" : byId("employee-active").value === "No" ? "Save team member" : "Save and create workspace";
  }

  function beginEmployeeEdit(event) {
    const button = event.target.closest("[data-edit-employee]");
    if (!button || !isAdmin()) return;
    const email = button.dataset.editEmployee || "";
    const employee = state.employees.find((member) => (member.Email || "").toLowerCase() === email.toLowerCase());
    if (!employee) return toast("Team member could not be found.", "error");
    editingEmployeeEmail = employee.Email.toLowerCase();
    byId("employee-name").value = employee.DisplayName || employee.Title || "";
    byId("employee-email").value = employee.Email || "";
    byId("employee-email").readOnly = true;
    byId("employee-designation").value = employee.Designation || "BIM Modeler";
    byId("employee-discipline").value = employee.Discipline || "Landscape";
    byId("employee-role").value = employee.Role || "Staff";
    byId("employee-active").value = employee.Active || "Yes";
    byId("employee-form-heading").textContent = "Update team member";
    byId("employee-cancel-button").classList.remove("hidden");
    updateEmployeeSaveLabel();
    byId("employee-name").focus();
  }

  function renderSetup() {
    if (!isAdmin()) return;
    byId("setup-list").innerHTML = Object.entries(FOLDERS).map(([key, name]) => {
      const ready = CORE_FOLDERS.includes(key) ? !state.missingFolders.includes(key) : key === "workspaces" ? !state.missingFolders.includes(key) : state.bimStorage[key];
      return `<li><strong>${esc(name)}</strong> — ${ready ? "ready" : key === "issues" || key === "registers" ? "prepare BIM registers" : "not created"}</li>`;
    }).join("");
    updateResetButtonState();
  }

  function renderAll() {
    byId("today-label").textContent = new Intl.DateTimeFormat("en-GB", { timeZone: "Asia/Dubai", weekday: "long", year: "numeric", month: "long", day: "numeric" }).format(new Date());
    setProfileUI(); renderMetrics(); renderWorkForm(); renderRecentWork(); renderAdminMonitor(); renderTaskFilters(); renderTasks(); renderManagers(); renderSetup();
  }
  function showView(name) {
    document.querySelectorAll(".page-view").forEach((view) => view.classList.toggle("hidden", view.id !== `${name}-view`));
    document.querySelectorAll(".nav-item").forEach((button) => button.classList.toggle("active", button.dataset.view === name));
    document.querySelector(".sidebar")?.classList.remove("menu-open");
    byId("mobile-nav-toggle")?.setAttribute("aria-expanded", "false");
  }

  function durationMinutes(start, end) {
    if (!start || !end) return 0;
    const [sh, sm] = start.split(":").map(Number); const [eh, em] = end.split(":").map(Number);
    const result = (eh * 60 + em) - (sh * 60 + sm);
    return result > 0 ? result : 0;
  }

  function updateDuration() {
    const duration = durationMinutes(byId("work-start").value, byId("work-end").value);
    byId("work-duration").textContent = duration ? `Duration: ${Math.floor(duration / 60)}h ${duration % 60}m` : "Duration: —";
  }

  async function refreshData(message = "SharePoint synced") {
    setSync("Syncing SharePoint…");
    await loadData();
    setRole(); renderAll();
    setSync(state.missingFolders.length ? "Storage setup required" : message, state.missingFolders.length > 0);
  }

  async function openPortal() {
    try {
      setSync("Connecting to SharePoint…");
      state.profile = await graph("/me?$select=displayName,mail,userPrincipalName");
      await refreshData("SharePoint connected");
      if (state.inactive) {
        byId("sign-in-status").textContent = "Your Asterwix portal access is inactive. Please contact an administrator.";
        setSync("Portal access inactive", true);
        toast("Your portal access is inactive. Please contact an administrator.", "error");
        return;
      }
      byId("sign-in-view").classList.add("hidden"); byId("app-view").classList.remove("hidden");
      if (state.missingFolders.length) toast("Portal storage needs one-time SharePoint folder setup.", "error");
    } catch (error) {
      console.error(error);
      setSync("SharePoint connection needs attention", true);
      byId("app-view").classList.add("hidden");
      byId("sign-in-view").classList.remove("hidden");
      byId("sign-in-status").textContent = "We could not load Asterwix SharePoint data. Check your access and try again.";
      toast("SharePoint connection needs attention. Please try signing in again.", "error");
    }
  }

  async function refreshSetup() {
    if (!isAdmin()) return;
    try { await refreshData("Storage status refreshed"); toast("SharePoint storage status refreshed.", "success"); }
    catch (error) { toast(error.message || "SharePoint storage could not be checked.", "error"); }
  }

  async function submitWork(event) {
    event.preventDefault();
    const workDate = byId("work-date").value;
    const startTime = byId("work-start").value;
    const endTime = byId("work-end").value;
    const duration = durationMinutes(startTime, endTime);
    const task = state.tasks.find((item) => item.id === byId("work-task").value);
    if (!byId("work-project").value || !task || !duration) return toast("Select project, assigned BIM task, and valid start/end time.", "error");
    if (task.ProjectCode !== byId("work-project").value || task.Status === "Completed" || isTaskInRecycleBin(task)) return toast("Choose an active BIM task from the selected project.", "error");
    const employeeEmail = accountEmail().toLowerCase();
    if ((task.AssigneeEmail || "").toLowerCase() !== employeeEmail) return toast("You can log work only against a BIM task assigned to you.", "error");
    const existing = editingWorkLogId ? state.workLogs.find((entry) => entry.id === editingWorkLogId && !isVoidedWorkLog(entry) && String(entry.EmployeeEmail || "").toLowerCase() === employeeEmail && entry.TaskId === task.id && entry.WorkDate === workDate) : null;
    const candidate = { EmployeeEmail: employeeEmail, WorkDate: workDate, ProjectCode: task.ProjectCode, TaskId: task.id, StartTime: startTime, EndTime: endTime };
    const existingDayLogs = state.workLogs.filter((entry) => !isVoidedWorkLog(entry) && entry.id !== existing?.id && String(entry.EmployeeEmail || "").toLowerCase() === employeeEmail && entry.WorkDate === workDate);
    if (existingDayLogs.some((entry) => workLogSignature(entry) === workLogSignature(candidate))) return toast("This work entry already exists for the same task, date, and time. It was not saved again.", "error");
    const overlappingEntry = existingDayLogs.find((entry) => workTimesOverlap(startTime, endTime, entry));
    if (overlappingEntry) return toast(`This time overlaps your saved work entry (${overlappingEntry.StartTime}–${overlappingEntry.EndTime}). Use non-overlapping time.`, "error");
    try {
      await ensureWorkspace();
      const now = new Date().toISOString();
      const action = existing ? "Daily work entry updated" : "Daily work entry created";
      const history = [...(Array.isArray(existing?.UpdateHistory) ? existing.UpdateHistory : []), { at: now, by: state.profile.displayName || employeeEmail, action, startTime, endTime }];
      await saveRecordAt(workspaceLogsPath(), existing?.id || recordId("work"), { ...(existing || {}), Title: `${workDate} · ${task.Title}`, WorkDate: workDate, TaskId: task.id, TaskTitle: task.Title, ProjectCode: task.ProjectCode, Discipline: task.Discipline || "", Deliverable: task.Deliverable || "", BIMStage: task.BIMStage || "", ModelDrawingNo: task.ModelDrawingNo || "", Revision: task.Revision || "", EmployeeEmail: employeeEmail, EmployeeName: state.profile.displayName, StartTime: startTime, EndTime: endTime, DurationMinutes: duration, WorkNote: byId("work-note").value.trim(), createdAt: existing?.createdAt || now, UpdatedBy: state.profile.displayName || employeeEmail, LastUpdateAction: action, UpdateHistory: history });
      let taskStarted = false;
      if ((task.Status || "Not started") === "Not started") {
        const updatedBy = state.profile.displayName || accountEmail();
        const history = [...(Array.isArray(task.UpdateHistory) ? task.UpdateHistory : []), { at: now, by: updatedBy, action: "Started from first work entry", status: "In progress" }];
        await saveRecord("tasks", task.id, { ...task, Status: "In progress", StatusUpdatedAt: now, StatusUpdatedBy: updatedBy, UpdatedBy: updatedBy, LastUpdateAction: "Started from first work entry", UpdateHistory: history });
        taskStarted = true;
      }
      await refreshData(existing ? "Work entry updated" : "Work entry saved");
      toast(taskStarted ? "Work entry saved and task marked In progress." : existing ? "Today’s work entry updated." : "Daily work entry saved to SharePoint.", "success");
    } catch (error) { toast(error.message || "Could not save work entry.", "error"); }
  }

  function beginWorkEntryEdit(event) {
    const button = event.target.closest("[data-edit-work]");
    if (!button) return false;
    const email = accountEmail().toLowerCase();
    const entry = state.workLogs.find((item) => item.id === button.dataset.editWork && !isVoidedWorkLog(item) && String(item.EmployeeEmail || "").toLowerCase() === email);
    if (!entry) { toast("This work entry is not available for editing.", "error"); return true; }
    showView("dashboard");
    byId("work-project").value = entry.ProjectCode || "";
    renderWorkTaskOptions();
    byId("work-task").value = entry.TaskId || "";
    byId("work-date").value = entry.WorkDate || dubaiDate();
    syncWorkEntryForSelection();
    byId("work-log-form").scrollIntoView({ behavior: "smooth", block: "start" });
    return true;
  }
  function resetProjectForm() {
    editingProjectId = "";
    byId("project-form").reset();
    byId("project-code").readOnly = false;
    byId("project-form-heading").textContent = "Create BIM project";
    byId("project-save-button").textContent = "Create project";
    byId("project-cancel-button").classList.add("hidden");
  }

  function beginProjectEdit(event) {
    const button = event.target.closest("[data-edit-project]");
    if (!button || !isAdmin()) return;
    const project = state.projects.find((item) => item.id === button.dataset.editProject);
    if (!project) return toast("Project could not be found.", "error");
    editingProjectId = project.id;
    byId("project-code").value = project.ProjectCode || "";
    byId("project-code").readOnly = false;
    byId("project-name").value = project.Title || "";
    byId("project-client").value = project.Client || "";
    byId("project-coordinator").value = coordinatorEmailFor(project);
    byId("project-start-date").value = project.StartDate || project.ProjectStartDate || "";
    byId("project-target-date").value = project.TargetDate || "";
    byId("project-status").value = project.Status || "Active";
    byId("project-form-heading").textContent = "Update BIM project";
    byId("project-save-button").textContent = "Save project changes";
    byId("project-cancel-button").classList.remove("hidden");
    byId("project-code").focus();
  }

  async function archiveProject(event) {
    const button = event.target.closest("[data-archive-project]");
    if (!button || !isAdmin()) return;
    const project = state.projects.find((item) => item.id === button.dataset.archiveProject);
    if (!project) return toast("Project could not be found.", "error");
    if (project.Status === "Archived") return toast("This project is already archived. Its records are retained.", "success");
    const label = button.textContent;
    button.disabled = true; button.textContent = "Archiving…";
    try {
      const now = new Date().toISOString();
      const history = [...(Array.isArray(project.UpdateHistory) ? project.UpdateHistory : []), { at: now, by: state.profile.displayName || accountEmail(), action: "Archived" }];
      await saveRecord("projects", project.id, { ...project, Status: "Archived", ArchivedAt: now, ArchivedBy: state.profile.displayName || accountEmail(), LastUpdateAction: "Archived", UpdateHistory: history });
      await refreshData("Project archived");
      toast("Project archived. No project data, tasks, logs, issues, or registers were deleted.", "success");
    } catch (error) {
      toast(error.message || "Could not archive this project.", "error");
    } finally {
      button.disabled = false; button.textContent = label;
    }
  }
  async function deleteProject(event) {
    const button = event.target.closest("[data-delete-project]");
    if (!button || !isAdmin()) return;
    toast("Permanent project deletion is disabled. Archive the project instead; all project information is retained.", "error");
  }

  async function submitProject(event) {
    event.preventDefault();
    if (!isAdmin()) return toast("Only an Admin can create or edit projects.", "error");
    try {
      const projectCode = projectCodeKey(byId("project-code").value);
      const coordinatorEmail = byId("project-coordinator").value.trim().toLowerCase();
      const coordinator = coordinatorEmployees().find((employee) => (employee.Email || "").toLowerCase() === coordinatorEmail);
      const startDate = byId("project-start-date").value || "";
      const targetDate = byId("project-target-date").value || "";
      const existing = state.projects.find((project) => project.id === editingProjectId);
      const id = editingProjectId || recordId("project");
      if (!projectCode) return toast("Enter a project code.", "error");
      if (!coordinator) return toast("Select an active BIM Coordinator or Team Lead.", "error");
      if (startDate && targetDate && targetDate < startDate) return toast("Target date must be on or after the project start date.", "error");
      if (state.projects.some((project) => project.id !== id && projectCodeKey(project.ProjectCode) === projectCode)) return toast(`Project code ${projectCode} already exists.`, "error");
      if (editingProjectId && !existing) return toast("This project could not be found. Refresh and try again.", "error");

      const oldProjectCode = projectCodeKey(existing?.ProjectCode);
      const isCodeChange = Boolean(existing && oldProjectCode && oldProjectCode !== projectCode);
      const references = isCodeChange ? projectCodeReferenceGroups(oldProjectCode) : [];
      const referenceSummary = projectCodeReferenceSummary(references);
      if (isCodeChange && !window.confirm(`Change project code from ${oldProjectCode} to ${projectCode}?${referenceSummary ? ` This will update ${referenceSummary}.` : ""} No records will be deleted.`)) return;

      const now = new Date().toISOString();
      const updatedBy = state.profile.displayName || accountEmail();
      const action = isCodeChange ? `Project code changed from ${oldProjectCode} to ${projectCode}` : editingProjectId ? "Project updated" : "Project created";
      const history = [...(Array.isArray(existing?.UpdateHistory) ? existing.UpdateHistory : []), { at: now, by: updatedBy, action }];
      const codeHistory = isCodeChange ? [...(Array.isArray(existing?.ProjectCodeHistory) ? existing.ProjectCodeHistory : []), { at: now, by: updatedBy, from: oldProjectCode, to: projectCode }] : existing?.ProjectCodeHistory;
      const projectFields = {
        ...(existing || {}),
        Title: byId("project-name").value.trim(),
        ProjectCode: projectCode,
        Client: byId("project-client").value.trim(),
        CoordinatorEmail: coordinatorEmail,
        CoordinatorName: coordinator.DisplayName || coordinator.Title || coordinatorEmail,
        StartDate: startDate,
        Status: byId("project-status").value,
        TargetDate: targetDate,
        createdAt: existing?.createdAt || now,
        UpdatedBy: updatedBy,
        LastUpdateAction: action,
        UpdateHistory: history,
        ...(isCodeChange ? { ProjectCodeHistory: codeHistory } : {})
      };

      let migratedReferences = [];
      try {
        if (isCodeChange) migratedReferences = await migrateProjectCodeReferences(references, projectCode, { from: oldProjectCode, to: projectCode, at: now, by: updatedBy });
        await saveRecord("projects", id, projectFields);
      } catch (error) {
        if (migratedReferences.length) {
          const rollbackFailures = await rollbackProjectCodeReferences([...migratedReferences].reverse());
          if (rollbackFailures) error.message = `${error.message || "Project update failed."} Some linked records could not be restored automatically; do not retry until an Admin checks SharePoint.`;
        }
        throw error;
      }

      const message = isCodeChange ? "Project code and linked records updated" : action;
      resetProjectForm(); await refreshData(message); toast(`${message}. No project information was deleted.`, "success");
    } catch (error) { toast(error.message || "Could not create project.", "error"); }
  }

  function resetTaskForm() {
    editingTaskId = "";
    byId("task-form").reset();
    byId("task-form-heading").textContent = isAdmin() ? "Assign BIM task" : "Delegate task to BIM modeller";
    byId("task-assignment-note").textContent = isAdmin() ? "Assign the project coordinator's package or a direct BIM task." : "You can assign tasks only within projects where you are the assigned Coordinator.";
    byId("task-save-button").textContent = "Assign task";
    byId("task-cancel-button").classList.add("hidden");
  }

  function beginTaskEdit(event) {
    const button = event.target.closest("[data-edit-task]");
    if (!button) return false;
    const task = state.tasks.find((item) => item.id === button.dataset.editTask);
    if (!task || !canRecycleTask(task) || isTaskInRecycleBin(task)) { toast("Only an Admin or the assigned Coordinator can edit this active task.", "error"); return true; }
    showView("projects");
    renderManagers();
    editingTaskId = task.id;
    byId("task-project").value = task.ProjectCode || "";
    byId("task-assignee").value = (task.AssigneeEmail || "").toLowerCase();
    byId("task-discipline").value = task.Discipline || "";
    byId("task-deliverable").value = task.Deliverable || "";
    byId("task-title").value = task.Title || "";
    byId("task-lod").value = task.BIMStage || "";
    byId("task-reference").value = task.ModelDrawingNo || "";
    byId("task-revision").value = task.Revision || "";
    byId("task-start-date").value = task.StartDate || "";
    byId("task-end-date").value = task.EndDate || "";
    byId("task-priority").value = task.Priority || "Medium";
    byId("task-status").value = task.Status || "Not started";
    byId("task-notes").value = task.Notes || "";
    byId("task-form-heading").textContent = "Edit or reassign BIM task";
    byId("task-assignment-note").textContent = "Task history is retained. If a work entry was logged by mistake, mark that entry incorrect before entering corrected time.";
    byId("task-save-button").textContent = "Save task changes";
    byId("task-cancel-button").classList.remove("hidden");
    byId("task-form").scrollIntoView({ behavior: "smooth", block: "start" });
    return true;
  }

  async function submitTask(event) {
    event.preventDefault();
    if (!isManager()) return toast("Only an Admin or assigned Coordinator can assign tasks.", "error");
    try {
      const existing = editingTaskId ? state.tasks.find((item) => item.id === editingTaskId) : null;
      const id = editingTaskId || recordId("task");
      if (editingTaskId && (!existing || !canRecycleTask(existing))) return toast("This task can no longer be edited by your account. Refresh and try again.", "error");
      const now = new Date().toISOString();
      const assigneeEmail = byId("task-assignee").value.trim().toLowerCase();
      const startDate = byId("task-start-date").value || "";
      const endDate = byId("task-end-date").value || "";
      const assignee = state.employees.find((employee) => (employee.Email || "").toLowerCase() === assigneeEmail && employee.Active !== "No");
      if (!assignee) return toast("Select an active employee.", "error");
      const project = projectByCode(byId("task-project").value);
      if (!project || ["Completed", "Archived"].includes(project.Status)) return toast("Select an active project.", "error");
      if (isCoordinator() && coordinatorEmailFor(project) !== accountEmail().toLowerCase()) return toast("You can assign tasks only in projects assigned to you as Coordinator.", "error");
      if (isCoordinator() && !modellerEmployees().some((employee) => (employee.Email || "").toLowerCase() === assigneeEmail)) return toast("A Coordinator can assign tasks only to active BIM Modelers or BIM Technicians.", "error");
      if (existing && existing.ProjectCode !== project.ProjectCode && state.workLogs.some((entry) => !isVoidedWorkLog(entry) && entry.TaskId === existing.id)) return toast("This task already has work entries. Keep the project code unchanged, or mark incorrect work entries before moving the task.", "error");
      if (startDate && endDate && endDate < startDate) return toast("End date must be on or after the start date.", "error");
      const discipline = byId("task-discipline").value;
      const deliverable = byId("task-deliverable").value;
      const bimStage = byId("task-lod").value;
      if (!discipline || !deliverable || !bimStage) return toast("Select discipline, deliverable, and BIM stage / LOD.", "error");
      const taskStatus = byId("task-status").value;
      const assignedBy = state.profile.displayName || accountEmail();
      const wasReassigned = Boolean(existing && (existing.AssigneeEmail || "").toLowerCase() !== assigneeEmail);
      const action = existing ? (wasReassigned ? `Task reassigned from ${existing.AssigneeEmail || "unassigned"} to ${assigneeEmail}` : "Task details updated") : "Task assigned";
      const history = [...(Array.isArray(existing?.UpdateHistory) ? existing.UpdateHistory : []), { at: now, by: assignedBy, action, status: taskStatus }];
      await saveRecord("tasks", id, { ...(existing || {}), Title: byId("task-title").value.trim(), ProjectCode: project.ProjectCode, Discipline: discipline, Deliverable: deliverable, BIMStage: bimStage, ModelDrawingNo: byId("task-reference").value.trim(), Revision: byId("task-revision").value.trim(), AssigneeEmail: assigneeEmail, StartDate: startDate, EndDate: endDate, Priority: byId("task-priority").value, Status: taskStatus, Notes: byId("task-notes").value.trim(), createdAt: existing?.createdAt || now, UpdatedBy: assignedBy, LastUpdateAction: action, UpdateHistory: history });
      try { await inviteToItem(filePath("tasks", id), assigneeEmail, "write"); }
      catch (error) { await refreshData("Task assigned"); toast(`Task saved, but ${assigneeEmail} could not be granted status-update access: ${error.message}`, "error"); return; }
      if (wasReassigned && existing?.AssigneeEmail) await revokeDirectAccess(filePath("tasks", id), existing.AssigneeEmail);
      resetTaskForm(); await refreshData(existing ? "Task updated" : "Task assigned"); toast(existing ? (wasReassigned ? "Task reassigned. Previous work records were retained." : "Task details updated.") : "Task assigned.", "success");
    } catch (error) { toast(error.message || "Could not assign task.", "error"); }
  }

  async function updateTaskStatus(event) {
    const button = event.target.closest("[data-task-id]");
    if (!button) return;
    const task = state.tasks.find((item) => item.id === button.dataset.taskId);
    if (!task || !canUpdateTask(task)) return toast("You cannot update this task.", "error");
    if (isTaskInRecycleBin(task)) return toast("Restore this task before updating its status.", "error");
    const status = button.closest(".task-card")?.querySelector(".task-status-select")?.value;
    if (!TASK_STATUSES.includes(status) || !allowedTaskStatuses(task).includes(status)) return toast("This status can only be set by a BIM manager or team lead.", "error");
    const label = button.textContent; button.disabled = true; button.textContent = "Saving…";
    try {
      const now = new Date().toISOString();
      const updatedBy = state.profile.displayName || accountEmail();
      const history = [...(Array.isArray(task.UpdateHistory) ? task.UpdateHistory : []), { at: now, by: updatedBy, action: `Status changed to ${status}`, status }];
      await saveRecord("tasks", task.id, { ...task, Status: status, StatusUpdatedAt: now, StatusUpdatedBy: updatedBy, UpdatedBy: updatedBy, LastUpdateAction: `Status changed to ${status}`, UpdateHistory: history });
      await refreshData("Task status updated"); toast("Task status updated.", "success");
    } catch (error) { toast(error.message || "Could not update task status.", "error"); }
    finally { button.disabled = false; button.textContent = label; }
  }

  async function recycleTask(event) {
    const button = event.target.closest("[data-delete-task]");
    if (!button) return;
    const task = state.tasks.find((item) => item.id === button.dataset.deleteTask);
    if (!task || !canRecycleTask(task)) return toast("Only an Admin or the assigned Coordinator can delete this task.", "error");
    if (isTaskInRecycleBin(task)) return toast("This task is already in the Recycle Bin.", "success");
    if (!window.confirm(`Move "${task.Title}" to the Task Recycle Bin? It will disappear from active My Tasks, but the task, its history, and every work log will be retained.`)) return;
    const label = button.textContent; button.disabled = true; button.textContent = "Moving…";
    try {
      const now = new Date().toISOString();
      const updatedBy = state.profile.displayName || accountEmail();
      const history = [...(Array.isArray(task.UpdateHistory) ? task.UpdateHistory : []), { at: now, by: updatedBy, action: "Moved to Task Recycle Bin", status: task.Status || "Not started" }];
      await saveRecord("tasks", task.id, { ...task, DeletedAt: now, DeletedBy: updatedBy, InRecycleBin: "Yes", LastUpdateAction: "Moved to Task Recycle Bin", UpdatedBy: updatedBy, UpdateHistory: history });
      await refreshData("Task moved to Recycle Bin");
      toast("Task moved to Recycle Bin. No task, history, or work-log data was deleted.", "success");
    } catch (error) { toast(error.message || "Could not move this task to the Recycle Bin.", "error"); }
    finally { button.disabled = false; button.textContent = label; }
  }

  async function restoreTask(event) {
    const button = event.target.closest("[data-restore-task]");
    if (!button) return;
    const task = state.tasks.find((item) => item.id === button.dataset.restoreTask);
    if (!task || !canRecycleTask(task)) return toast("Only an Admin or the assigned Coordinator can restore this task.", "error");
    if (!isTaskInRecycleBin(task)) return toast("This task is already active.", "success");
    const label = button.textContent; button.disabled = true; button.textContent = "Restoring…";
    try {
      const now = new Date().toISOString();
      const updatedBy = state.profile.displayName || accountEmail();
      const history = [...(Array.isArray(task.UpdateHistory) ? task.UpdateHistory : []), { at: now, by: updatedBy, action: "Restored from Task Recycle Bin", status: task.Status || "Not started" }];
      await saveRecord("tasks", task.id, { ...task, DeletedAt: "", DeletedBy: "", InRecycleBin: "No", RestoredAt: now, RestoredBy: updatedBy, LastUpdateAction: "Restored from Task Recycle Bin", UpdatedBy: updatedBy, UpdateHistory: history });
      await refreshData("Task restored");
      toast("Task restored to active My Tasks.", "success");
    } catch (error) { toast(error.message || "Could not restore this task.", "error"); }
    finally { button.disabled = false; button.textContent = label; }
  }

  async function markIncorrectWorkEntry(event) {
    const button = event.target.closest("[data-void-work]");
    if (!button) return false;
    const entry = state.workLogs.find((item) => item.id === button.dataset.voidWork);
    if (!entry || !isAdmin()) { toast("Only an Admin can mark a saved work entry incorrect. Coordinators can edit or reassign the task.", "error"); return true; }
    if (!window.confirm(`Mark the ${entry.WorkDate} ${entry.StartTime}–${entry.EndTime} work entry as incorrect? It will be kept in SharePoint audit history but excluded from Recent work, overlap checks, and workmanship reporting.`)) return true;
    const label = button.textContent; button.disabled = true; button.textContent = "Correcting…";
    try {
      const now = new Date().toISOString();
      const correctedBy = state.profile.displayName || accountEmail();
      const history = [...(Array.isArray(entry.CorrectionHistory) ? entry.CorrectionHistory : []), { at: now, by: correctedBy, action: "Marked incorrect and excluded from reporting" }];
      await saveRecordAt(workspaceLogsPath(entry.EmployeeEmail), entry.id, { ...entry, VoidedAt: now, VoidedBy: correctedBy, ExcludedFromReporting: "Yes", CorrectionHistory: history, LastCorrectionAction: "Marked incorrect and excluded from reporting" });
      await refreshData("Incorrect work entry excluded");
      toast("Incorrect work entry kept for audit and excluded from reporting. You can now enter corrected non-overlapping time.", "success");
    } catch (error) { toast(error.message || "Could not correct this work entry.", "error"); }
    finally { button.disabled = false; button.textContent = label; }
    return true;
  }

  async function handleTaskCardAction(event) {
    if (beginTaskEdit(event)) return;
    if (event.target.closest("[data-delete-task]")) return recycleTask(event);
    if (event.target.closest("[data-restore-task]")) return restoreTask(event);
    return updateTaskStatus(event);
  }

    async function submitEmployee(event) {
    event.preventDefault();
    if (!isAdmin()) return;
    try {
      const email = byId("employee-email").value.trim().toLowerCase();
      const name = byId("employee-name").value.trim();
      const updating = Boolean(editingEmployeeEmail);
      if (updating && email !== editingEmployeeEmail) return toast("Email cannot be changed while updating a team member.", "error");
      if (email === CONFIG.bootstrapAdminEmail.toLowerCase() && byId("employee-active").value === "No") return toast("The portal's bootstrap Admin cannot be marked inactive here.", "error");
      const previous = state.employees.find((member) => (member.Email || "").toLowerCase() === email);
      const isBootstrapAdmin = email === CONFIG.bootstrapAdminEmail.toLowerCase();
      const role = isBootstrapAdmin ? "Admin" : byId("employee-role").value;
      const employee = { ...(previous || {}), Title: name, Email: email, DisplayName: name, Designation: byId("employee-designation").value, Discipline: byId("employee-discipline").value, Role: role, Active: byId("employee-active").value, createdAt: previous?.createdAt || new Date().toISOString() };
      const accessChanged = previous && (previous.Active !== employee.Active || previous.Role !== employee.Role);
      if (accessChanged) await revokeEmployeePortalAccess(previous);
      await saveRecord("employees", emailKey(email), employee);
      await provisionEmployeeWorkspace(employee);
      resetEmployeeForm(); await refreshData("Team member saved"); toast(employee.Active === "No" && previous ? "Team member marked inactive and direct portal sharing removed." : employee.Active === "No" ? "Team member saved as inactive." : updating ? "Team member updated." : "Team member and personal SharePoint workspace created.", "success");
    } catch (error) { toast(error.message || "Could not save team member.", "error"); }
  }

  function updateResetButtonState() {
    const input = byId("reset-confirmation");
    const button = byId("reset-portal-data");
    if (!input || !button) return;
    button.disabled = input.value.trim().toUpperCase() !== "RESET";
  }

  async function resetPortalWorkData() {
    if (!isAdmin()) return;
    const input = byId("reset-confirmation");
    const button = byId("reset-portal-data");
    if (!input || !button) return;
    if (input.value.trim().toUpperCase() !== "RESET") return toast("Type RESET to enable this action.", "error");
    const label = button.textContent;
    try {
      const groups = await portalResetRecordGroups();
      const total = groups.reduce((sum, group) => sum + group.items.length, 0);
      const summary = groups.map((group) => `${group.items.length} ${group.label}${group.items.length === 1 ? "" : "s"}`).join(", ");
      if (!total) return toast("Task assignments and project work-hour logs are already clear. All other portal data is unchanged.", "success");
      const confirmed = window.confirm(`Remove ${total} task assignment or work-hour record(s) from the active portal?\n\n${summary}\n\nProjects, staff accounts, BIM issues, model / sheet registers, Admin access, SharePoint folders, and portal code will remain. Deleted files may be retained in the SharePoint Recycle Bin according to tenant policy.`);
      if (!confirmed) return;
      button.disabled = true;
      button.textContent = "Resetting…";
      const results = [];
    for (const group of groups) results.push({ ...group, ...(await deleteDriveItems(group.items)) });
      const deleted = results.reduce((sum, result) => sum + result.deleted, 0);
      const failures = results.flatMap((result) => result.failures);
      await refreshData(failures.length ? "Task/work-hour reset partly completed" : "Tasks and work hours cleared");
      if (failures.length) {
        input.value = "";
        throw new Error(`${deleted} task or work-hour records were cleared, but ${failures.length} could not be removed. Do not retry until an Admin checks SharePoint access.`);
      }
      input.value = "";
      toast(`Tasks and project work-hour logs cleared. ${deleted} records were removed from active portal data; all other portal data remains.`, "success");
    } catch (error) {
      toast(error.message || "Could not clear task assignments and work-hour logs.", "error");
    } finally {
      button.textContent = label;
      updateResetButtonState();
    }
  }

  async function prepareBimRegisters() {
    if (!isAdmin()) return;
    const button = byId("prepare-bim-registers");
    const label = button.textContent; button.disabled = true; button.textContent = "Preparing…";
    try {
      await Promise.all([ensureFolder(CONFIG.storageFolder, FOLDERS.issues), ensureFolder(CONFIG.storageFolder, FOLDERS.registers)]);
      const teamLeads = state.employees.filter((employee) => employee.Active !== "No" && employee.Role === "Team Lead");
      await Promise.all(teamLeads.flatMap((employee) => [inviteToFolder(folderPath("issues"), employee.Email, "write"), inviteToFolder(folderPath("registers"), employee.Email, "write")]));
      await refreshData("BIM registers ready"); toast("BIM issue and model/sheet registers are ready for managers.", "success");
    } catch (error) { toast(error.message || "Could not prepare BIM register storage.", "error"); }
    finally { button.disabled = false; button.textContent = label; }
  }

  async function ensureBimRegister(key) {
    if (state.bimStorage[key]) return;
    if (!isAdmin()) throw new Error("BIM register storage is not ready. Ask an Admin to prepare it from SharePoint setup.");
    await ensureFolder(CONFIG.storageFolder, FOLDERS[key]);
    state.bimStorage[key] = true;
  }

  async function submitIssue(event) {
    event.preventDefault();
    if (!isManager()) return;
    try {
      const project = projectByCode(byId("issue-project").value);
      if (!project || !canManageProject(project) || ["Completed", "Archived"].includes(project.Status)) return toast("Select an active project that you manage.", "error");
      await ensureBimRegister("issues");
      const ownerEmail = byId("issue-owner").value;
      const owner = state.employees.find((employee) => (employee.Email || "").toLowerCase() === ownerEmail.toLowerCase());
      await saveRecord("issues", recordId("issue"), { Title: byId("issue-title").value.trim(), ProjectCode: project.ProjectCode, Discipline: byId("issue-discipline").value, IssueType: byId("issue-type").value, Reference: byId("issue-reference").value.trim(), OwnerEmail: ownerEmail, OwnerName: owner?.DisplayName || owner?.Title || ownerEmail, DueDate: byId("issue-due-date").value || "", Priority: byId("issue-priority").value, Status: byId("issue-status").value, Notes: byId("issue-notes").value.trim(), ReportedBy: state.profile.displayName || accountEmail(), createdAt: new Date().toISOString() });
      event.target.reset(); await refreshData("BIM issue logged"); toast("BIM issue logged.", "success");
    } catch (error) { toast(error.message || "Could not log BIM issue.", "error"); }
  }
  async function submitRegister(event) {
    event.preventDefault();
    if (!isManager()) return;
    try {
      const project = projectByCode(byId("register-project").value);
      if (!project || !canManageProject(project) || ["Completed", "Archived"].includes(project.Status)) return toast("Select an active project that you manage.", "error");
      await ensureBimRegister("registers");
      await saveRecord("registers", recordId("register"), { RecordType: byId("register-type").value, ProjectCode: project.ProjectCode, Discipline: byId("register-discipline").value, Number: byId("register-number").value.trim(), Title: byId("register-title").value.trim(), Revision: byId("register-revision").value.trim(), BIMStage: byId("register-lod").value, Status: byId("register-status").value, PlannedDate: byId("register-date").value || "", SharePointLink: byId("register-link").value.trim(), Notes: byId("register-notes").value.trim(), RegisteredBy: state.profile.displayName || accountEmail(), createdAt: new Date().toISOString() });
      event.target.reset(); await refreshData("Model or sheet registered"); toast("Model or sheet delivery added to the register.", "success");
    } catch (error) { toast(error.message || "Could not add model or sheet register entry.", "error"); }
  }
  async function signOut() {
    try { await msalInstance.clearCache({ account: state.account }); }
    catch (error) { console.error("Could not clear the local portal session.", error); }
    msalInstance.setActiveAccount(null);
    state.account = null; state.profile = null; state.role = "Staff";
    state.projects = []; state.tasks = []; state.workLogs = []; state.employees = []; state.issues = []; state.registers = []; state.missingFolders = []; state.bimStorage = { issues: false, registers: false }; state.inactive = false;
    byId("app-view").classList.add("hidden");
    byId("sign-in-view").classList.remove("hidden");
    byId("sign-in-status").textContent = "You have exited the Work Portal.";
  }

  async function boot() {
    byId("sign-in-button").addEventListener("click", signIn);
    byId("sign-out-button").addEventListener("click", signOut);
    byId("work-log-form").addEventListener("submit", submitWork);
    byId("project-form").addEventListener("submit", submitProject);
    byId("project-cancel-button").addEventListener("click", resetProjectForm);
    byId("task-form").addEventListener("submit", submitTask);
    byId("task-cancel-button").addEventListener("click", resetTaskForm);
    byId("issue-form").addEventListener("submit", submitIssue);
    byId("register-form").addEventListener("submit", submitRegister);
    byId("tasks-list").addEventListener("click", handleTaskCardAction);
    byId("recent-work").addEventListener("click", (event) => { if (beginWorkEntryEdit(event)) return; markIncorrectWorkEntry(event); });
    byId("admin-activity-timeline").addEventListener("click", markIncorrectWorkEntry);
    byId("projects-list").addEventListener("click", (event) => { beginProjectEdit(event); archiveProject(event); deleteProject(event); });
    byId("employee-form").addEventListener("submit", submitEmployee);
    byId("team-list").addEventListener("click", beginEmployeeEdit);
    byId("employee-cancel-button").addEventListener("click", resetEmployeeForm);
    byId("employee-active").addEventListener("change", updateEmployeeSaveLabel);
    byId("refresh-setup").addEventListener("click", refreshSetup);
    byId("prepare-bim-registers").addEventListener("click", prepareBimRegisters);
    byId("reset-confirmation").addEventListener("input", updateResetButtonState);
    byId("reset-portal-data").addEventListener("click", resetPortalWorkData);
    byId("work-project").addEventListener("change", () => { renderWorkTaskOptions(); syncWorkEntryForSelection(); });
    byId("work-task").addEventListener("change", syncWorkEntryForSelection);
    byId("work-date").addEventListener("change", syncWorkEntryForSelection);
    byId("admin-monitor-project")?.addEventListener("change", renderAdminMonitor);
    byId("task-filter-form").addEventListener("submit", (event) => event.preventDefault());
    ["task-filter-project", "task-filter-discipline", "task-filter-status", "task-filter-assignee", "task-filter-due", "task-filter-record-state"].forEach((id) => byId(id).addEventListener("change", () => { renderTaskFilters(); renderTasks(); }));
    byId("clear-task-filters").addEventListener("click", () => { byId("task-filter-form").reset(); renderTaskFilters(); renderTasks(); });
    [byId("work-start"), byId("work-end")].forEach((input) => input.addEventListener("input", updateDuration));
    document.querySelectorAll(".nav-item").forEach((button) => button.addEventListener("click", () => showView(button.dataset.view)));
    const mobileNavToggle = byId("mobile-nav-toggle");
    mobileNavToggle?.addEventListener("click", () => {
      const sidebar = document.querySelector(".sidebar");
      const open = sidebar?.classList.toggle("menu-open") || false;
      mobileNavToggle.setAttribute("aria-expanded", String(open));
      mobileNavToggle.textContent = open ? "Close menu" : "Menu";
    });
    try { await initialiseAuth(); if (state.account) await openPortal(); }
    catch (error) { console.error(error); byId("sign-in-status").textContent = "Microsoft login configuration needs attention."; }
  }

  boot();
})();
