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
      toast(error.message || "Microsoft sign-in could not complete.", "error");
    }
  }

  async function accessToken() {
    if (!state.account) throw new Error("Sign in is required.");
    try { return (await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: state.account })).accessToken; }
    catch { return (await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES, account: state.account })).accessToken; }
  }

  async function graph(path, options = {}) {
    const token = await accessToken();
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, { ...options, headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(options.headers || {}) } });
    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      const error = new Error(body?.error?.message || `SharePoint request failed (${response.status}).`);
      error.status = response.status;
      throw error;
    }
    if (response.status === 204) return null;
    return (response.headers.get("content-type") || "").includes("json") ? response.json() : response.text();
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
    const response = await graph(`/drives/${CONFIG.driveId}/items/${item.id}/permissions`);
    const matching = (response.value || []).filter((permission) => permissionEmails(permission).includes(String(email).toLowerCase()));
    for (const permission of matching) {
      try { await graph(`/drives/${CONFIG.driveId}/items/${item.id}/permissions/${permission.id}`, { method: "DELETE" }); }
      catch (error) {
        if (error.status !== 404 && !/inherited/i.test(error.message || "")) throw error;
      }
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
    await Promise.all(state.tasks.filter((task) => (task.AssigneeEmail || "").toLowerCase() === email && task.id).map((task) => inviteToItem(filePath("tasks", task.id), email, "write")));
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
    const response = await graph(`${drivePath(path, "/children")}?$top=999`);
    const records = await Promise.all((response.value || []).filter((item) => item.file && item.name.endsWith(".json")).map(readRecord));
    return records.filter(Boolean);
  }

  async function listRecords(key) { return listRecordsAt(folderPath(key)); }

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
    const response = await graph(`${drivePath(folderPath("workspaces"), "/children")}?$top=999`);
    const folders = (response.value || []).filter((item) => item.folder);
    const logs = await Promise.all(folders.map(async (folder) => {
      try { return await listRecordsAt(`${folderPath("workspaces")}/${folder.name}/work-logs`); }
      catch (error) { console.warn("Skipped unreadable employee workspace", folder.name, error); return []; }
    }));
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
    state.workLogs = isManager() ? await listAllWorkspaceLogs() : await listRecordsAt(workspaceLogsPath());
    if (isManager()) [state.issues, state.registers] = await Promise.all([listOptionalRecords("issues"), listOptionalRecords("registers")]);
  }

  function accountEmail() { return state.profile?.mail || state.profile?.userPrincipalName || state.account?.username || ""; }

  function setRole() {
    const email = accountEmail().toLowerCase();
    const matchingEmployee = state.employees.find((employee) => (employee.Email || "").toLowerCase() === email);
    state.inactive = matchingEmployee?.Active === "No";
    const entry = state.inactive ? null : matchingEmployee;
    state.role = entry?.Role || (state.inactive ? "Inactive" : email === CONFIG.bootstrapAdminEmail.toLowerCase() ? "Admin" : "Staff");
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

  function managedProjects() {
    const email = accountEmail().toLowerCase();
    return isAdmin() ? state.projects : state.projects.filter((project) => (project.CoordinatorEmail || "").toLowerCase() === email);
  }

  function visibleTasks() {
    const email = accountEmail().toLowerCase();
    if (isAdmin()) return state.tasks;
    if (isCoordinator()) {
      const codes = new Set(managedProjects().map((project) => project.ProjectCode));
      return state.tasks.filter((task) => codes.has(task.ProjectCode) || (task.AssigneeEmail || "").toLowerCase() === email);
    }
    return state.tasks.filter((task) => (task.AssigneeEmail || "").toLowerCase() === email);
  }

  function canUpdateTask(task) { return isManager() || (task.AssigneeEmail || "").toLowerCase() === accountEmail().toLowerCase(); }

  function projectByCode(code) { return state.projects.find((project) => project.ProjectCode === code); }
  function coordinatorEmployees() { return state.employees.filter((employee) => employee.Active !== "No" && (employee.Role === "Team Lead" || /BIM Coordinator|BIM Team Leader/i.test(employee.Designation || ""))); }
  function modellerEmployees() { return state.employees.filter((employee) => employee.Active !== "No" && employee.Role !== "Admin" && /Modeler|Technician/i.test(employee.Designation || "")); }

  function renderMetrics() {
    const today = dubaiDate();
    const tasks = visibleTasks();
    const activeProjects = state.projects.filter((project) => project.Status === "Active").length;
    const openTasks = tasks.filter((task) => task.Status !== "Completed");
    const dueTasks = openTasks.filter((task) => task.EndDate && task.EndDate <= today).length;
    const modelOrSheetDeliveries = isManager() ? state.registers.filter((item) => !["Approved", "Superseded"].includes(item.Status || "")).length : openTasks.filter((task) => /model|drawing/i.test(task.Deliverable || "")).length;
    const coordinationIssues = isManager() ? state.issues.filter((issue) => issue.Status !== "Closed").length : openTasks.filter((task) => task.Status === "Blocked" || ["Clash Coordination", "RFI"].includes(task.Deliverable)).length;
    byId("metrics").innerHTML = [[String(activeProjects), "Active BIM projects"], [String(dueTasks), "Tasks due / overdue"], [String(modelOrSheetDeliveries), "Model / sheet deliveries"], [String(coordinationIssues), "Open coordination issues"]].map(([value, label]) => `<div class="metric"><div class="metric-value">${esc(value)}</div><div class="metric-label">${esc(label)}</div></div>`).join("");
  }

  function renderWorkForm() {
    const select = byId("work-project");
    const selectedProject = select.value;
    select.innerHTML = `<option value="">Select a project</option>${state.projects.filter((project) => project.Status !== "Completed").map((project) => `<option value="${esc(project.ProjectCode)}">${esc(project.ProjectCode)} · ${esc(project.Title)}</option>`).join("")}`;
    if ([...select.options].some((option) => option.value === selectedProject)) select.value = selectedProject;
    renderWorkTaskOptions();
    if (!byId("work-date").value) byId("work-date").value = dubaiDate();
  }

  function workLogTasks() {
    const projectCode = byId("work-project").value;
    const email = accountEmail().toLowerCase();
    return visibleTasks().filter((task) => task.ProjectCode === projectCode && task.Status !== "Completed" && (isManager() || (task.AssigneeEmail || "").toLowerCase() === email));
  }

  function renderWorkTaskOptions() {
    const select = byId("work-task");
    const selectedTaskId = select.value;
    const tasks = workLogTasks();
    select.innerHTML = `<option value="">${byId("work-project").value ? "Select an active BIM task" : "Select a project first"}</option>${tasks.map((task) => `<option value="${esc(task.id || "")}">${esc(task.Title)} · ${esc(task.Discipline || "BIM")} · ${esc(task.Deliverable || "Task")}</option>`).join("")}`;
    if ([...select.options].some((option) => option.value === selectedTaskId)) select.value = selectedTaskId;
  }

  function renderRecentWork() {
    const email = accountEmail().toLowerCase();
    const entries = state.workLogs.filter((entry) => isManager() || (entry.EmployeeEmail || "").toLowerCase() === email).sort((a, b) => `${b.WorkDate || ""}${b.StartTime || ""}`.localeCompare(`${a.WorkDate || ""}${a.StartTime || ""}`)).slice(0, 7);
    byId("recent-work").innerHTML = entries.length ? entries.map((entry) => `<div class="activity-row"><strong>${esc(entry.TaskTitle || "Work entry")}</strong><span>${esc(entry.ProjectCode || "—")} · ${esc(entry.WorkDate || "")} · ${esc(entry.StartTime || "")}–${esc(entry.EndTime || "")} · ${esc(entry.EmployeeName || "")}</span></div>`).join("") : `<p class="muted">No work entries have been logged yet.</p>`;
  }

  function renderTaskFilters() {
    const setOptions = (id, options, placeholder) => {
      const select = byId(id);
      const previous = select.value;
      select.innerHTML = `<option value="">${esc(placeholder)}</option>${options.map(([value, label]) => `<option value="${esc(value)}">${esc(label)}</option>`).join("")}`;
      if ([...select.options].some((option) => option.value === previous)) select.value = previous;
    };
    setOptions("task-filter-project", state.projects.map((project) => [project.ProjectCode, `${project.ProjectCode} · ${project.Title}`]), "All projects");
    setOptions("task-filter-discipline", [...new Set(visibleTasks().map((task) => task.Discipline).filter(Boolean))].sort().map((discipline) => [discipline, discipline]), "All disciplines");
    setOptions("task-filter-status", TASK_STATUSES.map((status) => [status, status]), "All statuses");
    setOptions("task-filter-assignee", state.employees.filter((employee) => employee.Active !== "No").map((employee) => [employee.Email, `${employee.DisplayName || employee.Email} · ${employee.Designation || "BIM team member"}`]), "All assignees");
  }

  function filteredTasks(tasks = visibleTasks()) {
    const project = byId("task-filter-project").value;
    const discipline = byId("task-filter-discipline").value;
    const status = byId("task-filter-status").value;
    const assignee = byId("task-filter-assignee").value.toLowerCase();
    const due = byId("task-filter-due").value;
    const today = dubaiDate();
    const nextWeek = new Date(`${today}T00:00:00Z`); nextWeek.setUTCDate(nextWeek.getUTCDate() + 7);
    const nextWeekDate = nextWeek.toISOString().slice(0, 10);
    return tasks.filter((task) => {
      if (project && task.ProjectCode !== project) return false;
      if (discipline && task.Discipline !== discipline) return false;
      if (status && (task.Status || "Not started") !== status) return false;
      if (assignee && (task.AssigneeEmail || "").toLowerCase() !== assignee) return false;
      if (due === "overdue" && !(task.EndDate && task.EndDate < today && task.Status !== "Completed")) return false;
      if (due === "due" && !(task.EndDate && task.EndDate <= today && task.Status !== "Completed")) return false;
      if (due === "next-7" && !(task.EndDate && task.EndDate >= today && task.EndDate <= nextWeekDate && task.Status !== "Completed")) return false;
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
      const controls = canUpdateTask(task) && task.id && choices.length ? `<div class="task-actions"><label class="task-status-control">Update status<select class="task-status-select">${choices.map((option) => `<option${option === status ? " selected" : ""}>${esc(option)}</option>`).join("")}</select></label><button class="button button-primary" type="button" data-task-id="${esc(task.id)}">Save status</button></div>` : "";
      const deliveryDetails = [task.Discipline, task.Deliverable].filter(Boolean).join(" · ");
      const referenceDetails = [task.BIMStage, task.ModelDrawingNo, task.Revision ? `Rev ${task.Revision}` : ""].filter(Boolean).join(" · ");
      const qualityDetails = task.StatusUpdatedAt ? `Last updated by ${task.StatusUpdatedBy || "Asterwix team"} · ${new Date(task.StatusUpdatedAt).toLocaleDateString("en-GB")}` : "";
      return `<article class="task-card"><p class="eyebrow">${esc(task.ProjectCode || "NO PROJECT")}</p><h2>${esc(task.Title)}</h2><p>${esc(deliveryDetails || "BIM delivery details not set")}</p><p>${esc(referenceDetails || project?.Client || "Asterwix project")}</p><p>${esc(task.AssigneeEmail || "")}</p><div class="task-meta"><span class="badge">${esc(status)}</span><span>Due: ${esc(task.EndDate || "—")}</span></div>${qualityDetails ? `<p class="task-audit">${esc(qualityDetails)}</p>` : ""}${controls}</article>`;
    }).join("") : `<section class="card"><p class="muted">No BIM tasks match the current filters.</p></section>`;
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
    byId("issue-project").innerHTML = `<option value="">Select project</option>${state.projects.map((project) => `<option value="${esc(project.ProjectCode)}">${esc(project.ProjectCode)} · ${esc(project.Title)}</option>`).join("")}`;
    byId("issue-owner").innerHTML = `<option value="">Select responsible person</option>${activeEmployees.map((employee) => `<option value="${esc(employee.Email)}">${esc(employee.DisplayName || employee.Email)} · ${esc(employee.Designation || "BIM team member")}</option>`).join("")}`;
    byId("register-project").innerHTML = `<option value="">Select project</option>${state.projects.map((project) => `<option value="${esc(project.ProjectCode)}">${esc(project.ProjectCode)} · ${esc(project.Title)}</option>`).join("")}`;
    const projectRows = isAdmin() ? state.projects : managedProjects();
    const projectActions = isAdmin() ? "<th>Action</th>" : "";
    byId("projects-list").innerHTML = projectRows.length ? `<table class="data-table"><thead><tr><th>Code</th><th>Project</th><th>Coordinator</th><th>Client</th><th>Status</th><th>Target</th>${projectActions}</tr></thead><tbody>${projectRows.map((project) => `<tr><td>${esc(project.ProjectCode)}</td><td>${esc(project.Title)}</td><td>${esc(project.CoordinatorName || project.CoordinatorEmail || "—")}</td><td>${esc(project.Client || "—")}</td><td>${esc(project.Status || "—")}</td><td>${esc(project.TargetDate || "—")}</td>${isAdmin() ? `<td><div class="table-actions"><button class="button button-quiet" type="button" data-edit-project="${esc(project.id || "")}">Edit</button><button class="button button-quiet" type="button" data-archive-project="${esc(project.id || "")}">Archive</button><button class="button button-danger" type="button" data-delete-project="${esc(project.id || "")}">Delete</button></div></td>` : ""}</tr>`).join("")}</tbody></table>` : `<p class="muted">No projects created yet.</p>`;
    const teamActionHeader = isAdmin() ? "<th>Action</th>" : "";
    byId("team-list").innerHTML = state.employees.length ? `<table class="data-table"><thead><tr><th>Name</th><th>Designation</th><th>Discipline</th><th>Email</th><th>Portal role</th><th>Active</th>${teamActionHeader}</tr></thead><tbody>${state.employees.map((employee) => `<tr><td>${esc(employee.DisplayName || "—")}</td><td>${esc(employee.Designation || "—")}</td><td>${esc(employee.Discipline || "—")}</td><td>${esc(employee.Email || "—")}</td><td>${esc(employee.Role || "Staff")}</td><td>${esc(employee.Active || "Yes")}</td>${isAdmin() ? `<td><button class="button button-quiet" type="button" data-edit-employee="${esc(employee.Email || "")}">Edit</button></td>` : ""}</tr>`).join("")}</tbody></table>` : `<p class="muted">Add BIM team members after portal storage is ready.</p>`;
    renderIssueRegister();
    renderModelSheetRegister();
  }

  function renderIssueRegister() {
    if (!isManager()) return;
    const storageNote = state.bimStorage.issues ? "" : `<p class="muted register-note">Issue register storage is not ready yet. An Admin can prepare it from SharePoint setup.</p>`;
    const issues = [...state.issues].sort((a, b) => `${a.Status === "Closed" ? 1 : 0}${a.DueDate || "9999"}`.localeCompare(`${b.Status === "Closed" ? 1 : 0}${b.DueDate || "9999"}`));
    byId("issues-list").innerHTML = storageNote || (issues.length ? `<table class="data-table"><thead><tr><th>Issue</th><th>Project</th><th>Discipline</th><th>Reference</th><th>Responsible</th><th>Due</th><th>Status</th></tr></thead><tbody>${issues.map((issue) => `<tr><td><strong>${esc(issue.Title)}</strong><br><span class="muted">${esc(issue.IssueType || "Issue")} · ${esc(issue.Priority || "Medium")}</span></td><td>${esc(issue.ProjectCode || "—")}</td><td>${esc(issue.Discipline || "—")}</td><td>${esc(issue.Reference || "—")}</td><td>${esc(issue.OwnerName || issue.OwnerEmail || "—")}</td><td>${esc(issue.DueDate || "—")}</td><td>${esc(issue.Status || "Open")}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">No BIM issues have been logged yet.</p>`);
  }

  function renderModelSheetRegister() {
    if (!isManager()) return;
    const storageNote = state.bimStorage.registers ? "" : `<p class="muted register-note">Model and sheet register storage is not ready yet. An Admin can prepare it from SharePoint setup.</p>`;
    const records = [...state.registers].sort((a, b) => `${a.ProjectCode || ""}${a.Number || ""}`.localeCompare(`${b.ProjectCode || ""}${b.Number || ""}`));
    byId("register-list").innerHTML = storageNote || (records.length ? `<table class="data-table"><thead><tr><th>Type</th><th>Project</th><th>Discipline</th><th>Number / title</th><th>Revision</th><th>Stage</th><th>Status</th></tr></thead><tbody>${records.map((record) => `<tr><td>${esc(record.RecordType || "—")}</td><td>${esc(record.ProjectCode || "—")}</td><td>${esc(record.Discipline || "—")}</td><td><strong>${esc(record.Number || "—")}</strong><br><span class="muted">${esc(record.Title || "—")}</span></td><td>${esc(record.Revision || "—")}</td><td>${esc(record.BIMStage || "—")}</td><td>${esc(record.Status || "WIP")}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">No model or sheet delivery records have been added yet.</p>`);
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
  }

  function renderAll() {
    byId("today-label").textContent = new Intl.DateTimeFormat("en-GB", { timeZone: "Asia/Dubai", weekday: "long", year: "numeric", month: "long", day: "numeric" }).format(new Date());
    setProfileUI(); renderMetrics(); renderWorkForm(); renderRecentWork(); renderTaskFilters(); renderTasks(); renderManagers(); renderSetup();
  }

  function showView(name) {
    document.querySelectorAll(".page-view").forEach((view) => view.classList.toggle("hidden", view.id !== `${name}-view`));
    document.querySelectorAll(".nav-item").forEach((button) => button.classList.toggle("active", button.dataset.view === name));
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
      byId("sign-in-view").classList.add("hidden"); byId("app-view").classList.remove("hidden");
      state.profile = state.profile || { displayName: state.account?.name || "Asterwix User", mail: state.account?.username || "" };
      state.role = accountEmail().toLowerCase() === CONFIG.bootstrapAdminEmail.toLowerCase() ? "Admin" : "Staff";
      renderAll();
      toast(`SharePoint connection needs attention: ${error.message}`, "error");
    }
  }

  async function refreshSetup() {
    if (!isAdmin()) return;
    try { await refreshData("Storage status refreshed"); toast("SharePoint storage status refreshed.", "success"); }
    catch (error) { toast(error.message || "SharePoint storage could not be checked.", "error"); }
  }

  async function submitWork(event) {
    event.preventDefault();
    const duration = durationMinutes(byId("work-start").value, byId("work-end").value);
    const task = state.tasks.find((item) => item.id === byId("work-task").value);
    if (!byId("work-project").value || !task || !duration) return toast("Select project, assigned BIM task, and valid start/end time.", "error");
    if (task.ProjectCode !== byId("work-project").value || task.Status === "Completed") return toast("Choose an active BIM task from the selected project.", "error");
    try {
      await ensureWorkspace();
      await saveRecordAt(workspaceLogsPath(), recordId("work"), { Title: `${byId("work-date").value} · ${task.Title}`, WorkDate: byId("work-date").value, TaskId: task.id, TaskTitle: task.Title, ProjectCode: task.ProjectCode, Discipline: task.Discipline || "", Deliverable: task.Deliverable || "", BIMStage: task.BIMStage || "", ModelDrawingNo: task.ModelDrawingNo || "", Revision: task.Revision || "", EmployeeEmail: accountEmail(), EmployeeName: state.profile.displayName, StartTime: byId("work-start").value, EndTime: byId("work-end").value, DurationMinutes: duration, WorkNote: byId("work-note").value.trim(), createdAt: new Date().toISOString() });
      event.target.reset(); byId("work-date").value = dubaiDate(); updateDuration(); await refreshData("Work entry saved"); toast("Daily work entry saved to SharePoint.", "success");
    } catch (error) { toast(error.message || "Could not save work entry.", "error"); }
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
    byId("project-code").readOnly = true;
    byId("project-name").value = project.Title || "";
    byId("project-client").value = project.Client || "";
    byId("project-coordinator").value = project.CoordinatorEmail || "";
    byId("project-target-date").value = project.TargetDate || "";
    byId("project-status").value = project.Status || "Active";
    byId("project-form-heading").textContent = "Update BIM project";
    byId("project-save-button").textContent = "Save project changes";
    byId("project-cancel-button").classList.remove("hidden");
    byId("project-name").focus();
  }

  async function archiveProject(event) {
    const button = event.target.closest("[data-archive-project]");
    if (!button || !isAdmin()) return;
    const project = state.projects.find((item) => item.id === button.dataset.archiveProject);
    if (!project) return toast("Project could not be found.", "error");
    await saveRecord("projects", project.id, { ...project, Status: "Archived", ArchivedAt: new Date().toISOString(), ArchivedBy: state.profile.displayName || accountEmail() });
    await refreshData("Project archived"); toast("Project archived. Its records are retained.", "success");
  }

  async function deleteProject(event) {
    const button = event.target.closest("[data-delete-project]");
    if (!button || !isAdmin()) return;
    const project = state.projects.find((item) => item.id === button.dataset.deleteProject);
    if (!project) return toast("Project could not be found.", "error");
    const linked = state.tasks.some((task) => task.ProjectCode === project.ProjectCode) || state.workLogs.some((log) => log.ProjectCode === project.ProjectCode) || state.issues.some((issue) => issue.ProjectCode === project.ProjectCode) || state.registers.some((record) => record.ProjectCode === project.ProjectCode);
    if (linked) return toast("This project has linked records. Archive it instead of deleting it.", "error");
    if (!window.confirm(`Delete project ${project.ProjectCode}? This cannot be undone.`)) return;
    await graph(drivePath(filePath("projects", project.id)), { method: "DELETE" });
    if (editingProjectId === project.id) resetProjectForm();
    await refreshData("Project deleted"); toast("Project deleted.", "success");
  }

  async function submitProject(event) {
    event.preventDefault();
    if (!isAdmin()) return toast("Only an Admin can create or edit projects.", "error");
    try {
      const projectCode = byId("project-code").value.trim().toUpperCase();
      const coordinatorEmail = byId("project-coordinator").value.trim().toLowerCase();
      const coordinator = coordinatorEmployees().find((employee) => (employee.Email || "").toLowerCase() === coordinatorEmail);
      if (!projectCode) return toast("Enter a project code.", "error");
      if (!coordinator) return toast("Select an active BIM Coordinator or Team Lead.", "error");
      if (!editingProjectId && state.projects.some((project) => String(project.ProjectCode || "").trim().toUpperCase() === projectCode)) return toast(`Project code ${projectCode} already exists.`, "error");
      const existing = state.projects.find((project) => project.id === editingProjectId);
      const id = editingProjectId || recordId("project");
      await saveRecord("projects", id, { Title: byId("project-name").value.trim(), ProjectCode: projectCode, Client: byId("project-client").value.trim(), CoordinatorEmail: coordinatorEmail, CoordinatorName: coordinator.DisplayName || coordinator.Title || coordinatorEmail, Status: byId("project-status").value, TargetDate: byId("project-target-date").value || "", createdAt: existing?.createdAt || new Date().toISOString() });
      const message = editingProjectId ? "Project updated" : "Project created";
      resetProjectForm(); await refreshData(message); toast(`${message}.`, "success");
    } catch (error) { toast(error.message || "Could not create project.", "error"); }
  }

  async function submitTask(event) {
    event.preventDefault();
    if (!isManager()) return toast("Only an Admin or assigned Coordinator can assign tasks.", "error");
    try {
      const id = recordId("task");
      const assigneeEmail = byId("task-assignee").value.trim().toLowerCase();
      const startDate = byId("task-start-date").value || "";
      const endDate = byId("task-end-date").value || "";
      const assignee = state.employees.find((employee) => (employee.Email || "").toLowerCase() === assigneeEmail && employee.Active !== "No");
      if (!assignee) return toast("Select an active employee.", "error");
      const project = projectByCode(byId("task-project").value);
      if (!project || ["Completed", "Archived"].includes(project.Status)) return toast("Select an active project.", "error");
      if (isCoordinator() && (project.CoordinatorEmail || "").toLowerCase() !== accountEmail().toLowerCase()) return toast("You can assign tasks only in projects assigned to you as Coordinator.", "error");
      if (isCoordinator() && !modellerEmployees().some((employee) => (employee.Email || "").toLowerCase() === assigneeEmail)) return toast("A Coordinator can assign tasks only to active BIM Modelers or BIM Technicians.", "error");
      if (startDate && endDate && endDate < startDate) return toast("End date must be on or after the start date.", "error");
      const discipline = byId("task-discipline").value;
      const deliverable = byId("task-deliverable").value;
      const bimStage = byId("task-lod").value;
      if (!discipline || !deliverable || !bimStage) return toast("Select discipline, deliverable, and BIM stage / LOD.", "error");
      await saveRecord("tasks", id, { Title: byId("task-title").value.trim(), ProjectCode: byId("task-project").value, Discipline: discipline, Deliverable: deliverable, BIMStage: bimStage, ModelDrawingNo: byId("task-reference").value.trim(), Revision: byId("task-revision").value.trim(), AssigneeEmail: assigneeEmail, StartDate: startDate, EndDate: endDate, Priority: byId("task-priority").value, Status: byId("task-status").value, Notes: byId("task-notes").value.trim(), createdAt: new Date().toISOString() });
      try { await inviteToItem(filePath("tasks", id), assigneeEmail, "write"); }
      catch (error) { await refreshData("Task assigned"); toast(`Task saved, but ${assigneeEmail} could not be granted status-update access: ${error.message}`, "error"); return; }
      event.target.reset(); await refreshData("Task assigned"); toast("Task assigned.", "success");
    } catch (error) { toast(error.message || "Could not assign task.", "error"); }
  }

  async function updateTaskStatus(event) {
    const button = event.target.closest("[data-task-id]");
    if (!button) return;
    const task = state.tasks.find((item) => item.id === button.dataset.taskId);
    if (!task || !canUpdateTask(task)) return toast("You cannot update this task.", "error");
    const status = button.closest(".task-card")?.querySelector(".task-status-select")?.value;
    if (!TASK_STATUSES.includes(status) || !allowedTaskStatuses(task).includes(status)) return toast("This status can only be set by a BIM manager or team lead.", "error");
    const label = button.textContent; button.disabled = true; button.textContent = "Saving…";
    try {
      await saveRecord("tasks", task.id, { ...task, Status: status, StatusUpdatedAt: new Date().toISOString(), StatusUpdatedBy: state.profile.displayName || accountEmail() });
      await refreshData("Task status updated"); toast("Task status updated.", "success");
    } catch (error) { toast(error.message || "Could not update task status.", "error"); }
    finally { button.disabled = false; button.textContent = label; }
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
      const employee = { Title: name, Email: email, DisplayName: name, Designation: byId("employee-designation").value, Discipline: byId("employee-discipline").value, Role: byId("employee-role").value, Active: byId("employee-active").value, createdAt: state.employees.find((member) => (member.Email || "").toLowerCase() === email)?.createdAt || new Date().toISOString() };
      const previous = state.employees.find((member) => (member.Email || "").toLowerCase() === email);
      const accessChanged = previous && (previous.Active !== employee.Active || previous.Role !== employee.Role);
      if (accessChanged) await revokeEmployeePortalAccess(previous);
      await saveRecord("employees", emailKey(email), employee);
      await provisionEmployeeWorkspace(employee);
      resetEmployeeForm(); await refreshData("Team member saved"); toast(employee.Active === "No" && previous ? "Team member marked inactive and direct portal sharing removed." : employee.Active === "No" ? "Team member saved as inactive." : updating ? "Team member updated." : "Team member and personal SharePoint workspace created.", "success");
    } catch (error) { toast(error.message || "Could not save team member.", "error"); }
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
      await ensureBimRegister("issues");
      const ownerEmail = byId("issue-owner").value;
      const owner = state.employees.find((employee) => (employee.Email || "").toLowerCase() === ownerEmail.toLowerCase());
      await saveRecord("issues", recordId("issue"), { Title: byId("issue-title").value.trim(), ProjectCode: byId("issue-project").value, Discipline: byId("issue-discipline").value, IssueType: byId("issue-type").value, Reference: byId("issue-reference").value.trim(), OwnerEmail: ownerEmail, OwnerName: owner?.DisplayName || owner?.Title || ownerEmail, DueDate: byId("issue-due-date").value || "", Priority: byId("issue-priority").value, Status: byId("issue-status").value, Notes: byId("issue-notes").value.trim(), ReportedBy: state.profile.displayName || accountEmail(), createdAt: new Date().toISOString() });
      event.target.reset(); await refreshData("BIM issue logged"); toast("BIM issue logged.", "success");
    } catch (error) { toast(error.message || "Could not log BIM issue.", "error"); }
  }

  async function submitRegister(event) {
    event.preventDefault();
    if (!isManager()) return;
    try {
      await ensureBimRegister("registers");
      await saveRecord("registers", recordId("register"), { RecordType: byId("register-type").value, ProjectCode: byId("register-project").value, Discipline: byId("register-discipline").value, Number: byId("register-number").value.trim(), Title: byId("register-title").value.trim(), Revision: byId("register-revision").value.trim(), BIMStage: byId("register-lod").value, Status: byId("register-status").value, PlannedDate: byId("register-date").value || "", SharePointLink: byId("register-link").value.trim(), Notes: byId("register-notes").value.trim(), RegisteredBy: state.profile.displayName || accountEmail(), createdAt: new Date().toISOString() });
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
    byId("issue-form").addEventListener("submit", submitIssue);
    byId("register-form").addEventListener("submit", submitRegister);
    byId("tasks-list").addEventListener("click", updateTaskStatus);
    byId("projects-list").addEventListener("click", (event) => { beginProjectEdit(event); archiveProject(event); deleteProject(event); });
    byId("employee-form").addEventListener("submit", submitEmployee);
    byId("team-list").addEventListener("click", beginEmployeeEdit);
    byId("employee-cancel-button").addEventListener("click", resetEmployeeForm);
    byId("employee-active").addEventListener("change", updateEmployeeSaveLabel);
    byId("refresh-setup").addEventListener("click", refreshSetup);
    byId("prepare-bim-registers").addEventListener("click", prepareBimRegisters);
    byId("work-project").addEventListener("change", renderWorkTaskOptions);
    byId("task-filter-form").addEventListener("submit", (event) => event.preventDefault());
    ["task-filter-project", "task-filter-discipline", "task-filter-status", "task-filter-assignee", "task-filter-due"].forEach((id) => byId(id).addEventListener("change", renderTasks));
    byId("clear-task-filters").addEventListener("click", () => { byId("task-filter-form").reset(); renderTaskFilters(); renderTasks(); });
    [byId("work-start"), byId("work-end")].forEach((input) => input.addEventListener("input", updateDuration));
    document.querySelectorAll(".nav-item").forEach((button) => button.addEventListener("click", () => showView(button.dataset.view)));
    try { await initialiseAuth(); if (state.account) await openPortal(); }
    catch (error) { console.error(error); byId("sign-in-status").textContent = "Microsoft login configuration needs attention."; }
  }

  boot();
})();
