(() => {
  "use strict";

  const CONFIG = window.ASTERWIX_PORTAL_CONFIG;
  const GRAPH_SCOPES = ["User.Read", "Sites.ReadWrite.All"];
  const FOLDERS = { employees: "employees", projects: "projects", tasks: "tasks", workspaces: "employee-workspaces" };
  const state = { account: null, profile: null, role: "Staff", projects: [], tasks: [], workLogs: [], employees: [], missingFolders: [] };
  const byId = (id) => document.getElementById(id);
  const isManager = () => state.role === "Admin" || state.role === "Team Lead";
  const isAdmin = () => state.role === "Admin";
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
    await checkPaths(Object.entries(FOLDERS).filter(([key]) => key !== "workspaces").map(([key]) => ({ label: key, path: folderPath(key) })));
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

  async function provisionEmployeeWorkspace(employee) {
    const email = String(employee.Email || "").toLowerCase();
    if (!email || employee.Active === "No") return;
    await ensureWorkspace(email);
    if (employee.Role === "Admin") return;
    await inviteToFolder(folderPath("employees"), email, "read");
    await inviteToFolder(folderPath("projects"), email, employee.Role === "Team Lead" ? "write" : "read");
    await inviteToFolder(folderPath("tasks"), email, employee.Role === "Team Lead" ? "write" : "read");
    await inviteToFolder(workspacePath(email), email, "write");
    if (employee.Role === "Team Lead") await inviteToFolder(folderPath("workspaces"), email, "read");
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
    await checkFolders();
    if (state.missingFolders.length) return;
    [state.projects, state.tasks, state.employees] = await Promise.all([listRecords("projects"), listRecords("tasks"), listRecords("employees")]);
    setRole();
    await checkPaths([{ label: isManager() ? FOLDERS.workspaces : "your personal work folder", path: isManager() ? folderPath("workspaces") : workspaceLogsPath() }], true);
    if (state.missingFolders.length) return;
    state.workLogs = isManager() ? await listAllWorkspaceLogs() : await listRecordsAt(workspaceLogsPath());
  }

  function accountEmail() { return state.profile?.mail || state.profile?.userPrincipalName || state.account?.username || ""; }

  function setRole() {
    const email = accountEmail().toLowerCase();
    const entry = state.employees.find((employee) => (employee.Email || "").toLowerCase() === email && employee.Active !== "No");
    state.role = entry?.Role || (email === CONFIG.bootstrapAdminEmail.toLowerCase() ? "Admin" : "Staff");
  }

  function setProfileUI() {
    const name = state.profile.displayName || state.account.name || state.account.username;
    const email = accountEmail();
    byId("profile-name").textContent = name;
    byId("profile-role").textContent = `${state.role} · ${email}`;
    byId("profile-initials").textContent = name.split(/\s+/).slice(0, 2).map((part) => part[0]).join("").toUpperCase();
    byId("role-badge").textContent = state.role;
    document.querySelectorAll(".manager-only").forEach((element) => element.classList.toggle("hidden", !isManager()));
    document.querySelectorAll(".admin-only").forEach((element) => element.classList.toggle("hidden", !isAdmin()));
  }

  function visibleTasks() {
    const email = accountEmail().toLowerCase();
    return isManager() ? state.tasks : state.tasks.filter((task) => (task.AssigneeEmail || "").toLowerCase() === email);
  }

  function projectByCode(code) { return state.projects.find((project) => project.ProjectCode === code); }

  function renderMetrics() {
    const today = dubaiDate();
    const email = accountEmail().toLowerCase();
    const todayEntries = state.workLogs.filter((entry) => entry.WorkDate === today && (isManager() || (entry.EmployeeEmail || "").toLowerCase() === email));
    const totalMinutes = todayEntries.reduce((sum, entry) => sum + Number(entry.DurationMinutes || 0), 0);
    const hours = totalMinutes ? `${Math.floor(totalMinutes / 60)}h ${totalMinutes % 60}m` : "0h";
    const tasks = visibleTasks();
    const activeProjects = state.projects.filter((project) => project.Status === "Active").length;
    byId("metrics").innerHTML = [[hours, "Logged today"], [String(todayEntries.length), "Today’s entries"], [String(tasks.filter((task) => task.Status !== "Completed").length), "Open tasks"], [String(activeProjects), "Active projects"]].map(([value, label]) => `<div class="metric"><div class="metric-value">${esc(value)}</div><div class="metric-label">${esc(label)}</div></div>`).join("");
  }

  function renderWorkForm() {
    const select = byId("work-project");
    select.innerHTML = `<option value="">Select a project</option>${state.projects.filter((project) => project.Status !== "Completed").map((project) => `<option value="${esc(project.ProjectCode)}">${esc(project.ProjectCode)} · ${esc(project.Title)}</option>`).join("")}`;
    byId("work-date").value = dubaiDate();
  }

  function renderRecentWork() {
    const email = accountEmail().toLowerCase();
    const entries = state.workLogs.filter((entry) => isManager() || (entry.EmployeeEmail || "").toLowerCase() === email).sort((a, b) => `${b.WorkDate || ""}${b.StartTime || ""}`.localeCompare(`${a.WorkDate || ""}${a.StartTime || ""}`)).slice(0, 7);
    byId("recent-work").innerHTML = entries.length ? entries.map((entry) => `<div class="activity-row"><strong>${esc(entry.TaskTitle || "Work entry")}</strong><span>${esc(entry.ProjectCode || "—")} · ${esc(entry.WorkDate || "")} · ${esc(entry.StartTime || "")}–${esc(entry.EndTime || "")} · ${esc(entry.EmployeeName || "")}</span></div>`).join("") : `<p class="muted">No work entries have been logged yet.</p>`;
  }

  function renderTasks() {
    const tasks = visibleTasks();
    byId("tasks-list").innerHTML = tasks.length ? tasks.map((task) => {
      const project = projectByCode(task.ProjectCode);
      return `<article class="task-card"><p class="eyebrow">${esc(task.ProjectCode || "NO PROJECT")}</p><h2>${esc(task.Title)}</h2><p>${esc(project?.Client || "Asterwix project")}</p><p>${esc(task.AssigneeEmail || "")}</p><div class="task-meta"><span class="badge">${esc(task.Status || "Not started")}</span><span>Due: ${esc(task.EndDate || "—")}</span></div></article>`;
    }).join("") : `<section class="card"><p class="muted">No task is assigned to your Asterwix account yet.</p></section>`;
  }

  function renderManagers() {
    if (!isManager()) return;
    byId("task-project").innerHTML = `<option value="">Select project</option>${state.projects.map((project) => `<option value="${esc(project.ProjectCode)}">${esc(project.ProjectCode)} · ${esc(project.Title)}</option>`).join("")}`;
    byId("projects-list").innerHTML = state.projects.length ? `<table class="data-table"><thead><tr><th>Code</th><th>Project</th><th>Client</th><th>Status</th><th>Target</th></tr></thead><tbody>${state.projects.map((project) => `<tr><td>${esc(project.ProjectCode)}</td><td>${esc(project.Title)}</td><td>${esc(project.Client || "—")}</td><td>${esc(project.Status || "—")}</td><td>${esc(project.TargetDate || "—")}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">No projects created yet.</p>`;
    byId("team-list").innerHTML = state.employees.length ? `<table class="data-table"><thead><tr><th>Name</th><th>Email</th><th>Role</th><th>Active</th></tr></thead><tbody>${state.employees.map((employee) => `<tr><td>${esc(employee.DisplayName || "—")}</td><td>${esc(employee.Email || "—")}</td><td>${esc(employee.Role || "Staff")}</td><td>${esc(employee.Active || "Yes")}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">Add staff after portal storage is ready.</p>`;
  }

  function renderSetup() {
    if (!isAdmin()) return;
    byId("setup-list").innerHTML = Object.entries(FOLDERS).map(([key, name]) => `<li><strong>${esc(name)}</strong> — ${state.missingFolders.includes(key) ? "not created" : "ready"}</li>`).join("");
  }

  function renderAll() {
    byId("today-label").textContent = new Intl.DateTimeFormat("en-GB", { timeZone: "Asia/Dubai", weekday: "long", year: "numeric", month: "long", day: "numeric" }).format(new Date());
    setProfileUI(); renderMetrics(); renderWorkForm(); renderRecentWork(); renderTasks(); renderManagers(); renderSetup();
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
    if (!byId("work-project").value || !byId("work-task-title").value.trim() || !duration) return toast("Enter project, task, and valid start/end time.", "error");
    try {
      const taskTitle = byId("work-task-title").value.trim();
      await ensureWorkspace();
      await saveRecordAt(workspaceLogsPath(), recordId("work"), { Title: `${byId("work-date").value} · ${taskTitle}`, WorkDate: byId("work-date").value, TaskTitle: taskTitle, ProjectCode: byId("work-project").value, EmployeeEmail: accountEmail(), EmployeeName: state.profile.displayName, StartTime: byId("work-start").value, EndTime: byId("work-end").value, DurationMinutes: duration, WorkNote: byId("work-note").value.trim(), createdAt: new Date().toISOString() });
      event.target.reset(); byId("work-date").value = dubaiDate(); updateDuration(); await refreshData("Work entry saved"); toast("Daily work entry saved to SharePoint.", "success");
    } catch (error) { toast(error.message || "Could not save work entry.", "error"); }
  }

  async function submitProject(event) {
    event.preventDefault();
    try {
      await saveRecord("projects", recordId("project"), { Title: byId("project-name").value.trim(), ProjectCode: byId("project-code").value.trim().toUpperCase(), Client: byId("project-client").value.trim(), Status: byId("project-status").value, TargetDate: byId("project-target-date").value || "", createdAt: new Date().toISOString() });
      event.target.reset(); await refreshData("Project created"); toast("Project created.", "success");
    } catch (error) { toast(error.message || "Could not create project.", "error"); }
  }

  async function submitTask(event) {
    event.preventDefault();
    try {
      await saveRecord("tasks", recordId("task"), { Title: byId("task-title").value.trim(), ProjectCode: byId("task-project").value, AssigneeEmail: byId("task-assignee").value.trim().toLowerCase(), StartDate: byId("task-start-date").value || "", EndDate: byId("task-end-date").value || "", Priority: byId("task-priority").value, Status: byId("task-status").value, Notes: byId("task-notes").value.trim(), createdAt: new Date().toISOString() });
      event.target.reset(); await refreshData("Task assigned"); toast("Task assigned.", "success");
    } catch (error) { toast(error.message || "Could not assign task.", "error"); }
  }

  async function submitEmployee(event) {
    event.preventDefault();
    if (!isAdmin()) return;
    try {
      const email = byId("employee-email").value.trim().toLowerCase();
      const name = byId("employee-name").value.trim();
      const employee = { Title: name, Email: email, DisplayName: name, Role: byId("employee-role").value, Active: byId("employee-active").value, createdAt: state.employees.find((member) => (member.Email || "").toLowerCase() === email)?.createdAt || new Date().toISOString() };
      await saveRecord("employees", emailKey(email), employee);
      await provisionEmployeeWorkspace(employee);
      event.target.reset(); await refreshData("Team member and workspace saved"); toast(employee.Active === "No" ? "Team member saved as inactive." : "Team member and personal SharePoint workspace created.", "success");
    } catch (error) { toast(error.message || "Could not save team member.", "error"); }
  }

  async function signOut() { await msalInstance.logoutPopup({ account: state.account, postLogoutRedirectUri: window.location.origin + window.location.pathname }); }

  async function boot() {
    byId("sign-in-button").addEventListener("click", signIn);
    byId("sign-out-button").addEventListener("click", signOut);
    byId("work-log-form").addEventListener("submit", submitWork);
    byId("project-form").addEventListener("submit", submitProject);
    byId("task-form").addEventListener("submit", submitTask);
    byId("employee-form").addEventListener("submit", submitEmployee);
    byId("refresh-setup").addEventListener("click", refreshSetup);
    [byId("work-start"), byId("work-end")].forEach((input) => input.addEventListener("input", updateDuration));
    document.querySelectorAll(".nav-item").forEach((button) => button.addEventListener("click", () => showView(button.dataset.view)));
    try { await initialiseAuth(); if (state.account) await openPortal(); }
    catch (error) { console.error(error); byId("sign-in-status").textContent = "Microsoft login configuration needs attention."; }
  }

  boot();
})();
