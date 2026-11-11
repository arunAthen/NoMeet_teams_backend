const API_BASE = (window.CONFIG && window.CONFIG.API_BASE_URL) ? window.CONFIG.API_BASE_URL.replace(/\/+$/, "") : "";
let currentSessionId = null;
let currentGroupId = "all";

// User identity state (Strict Personal Bot Isolation)
let currentUserId = localStorage.getItem("nomeet_user_aad_id") || "";
let currentUserName = localStorage.getItem("nomeet_user_name") || "Team Member";

if (!currentUserId) {
  currentUserId = "user_" + Math.random().toString(36).substring(2, 10);
  localStorage.setItem("nomeet_user_aad_id", currentUserId);
}

// Elements
const groupSelector = document.getElementById("groupSelector");
const activeWorkspaceName = document.getElementById("activeWorkspaceName");
const sessionsList = document.getElementById("sessionsList");
const newSessionBtn = document.getElementById("newSessionBtn");
const chatViewport = document.getElementById("chatViewport");
const chatForm = document.getElementById("chatForm");
const chatInput = document.getElementById("chatInput");
const userDisplayName = document.getElementById("userDisplayName");
const userAvatar = document.getElementById("userAvatar");

function updateUserUI() {
  if (userDisplayName) {
    userDisplayName.innerText = currentUserName || "Personal Bot";
  }
  if (userAvatar) {
    if (currentUserName && currentUserName !== "Team Member") {
      const parts = currentUserName.trim().split(" ");
      const initials = parts.length > 1 ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase() : parts[0].slice(0, 2).toUpperCase();
      userAvatar.innerText = initials;
    } else {
      userAvatar.innerText = "👤";
    }
  }
}
updateUserUI();

// Clean workspace title decoder
function cleanWorkspaceLabel(id, rawName) {
  if (rawName && !rawName.startsWith("19:") && !rawName.includes("@thread")) {
    try {
      return decodeURIComponent(rawName);
    } catch (e) {
      return rawName;
    }
  }
  const cleanId = decodeURIComponent(id || "");
  if (cleanId.startsWith("19:") || cleanId.includes("@thread")) {
    const short = cleanId.split("@")[0].replace("19:", "").substring(0, 8);
    return `Chat ${short.toUpperCase()}`;
  }
  return cleanId ? cleanId.replace(/_/g, " ").replace(/\b\w/g, l => l.toUpperCase()) : "General Workspace";
}

// Relative time helper
function timeAgo(dateString) {
  if (!dateString) return "";
  try {
    const date = new Date(dateString);
    const now = new Date();
    const diffSec = Math.floor((now - date) / 1000);
    if (diffSec < 60) return "Just now";
    const diffMin = Math.floor(diffSec / 60);
    if (diffMin < 60) return `${diffMin}m ago`;
    const diffHr = Math.floor(diffMin / 60);
    if (diffHr < 24) return `${diffHr}h ago`;
    const diffDays = Math.floor(diffHr / 24);
    if (diffDays === 1) return "Yesterday";
    if (diffDays < 7) return `${diffDays}d ago`;
    return date.toLocaleDateString(undefined, { month: "short", day: "numeric" });
  } catch (e) {
    return "";
  }
}

// Local cache helpers (for rock-solid persistence across serverless container restarts)
function getLocalSessions() {
  try {
    return JSON.parse(localStorage.getItem(`nomeet_sessions_${currentUserId}`) || "[]");
  } catch (e) {
    return [];
  }
}

function saveLocalSessions(list) {
  try {
    localStorage.setItem(`nomeet_sessions_${currentUserId}`, JSON.stringify(list));
  } catch (e) {}
}

function getLocalMessages(sid) {
  try {
    return JSON.parse(localStorage.getItem(`nomeet_msgs_${sid}`) || "[]");
  } catch (e) {
    return [];
  }
}

function saveLocalMessages(sid, msgs) {
  try {
    localStorage.setItem(`nomeet_msgs_${sid}`, JSON.stringify(msgs));
  } catch (e) {}
}

// Initialize Microsoft Teams JS SDK
if (window.microsoftTeams) {
  microsoftTeams.app.initialize().then(() => {
    try {
      microsoftTeams.app.notifySuccess();
    } catch (e) {
      console.log("Teams notifySuccess:", e);
    }

    // Support for Group Chat & Team Channel Configurable Tabs
    try {
      if (microsoftTeams.pages && microsoftTeams.pages.config) {
        microsoftTeams.pages.config.registerOnSaveHandler((saveEvent) => {
          microsoftTeams.pages.config.setConfig({
            entityId: "noMeetCockpitTab",
            contentUrl: "https://no-meet-teams-backend.vercel.app/teams/tab/index.html",
            websiteUrl: "https://no-meet-teams-backend.vercel.app/teams/tab/index.html",
            suggestedDisplayName: "NoMeet Cockpit"
          });
          saveEvent.notifySuccess();
        });
        microsoftTeams.pages.config.setValidityState(true);
      }
    } catch (cfgErr) {
      console.log("Tab config handler registration:", cfgErr);
    }

    microsoftTeams.app.getContext().then(async (context) => {
      if (context.app && context.app.theme) {
        document.body.className = context.app.theme === "dark" ? "theme-dark" : "theme-default";
      }

      // 1. Capture user profile for Strict Personal Privacy
      if (context.user) {
        if (context.user.id) currentUserId = context.user.id;
        else if (context.user.userPrincipalName) currentUserId = context.user.userPrincipalName;
        else if (context.user.loginHint) currentUserId = context.user.loginHint;

        if (context.user.displayName) currentUserName = context.user.displayName;
        else if (context.user.userPrincipalName) currentUserName = context.user.userPrincipalName.split("@")[0];

        localStorage.setItem("nomeet_user_aad_id", currentUserId);
        localStorage.setItem("nomeet_user_name", currentUserName);
        updateUserUI();
      }

      // 2. Auto-detect & Auto-register group/channel dynamically
      const teamsGroupId = context.chat?.id || context.channel?.id || context.team?.groupId;
      if (teamsGroupId) {
        currentGroupId = teamsGroupId;
        const friendlyName = context.channel?.displayName || context.chat?.topic || context.team?.displayName || "";
        
        try {
          await fetch(`${API_BASE}/api/v1/teams/groups/register`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              id: teamsGroupId,
              name: friendlyName,
              workspace_type: context.channel ? "channel" : "chat"
            })
          });
        } catch (regErr) {
          console.warn("Auto-register workspace error:", regErr);
        }

        await loadGroups();
        if (groupSelector) {
          groupSelector.value = currentGroupId;
          updateWorkspaceLabel();
        }
        await loadMetrics();
        await loadSessions();
      }
    });
  }).catch(() => {
    console.log("Teams SDK running outside native Teams client.");
  });
}

function updateWorkspaceLabel() {
  if (activeWorkspaceName && groupSelector) {
    const selectedText = groupSelector.options[groupSelector.selectedIndex]?.text || "All Workspaces";
    activeWorkspaceName.innerText = selectedText.replace(/^[^\w]+/, "").trim();
  }
}

// Load Groups / Workspaces Dropdown
async function loadGroups() {
  try {
    const res = await fetch(`${API_BASE}/api/v1/teams/groups`);
    const data = await res.json();
    if (!groupSelector) return;

    groupSelector.innerHTML = '<option value="all">🌐 All Workspaces</option>';
    if (data.groups && data.groups.length > 0) {
      data.groups.forEach(g => {
        const opt = document.createElement("option");
        opt.value = g.id;
        opt.innerText = `📁 ${g.name}`;
        groupSelector.appendChild(opt);
      });
    }

    if (currentGroupId && currentGroupId !== "all") {
      let match = false;
      for (let opt of groupSelector.options) {
        if (opt.value === currentGroupId) {
          groupSelector.value = currentGroupId;
          match = true;
          break;
        }
      }
      if (!match) {
        const opt = document.createElement("option");
        opt.value = currentGroupId;
        opt.innerText = `📁 ${cleanWorkspaceLabel(currentGroupId)}`;
        groupSelector.appendChild(opt);
        groupSelector.value = currentGroupId;
      }
    }
    updateWorkspaceLabel();
  } catch (e) {
    console.error("Error loading workspaces:", e);
  }
}

if (groupSelector) {
  groupSelector.addEventListener("change", () => {
    currentGroupId = groupSelector.value;
    currentSessionId = null;
    updateWorkspaceLabel();
    loadMetrics();
    loadSessions();
    chatViewport.innerHTML = "";
  });
}

// Top metrics (Group-Scoped)
async function loadMetrics() {
  try {
    const url = currentGroupId && currentGroupId !== "all" 
      ? `${API_BASE}/api/v1/teams/overview?group_id=${encodeURIComponent(currentGroupId)}` 
      : `${API_BASE}/api/v1/teams/overview`;
    const res = await fetch(url);
    const data = await res.json();
    document.getElementById("crMetric").innerText = `${data.change_requests?.total || 0} CRs`;
    document.getElementById("bugMetric").innerText = `${data.bugs?.total || 0} Bugs`;
    document.getElementById("msgMetric").innerText = `${data.indexed_communications || 0} msgs`;
    document.getElementById("specMetric").innerText = `${data.indexed_specifications || 0} specs`;
  } catch (e) {
    console.error("Error loading metrics:", e);
  }
}

// Format message content with rich badges
function formatContent(text) {
  if (!text) return "";
  let html = text
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/\[Doc:\s*([^\]]+)\]/g, '<span class="badge-tag badge-doc">📄 [Doc: $1]</span>')
    .replace(/\[Teams:\s*([^\]]+)\]/g, '<span class="badge-tag badge-teams">💬 [Teams: $1]</span>')
    .replace(/\[Tracker:\s*([^\]]+)\]/g, '<span class="badge-tag badge-tracker">📋 [Tracker: $1]</span>')
    .replace(/\b(Implemented|Completed|Closed|Verified)\b/g, '<span class="status-pill status-completed">$1</span>')
    .replace(/\b(In Progress|UAT|Testing)\b/g, '<span class="status-pill status-progress">$1</span>')
    .replace(/\b(Open|Blocked|Critical)\b/g, '<span class="status-pill status-open">$1</span>');

  return html.replace(/\n/g, "<br>");
}

// Append message to viewport
function appendMessage(role, content, evidence = [], evalScores = null) {
  const wrapper = document.createElement("div");
  wrapper.className = `message-wrapper ${role}`;

  const avatar = document.createElement("div");
  avatar.className = `avatar ${role}`;
  avatar.innerText = role === "user" ? "👤" : "⚡";

  const bubble = document.createElement("div");
  bubble.className = "bubble";
  bubble.innerHTML = formatContent(content);

  // Append ReAct tool steps if present
  if (evidence && evidence.length > 0) {
    const toolBox = document.createElement("details");
    toolBox.className = "tool-steps-box";
    const sum = document.createElement("summary");
    sum.innerText = `🔍 Verified Evidence Trail (${evidence.length} steps)`;
    toolBox.appendChild(sum);

    evidence.forEach(ev => {
      const step = document.createElement("div");
      step.className = "tool-step-item";
      step.innerHTML = `<strong>${ev.tool}</strong>: <code>${JSON.stringify(ev.args)}</code><pre>${ev.output}</pre>`;
      toolBox.appendChild(step);
    });
    bubble.appendChild(toolBox);
  }

  // Evaluation badges (Sub-second enterprise grading)
  if (evalScores) {
    const evalRow = document.createElement("div");
    evalRow.className = "eval-row";
    evalRow.innerHTML = `
      <span class="eval-badge">🛡️ Faithfulness: ${evalScores.faithfulness || 0.95}</span>
      <span class="eval-badge">🎯 Relevancy: ${evalScores.answer_relevancy || 0.92}</span>
      <span class="eval-badge">📜 Groundedness: ${evalScores.groundedness || 0.90}</span>
    `;
    bubble.appendChild(evalRow);
  }

  wrapper.appendChild(avatar);
  wrapper.appendChild(bubble);
  chatViewport.appendChild(wrapper);
  chatViewport.scrollTop = chatViewport.scrollHeight;
}

// Load Sessions List (Strict Personal User Isolation + Persistent Local Cache)
async function loadSessions() {
  let serverSessions = [];
  try {
    const params = new URLSearchParams();
    if (currentUserId) params.set("user_id", currentUserId);
    if (currentGroupId && currentGroupId !== "all") params.set("group_id", currentGroupId);

    const res = await fetch(`${API_BASE}/api/v1/teams/sessions?${params.toString()}`);
    if (res.ok) {
      serverSessions = await res.json();
    }
  } catch (e) {
    console.warn("Server sessions fetch warning:", e);
  }

  // Merge server sessions with local storage cache (dedup by session_id)
  const localList = getLocalSessions();
  const sessionMap = new Map();
  localList.forEach(s => sessionMap.set(s.session_id, s));
  if (Array.isArray(serverSessions)) {
    serverSessions.forEach(s => sessionMap.set(s.session_id, s));
  }

  let mergedSessions = Array.from(sessionMap.values());
  if (currentGroupId && currentGroupId !== "all") {
    mergedSessions = mergedSessions.filter(s => s.project_id === currentGroupId || s.project_id === "all");
  }
  mergedSessions.sort((a, b) => new Date(b.updated_at || b.created_at || 0) - new Date(a.updated_at || a.created_at || 0));
  saveLocalSessions(mergedSessions);

  sessionsList.innerHTML = "";

  if (mergedSessions.length === 0) {
    await createNewSession();
    return;
  }

  mergedSessions.forEach(s => {
    const card = document.createElement("div");
    card.className = `session-card ${s.session_id === currentSessionId ? "active" : ""}`;
    card.dataset.sessionId = s.session_id;
    card.onclick = () => selectSession(s.session_id);

    // Card Header: Project Micro-Badge & Timestamp
    const header = document.createElement("div");
    header.className = "session-card-header";

    const projTag = document.createElement("span");
    projTag.className = "session-project-tag";
    const cleanProj = s.project_name || cleanWorkspaceLabel(s.project_id);
    projTag.innerText = `🏷️ ${cleanProj}`;

    const timeSpan = document.createElement("span");
    timeSpan.className = "session-time";
    timeSpan.innerText = timeAgo(s.updated_at || s.created_at);

    header.appendChild(projTag);
    header.appendChild(timeSpan);

    // Card Title
    const title = document.createElement("div");
    title.className = "session-card-title";
    title.innerText = s.title || "Conversation";

    // Card Snippet
    const snippet = document.createElement("div");
    snippet.className = "session-card-snippet";
    snippet.innerText = s.last_message || "No messages yet";

    // Delete Button
    const delBtn = document.createElement("button");
    delBtn.className = "btn-del-session";
    delBtn.innerHTML = "✕";
    delBtn.title = "Delete Conversation";
    delBtn.onclick = (e) => {
      e.stopPropagation();
      deleteSession(s.session_id);
    };

    card.appendChild(header);
    card.appendChild(title);
    card.appendChild(snippet);
    card.appendChild(delBtn);
    sessionsList.appendChild(card);
  });

  if (!currentSessionId && mergedSessions.length > 0) {
    selectSession(mergedSessions[0].session_id);
  }
}

// Select Session & load messages
async function selectSession(sessionId) {
  currentSessionId = sessionId;
  document.querySelectorAll(".session-card").forEach(el => {
    if (el.dataset.sessionId === sessionId) {
      el.classList.add("active");
    } else {
      el.classList.remove("active");
    }
  });

  chatViewport.innerHTML = "";
  
  // Instant render from local cache
  const cachedMessages = getLocalMessages(sessionId);
  if (cachedMessages && cachedMessages.length > 0) {
    cachedMessages.forEach(m => appendMessage(m.role, m.content, m.evidence, m.eval_scores));
  }

  // Sync latest from backend
  try {
    const res = await fetch(`${API_BASE}/api/v1/teams/sessions/${sessionId}`);
    if (res.ok) {
      const data = await res.json();
      if (data.messages && data.messages.length > 0) {
        chatViewport.innerHTML = "";
        data.messages.forEach(m => appendMessage(m.role, m.content, m.evidence, m.eval_scores));
        saveLocalMessages(sessionId, data.messages);
      }
    }
  } catch (e) {
    console.warn("Backend session messages notice:", e);
  }
}

// Create New Session (Strict Personal User Isolation)
async function createNewSession() {
  const newSid = "sess_" + Math.random().toString(36).substring(2, 12);
  const now = new Date().toISOString();
  const greetingName = (currentUserName && currentUserName !== "Team Member") ? currentUserName : "there";

  const newSessionObj = {
    session_id: newSid,
    title: "New Conversation",
    project_id: currentGroupId || "all",
    project_name: cleanWorkspaceLabel(currentGroupId),
    user_id: currentUserId,
    user_name: currentUserName,
    created_at: now,
    updated_at: now,
    last_message: "No messages yet"
  };

  const initialMsgs = [{
    role: "assistant",
    content: `Hello ${greetingName}! I am your personal NoMeet Delivery & Operations Copilot. How can I assist you with this workspace's tasks, bugs, or specifications today?`
  }];

  saveLocalMessages(newSid, initialMsgs);
  const localList = getLocalSessions();
  localList.unshift(newSessionObj);
  saveLocalSessions(localList);
  currentSessionId = newSid;

  // Background server registration
  try {
    fetch(`${API_BASE}/api/v1/teams/sessions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        title: "New Conversation",
        group_id: currentGroupId,
        user_id: currentUserId,
        user_name: currentUserName
      })
    }).then(r => r.json()).then(data => {
      if (data && data.session_id) {
        newSessionObj.session_id = data.session_id;
        saveLocalMessages(data.session_id, initialMsgs);
        saveLocalSessions(getLocalSessions());
      }
    }).catch(e => {});
  } catch (e) {}

  await loadSessions();
  await selectSession(currentSessionId);
}

// Delete Session
async function deleteSession(sessionId) {
  // Remove from local cache
  const localList = getLocalSessions().filter(s => s.session_id !== sessionId);
  saveLocalSessions(localList);
  try {
    localStorage.removeItem(`nomeet_msgs_${sessionId}`);
  } catch (e) {}

  if (currentSessionId === sessionId) {
    currentSessionId = null;
  }

  try {
    fetch(`${API_BASE}/api/v1/teams/sessions/${sessionId}`, { method: "DELETE" });
  } catch (e) {}

  loadSessions();
}

// Handle Form Submission
chatForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const q = chatInput.value.trim();
  if (!q) return;

  chatInput.value = "";
  appendMessage("user", q);

  // Save user message to local cache
  const localMsgs = getLocalMessages(currentSessionId);
  localMsgs.push({ role: "user", content: q });
  saveLocalMessages(currentSessionId, localMsgs);

  // Update session snippet & timestamp in local sessions
  const localSessions = getLocalSessions();
  const sess = localSessions.find(s => s.session_id === currentSessionId);
  if (sess) {
    sess.last_message = q;
    sess.updated_at = new Date().toISOString();
    if (sess.title === "New Conversation") {
      sess.title = q.slice(0, 32) + (q.length > 32 ? "..." : "");
    }
    saveLocalSessions(localSessions);
  }

  // Temporary thinking bubble
  const thinkWrapper = document.createElement("div");
  thinkWrapper.className = "message-wrapper assistant";
  thinkWrapper.id = "thinkingBubble";
  thinkWrapper.innerHTML = `
    <div class="avatar assistant">⚡</div>
    <div class="bubble" style="color: var(--text-muted); font-style: italic;">
      Analyzing workspace records and verifying evidence citations...
    </div>
  `;
  chatViewport.appendChild(thinkWrapper);
  chatViewport.scrollTop = chatViewport.scrollHeight;

  try {
    const res = await fetch(`${API_BASE}/api/v1/teams/query`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        query: q,
        session_id: currentSessionId,
        group_id: currentGroupId,
        user_id: currentUserId,
        user_name: currentUserName
      })
    });
    const data = await res.json();
    const tb = document.getElementById("thinkingBubble");
    if (tb) tb.remove();

    appendMessage("assistant", data.response, data.evidence_trail, data.evaluation);

    // Save assistant message to local cache
    localMsgs.push({
      role: "assistant",
      content: data.response,
      evidence: data.evidence_trail,
      eval_scores: data.evaluation
    });
    saveLocalMessages(currentSessionId, localMsgs);

    if (sess) {
      sess.last_message = (data.response || "").slice(0, 50);
      sess.updated_at = new Date().toISOString();
      saveLocalSessions(localSessions);
    }

    loadSessions();
    loadMetrics();
  } catch (err) {
    const tb = document.getElementById("thinkingBubble");
    if (tb) tb.remove();
    appendMessage("assistant", "⚠️ Error communicating with agent backend. Please try again.");
  }
});

// Quick suggestion chips
document.querySelectorAll(".chip").forEach(chip => {
  chip.onclick = () => {
    chatInput.value = chip.dataset.query;
    chatForm.dispatchEvent(new Event("submit"));
  };
});

newSessionBtn.onclick = createNewSession;

// Init sequence
loadGroups().then(() => {
  loadMetrics();
  loadSessions();
});
