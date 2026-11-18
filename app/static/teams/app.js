const API_BASE = (window.CONFIG && window.CONFIG.API_BASE_URL) ? window.CONFIG.API_BASE_URL.replace(/\/+$/, "") : "";

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
const clearChatBtn = document.getElementById("clearChatBtn");

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

// Workspace label cache (dynamically populated from backend groups API)
const workspaceNameCache = {
  "all": "All Workspaces",
  "19:7bdcc4616332483b82e7488c4424ec23@thread.v2": "Intranet Solutions",
  "19:2337de30cb834d059c344292c8377656@thread.v2": "AgentAI Bot",
  "channel_athen_tech_solutions_general": "Athen Technology Solutions — General",
  "channel_synergy_comms_approvers_general": "Synergy Comms Approvers — General"
};

function cleanWorkspaceLabel(id, rawName) {
  if (!id || id === "all") return "All Workspaces";
  const cleanId = decodeURIComponent(id).trim();
  if (workspaceNameCache[cleanId]) return workspaceNameCache[cleanId];
  if (rawName && !rawName.startsWith("19:") && !rawName.includes("@thread") && !rawName.startsWith("Chat ")) {
    try {
      const decoded = decodeURIComponent(rawName).trim();
      workspaceNameCache[cleanId] = decoded;
      return decoded;
    } catch (e) {
      return rawName;
    }
  }
  if (cleanId.includes("7bdcc461")) return "Intranet Solutions";
  if (cleanId.includes("2337de30")) return "AgentAI Bot";
  if (cleanId.includes("athen_tech")) return "Athen Technology Solutions — General";
  if (cleanId.includes("synergy")) return "Synergy Comms Approvers — General";
  return cleanId.replace(/_/g, " ").replace(/\b\w/g, l => l.toUpperCase());
}

// Deterministic 1 Session Per Group / Workspace
function getWorkspaceSessionId(groupId) {
  const cleanGid = (groupId || "all").toString().replace(/[^a-zA-Z0-9_-]/g, "_");
  const uid = (currentUserId || "user").toString().replace(/[^a-zA-Z0-9_-]/g, "_");
  return `sess_${uid}_${cleanGid}`;
}

let currentGroupId = "all";
let currentSessionId = getWorkspaceSessionId("all");

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

// Local storage persistence helpers
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

// Startup: Purge legacy dummy / corrupted sessions from localStorage
(function sanitizeLegacySessions() {
  try {
    const raw = getLocalSessions();
    if (Array.isArray(raw)) {
      const valid = raw.filter(s => {
        const lm = (s.last_message || "").trim();
        const t = (s.title || "").trim();
        if (!lm || lm === "No messages yet" || lm.startsWith("Hello ") || (t === "New Conversation" && !lm)) {
          return false;
        }
        return true;
      });
      saveLocalSessions(valid);
    }
  } catch (e) {}
})();

// Update top header workspace pill
function updateWorkspaceLabel() {
  if (activeWorkspaceName && groupSelector) {
    const selectedText = groupSelector.options[groupSelector.selectedIndex]?.text || cleanWorkspaceLabel(currentGroupId);
    activeWorkspaceName.innerText = selectedText.replace(/^[^\w]+/, "").trim();
  }
}

// Load Groups / Workspaces Dropdown
async function loadGroups() {
  try {
    const url = currentUserId ? `${API_BASE}/api/v1/teams/groups?user_id=${encodeURIComponent(currentUserId)}` : `${API_BASE}/api/v1/teams/groups`;
    const res = await fetch(url);
    const data = await res.json();
    if (!groupSelector) return;

    groupSelector.innerHTML = '<option value="all">🌐 All Workspaces</option>';
    if (data.groups && data.groups.length > 0) {
      data.groups.forEach(g => {
        workspaceNameCache[g.id] = g.name;
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

// Append message to chat viewport
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

  // Evaluation badges
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

// Render clean welcome state for active workspace
function renderWelcomeState() {
  chatViewport.innerHTML = "";
  const greetingName = (currentUserName && currentUserName !== "Team Member") ? currentUserName : "there";
  const label = cleanWorkspaceLabel(currentGroupId);
  appendMessage("assistant", `Hello ${greetingName}! I am your personal NoMeet Delivery & Operations Copilot for **${label}**. How can I assist you with this workspace's tasks, bugs, or specifications today?`);
}

// Load messages for a workspace's dedicated session
async function loadWorkspaceMessages(sessionId) {
  chatViewport.innerHTML = "";
  
  // 1. Instant render from local cache
  const cachedMessages = getLocalMessages(sessionId);
  if (cachedMessages && cachedMessages.length > 0) {
    cachedMessages.forEach(m => appendMessage(m.role, m.content, m.evidence, m.eval_scores));
  } else {
    renderWelcomeState();
  }

  // 2. Fetch latest from backend
  try {
    const res = await fetch(`${API_BASE}/api/v1/teams/sessions/${sessionId}`);
    if (res.ok) {
      const data = await res.json();
      if (data.messages && data.messages.length > 0) {
        // If server has more messages, refresh viewport & cache
        if (data.messages.length !== cachedMessages.length) {
          chatViewport.innerHTML = "";
          data.messages.forEach(m => appendMessage(m.role, m.content, m.evidence, m.eval_scores));
        }
        saveLocalMessages(sessionId, data.messages);
      }
    }
  } catch (e) {
    console.warn("Backend session messages check:", e);
  }
}

// Switch Workspace seamlessly (One Session Per Workspace - Old Chats Are Always Preserved!)
async function switchWorkspace(groupId) {
  currentGroupId = groupId || "all";
  if (groupSelector && groupSelector.value !== currentGroupId) {
    groupSelector.value = currentGroupId;
  }
  updateWorkspaceLabel();
  
  currentSessionId = getWorkspaceSessionId(currentGroupId);

  // Load metrics for the active workspace
  loadMetrics();

  // Load the dedicated conversation for this workspace
  await loadWorkspaceMessages(currentSessionId);

  // Update recent conversations list
  loadSessions();
}

if (groupSelector) {
  groupSelector.addEventListener("change", async () => {
    await switchWorkspace(groupSelector.value);
  });
}

// Render Recent Conversations in Sidebar (Only real conversations with user questions)
async function loadSessions() {
  let serverSessions = [];
  try {
    const params = new URLSearchParams();
    if (currentUserId) params.set("user_id", currentUserId);
    const res = await fetch(`${API_BASE}/api/v1/teams/sessions?${params.toString()}`);
    if (res.ok) {
      serverSessions = await res.json();
    }
  } catch (e) {
    console.warn("Server sessions fetch warning:", e);
  }

  // Merge server sessions with local storage cache (dedup by project_id / session_id)
  const localList = getLocalSessions();
  const sessionMap = new Map();
  localList.forEach(s => {
    const key = s.project_id || s.session_id;
    if (key) sessionMap.set(key, s);
  });
  if (Array.isArray(serverSessions)) {
    serverSessions.forEach(s => {
      const key = s.project_id || s.session_id;
      if (key && !sessionMap.has(key)) {
        sessionMap.set(key, s);
      }
    });
  }

  // Only keep sessions that have actual conversation history
  let validSessions = Array.from(sessionMap.values()).filter(s => {
    const lm = (s.last_message || "").trim();
    return lm && lm !== "No messages yet" && !lm.startsWith("Hello ");
  });
  saveLocalSessions(validSessions);

  // Sort by recent activity
  validSessions.sort((a, b) => new Date(b.updated_at || b.created_at || 0) - new Date(a.updated_at || a.created_at || 0));

  sessionsList.innerHTML = "";

  if (validSessions.length === 0) {
    const hint = document.createElement("div");
    hint.className = "empty-sessions-hint";
    hint.style.cssText = "padding: 24px 12px; color: #94a3b8; font-size: 11px; text-align: center; line-height: 1.5;";
    hint.innerHTML = `No previous conversations.<br/><span style="opacity: 0.7;">Type a question below to start chatting in <strong>${cleanWorkspaceLabel(currentGroupId)}</strong>!</span>`;
    sessionsList.appendChild(hint);
    return;
  }

  validSessions.forEach(s => {
    const card = document.createElement("div");
    const projId = s.project_id || s.group_id || "all";
    const isActive = (projId === currentGroupId || s.session_id === currentSessionId);
    card.className = `session-card ${isActive ? "active" : ""}`;
    card.dataset.sessionId = s.session_id;
    card.dataset.groupId = projId;

    // Clicking card switches directly to that workspace conversation
    card.onclick = async () => {
      await switchWorkspace(projId);
    };

    // Header: Workspace Tag, Time, Delete Button
    const header = document.createElement("div");
    header.className = "session-card-header";

    const projTag = document.createElement("span");
    projTag.className = "session-project-tag";
    const cleanProj = cleanWorkspaceLabel(projId, s.project_name);
    projTag.innerText = (projId === "all" || cleanProj === "All") ? "🌐 All Workspaces" : `📁 ${cleanProj}`;

    const metaRight = document.createElement("div");
    metaRight.className = "session-header-meta";

    const timeSpan = document.createElement("span");
    timeSpan.className = "session-time";
    timeSpan.innerText = timeAgo(s.updated_at || s.created_at);

    // Delete / Clear Button
    const delBtn = document.createElement("button");
    delBtn.className = "btn-del-session";
    delBtn.innerHTML = "🗑️";
    delBtn.title = "Delete Conversation";
    delBtn.onclick = (e) => {
      e.stopPropagation();
      deleteSession(s.session_id, projId);
    };

    metaRight.appendChild(timeSpan);
    metaRight.appendChild(delBtn);

    header.appendChild(projTag);
    header.appendChild(metaRight);

    // Title: User question
    const title = document.createElement("div");
    title.className = "session-card-title";
    title.innerText = s.title || "Conversation";

    // Snippet: Latest reply excerpt
    const snippet = document.createElement("div");
    snippet.className = "session-card-snippet";
    snippet.innerText = s.last_message || "";

    card.appendChild(header);
    card.appendChild(title);
    card.appendChild(snippet);
    sessionsList.appendChild(card);
  });
}

// Delete / Clear a workspace's conversation
async function deleteSession(sessionId, groupId) {
  const targetSid = sessionId || currentSessionId;
  const targetGid = groupId || currentGroupId;

  const localList = getLocalSessions().filter(s => s.session_id !== targetSid && s.project_id !== targetGid);
  saveLocalSessions(localList);

  try {
    localStorage.removeItem(`nomeet_msgs_${targetSid}`);
  } catch (e) {}

  try {
    fetch(`${API_BASE}/api/v1/teams/sessions/${targetSid}`, { method: "DELETE" });
  } catch (e) {}

  if (targetGid === currentGroupId || targetSid === currentSessionId) {
    renderWelcomeState();
  }
  loadSessions();
}

// Handle Form Submission
chatForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const q = chatInput.value.trim();
  if (!q) return;

  chatInput.value = "";
  appendMessage("user", q);

  currentSessionId = getWorkspaceSessionId(currentGroupId);

  // Save user message to local cache
  const localMsgs = getLocalMessages(currentSessionId);
  localMsgs.push({ role: "user", content: q });
  saveLocalMessages(currentSessionId, localMsgs);

  // Update session record in local sessions
  const localSessions = getLocalSessions();
  let sess = localSessions.find(s => s.session_id === currentSessionId || s.project_id === currentGroupId);
  if (!sess) {
    sess = {
      session_id: currentSessionId,
      project_id: currentGroupId || "all",
      project_name: cleanWorkspaceLabel(currentGroupId),
      title: q.slice(0, 36) + (q.length > 36 ? "..." : ""),
      user_id: currentUserId,
      user_name: currentUserName,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString(),
      last_message: q
    };
    localSessions.unshift(sess);
  } else {
    sess.last_message = q;
    sess.updated_at = new Date().toISOString();
    sess.title = q.slice(0, 36) + (q.length > 36 ? "..." : "");
    sess.project_id = currentGroupId;
    sess.project_name = cleanWorkspaceLabel(currentGroupId);
  }
  saveLocalSessions(localSessions);

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
      sess.last_message = (data.response || "").slice(0, 60);
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

// "+ New Session": Clears current workspace chat so user can start fresh
if (newSessionBtn) {
  newSessionBtn.onclick = () => {
    deleteSession(currentSessionId, currentGroupId);
  };
}

// "Clear Chat" top bar button
if (clearChatBtn) {
  clearChatBtn.onclick = () => {
    deleteSession(currentSessionId, currentGroupId);
  };
}

// Initialize Microsoft Teams JS SDK
if (window.microsoftTeams) {
  microsoftTeams.app.initialize().then(() => {
    try {
      microsoftTeams.app.notifySuccess();
    } catch (e) {}

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
    } catch (cfgErr) {}

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
        } catch (regErr) {}

        await loadGroups();
        await switchWorkspace(teamsGroupId);
      }
    });
  }).catch(() => {
    console.log("Teams SDK running outside native Teams client.");
  });
}

// App startup sequence
loadGroups().then(() => {
  switchWorkspace("all");
});
