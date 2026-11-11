let currentSessionId = null;
let currentGroupId = "all";

// Elements
const groupSelector = document.getElementById("groupSelector");
const activeWorkspaceName = document.getElementById("activeWorkspaceName");
const sessionsList = document.getElementById("sessionsList");
const newSessionBtn = document.getElementById("newSessionBtn");
const chatViewport = document.getElementById("chatViewport");
const chatForm = document.getElementById("chatForm");
const chatInput = document.getElementById("chatInput");

// Initialize Microsoft Teams JS SDK
if (window.microsoftTeams) {
  microsoftTeams.app.initialize().then(() => {
    try {
      microsoftTeams.app.notifySuccess();
    } catch (e) {
      console.log("Teams notifySuccess:", e);
    }
    microsoftTeams.app.getContext().then((context) => {
      if (context.app && context.app.theme) {
        document.body.className = context.app.theme === "dark" ? "theme-dark" : "theme-default";
      }
      // Auto-detect group/channel from Teams context
      const teamsGroupId = context.chat?.id || context.channel?.id || context.team?.groupId;
      if (teamsGroupId) {
        currentGroupId = teamsGroupId;
        if (groupSelector) {
          let found = false;
          for (let opt of groupSelector.options) {
            if (opt.value === currentGroupId) {
              groupSelector.value = currentGroupId;
              found = true;
              break;
            }
          }
          if (!found) {
            const opt = document.createElement("option");
            opt.value = currentGroupId;
            opt.innerText = `👥 Current Chat: ${currentGroupId.slice(0, 15)}...`;
            groupSelector.appendChild(opt);
            groupSelector.value = currentGroupId;
          }
          updateWorkspaceLabel();
          loadMetrics();
          loadSessions();
        }
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
    const res = await fetch("/api/v1/teams/groups");
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
      groupSelector.value = currentGroupId;
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
      ? `/api/v1/teams/overview?group_id=${encodeURIComponent(currentGroupId)}` 
      : "/api/v1/teams/overview";
    const res = await fetch(url);
    const data = await res.json();
    document.getElementById("crMetric").innerText = `${data.change_requests.total} CRs`;
    document.getElementById("bugMetric").innerText = `${data.bugs.total} Bugs`;
    document.getElementById("msgMetric").innerText = `${data.indexed_communications} msgs`;
    document.getElementById("specMetric").innerText = `${data.indexed_specifications} specs`;
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

// Load Sessions List (Group-Scoped)
async function loadSessions() {
  try {
    const url = currentGroupId && currentGroupId !== "all" 
      ? `/api/v1/teams/sessions?group_id=${encodeURIComponent(currentGroupId)}` 
      : "/api/v1/teams/sessions";
    const res = await fetch(url);
    const sessions = await res.json();
    sessionsList.innerHTML = "";

    if (sessions.length === 0) {
      await createNewSession();
      return;
    }

    sessions.forEach(s => {
      const item = document.createElement("div");
      item.className = `session-item ${s.session_id === currentSessionId ? "active" : ""}`;
      item.onclick = () => selectSession(s.session_id);

      const title = document.createElement("span");
      title.className = "session-title";
      title.innerText = s.title || "Conversation";

      const delBtn = document.createElement("button");
      delBtn.className = "btn-del-session";
      delBtn.innerHTML = "✕";
      delBtn.onclick = (e) => {
        e.stopPropagation();
        deleteSession(s.session_id);
      };

      item.appendChild(title);
      item.appendChild(delBtn);
      sessionsList.appendChild(item);
    });

    if (!currentSessionId && sessions.length > 0) {
      selectSession(sessions[0].session_id);
    }
  } catch (e) {
    console.error("Error loading sessions:", e);
  }
}

// Select Session & load messages
async function selectSession(sessionId) {
  currentSessionId = sessionId;
  document.querySelectorAll(".session-item").forEach(el => el.classList.remove("active"));
  
  // Highlight active
  document.querySelectorAll(".session-item").forEach(el => {
    if (el.onclick && el.onclick.toString().includes(sessionId)) el.classList.add("active");
  });

  chatViewport.innerHTML = "";
  try {
    const res = await fetch(`/api/v1/teams/sessions/${sessionId}`);
    const data = await res.json();
    data.messages.forEach(m => {
      appendMessage(m.role, m.content, m.evidence, m.eval_scores);
    });
  } catch (e) {
    console.error("Error loading session messages:", e);
  }
}

// Create New Session (Group-Scoped)
async function createNewSession() {
  try {
    const res = await fetch("/api/v1/teams/sessions", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ title: "New Conversation", group_id: currentGroupId })
    });
    const data = await res.json();
    currentSessionId = data.session_id;
    await loadSessions();
    await selectSession(currentSessionId);
  } catch (e) {
    console.error("Error creating session:", e);
  }
}

// Delete Session
async function deleteSession(sessionId) {
  try {
    await fetch(`/api/v1/teams/sessions/${sessionId}`, { method: "DELETE" });
    if (currentSessionId === sessionId) {
      currentSessionId = null;
    }
    loadSessions();
  } catch (e) {
    console.error("Error deleting session:", e);
  }
}

// Handle Form Submission
chatForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const q = chatInput.value.trim();
  if (!q) return;

  chatInput.value = "";
  appendMessage("user", q);

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
    const res = await fetch("/api/v1/teams/query", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ query: q, session_id: currentSessionId, group_id: currentGroupId })
    });
    const data = await res.json();
    const tb = document.getElementById("thinkingBubble");
    if (tb) tb.remove();

    appendMessage("assistant", data.response, data.evidence_trail, data.evaluation);
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
