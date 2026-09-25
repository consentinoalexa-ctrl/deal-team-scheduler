/* Deal Team Scheduler task pane.
 * Deal teams are saved in Outlook roaming settings (stored in your own mailbox).
 * No backend, no Microsoft Graph, no external calls.
 */
(function () {
  "use strict";

  const SETTINGS_KEY = "dealTeams";
  const SIZE_WARNING = 28000; // roaming settings cap is about 32 KB

  const state = {
    teams: [],
    view: "list", // list | team | edit
    teamId: null,
    draft: null
  };

  // ---------- Helpers ----------

  function el(tag, attrs, children) {
    const node = document.createElement(tag);
    Object.entries(attrs || {}).forEach(([key, value]) => {
      if (value === undefined || value === null || value === false) return;
      if (key === "text") node.textContent = value;
      else if (key === "class") node.className = value;
      else if (key.startsWith("on")) node.addEventListener(key.slice(2), value);
      else if (value === true) node.setAttribute(key, "");
      else node.setAttribute(key, value);
    });
    (children || []).forEach((child) => {
      if (child) node.appendChild(typeof child === "string" ? document.createTextNode(child) : child);
    });
    return node;
  }

  function newId() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  }

  function initials(member) {
    const source = member.displayName && member.displayName !== member.email ? member.displayName : member.email;
    const parts = source.replace(/@.*/, "").split(/[\s._,-]+/).filter(Boolean);
    const letters = parts.length > 1 ? parts[0][0] + parts[parts.length - 1][0] : (parts[0] || "?").slice(0, 2);
    return letters.toUpperCase();
  }

  function isValidEmail(email) {
    return /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]{2,}$/.test(email);
  }

  // Accepts "a@x.com, b@x.com", one per line, or Outlook style "Doe, Jane <jane@x.com>; ..."
  function parseMembers(text) {
    const re = /<\s*([^<>\s]+@[^<>\s]+?)\s*>|([A-Z0-9._%+'-]+@[A-Z0-9.-]+\.[A-Z]{2,})/gi;
    const members = [];
    const seen = new Set();
    let last = 0;
    let match;
    while ((match = re.exec(text)) !== null) {
      const email = (match[1] || match[2]).trim().toLowerCase();
      let name = "";
      if (match[1]) {
        name = text.slice(last, match.index)
          .replace(/^[\s,;]+/, "")
          .replace(/["']/g, "")
          .replace(/\s+/g, " ")
          .trim();
      }
      last = re.lastIndex;
      if (!isValidEmail(email) || seen.has(email)) continue;
      seen.add(email);
      members.push({ displayName: name || email, email: email });
    }
    return members;
  }

  function currentTeam() {
    return state.teams.find((t) => t.id === state.teamId) || null;
  }

  function sortedTeams() {
    return state.teams.slice().sort((a, b) => a.name.localeCompare(b.name));
  }

  // ---------- Status ----------

  let statusTimer = null;
  function showStatus(message, isError) {
    const box = document.getElementById("status");
    box.textContent = message;
    box.className = "status" + (isError ? " error" : "");
    box.hidden = false;
    clearTimeout(statusTimer);
    if (!isError) statusTimer = setTimeout(() => { box.hidden = true; }, 6000);
  }
  function clearStatus() {
    document.getElementById("status").hidden = true;
  }

  // ---------- Storage ----------

  function loadTeams() {
    try {
      const raw = Office.context.roamingSettings.get(SETTINGS_KEY);
      const data = typeof raw === "string" ? JSON.parse(raw) : raw;
      return Array.isArray(data) ? data : [];
    } catch (e) {
      return [];
    }
  }

  function saveTeams() {
    const json = JSON.stringify(state.teams);
    return new Promise((resolve, reject) => {
      Office.context.roamingSettings.set(SETTINGS_KEY, json);
      Office.context.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          if (json.length > SIZE_WARNING) {
            showStatus("Saved. You are close to Outlook's storage limit for add-in settings, so consider removing teams for closed deals.", false);
          }
          resolve();
        } else {
          reject(result.error);
        }
      });
    });
  }

  // ---------- Meeting attendees ----------

  function meetingItem() {
    const item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item || item.itemType !== Office.MailboxEnums.ItemType.Appointment) return null;
    if (!item.requiredAttendees || typeof item.requiredAttendees.addAsync !== "function") return null;
    return item;
  }

  function getAttendees(field) {
    return new Promise((resolve, reject) => {
      field.getAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value || []) : reject(r.error)));
    });
  }

  function addToField(field, recipients) {
    return new Promise((resolve, reject) => {
      field.addAsync(recipients, (r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)));
    });
  }

  async function addTeamToMeeting(members, asOptional) {
    const item = meetingItem();
    if (!item) {
      showStatus("Open a new meeting or a meeting you organize, then use the Deal Teams button from there.", true);
      return;
    }
    if (members.length === 0) {
      showStatus("Select at least one person to add.", true);
      return;
    }
    try {
      const [required, optional] = await Promise.all([
        getAttendees(item.requiredAttendees),
        getAttendees(item.optionalAttendees)
      ]);
      const existing = new Set(required.concat(optional).map((r) => (r.emailAddress || "").toLowerCase()));
      const toAdd = members
        .filter((m) => !existing.has(m.email.toLowerCase()))
        .map((m) => ({ displayName: m.displayName, emailAddress: m.email }));
      const skipped = members.length - toAdd.length;

      if (toAdd.length === 0) {
        showStatus("Everyone selected is already on this invite.", false);
        return;
      }
      await addToField(asOptional ? item.optionalAttendees : item.requiredAttendees, toAdd);
      let message = `Added ${toAdd.length} ${toAdd.length === 1 ? "person" : "people"} as ${asOptional ? "optional" : "required"}.`;
      if (skipped > 0) message += ` Skipped ${skipped} already on the invite.`;
      message += " Open the Scheduling Assistant to see everyone's availability.";
      showStatus(message, false);
    } catch (err) {
      showStatus("Outlook could not update the attendees: " + (err && err.message ? err.message : "unknown error") + ".", true);
    }
  }

  // ---------- Views ----------

  function render() {
    const app = document.getElementById("app");
    app.replaceChildren();
    if (state.view === "team" && currentTeam()) app.appendChild(renderTeam(currentTeam()));
    else if (state.view === "edit" && state.draft) app.appendChild(renderEdit());
    else {
      state.view = "list";
      app.appendChild(renderList());
    }
  }

  function go(view, teamId) {
    state.view = view;
    if (teamId !== undefined) state.teamId = teamId;
    render();
    window.scrollTo(0, 0);
  }

  function renderList() {
    const wrap = el("div");
    const teams = sortedTeams();

    if (teams.length === 0) {
      wrap.appendChild(el("p", { class: "empty", text: "No deal teams yet. Create one for each deal, then add the whole team to a meeting in one click." }));
    } else {
      const list = el("ul", { class: "team-list" });
      teams.forEach((team) => {
        const faces = el("span", { class: "faces", "aria-hidden": "true" },
          team.members.slice(0, 3).map((m) => el("span", { class: "face", text: initials(m) })));
        const count = team.members.length;
        list.appendChild(el("li", {}, [
          el("button", { class: "team-row", type: "button", onclick: () => { clearStatus(); go("team", team.id); } }, [
            el("span", { class: "team-name", text: team.name }),
            faces,
            el("span", { class: "count", text: `${count} ${count === 1 ? "person" : "people"}` })
          ])
        ]));
      });
      wrap.appendChild(list);
    }

    wrap.appendChild(el("div", { class: "actions" }, [
      el("button", { class: "btn primary", type: "button", onclick: () => startEdit(null) }, ["New team"])
    ]));

    wrap.appendChild(renderBackup());
    return wrap;
  }

  function renderTeam(team) {
    const wrap = el("div");
    wrap.appendChild(el("button", { class: "link back", type: "button", onclick: () => go("list") }, ["All teams"]));
    wrap.appendChild(el("h2", { text: team.name }));

    if (team.members.length === 0) {
      wrap.appendChild(el("p", { class: "empty", text: "This team has no members yet. Edit the team to add people." }));
    } else {
      const list = el("ul", { class: "member-list" });
      team.members.forEach((m, i) => {
        const id = "m-" + i;
        list.appendChild(el("li", {}, [
          el("label", { for: id }, [
            el("input", { type: "checkbox", id: id, checked: true, "data-index": String(i) }),
            el("span", { class: "member-text" }, [
              el("span", { class: "member-name", text: m.displayName }),
              m.displayName !== m.email ? el("span", { class: "member-email", text: m.email }) : null
            ])
          ])
        ]));
      });
      wrap.appendChild(list);

      const selected = () => Array.from(list.querySelectorAll("input[type=checkbox]:checked"))
        .map((box) => team.members[Number(box.getAttribute("data-index"))]);

      wrap.appendChild(el("div", { class: "actions" }, [
        el("button", { class: "btn primary", type: "button", onclick: () => addTeamToMeeting(selected(), false) }, ["Add as required"]),
        el("button", { class: "btn", type: "button", onclick: () => addTeamToMeeting(selected(), true) }, ["Add as optional"])
      ]));

      if (!meetingItem()) {
        wrap.appendChild(el("p", { class: "hint", text: "To add a team, open this pane from a new meeting or a meeting you organize." }));
      }
    }

    wrap.appendChild(el("div", { class: "actions" }, [
      el("button", { class: "btn", type: "button", onclick: () => startEdit(team) }, ["Edit team"])
    ]));
    return wrap;
  }

  function startEdit(team) {
    clearStatus();
    state.draft = team
      ? { id: team.id, name: team.name, members: team.members.map((m) => Object.assign({}, m)), isNew: false }
      : { id: newId(), name: "", members: [], isNew: true };
    go("edit");
  }

  function renderEdit() {
    const draft = state.draft;
    const wrap = el("div");

    wrap.appendChild(el("button", {
      class: "link back", type: "button",
      onclick: () => { state.draft = null; go(draft.isNew ? "list" : "team"); }
    }, ["Cancel"]));
    wrap.appendChild(el("h2", { text: draft.isNew ? "New deal team" : "Edit deal team" }));

    const nameInput = el("input", { type: "text", id: "team-name", value: draft.name, placeholder: "Example: Victory Foam", maxlength: "80" });
    nameInput.addEventListener("input", () => { draft.name = nameInput.value; });
    wrap.appendChild(el("label", { class: "field", for: "team-name", text: "Team name" }));
    wrap.appendChild(nameInput);

    wrap.appendChild(el("label", { class: "field", text: `Members (${draft.members.length})` }));
    if (draft.members.length === 0) {
      wrap.appendChild(el("p", { class: "hint", text: "No members yet. Add people below." }));
    } else {
      const list = el("ul", { class: "member-list" });
      draft.members.forEach((m, i) => {
        list.appendChild(el("li", {}, [
          el("span", { class: "member-text", style: "flex:1" }, [
            el("span", { class: "member-name", text: m.displayName }),
            m.displayName !== m.email ? el("span", { class: "member-email", text: m.email }) : null
          ]),
          el("button", {
            class: "link remove", type: "button", "aria-label": "Remove " + m.displayName,
            onclick: () => { draft.members.splice(i, 1); render(); }
          }, ["Remove"])
        ]));
      });
      wrap.appendChild(list);
    }

    const addBox = el("textarea", {
      id: "add-members",
      placeholder: "Paste emails separated by commas, semicolons, or new lines. You can also paste straight from an Outlook To line."
    });
    wrap.appendChild(el("label", { class: "field", for: "add-members", text: "Add people" }));
    wrap.appendChild(addBox);
    wrap.appendChild(el("div", { class: "actions" }, [
      el("button", {
        class: "btn", type: "button",
        onclick: () => {
          const found = parseMembers(addBox.value);
          if (found.length === 0) {
            showStatus("No valid email addresses found. Check the format and try again.", true);
            return;
          }
          const existing = new Set(draft.members.map((m) => m.email));
          const fresh = found.filter((m) => !existing.has(m.email));
          draft.members = draft.members.concat(fresh);
          const dupes = found.length - fresh.length;
          showStatus(`Added ${fresh.length} to the list.` + (dupes ? ` ${dupes} already on the team.` : "") + " Save the team to keep changes.", false);
          render();
        }
      }, ["Add to list"])
    ]));

    const saveButton = el("button", {
      class: "btn primary", type: "button",
      onclick: async () => {
        const name = draft.name.trim();
        if (!name) { showStatus("Give the team a name before saving.", true); nameInput.focus(); return; }
        const clash = state.teams.find((t) => t.id !== draft.id && t.name.toLowerCase() === name.toLowerCase());
        if (clash) { showStatus(`You already have a team called "${clash.name}". Pick a different name.`, true); return; }

        const saved = { id: draft.id, name: name, members: draft.members };
        const snapshot = state.teams.slice();
        const index = state.teams.findIndex((t) => t.id === draft.id);
        if (index >= 0) state.teams[index] = saved; else state.teams.push(saved);

        saveButton.disabled = true;
        try {
          await saveTeams();
          state.draft = null;
          go("team", saved.id);
          showStatus(`Saved ${name}.`, false);
        } catch (err) {
          state.teams = snapshot;
          saveButton.disabled = false;
          showStatus("Outlook could not save the team: " + (err && err.message ? err.message : "unknown error") + ".", true);
        }
      }
    }, ["Save team"]);

    const actions = [saveButton];
    if (!draft.isNew) {
      let armed = false;
      const deleteButton = el("button", {
        class: "btn danger", type: "button",
        onclick: async () => {
          if (!armed) {
            armed = true;
            deleteButton.textContent = "Click again to delete";
            deleteButton.classList.add("armed");
            return;
          }
          const snapshot = state.teams.slice();
          const name = draft.name;
          state.teams = state.teams.filter((t) => t.id !== draft.id);
          try {
            await saveTeams();
            state.draft = null;
            go("list", null);
            showStatus(`Deleted ${name}.`, false);
          } catch (err) {
            state.teams = snapshot;
            showStatus("Outlook could not delete the team: " + (err && err.message ? err.message : "unknown error") + ".", true);
          }
        }
      }, ["Delete team"]);
      actions.push(deleteButton);
    }
    wrap.appendChild(el("div", { class: "actions" }, actions));
    return wrap;
  }

  function renderBackup() {
    const box = el("textarea", { class: "code", id: "backup-json", placeholder: "Paste a backup here to import it." });

    const copyButton = el("button", {
      class: "btn", type: "button",
      onclick: async () => {
        box.value = JSON.stringify({ app: "deal-team-scheduler", version: 1, teams: state.teams }, null, 2);
        box.select();
        try {
          await navigator.clipboard.writeText(box.value);
          showStatus("Backup copied to your clipboard. Paste it somewhere safe, like a note or email to yourself.", false);
        } catch (e) {
          try {
            document.execCommand("copy");
            showStatus("Backup copied to your clipboard.", false);
          } catch (e2) {
            showStatus("Backup is shown in the box below. Select it and copy it manually.", false);
          }
        }
      }
    }, ["Export backup"]);

    const importButton = el("button", {
      class: "btn", type: "button",
      onclick: async () => {
        let incoming;
        try {
          const data = JSON.parse(box.value);
          incoming = Array.isArray(data) ? data : data.teams;
          if (!Array.isArray(incoming)) throw new Error("no teams");
        } catch (e) {
          showStatus("That backup could not be read. Paste the full text from Export backup.", true);
          return;
        }
        const clean = incoming
          .filter((t) => t && typeof t.name === "string" && Array.isArray(t.members))
          .map((t) => ({
            id: newId(),
            name: t.name.trim(),
            members: t.members
              .filter((m) => m && typeof m.email === "string" && isValidEmail(m.email.trim().toLowerCase()))
              .map((m) => ({ displayName: (m.displayName || m.email).trim(), email: m.email.trim().toLowerCase() }))
          }))
          .filter((t) => t.name);

        const snapshot = state.teams.slice();
        let replaced = 0;
        clean.forEach((t) => {
          const i = state.teams.findIndex((x) => x.name.toLowerCase() === t.name.toLowerCase());
          if (i >= 0) { t.id = state.teams[i].id; state.teams[i] = t; replaced++; } else state.teams.push(t);
        });
        try {
          await saveTeams();
          render();
          showStatus(`Imported ${clean.length} ${clean.length === 1 ? "team" : "teams"}` + (replaced ? `, replacing ${replaced} with the same name.` : "."), false);
        } catch (err) {
          state.teams = snapshot;
          showStatus("Outlook could not save the imported teams: " + (err && err.message ? err.message : "unknown error") + ".", true);
        }
      }
    }, ["Import backup"]);

    return el("details", { class: "backup" }, [
      el("summary", { text: "Back up or move your teams" }),
      el("p", { class: "hint", text: "Teams are saved in your mailbox and follow you across Outlook on the web, new Outlook, and classic Outlook. Export a backup before reinstalling the add-in." }),
      el("div", { class: "actions" }, [copyButton, importButton]),
      box
    ]);
  }

  // ---------- Start ----------

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) {
      document.getElementById("app").replaceChildren(
        el("p", { class: "empty", text: "This add-in runs inside Outlook. Upload the manifest to Outlook to use it." })
      );
      return;
    }
    state.teams = loadTeams();
    render();
  });

  // Exposed for local testing only.
  if (typeof window !== "undefined") window.__dealTeamParse = parseMembers;
})();
