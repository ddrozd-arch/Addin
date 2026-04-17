import { getAllowedDomains } from "../common/whitelist.js";

Office.onReady(init);

// ---------- TABS ----------
function showTab(tabId) {
  document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
  document.getElementById(tabId).classList.add("active");
}
window.showTab = showTab;

// ---------- HELPERS ----------
function getDomain(email) {
  return email.split("@")[1]?.toLowerCase();
}

function groupByDomain(list) {
  const map = {};

  list.forEach(email => {
    const domain = getDomain(email) || "unknown";
    if (!map[domain]) map[domain] = [];
    map[domain].push(email);
  });

  return map;
}

function getTrustedRecipients() {
  const trusted = Office.context.roamingSettings.get("trustedRecipients");
  return trusted ? JSON.parse(trusted) : [];
}

function saveTrustedRecipients(list) {
  Office.context.roamingSettings.set("trustedRecipients", JSON.stringify(list));
  Office.context.roamingSettings.saveAsync();
}

function getAllRecipients() {
  const item = Office.context.mailbox.item;

  return new Promise((resolve) => {
    item.to.getAsync(toRes => {
      item.cc.getAsync(ccRes => {
        item.bcc.getAsync(bccRes => {
          resolve([
            ...(toRes.value || []),
            ...(ccRes.value || []),
            ...(bccRes.value || [])
          ].map(r => r.emailAddress.toLowerCase()));
        });
      });
    });
  });
}

// ---------- STATE ----------
let externalCache = [];
let trustedCache = [];

// ---------- INIT ----------
async function init() {
  const allowedDomains = await getAllowedDomains();
  const recipients = await getAllRecipients();
  const trusted = getTrustedRecipients();

  trustedCache = trusted;

  // DOMAINS
  const allowedList = document.getElementById("allowedList");
  allowedDomains.forEach(d => {
    const li = document.createElement("li");
    li.innerText = d;
    allowedList.appendChild(li);
  });

  // EXTERNAL
  externalCache = recipients.filter(e => {
    const domain = getDomain(e);
    return !allowedDomains.includes(domain) && !trusted.includes(e);
  });

  renderExternalGrouped(externalCache);

  // TRUSTED
  renderTrustedGrouped(trustedCache);

  document.getElementById("searchBox").addEventListener("input", (e) => {
    const val = e.target.value.toLowerCase();
    renderTrustedGrouped(trustedCache.filter(x => x.includes(val)));
  });

  document.getElementById("saveBtn").onclick = () => {
    const selected = [...document.querySelectorAll("#externalList input:checked")]
      .map(cb => cb.value);

    const updated = [...new Set([...trustedCache, ...selected])];
    saveTrustedRecipients(updated);
    location.reload();
  };
}

// ---------- RENDER ----------
function renderExternalGrouped(list) {
  const container = document.getElementById("externalList");
  container.innerHTML = "";

  const grouped = groupByDomain(list);

  Object.keys(grouped).forEach(domain => {
    const groupDiv = document.createElement("div");
    groupDiv.className = "group";

    const title = document.createElement("h4");
    title.innerText = `${domain} (${grouped[domain].length})`;

    groupDiv.appendChild(title);

    grouped[domain].forEach(email => {
      const row = document.createElement("div");
      row.className = "row";

      row.innerHTML = `
        <label>
          <input type="checkbox" value="${email}" />
          ${email}
        </label>
      `;

      groupDiv.appendChild(row);
    });

    container.appendChild(groupDiv);
  });
}

function renderTrustedGrouped(list) {
  const container = document.getElementById("trustedList");
  container.innerHTML = "";

  const grouped = groupByDomain(list);

  Object.keys(grouped).forEach(domain => {
    const groupDiv = document.createElement("div");
    groupDiv.className = "group";

    const title = document.createElement("h4");
    title.innerText = `${domain} (${grouped[domain].length})`;

    groupDiv.appendChild(title);

    grouped[domain].forEach(email => {
      const row = document.createElement("div");
      row.className = "row";

      const label = document.createElement("span");
      label.innerText = email;

      const remove = document.createElement("span");
      remove.innerText = "Usuń";
      remove.className = "danger";

      remove.onclick = () => {
        trustedCache = trustedCache.filter(e => e !== email);
        saveTrustedRecipients(trustedCache);
        renderTrustedGrouped(trustedCache);
      };

      row.appendChild(label);
      row.appendChild(remove);

      groupDiv.appendChild(row);
    });

    container.appendChild(groupDiv);
  });
}
