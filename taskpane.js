/*
 * File on Send — folder picker task pane.
 *
 * Choosing a folder tags the message with a category "File: <folder path>".
 * After the message lands in Sent Items, the Power Automate "File sent mail" flow
 * reads that category, moves the message to the folder (creating it if needed),
 * and removes the tag.
 *
 * The folder list lives in the add-in's roaming settings (stored in your mailbox,
 * so it follows you to every Outlook you use). Refresh it by pasting the list that
 * the "Export folder list" flow emails you.
 */
(function () {
  "use strict";

  var FILE_PREFIX = "File: ";
  var CATEGORY_COLOR = "Preset7"; // blue
  var MAX_RECENTS = 8;
  var MAX_RESULTS = 150;

  var store;          // adapter: Outlook or demo
  var folders = [];   // ["Inbox/Clients/Acme", ...]
  var recents = [];
  var current = null; // chosen path or null
  var activeIndex = -1;
  var rows = [];      // currently rendered selectable rows

  // ---------- helpers ----------
  function $(id) { return document.getElementById(id); }

  function normalizePath(p) {
    return String(p || "")
      .split("/")
      .map(function (s) { return s.trim(); })
      .filter(Boolean)
      .join("/");
  }

  function parseList(text) {
    var seen = {};
    var out = [];
    String(text || "").split(/\r?\n/).forEach(function (line) {
      var p = normalizePath(line);
      var key = p.toLowerCase();
      if (p && !seen[key]) { seen[key] = true; out.push(p); }
    });
    out.sort(function (a, b) { return a.localeCompare(b, undefined, { sensitivity: "base" }); });
    return out;
  }

  function leaf(p) { var i = p.lastIndexOf("/"); return i < 0 ? p : p.slice(i + 1); }
  function parent(p) { var i = p.lastIndexOf("/"); return i < 0 ? "" : p.slice(0, i); }

  function setStatus(msg, kind) {
    var el = $("status");
    el.textContent = msg || "";
    el.className = "status" + (kind ? " " + kind : "");
  }

  // ---------- storage adapters ----------
  function outlookStore() {
    var mb = Office.context.mailbox;
    var settings = Office.context.roamingSettings;
    function cb(resolve, reject) {
      return function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
        else reject(r.error);
      };
    }
    return {
      demo: false,
      getSetting: function (k) { return settings.get(k); },
      setSettings: function (obj) {
        Object.keys(obj).forEach(function (k) { settings.set(k, obj[k]); });
        return new Promise(function (res, rej) { settings.saveAsync(cb(res, rej)); });
      },
      getItemCategories: function () {
        return new Promise(function (res, rej) { mb.item.categories.getAsync(cb(res, rej)); })
          .then(function (v) { return (v || []).map(function (c) { return c.displayName; }); });
      },
      removeItemCategories: function (names) {
        if (!names.length) return Promise.resolve();
        return new Promise(function (res, rej) { mb.item.categories.removeAsync(names, cb(res, rej)); });
      },
      addItemCategory: function (name) {
        // A category must exist in the mailbox's master list before it can be applied.
        return new Promise(function (res, rej) { mb.masterCategories.getAsync(cb(res, rej)); })
          .then(function (master) {
            var exists = (master || []).some(function (c) {
              return c.displayName.toLowerCase() === name.toLowerCase();
            });
            if (exists) return;
            return new Promise(function (res, rej) {
              mb.masterCategories.addAsync(
                [{ displayName: name, color: Office.MailboxEnums.CategoryColor[CATEGORY_COLOR] }],
                function (r) {
                  // Ignore "already exists" races.
                  if (r.status === Office.AsyncResultStatus.Succeeded ||
                      (r.error && /Duplicate/i.test(r.error.name || r.error.code || ""))) res();
                  else rej(r.error);
                });
            });
          })
          .then(function () {
            return new Promise(function (res, rej) { mb.item.categories.addAsync([name], cb(res, rej)); });
          });
      }
    };
  }

  function demoStore() {
    // Lets the pane be opened in a normal browser for testing. Nothing touches Outlook.
    var mem = {};
    var itemCats = [];
    function ls(k, v) {
      try {
        if (v === undefined) { var s = localStorage.getItem("fos-" + k); return s ? JSON.parse(s) : undefined; }
        localStorage.setItem("fos-" + k, JSON.stringify(v));
      } catch (e) { if (v === undefined) return mem[k]; mem[k] = v; }
    }
    if (!ls("folders")) {
      ls("folders", [
        "Inbox", "Inbox/Students", "Inbox/Students/Advising", "Archive",
        "Courses/CHEG 101", "Courses/CHEG 101/Grading", "Courses/CHEG 345",
        "Research/Para-aramid", "Research/Spandex", "Research/Polymer recycling",
        "Admin/Department", "Admin/Travel", "Committees/Curriculum"
      ].join("\n"));
    }
    return {
      demo: true,
      getSetting: function (k) { return ls(k); },
      setSettings: function (obj) { Object.keys(obj).forEach(function (k) { ls(k, obj[k]); }); return Promise.resolve(); },
      getItemCategories: function () { return Promise.resolve(itemCats.slice()); },
      removeItemCategories: function (names) {
        itemCats = itemCats.filter(function (c) { return names.indexOf(c) < 0; });
        return Promise.resolve();
      },
      addItemCategory: function (name) { itemCats.push(name); return Promise.resolve(); }
    };
  }

  // ---------- state ----------
  function loadSettings() {
    folders = parseList(store.getSetting("folders") || "");
    recents = (store.getSetting("recents") || []).filter(Boolean);
    var updated = store.getSetting("foldersUpdated");
    $("listInfo").textContent = folders.length
      ? folders.length + " folders" + (updated ? " · updated " + new Date(updated).toLocaleDateString() : "")
      : "No folder list yet";
  }

  function readCurrent() {
    return store.getItemCategories().then(function (names) {
      var c = names.filter(function (n) { return n.indexOf(FILE_PREFIX) === 0; })[0];
      current = c ? c.slice(FILE_PREFIX.length) : null;
      renderCurrent();
    });
  }

  function renderCurrent() {
    var box = $("current");
    if (current) {
      box.className = "current chosen";
      $("currentLeaf").textContent = leaf(current);
      $("currentPath").textContent = parent(current) || "(top level)";
      $("clearBtn").hidden = false;
    } else {
      box.className = "current";
      $("currentLeaf").textContent = "Not chosen yet";
      $("currentPath").textContent = "Message will stay in Sent Items";
      $("clearBtn").hidden = true;
    }
  }

  // ---------- choosing ----------
  function choose(path, isNew) {
    path = normalizePath(path);
    if (!path) return;
    setStatus("Saving…");
    var name = FILE_PREFIX + path;
    store.getItemCategories()
      .then(function (names) {
        var old = names.filter(function (n) { return n.indexOf(FILE_PREFIX) === 0 && n !== name; });
        return store.removeItemCategories(old);
      })
      .then(function () { return store.addItemCategory(name); })
      .then(function () {
        current = path;
        recents = [path].concat(recents.filter(function (r) { return r.toLowerCase() !== path.toLowerCase(); }))
          .slice(0, MAX_RECENTS);
        var upd = { recents: recents };
        if (isNew && !folders.some(function (f) { return f.toLowerCase() === path.toLowerCase(); })) {
          folders = parseList(folders.concat([path]).join("\n"));
          upd.folders = folders.join("\n");
        }
        return store.setSettings(upd);
      })
      .then(function () {
        loadSettings();
        renderCurrent();
        $("search").value = "";
        renderList();
        setStatus(isNew
          ? "Will file in “" + path + "” — the folder will be created when the message is filed."
          : "Will file in “" + leaf(path) + "”. You can send now.", "ok");
      })
      .catch(function (e) {
        setStatus("Couldn't set the folder: " + (e && (e.message || e.name) || e), "err");
      });
  }

  function clearChoice() {
    store.getItemCategories()
      .then(function (names) {
        return store.removeItemCategories(names.filter(function (n) { return n.indexOf(FILE_PREFIX) === 0; }));
      })
      .then(function () { current = null; renderCurrent(); setStatus("Cleared. The message will stay in Sent Items."); })
      .catch(function (e) { setStatus("Couldn't clear: " + (e && e.message || e), "err"); });
  }

  // ---------- list rendering ----------
  function matches(query) {
    var terms = query.toLowerCase().split(/[\s/]+/).filter(Boolean);
    if (!terms.length) return [];
    var scored = [];
    folders.forEach(function (f) {
      var lower = f.toLowerCase();
      for (var i = 0; i < terms.length; i++) if (lower.indexOf(terms[i]) < 0) return;
      var lf = leaf(f).toLowerCase();
      var score = 0;
      if (lf === terms[terms.length - 1]) score += 100;
      else if (lf.indexOf(terms[terms.length - 1]) === 0) score += 50;
      else if (lf.indexOf(terms[terms.length - 1]) >= 0) score += 20;
      if (recents.some(function (r) { return r.toLowerCase() === lower; })) score += 30;
      score -= f.split("/").length; // prefer shallower
      scored.push({ f: f, s: score });
    });
    scored.sort(function (a, b) { return b.s - a.s || a.f.localeCompare(b.f); });
    return scored.slice(0, MAX_RESULTS).map(function (x) { return x.f; });
  }

  function rowEl(path, opts) {
    opts = opts || {};
    var li = document.createElement("li");
    li.className = "row" + (opts.isNew ? " new" : "");
    li.setAttribute("role", "option");
    var a = document.createElement("div");
    a.className = "leaf";
    a.textContent = opts.isNew ? "Create and file in “" + leaf(path) + "”" : leaf(path);
    var b = document.createElement("div");
    b.className = "path";
    b.textContent = opts.isNew ? (path.indexOf("/") < 0 ? "New top-level folder" : "inside " + parent(path)) : (parent(path) || "top level");
    li.appendChild(a); li.appendChild(b);
    li.addEventListener("click", function () { choose(path, !!opts.isNew); });
    return li;
  }

  function renderList() {
    var q = $("search").value.trim();
    var ul = $("list");
    ul.innerHTML = "";
    rows = [];
    activeIndex = -1;
    var heading = $("listHeading");

    if (!q) {
      if (recents.length) {
        heading.textContent = "Recent";
        recents.forEach(function (p) { var el = rowEl(p); ul.appendChild(el); rows.push(el); });
      } else if (folders.length) {
        heading.textContent = "Type to search " + folders.length + " folders";
      } else {
        heading.textContent = "Add your folder list to get started (Update list, below)";
      }
      return;
    }

    var found = matches(q);
    heading.textContent = found.length ? found.length + (found.length === MAX_RESULTS ? "+" : "") + " matching" : "No matching folders";
    found.forEach(function (p) { var el = rowEl(p); ul.appendChild(el); rows.push(el); });

    // Offer to create a folder when the typed path isn't an existing folder.
    var typed = normalizePath(q);
    var exact = folders.some(function (f) { return f.toLowerCase() === typed.toLowerCase(); });
    if (typed && !exact && q.indexOf("/") >= 0) {
      var el = rowEl(typed, { isNew: true });
      ul.appendChild(el); rows.push(el);
      if (!found.length) heading.textContent = "Not in your folder list yet";
    } else if (typed && !exact && !found.length) {
      $("listHeading").textContent = "No matching folders — type a full path like Courses/CHEG 345/Exams to create one";
    }
    if (rows.length) setActive(0);
  }

  function setActive(i) {
    if (activeIndex >= 0 && rows[activeIndex]) rows[activeIndex].classList.remove("active");
    activeIndex = i;
    if (rows[i]) { rows[i].classList.add("active"); rows[i].scrollIntoView({ block: "nearest" }); }
  }

  // ---------- settings panel ----------
  function openSettings() {
    $("settings").hidden = false;
    $("folderText").value = folders.join("\n");
    $("folderText").focus();
  }

  function saveFolderList() {
    var list = parseList($("folderText").value);
    store.setSettings({ folders: list.join("\n"), foldersUpdated: new Date().toISOString() })
      .then(function () {
        loadSettings();
        $("settings").hidden = true;
        renderList();
        setStatus("Folder list saved (" + list.length + " folders).", "ok");
      })
      .catch(function (e) {
        setStatus("Couldn't save the list: " + (e && e.message || e) +
          " (the list may be too large; remove folders you never file to).", "err");
      });
  }

  // ---------- wiring ----------
  function wire() {
    $("search").addEventListener("input", function () { setStatus(""); renderList(); });
    $("search").addEventListener("keydown", function (e) {
      if (e.key === "ArrowDown") { e.preventDefault(); if (rows.length) setActive(Math.min(rows.length - 1, activeIndex + 1)); }
      else if (e.key === "ArrowUp") { e.preventDefault(); if (rows.length) setActive(Math.max(0, activeIndex - 1)); }
      else if (e.key === "Enter") { e.preventDefault(); if (rows[activeIndex]) rows[activeIndex].click(); }
      else if (e.key === "Escape") { $("search").value = ""; renderList(); }
    });
    $("clearBtn").addEventListener("click", clearChoice);
    $("updateBtn").addEventListener("click", openSettings);
    $("cancelBtn").addEventListener("click", function () { $("settings").hidden = true; });
    $("saveBtn").addEventListener("click", saveFolderList);
  }

  function start(inOutlook) {
    store = inOutlook ? outlookStore() : demoStore();
    if (store.demo) $("demoBanner").hidden = false;
    wire();
    loadSettings();
    renderList();
    readCurrent().catch(function () { renderCurrent(); });
    $("search").focus();
  }

  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      start(!!(info && info.host === Office.HostType.Outlook && Office.context && Office.context.mailbox));
    });
  } else {
    document.addEventListener("DOMContentLoaded", function () { start(false); });
  }
})();
