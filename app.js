(function () {
  "use strict";

  var state = {
    participants: null, // [{name, dept, isFK}]
    groups: null
  };

  var els = {
    fileInput: document.getElementById("file-input"),
    fileDropLabel: document.getElementById("file-drop-label"),
    fileDrop: document.querySelector(".file-drop"),
    fileStatus: document.getElementById("file-status"),
    btnCreate: document.getElementById("btn-create"),
    sectionReroll: document.getElementById("section-reroll"),
    btnReroll: document.getElementById("btn-reroll"),
    sectionQuality: document.getElementById("section-quality"),
    qualityDept: document.getElementById("quality-dept"),
    qualityFk: document.getElementById("quality-fk"),
    qualitySize: document.getElementById("quality-size"),
    sectionResult: document.getElementById("section-result"),
    deptLegend: document.getElementById("dept-legend"),
    groupsContainer: document.getElementById("groups-container"),
    btnCopy: document.getElementById("btn-copy"),
    copyStatus: document.getElementById("copy-status"),
    errorMessage: document.getElementById("error-message")
  };

  var DEPT_COLORS = ["dept-a", "dept-b", "dept-c", "dept-d", "dept-e", "dept-f"];
  var deptColorMap = {};

  function showError(msg) {
    els.errorMessage.textContent = msg;
    els.errorMessage.hidden = !msg;
  }

  function clearError() {
    showError("");
  }

  // ---------- File parsing ----------

  els.fileInput.addEventListener("change", function (e) {
    var file = e.target.files && e.target.files[0];
    clearError();
    state.participants = null;
    els.btnCreate.disabled = true;
    els.sectionReroll.hidden = true;
    els.sectionQuality.hidden = true;
    els.sectionResult.hidden = true;

    if (!file) {
      els.fileDropLabel.textContent = "📄 Excel-Datei auswählen (.xlsx)";
      els.fileDrop.classList.remove("has-file");
      return;
    }

    els.fileDropLabel.textContent = "📄 " + file.name;
    els.fileDrop.classList.add("has-file");
    els.fileStatus.textContent = "Datei wird gelesen …";

    var reader = new FileReader();
    reader.onload = function (evt) {
      try {
        var data = new Uint8Array(evt.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        var sheetName = workbook.SheetNames[0];
        var sheet = workbook.Sheets[sheetName];
        var rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        var participants = parseParticipants(rows);

        if (participants.length < 3) {
          throw new Error("Es werden mindestens 3 Teilnehmende benötigt (gefunden: " + participants.length + ").");
        }

        state.participants = participants;
        var fkCount = participants.filter(function (p) { return p.isFK; }).length;
        var deptCount = {};
        participants.forEach(function (p) { deptCount[p.dept] = (deptCount[p.dept] || 0) + 1; });
        var deptSummary = Object.keys(deptCount).map(function (d) { return d + ": " + deptCount[d]; }).join(", ");

        els.fileStatus.textContent =
          participants.length + " Teilnehmende geladen · " + fkCount + " FK · " + deptSummary;
        els.btnCreate.disabled = false;
      } catch (err) {
        console.error(err);
        showError("Fehler beim Einlesen der Datei: " + err.message);
        els.fileStatus.textContent = "";
      }
    };
    reader.onerror = function () {
      showError("Die Datei konnte nicht gelesen werden.");
    };
    reader.readAsArrayBuffer(file);
  });

  function normalizeHeader(h) {
    return String(h).trim().toLowerCase();
  }

  function parseParticipants(rows) {
    if (!rows.length) throw new Error("Die Datei enthält keine Zeilen.");

    var headers = Object.keys(rows[0]);
    var nameKey = null;
    var deptKey = null;

    headers.forEach(function (h) {
      var norm = normalizeHeader(h);
      if (!nameKey && (norm === "name" || norm.indexOf("name") !== -1)) nameKey = h;
      if (!deptKey && (norm === "abteilung" || norm.indexOf("abteilung") !== -1 || norm.indexOf("department") !== -1)) deptKey = h;
    });

    if (!nameKey) throw new Error('Spalte "Name" wurde nicht gefunden.');
    if (!deptKey) throw new Error('Spalte "Abteilung" wurde nicht gefunden.');

    var participants = [];
    rows.forEach(function (row) {
      var rawName = String(row[nameKey] || "").trim();
      var rawDept = String(row[deptKey] || "").trim();
      if (!rawName) return;

      var isFK = /\bFK\b/i.test(rawName) || /\(FK\)/i.test(rawName);
      var cleanName = rawName
        .replace(/\(\s*FK\s*\)/gi, "")
        .replace(/\bFK\b/gi, "")
        .replace(/\s{2,}/g, " ")
        .trim();

      participants.push({
        name: cleanName || rawName,
        dept: rawDept || "Unbekannt",
        isFK: isFK
      });
    });

    return participants;
  }

  // ---------- Group size calculation ----------

  function computeGroupSizes(n) {
    var kMin = Math.ceil(n / 6);
    var kMax = Math.floor(n / 3);
    if (kMin > kMax) return null;

    // Prefer the smallest number of groups whose sizes stay mostly within the
    // preferred 4-5 range, tolerating at most one group outside it (a single
    // 3 or 6) before adding another group.
    var fallback = null;
    for (var k = kMin; k <= kMax; k++) {
      var base = Math.floor(n / k);
      var rem = n % k;
      var sizes = [];
      for (var i = 0; i < k; i++) sizes.push(base + (i < rem ? 1 : 0));

      var penalty = sizes.filter(function (s) { return s < 4 || s > 5; }).length;
      if (penalty <= 1) return sizes;
      if (!fallback || penalty < fallback.penalty) fallback = { sizes: sizes, penalty: penalty };
    }
    return fallback.sizes;
  }

  // ---------- Randomization helpers ----------

  function shuffle(arr) {
    var a = arr.slice();
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var tmp = a[i];
      a[i] = a[j];
      a[j] = tmp;
    }
    return a;
  }

  // Interleave a list round-robin across buckets formed by one or more
  // categorical attributes (e.g. department, FK-status), each shuffled
  // internally. All attributes are treated equally as mixing criteria.
  function interleaveByAttributes(list, keyFns) {
    var buckets = {};
    list.forEach(function (p) {
      var key = keyFns.map(function (fn) { return fn(p); }).join("||");
      if (!buckets[key]) buckets[key] = [];
      buckets[key].push(p);
    });
    var bucketKeys = shuffle(Object.keys(buckets));
    var queues = bucketKeys.map(function (k) { return shuffle(buckets[k]); });

    var result = [];
    var remaining = true;
    while (remaining) {
      remaining = false;
      for (var i = 0; i < queues.length; i++) {
        if (queues[i].length) {
          result.push(queues[i].shift());
          remaining = true;
        }
      }
    }
    return result;
  }

  // ---------- Group assignment ----------

  function assignGroups(participants) {
    var sizes = computeGroupSizes(participants.length);
    if (!sizes) throw new Error("Für " + participants.length + " Teilnehmende kann keine gültige Gruppengröße (3–6) gebildet werden.");

    var groupCount = sizes.length;
    var groups = sizes.map(function () { return []; });
    var capacities = sizes.slice();

    var ordered = interleaveByAttributes(participants, [
      function (p) { return p.dept; },
      function (p) { return p.isFK ? "FK" : "no-FK"; }
    ]);

    var order = shuffle(Array.from({ length: groupCount }, function (_, i) { return i; }));
    var cursor = 0;

    function place(person) {
      var attempts = 0;
      while (attempts < groupCount) {
        var gi = order[cursor % groupCount];
        cursor++;
        attempts++;
        if (groups[gi].length < capacities[gi]) {
          groups[gi].push(person);
          return;
        }
      }
      // Fallback: force into first group with space (should not normally happen)
      for (var i = 0; i < groupCount; i++) {
        if (groups[i].length < capacities[i]) {
          groups[i].push(person);
          return;
        }
      }
    }

    ordered.forEach(place);

    return groups;
  }

  // ---------- Quality evaluation ----------

  function evaluateQuality(groups) {
    var allDepts = {};
    var totalFK = 0;
    var totalPeople = 0;

    groups.forEach(function (g) {
      g.forEach(function (p) {
        allDepts[p.dept] = (allDepts[p.dept] || 0) + 1;
        if (p.isFK) totalFK++;
        totalPeople++;
      });
    });

    var deptKeys = Object.keys(allDepts);
    var globalRatio = {};
    deptKeys.forEach(function (d) { globalRatio[d] = allDepts[d] / totalPeople; });

    // Department mixing: average absolute deviation of per-group dept ratio vs global ratio.
    var deptRating;
    if (deptKeys.length < 2) {
      deptRating = "optimal";
    } else {
      var totalDeviation = 0;
      groups.forEach(function (g) {
        var counts = {};
        deptKeys.forEach(function (d) { counts[d] = 0; });
        g.forEach(function (p) { counts[p.dept]++; });
        var dev = 0;
        deptKeys.forEach(function (d) {
          var groupRatio = counts[d] / g.length;
          dev += Math.abs(groupRatio - globalRatio[d]);
        });
        totalDeviation += dev;
      });
      var avgDeviation = totalDeviation / groups.length;
      if (avgDeviation < 0.15) deptRating = "optimal";
      else if (avgDeviation < 0.4) deptRating = "sehr-gut";
      else deptRating = "gut";
    }

    // FK distribution: spread between max and min FK count per group.
    var fkRating;
    if (totalFK === 0) {
      fkRating = "optimal";
    } else {
      var fkCounts = groups.map(function (g) { return g.filter(function (p) { return p.isFK; }).length; });
      var spread = Math.max.apply(null, fkCounts) - Math.min.apply(null, fkCounts);
      if (spread <= 0) fkRating = "optimal";
      else if (spread === 1) fkRating = "sehr-gut";
      else fkRating = "gut";
    }

    var sizeValid = groups.every(function (g) { return g.length >= 3 && g.length <= 6; });

    return { deptRating: deptRating, fkRating: fkRating, sizeValid: sizeValid };
  }

  var RATING_LABEL = { optimal: "Optimal", "sehr-gut": "Sehr gut", gut: "Gut" };

  function applyBadge(el, ratingKey) {
    el.textContent = RATING_LABEL[ratingKey] || ratingKey;
    el.className = "badge " + ratingKey;
  }

  // ---------- Rendering ----------

  function buildDeptColorMap(groups) {
    var depts = {};
    groups.forEach(function (g) {
      g.forEach(function (p) { depts[p.dept] = true; });
    });
    var sortedDepts = Object.keys(depts).sort(function (a, b) { return a.localeCompare(b, "de"); });
    var map = {};
    sortedDepts.forEach(function (d, i) {
      map[d] = DEPT_COLORS[i % DEPT_COLORS.length];
    });
    return map;
  }

  function renderLegend(colorMap) {
    els.deptLegend.innerHTML = "";
    Object.keys(colorMap).forEach(function (dept) {
      var item = document.createElement("span");
      item.className = "dept-legend-item";

      var dot = document.createElement("span");
      dot.className = "dept-dot " + colorMap[dept];
      item.appendChild(dot);

      var label = document.createElement("span");
      label.textContent = dept;
      item.appendChild(label);

      els.deptLegend.appendChild(item);
    });
  }

  function renderGroups(groups) {
    deptColorMap = buildDeptColorMap(groups);
    renderLegend(deptColorMap);
    els.groupsContainer.innerHTML = "";

    groups.forEach(function (group, idx) {
      var card = document.createElement("div");
      card.className = "group-card";

      var heading = document.createElement("h3");
      var titleSpan = document.createElement("span");
      titleSpan.textContent = "Gruppe " + (idx + 1);
      var countSpan = document.createElement("span");
      countSpan.className = "group-count";
      countSpan.textContent = group.length + " Personen";
      heading.appendChild(titleSpan);
      heading.appendChild(countSpan);
      card.appendChild(heading);

      var ul = document.createElement("ul");
      group.forEach(function (p) {
        var li = document.createElement("li");

        var nameSpan = document.createElement("span");
        nameSpan.className = "person-name";
        nameSpan.textContent = p.name;
        li.appendChild(nameSpan);

        var chip = document.createElement("span");
        chip.className = "dept-chip " + deptColorMap[p.dept];
        chip.textContent = p.dept;
        chip.title = p.dept;
        li.appendChild(chip);

        if (p.isFK) {
          var tag = document.createElement("span");
          tag.className = "fk-tag";
          tag.textContent = "FK";
          li.appendChild(tag);
        }

        ul.appendChild(li);
      });
      card.appendChild(ul);
      els.groupsContainer.appendChild(card);
    });
  }

  function renderQuality(quality) {
    applyBadge(els.qualityDept, quality.deptRating);
    applyBadge(els.qualityFk, quality.fkRating);
    els.qualitySize.textContent = quality.sizeValid ? "Gültig" : "Ungültig";
    els.qualitySize.className = "badge " + (quality.sizeValid ? "optimal" : "gut");
  }

  function runAssignment() {
    clearError();
    try {
      var groups = assignGroups(state.participants);
      state.groups = groups;
      renderGroups(groups);
      renderQuality(evaluateQuality(groups));
      els.sectionReroll.hidden = false;
      els.sectionQuality.hidden = false;
      els.sectionResult.hidden = false;
      els.copyStatus.textContent = "";
    } catch (err) {
      console.error(err);
      showError(err.message);
      els.sectionReroll.hidden = true;
      els.sectionQuality.hidden = true;
      els.sectionResult.hidden = true;
    }
  }

  els.btnCreate.addEventListener("click", function () {
    if (!state.participants) return;
    runAssignment();
    els.sectionResult.scrollIntoView({ behavior: "smooth", block: "start" });
  });

  els.btnReroll.addEventListener("click", function () {
    runAssignment();
  });

  // ---------- Copy to clipboard ----------

  function buildCopyText(groups) {
    var lines = [];
    groups.forEach(function (group, idx) {
      lines.push("Gruppe " + (idx + 1));
      group.forEach(function (p) {
        lines.push("- " + p.name + (p.isFK ? " (FK)" : ""));
      });
      lines.push("");
    });
    return lines.join("\n").trim();
  }

  els.btnCopy.addEventListener("click", function () {
    if (!state.groups) return;
    var text = buildCopyText(state.groups);

    function done(ok) {
      els.copyStatus.textContent = ok ? "✅ In die Zwischenablage kopiert." : "Kopieren fehlgeschlagen. Bitte manuell markieren.";
    }

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () { done(true); }, function () { fallbackCopy(text, done); });
    } else {
      fallbackCopy(text, done);
    }
  });

  function fallbackCopy(text, cb) {
    try {
      var ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      var ok = document.execCommand("copy");
      document.body.removeChild(ta);
      cb(ok);
    } catch (e) {
      cb(false);
    }
  }
})();
