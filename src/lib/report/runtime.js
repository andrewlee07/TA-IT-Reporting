import { MAP_VIEWBOX, resolveOfficeMapPoint } from "./office-map";

export function initReportApp(root, options) {
  if (!root) {
    throw new Error("A report root element is required.");
  }

  var nativeDocument = root.ownerDocument;
  var window = nativeDocument.defaultView || globalThis.window;
  var Chart = options.ChartLib || window.Chart;
  var D = options.data;
  var ACTIVE_MONTH = options.activeMonth;
  var INITIAL_PAGE_ID = options.initialPageId || "p-exec";
  var SHOW_ALL_PAGES = Boolean(options.showAllPages);
  var document = {
    getElementById: function getElementById(id) {
      return root.querySelector("#" + id);
    },
    querySelector: function querySelector(selector) {
      return root.querySelector(selector);
    },
    querySelectorAll: function querySelectorAll(selector) {
      return root.querySelectorAll(selector);
    },
    get title() {
      return nativeDocument.title;
    },
    set title(value) {
      nativeDocument.title = value;
    },
  };

  Chart.defaults.font.family = "Arial, sans-serif";
  Chart.defaults.font.size = 10;
  Chart.defaults.color = "#9CA3AF";

  var GRID = { color: "#F3F4F6", lineWidth: 0.8 };
  var TIP = {
    backgroundColor: "#fff",
    borderColor: "#E5E7EB",
    borderWidth: 1,
    titleColor: "#111827",
    bodyColor: "#4B5563",
    padding: 10,
    cornerRadius: 4,
  };
  var NO_LEGEND = { legend: { display: false }, tooltip: TIP };
  var COLORS = {
    blue: "#005292",
    orange: "#F57D00",
    teal: "#219D98",
    grey: "#A9A9AA",
    red: "#003D6E",
    alert: "#C0392B",
  };
  var CHARTS = Object.create(null);

  var orderedMonths = D.meta.availableMonths || [];
  var activeMonth = orderedMonths.indexOf(ACTIVE_MONTH) >= 0 ? ACTIVE_MONTH : D.meta.activeMonth;
  var visibleMonths = orderedMonths.filter(function filterMonths(month) {
    return month <= activeMonth;
  });
  var visibleLabels = visibleMonths.map(function mapMonth(month) {
    return D.meta.monthLabels[month] || month;
  });
  var monthLabel = D.meta.activeMonthLabel || activeMonth;
  var previousMonth = visibleMonths[visibleMonths.length - 2] || activeMonth;

  function fmt(value) {
    return value >= 1000 ? (value / 1000).toFixed(0) + "k" : String(value);
  }

  function pctNum(value) {
    return parseFloat(String(value).replace("%", ""));
  }

  function svcClass(pct, target) {
    var p = pctNum(pct);
    var t = pctNum(target);

    if (p >= t) {
      return "up";
    }

    if (p >= t - 1) {
      return "warn";
    }

    return "down";
  }

  function svcPctClass(pct, target) {
    var p = pctNum(pct);
    var t = pctNum(target);

    if (p >= t) {
      return "full";
    }

    if (p >= t - 1) {
      return "warn";
    }

    return "down";
  }

  function svcStatusHTML(pct, target) {
    var p = pctNum(pct);
    var t = pctNum(target);

    if (p >= t) {
      return '<span class="svc-status pill-green">Above target</span>';
    }

    if (p >= t - 1) {
      return '<span class="svc-status pill-amber">Within tolerance</span>';
    }

    return '<span class="svc-status pill-red">Below target</span>';
  }

  function ragColor(value) {
    if (value === "Green") {
      return "green";
    }

    if (value === "Amber") {
      return "amber";
    }

    return "red";
  }

  function byMonth(rows, month) {
    return rows.filter(function filterRow(row) {
      return row.Month === month;
    });
  }

  function latest(rows) {
    return rows.find(function findRow(row) {
      return row.Month === activeMonth;
    }) || rows[rows.length - 1];
  }

  function previous(rows) {
    return rows.find(function findRow(row) {
      return row.Month === previousMonth;
    }) || latest(rows);
  }

  function monthSeries(rows, metric) {
    return visibleMonths.map(function mapMonth(month) {
      var row = rows.find(function findRow(entry) {
        return entry.Month === month;
      });
      return row ? row[metric] : null;
    });
  }

  function slugify(value) {
    return String(value)
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .replace(/-{2,}/g, "-");
  }

  function escapeAttr(value) {
    return String(value).replace(/&/g, "&amp;").replace(/"/g, "&quot;");
  }

  function exportAttrs(id, label) {
    return ' id="' + escapeAttr(id) + '" data-export-id="' + escapeAttr(id) + '" data-export-label="' + escapeAttr(label) + '"';
  }

  function renderKpiCard(id, tone, label, valueHtml, deltaHtml, deltaClass) {
    return (
      '<div class="kc ' +
      tone +
      '"' +
      exportAttrs(id, label) +
      '><div class="kc-label">' +
      label +
      '</div><div class="kc-value">' +
      valueHtml +
      '</div><div class="kc-delta' +
      (deltaClass ? " " + deltaClass : "") +
      '">' +
      deltaHtml +
      "</div></div>"
    );
  }

  function updateStaticChrome() {
    document.title = "TeacherActive — IT Executive Report · " + monthLabel;

    document.querySelectorAll(".ph-period-val").forEach(function updatePeriod(el) {
      el.textContent = monthLabel;
    });

    var sidebarFooter = document.querySelector(".sidebar-footer");
    if (sidebarFooter) {
      sidebarFooter.innerHTML = "Data: TABS / internal systems<br>Period: " + monthLabel + " · Internal use only";
    }

    var pages = Array.prototype.slice.call(document.querySelectorAll(".report-page"));
    pages.forEach(function updateFooter(page, index) {
      var footer = page.querySelector(".pf-page");
      if (!footer) {
        return;
      }

      var prefix = String(footer.textContent || "").split("· PAGE")[0].trim();
      footer.textContent = prefix + " · PAGE " + (index + 1) + " OF " + pages.length;
    });

    if (SHOW_ALL_PAGES) {
      pages.forEach(function activatePage(page) {
        page.classList.add("active");
        page.style.display = "block";
      });
      document.querySelectorAll(".nav-link").forEach(function clearNav(link) {
        link.classList.remove("active");
      });
      return;
    }

    document.querySelectorAll(".report-page").forEach(function deactivatePage(page) {
      page.classList.remove("active");
      page.style.display = "";
    });
    document.querySelectorAll(".nav-link").forEach(function clearNav(link) {
      link.classList.remove("active");
    });

    var page = document.getElementById(INITIAL_PAGE_ID);
    if (page) {
      page.classList.add("active");
    }

    var nav = document.querySelector('.nav-link[data-page-id="' + INITIAL_PAGE_ID + '"]');
    if (nav) {
      nav.classList.add("active");
    }
  }

  function registerChart(id, createChart) {
    if (CHARTS[id]) {
      CHARTS[id].destroy();
    }

    var chart = createChart();
    CHARTS[id] = chart;
    return chart;
  }

  function resetChartCanvas(id) {
    var canvas = document.getElementById(id);
    var existingChart = canvas ? Chart.getChart(canvas) : null;

    if (existingChart) {
      existingChart.destroy();
    }

    return canvas;
  }

  function rebuildVisiblePage(id) {
    switch (id) {
      case "p-exec":
        buildExecutivePage();
        break;
      case "p-avail":
        buildAvailabilityPage();
        break;
      case "p-network":
        buildNetworkPage();
        break;
      case "p-support":
        buildSupportPage();
        break;
      case "p-security":
        buildSecurityPage();
        break;
      case "p-assets":
        buildAssetsPage();
        break;
      case "p-change":
        buildChangePage();
        break;
      case "p-dev":
        buildDevelopmentPage();
        break;
      case "p-projects":
        buildProjectsPage();
        break;
      case "p-roadmap":
        buildRoadmapPage();
        break;
      case "p-budget":
        buildBudgetPage();
        break;
      case "p-risks":
        buildRisksPage();
        break;
      default:
        break;
    }
  }

  function showPage(id, el, runtimeOptions) {
    if (SHOW_ALL_PAGES) {
      return;
    }

    document.querySelectorAll(".report-page").forEach(function deactivatePage(page) {
      page.classList.remove("active");
    });
    document.querySelectorAll(".nav-link").forEach(function clearNav(link) {
      link.classList.remove("active");
    });

    var page = document.getElementById(id);
    if (page) {
      page.classList.add("active");
    }

    var activeNav = el || document.querySelector('.nav-link[data-page-id="' + id + '"]');
    if (activeNav) {
      activeNav.classList.add("active");
    }

    rebuildVisiblePage(id);
    window.scrollTo(0, 0);

    if ((!runtimeOptions || runtimeOptions.silent !== true) && typeof options.onPageChange === "function") {
      options.onPageChange(id);
    }
  }

  if (options.attachGlobals !== false) {
    window.showPage = showPage;
  }

  function buildSvcTiles(containerId, rows) {
    var el = document.getElementById(containerId);
    if (!el) {
      return;
    }

    el.innerHTML = rows
      .map(function renderTile(service) {
        var tileId = containerId + "-" + slugify(service.Service);
        return (
          '<div class="svc-tile ' +
          svcClass(service.Availability, service.Target) +
          '"' +
          exportAttrs(tileId, service.Service + " service tile") +
          ">" +
          '<div class="svc-type">' +
          service.Type +
          "</div>" +
          '<div class="svc-name">' +
          service.Service +
          "</div>" +
          '<div class="svc-pct ' +
          svcPctClass(service.Availability, service.Target) +
          '">' +
          service.Availability +
          "</div>" +
          '<div class="svc-target">Target: ' +
          service.Target +
          "</div>" +
          svcStatusHTML(service.Availability, service.Target) +
          '<div class="svc-outage">' +
          (service.OutageMins > 0 ? service.OutageMins + " mins downtime" : "No downtime") +
          " · " +
          (service.MajorIncidents === 0 ? "No major incidents" : service.MajorIncidents + " major incident") +
          "</div>" +
          "</div>"
        );
      })
      .join("");
  }

  function buildExecutivePage() {
    var services = byMonth(D.service, activeMonth);
    var support = latest(D.support);
    var supportPrev = previous(D.support);
    var security = latest(D.security);
    var securityPrev = previous(D.security);
    var change = latest(D.change);
    var changePrev = previous(D.change);
    var dev = latest(D.dev);
    var devPrev = previous(D.dev);

    buildSvcTiles("exec-svc-grid", services);

    var execKpis = document.getElementById("exec-kpis");
    if (execKpis) {
      execKpis.innerHTML =
        renderKpiCard(
          "exec-kpi-support-sla",
          "blue",
          "Support SLA",
          support.ResolutionSLA,
          "▲ " + (pctNum(support.ResolutionSLA) - pctNum(supportPrev.ResolutionSLA)).toFixed(1) + " pts vs prior month",
          "up",
        ) +
        renderKpiCard(
          "exec-kpi-user-csat",
          "teal",
          "User CSAT",
          support.CSAT,
          "▲ " + (pctNum(support.CSAT) - pctNum(supportPrev.CSAT)).toFixed(1) + " vs prior month",
          "up",
        ) +
        renderKpiCard(
          "exec-kpi-critical-vulns",
          "green",
          "Critical Vulns",
          security.CritVulns,
          security.CritVulns <= securityPrev.CritVulns ? "▼ reduced backlog" : "▲ increased backlog",
          "up",
        ) +
        renderKpiCard(
          "exec-kpi-change-success",
          "orange",
          "Change Success",
          change.SuccessRate.replace("%", "") + '<span class="u">%</span>',
          "▲ " + (pctNum(change.SuccessRate) - pctNum(changePrev.SuccessRate)).toFixed(1) + " pts vs prior month",
          "up",
        ) +
        renderKpiCard(
          "exec-kpi-dev-backlog",
          "grey",
          "Dev Backlog",
          dev.BacklogEnd,
          dev.BacklogEnd <= devPrev.BacklogEnd ? "▼ reduced backlog" : "▲ increased backlog",
          "up",
        );
    }

    var narrative = document.getElementById("exec-narrative");
    if (narrative) {
      var typeClassMap = {
        Win: "win",
        Concern: "concern",
        "Decision needed": "decision",
      };

      narrative.innerHTML = byMonth(D.narrative, activeMonth)
        .map(function renderNote(note, index) {
          return (
            '<div class="nar-card"' +
            exportAttrs("exec-note-" + (index + 1), note.NoteType + " narrative card") +
            ">" +
            '<div class="nar-type ' +
            (typeClassMap[note.NoteType] || "win") +
            '">' +
            note.NoteType +
            "</div>" +
            '<div class="nar-headline">' +
            note.Headline +
            "</div>" +
            '<div class="nar-text">' +
            note.Narrative +
            "</div>" +
            '<div class="nar-owner">' +
            note.Owner +
            " · " +
            monthLabel +
            "</div>" +
            "</div>"
          );
        })
        .join("");
    }
  }

  function buildAvailabilityPage() {
    var services = byMonth(D.service, activeMonth);
    var network = services.find(function findService(service) {
      return service.Service === "Network";
    });
    var privateCloud = services.find(function findService(service) {
      return service.Service === "Private Cloud";
    });
    var keyServices = ["Network", "TABS / TAMS", "Private Cloud"];
    var outageData = {};
    var outageBars = document.getElementById("avail-outage-bars");
    var note = document.getElementById("avail-note");

    buildSvcTiles("avail-svc-grid", services);

    registerChart("c-avail-trend", function createAvailabilityTrendChart() {
      return new Chart(resetChartCanvas("c-avail-trend"), {
      type: "line",
      data: {
        labels: visibleLabels,
        datasets: keyServices.map(function createDataset(serviceName, index) {
          return {
            label: serviceName,
            data: visibleMonths.map(function mapMonth(month) {
              var row = D.service.find(function findRow(service) {
                return service.Month === month && service.Service === serviceName;
              });
              return row ? pctNum(row.Availability) : null;
            }),
            borderColor: [COLORS.blue, COLORS.teal, COLORS.orange][index],
            backgroundColor: "transparent",
            tension: 0.3,
            pointRadius: 4,
            borderWidth: 2,
          };
        }),
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: true, labels: { boxWidth: 10, font: { size: 10 }, color: "#4B5563" } },
          tooltip: TIP,
        },
        scales: {
          x: { grid: { display: false } },
          y: { grid: GRID, border: { display: false }, min: 97, max: 100.5, ticks: { callback: function callback(v) { return v + "%"; } } },
        },
      },
      });
    });

    D.service
      .filter(function filterService(service) {
        return visibleMonths.indexOf(service.Month) >= 0;
      })
      .forEach(function groupOutage(service) {
        outageData[service.Service] = (outageData[service.Service] || 0) + service.OutageMins;
      });
    var outageValues = Object.values(outageData);
    var outageMax = outageValues.length ? Math.max.apply(Math, outageValues) : 1;

    if (outageBars) {
      outageBars.innerHTML = Object.entries(outageData)
        .sort(function sortOutage(a, b) {
          return b[1] - a[1];
        })
        .map(function renderBar(entry) {
          var serviceName = entry[0];
          var minutes = entry[1];
          return (
            '<div class="bar-row cols-3" style="grid-template-columns:130px 1fr 56px;">' +
            '<div class="bar-name">' +
            serviceName +
            "</div>" +
            '<div class="bar-track"><div class="bar-fill" style="width:' +
            ((minutes / outageMax) * 100).toFixed(1) +
            "%;background:" +
            (minutes > 300 ? COLORS.orange : COLORS.blue) +
            '"></div></div>' +
            '<div class="bar-val">' +
            minutes +
            '<span style="font-size:9px;color:var(--text-3);font-family:Arial"> mins</span></div>' +
            "</div>"
          );
        })
        .join("");
    }

    if (note) {
      note.innerHTML =
        "<strong>" +
        monthLabel +
        "</strong> delivered " +
        (network && pctNum(network.Availability) >= pctNum(network.Target) ? "network performance above target" : "mixed network performance") +
        ". Overall network availability closed at <strong>" +
        (network ? network.Availability : "0.0%") +
        "</strong> with <strong>" +
        (network ? network.OutageMins : 0) +
        "</strong> outage minutes. Private Cloud closed at <strong>" +
        (privateCloud ? privateCloud.Availability : "0.0%") +
        "</strong> and TABS / TAMS remained above target for the selected month.";
    }
  }

  function buildNetworkPage() {
    var officeRows = byMonth(D.officeNetwork, activeMonth).sort(function sortRows(a, b) {
      return a.DisplayOrder - b.DisplayOrder;
    });
    var networkMetric = latest(D.derivedNetwork);
    var averageAvailability = pctNum(networkMetric.Availability);
    var perfect = networkMetric.PerfectOffices;
    var below99 = networkMetric.Below99Offices;
    var below999 = networkMetric.Below99_9Offices;
    var netKpis = document.getElementById("net-kpis");
    var mapBadge = document.getElementById("net-map-badge");
    var dots = document.getElementById("office-dots");
    var officeList = document.getElementById("office-list");

    if (netKpis) {
      netKpis.innerHTML =
        renderKpiCard("net-kpi-avg-availability", "teal", "Avg Availability", averageAvailability.toFixed(2) + '<span class="u">%</span>', "▲ " + perfect + " offices at 100%", "up") +
        renderKpiCard("net-kpi-total-offices", "blue", "Total Offices", officeRows.length, "England &amp; Wales") +
        renderKpiCard(
          "net-kpi-below-99",
          below99 > 0 ? "red" : "green",
          "Below 99%",
          below99,
          below99 === 0 ? "No materially impacted sites" : "Offices require attention",
        ) +
        renderKpiCard("net-kpi-below-99-9", "orange", "Below 99.9%", below999, "Minor or major issues");
    }

    if (mapBadge) {
      mapBadge.textContent = officeRows.length + " offices";
    }

    if (dots) {
      var unresolvedOffices = [];

      dots.innerHTML = officeRows
        .map(function renderDot(office) {
          var point = resolveOfficeMapPoint(office.OfficeName);

          if (!point) {
            unresolvedOffices.push(office.OfficeName);
            return "";
          }

          var fill = pctNum(office.Availability) === 100 ? COLORS.teal : pctNum(office.Availability) >= 99.9 ? COLORS.orange : COLORS.alert;
          var radius = pctNum(office.Availability) < 99 ? 7 : pctNum(office.Availability) < 99.9 ? 6 : 5;
          return (
            '<circle cx="' +
            point.x.toFixed(1) +
            '" cy="' +
            point.y.toFixed(1) +
            '" r="' +
            radius +
            '" fill="' +
            fill +
            '" stroke="white" stroke-width="1.5" opacity="0.92" data-office-name="' +
            office.OfficeName.replace(/"/g, "&quot;") +
            '" onmouseenter="showMapTip(\'' +
            office.OfficeName.replace(/'/g, "\\'") +
            "', '" +
            office.Availability +
            "', " +
            point.x.toFixed(1) +
            ", " +
            point.y.toFixed(1) +
            ')" onmouseleave="hideMapTip()" />'
          );
        })
        .join("");

      if (
        unresolvedOffices.length > 0 &&
        typeof console !== "undefined" &&
        typeof process !== "undefined" &&
        process.env.NODE_ENV !== "production"
      ) {
        console.warn("Missing office map coordinates for:", unresolvedOffices.join(", "));
      }
    }

    if (officeList) {
      officeList.innerHTML = officeRows
        .slice()
        .sort(function sortOffice(a, b) {
          return pctNum(b.Availability) - pctNum(a.Availability);
        })
        .map(function renderOffice(office) {
          var pct = pctNum(office.Availability);
          var color = pct === 100 ? COLORS.teal : pct >= 99.9 ? COLORS.orange : COLORS.alert;
          var label = pct === 100 ? "Perfect" : pct >= 99.9 ? "Good" : pct >= 99 ? "Minor" : "Impacted";
          return (
            '<div style="padding:8px 14px;border-bottom:1px solid var(--rule-l);display:grid;grid-template-columns:1fr 58px;gap:8px;align-items:center;">' +
            "<div>" +
            '<div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;">' +
            '<span style="width:7px;height:7px;border-radius:50%;background:' +
            color +
            ';flex-shrink:0;display:inline-block"></span>' +
            '<span style="font-size:11px;font-weight:700;color:var(--text)">' +
            office.OfficeName +
            "</span>" +
            '<span style="font-size:9px;color:var(--text-3)">' +
            office.Region +
            "</span>" +
            "</div>" +
            '<div style="height:3px;background:var(--rule);border-radius:2px;overflow:hidden;">' +
            '<div style="height:100%;width:' +
            Math.max(0, ((pct - 97) / 3) * 100).toFixed(1) +
            "%;background:" +
            color +
            ';border-radius:2px;"></div></div>' +
            "</div>" +
            '<div style="text-align:right;"><div style="font-family:\'Arial Black\',Arial;font-size:11px;font-weight:700;color:' +
            color +
            '">' +
            office.Availability +
            '</div><div style="font-size:9px;color:var(--text-3)">' +
            label +
            "</div></div>" +
            "</div>"
          );
        })
        .join("");
    }

    registerChart("c-net-trend", function createNetworkTrendChart() {
      return new Chart(resetChartCanvas("c-net-trend"), {
      type: "line",
      data: {
        labels: visibleLabels,
        datasets: [
          {
            label: "Network avg",
            data: visibleMonths.map(function mapMetric(month) {
              var metric = D.derivedNetwork.find(function findMetric(item) {
                return item.Month === month;
              });
              return metric ? pctNum(metric.Availability) : null;
            }),
            borderColor: COLORS.blue,
            backgroundColor: "rgba(0,82,146,0.07)",
            fill: true,
            tension: 0.3,
            pointRadius: 4,
            borderWidth: 2,
          },
          {
            label: "Worst office",
            data: visibleMonths.map(function mapWorst(month) {
              var metric = D.derivedNetwork.find(function findMetric(item) {
                return item.Month === month;
              });
              return metric ? pctNum(metric.WorstAvailability) : null;
            }),
            borderColor: COLORS.orange,
            backgroundColor: "transparent",
            tension: 0.3,
            pointRadius: 4,
            borderWidth: 1.5,
            borderDash: [4, 3],
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 10 }, color: "#4B5563" } }, tooltip: TIP },
        scales: {
          x: { grid: { display: false } },
          y: { grid: GRID, border: { display: false }, min: 97, max: 100.2, ticks: { callback: function callback(v) { return v + "%"; } } },
        },
      },
      });
    });
  }

  if (options.attachGlobals !== false) {
    window.showMapTip = function showMapTip(name, availability, x, y) {
      var tooltip = document.getElementById("map-tooltip");
      if (!tooltip) {
        return;
      }

      var pct = pctNum(availability);
      var color = pct === 100 ? COLORS.teal : pct >= 99.9 ? COLORS.orange : COLORS.alert;
      var label = pct === 100 ? "No incidents" : pct >= 99.9 ? "Minor issues" : pct >= 99 ? "Monitored" : "Impacted";
      var width = 150;
      var height = 52;
      var tx = Math.min(Math.max(x + 14, 12), MAP_VIEWBOX.width - width - 12);
      var ty = Math.min(Math.max(y - 58, 12), MAP_VIEWBOX.height - height - 12);

      document.getElementById("tt-bg").setAttribute("x", String(tx));
      document.getElementById("tt-bg").setAttribute("y", String(ty));
      document.getElementById("tt-bg").setAttribute("width", String(width));
      document.getElementById("tt-bg").setAttribute("height", String(height));
      document.getElementById("tt-name").setAttribute("x", String(tx + 8));
      document.getElementById("tt-name").setAttribute("y", String(ty + 16));
      document.getElementById("tt-name").textContent = name;
      document.getElementById("tt-avail").setAttribute("x", String(tx + 8));
      document.getElementById("tt-avail").setAttribute("y", String(ty + 32));
      document.getElementById("tt-avail").textContent = availability;
      document.getElementById("tt-avail").setAttribute("fill", color);
      document.getElementById("tt-status").setAttribute("x", String(tx + 52));
      document.getElementById("tt-status").setAttribute("y", String(ty + 32));
      document.getElementById("tt-status").textContent = label;
      document.getElementById("tt-status").setAttribute("fill", color);
      tooltip.style.display = "block";
    };

    window.hideMapTip = function hideMapTip() {
      var tooltip = document.getElementById("map-tooltip");
      if (tooltip) {
        tooltip.style.display = "none";
      }
    };
  }

  function buildSupportPage() {
    var support = latest(D.support);
    var supportPrev = previous(D.support);
    var categories = {};
    var tickets = byMonth(D.tickets, activeMonth);
    var oldestTicket = tickets.reduce(function getOldest(max, ticket) {
      return ticket.AgeDays > max ? ticket.AgeDays : max;
    }, 0);
    var categoryColors = [COLORS.blue, COLORS.orange, COLORS.teal, COLORS.grey];

    document.getElementById("sh-number").innerHTML = support.ResolutionSLA.replace("%", "") + "<sup>%</sup>";
    document.getElementById("sh-benchmarks").innerHTML =
      '<div class="hb-item"><div class="hb-label">Target</div><div class="hb-val">95.0%</div></div>' +
      '<div class="hb-item"><div class="hb-label">Prior month</div><div class="hb-val">' +
      supportPrev.ResolutionSLA +
      '</div></div><div class="hb-item"><div class="hb-label">CSAT</div><div class="hb-val">' +
      support.CSAT +
      "</div></div>";
    document.getElementById("sh-trend").innerHTML = D.support.filter(function filterSupport(row) {
      return visibleMonths.indexOf(row.Month) >= 0;
    })
      .map(function renderTrend(row) {
        var value = pctNum(row.ResolutionSLA);
        var high = value >= 99;
        return (
          '<div class="week-row"><div class="week-label">' +
          (D.meta.monthLabels[row.Month] || row.Month) +
          '</div><div class="week-track"><div class="week-fill" style="width:' +
          (((value - 94) / 6) * 100).toFixed(0) +
          "%;background:" +
          (high ? "var(--teal)" : "var(--blue)") +
          '"></div></div><div class="week-val ' +
          (high ? "hi" : "") +
          '">' +
          row.ResolutionSLA +
          "</div></div>"
        );
      })
      .join("");

    document.getElementById("support-kpis").innerHTML =
      renderKpiCard("support-kpi-opened", "blue", "Opened", support.Opened.toLocaleString(), "tickets this month") +
      renderKpiCard(
        "support-kpi-closed",
        "orange",
        "Closed",
        support.Closed.toLocaleString(),
        support.Closed >= support.Opened ? "▲ net positive flow" : "▼ net negative flow",
        "up",
      ) +
      renderKpiCard(
        "support-kpi-backlog",
        "teal",
        "Backlog End",
        support.Backlog,
        support.Backlog <= supportPrev.Backlog ? "▼ lower than prior month" : "▲ higher than prior month",
        "up",
      ) +
      renderKpiCard(
        "support-kpi-avg-resolution",
        "green",
        "Avg Resolution",
        support.AvgResolution + '<span class="u"> days</span>',
        support.AvgResolution <= supportPrev.AvgResolution ? "▲ improved turnaround" : "▼ slower turnaround",
        "up",
      ) +
      renderKpiCard(
        "support-kpi-major-incidents",
        "grey",
        "Major Incidents",
        support.MajorIncidents,
        support.MajorIncidents === 0 ? "clean month" : "incident activity recorded",
      );

    registerChart("c-support-vol", function createSupportVolumeChart() {
      return new Chart(resetChartCanvas("c-support-vol"), {
      type: "bar",
      data: {
        labels: visibleLabels,
        datasets: [
          { label: "Opened", data: monthSeries(D.support, "Opened"), backgroundColor: COLORS.blue, borderRadius: 3, barPercentage: 0.45 },
          { label: "Closed", data: monthSeries(D.support, "Closed"), backgroundColor: COLORS.orange, borderRadius: 3, barPercentage: 0.45 },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: NO_LEGEND,
        scales: { x: { grid: { display: false } }, y: { grid: GRID, border: { display: false }, ticks: { callback: fmt } } },
      },
      });
    });

    D.support
      .filter(function filterSupport(row) {
        return visibleMonths.indexOf(row.Month) >= 0;
      })
      .forEach(function countCategories(row) {
        categories[row.TopCategory] = (categories[row.TopCategory] || 0) + 1;
      });
    var categoryMax = Math.max.apply(Math, Object.values(categories).length ? Object.values(categories) : [1]);

    document.getElementById("support-cats").innerHTML = Object.entries(categories)
      .sort(function sortCategories(a, b) {
        return b[1] - a[1];
      })
      .map(function renderCategory(entry, index) {
        return (
          '<div class="bar-row cols-3" style="grid-template-columns:150px 1fr 48px;">' +
          '<div class="bar-name">' +
          entry[0] +
          "</div>" +
          '<div class="bar-track"><div class="bar-fill" style="width:' +
          ((entry[1] / categoryMax) * 100).toFixed(0) +
          "%;background:" +
          categoryColors[index % categoryColors.length] +
          '"></div></div><div class="bar-val" style="font-size:11px">' +
          entry[1] +
          "</div></div>"
        );
      })
      .join("");

    document.getElementById("tix-age-badge").textContent = "Oldest: " + oldestTicket + " days";
    document.getElementById("oldest-tickets").innerHTML =
      '<div class="tix-list"><div class="tix-row head"><div class="tix-cell">Ticket ID</div><div class="tix-cell">Title</div><div class="tix-cell">Category · Queue</div><div class="tix-cell">Age (days)</div><div class="tix-cell">Critical</div></div>' +
      tickets
        .map(function renderTicket(ticket) {
          return (
            '<div class="tix-row"><div class="tix-id">' +
            ticket.TicketID +
            '</div><div><div class="tix-title">' +
            ticket.Title +
            '</div></div><div><div class="tix-cat">' +
            ticket.Category +
            '</div><div style="font-size:10px;color:var(--text-3)">' +
            ticket.OwnerQueue +
            '</div></div><div class="tix-age ' +
            (ticket.AgeDays >= 60 ? "crit" : ticket.AgeDays >= 40 ? "warn" : "ok") +
            '">' +
            ticket.AgeDays +
            '</div><div><span class="tix-crit ' +
            (ticket.BusinessCritical === "Yes" ? "yes" : "no") +
            '"></span></div></div>'
          );
        })
        .join("") +
      "</div>";
  }

  function buildSecurityPage() {
    var security = latest(D.security);
    var previousSecurity = previous(D.security);

    document.getElementById("sec-kpis").innerHTML =
      renderKpiCard(
        "sec-kpi-critical-vulns",
        "green",
        "Critical Vulns",
        security.CritVulns,
        security.CritVulns <= previousSecurity.CritVulns ? "▼ reduced backlog" : "▲ increased backlog",
        "up",
      ) +
      renderKpiCard(
        "sec-kpi-workstation-patch",
        "blue",
        "Workstation Patch",
        security.WkstationPatch.replace("%", "") + '<span class="u">%</span>',
        pctNum(security.WkstationPatch) >= pctNum(previousSecurity.WkstationPatch) ? "▲ improving" : "▼ lower than prior month",
        "up",
      ) +
      renderKpiCard(
        "sec-kpi-mfa-coverage",
        "teal",
        "MFA Coverage",
        security.MFACoverage.replace("%", "") + '<span class="u">%</span>',
        "▲ near full coverage",
        "up",
      ) +
      renderKpiCard(
        "sec-kpi-overdue-remediation",
        "orange",
        "Overdue Remediation",
        security.OverdueRemediation,
        security.OverdueRemediation <= previousSecurity.OverdueRemediation ? "▼ reduced backlog" : "▲ higher than prior month",
        "up",
      );

    document.getElementById("sec-compliance-bars").innerHTML = [
      { label: "Workstation Patch", value: pctNum(security.WkstationPatch) },
      { label: "Server Patch", value: pctNum(security.ServerPatch) },
      { label: "Critical Patch", value: pctNum(security.CriticalPatch) },
      { label: "MFA Coverage", value: pctNum(security.MFACoverage) },
      { label: "Endpoint Coverage", value: pctNum(security.EndpointCoverage) },
    ]
      .map(function renderCompliance(item) {
        return (
          '<div class="comp-row"><div class="comp-label">' +
          item.label +
          '</div><div class="comp-track"><div class="comp-fill" style="width:' +
          item.value +
          "%;background:" +
          (item.value >= 97 ? COLORS.teal : item.value >= 93 ? COLORS.blue : COLORS.orange) +
          '"></div></div><div class="comp-val">' +
          item.value.toFixed(1) +
          "%</div></div>"
        );
      })
      .join("");

    registerChart("c-vuln-trend", function createVulnerabilityTrendChart() {
      return new Chart(resetChartCanvas("c-vuln-trend"), {
      type: "bar",
      data: {
        labels: visibleLabels,
        datasets: [
          { label: "Critical", data: monthSeries(D.security, "CritVulns"), backgroundColor: COLORS.red, stack: "s" },
          { label: "High", data: monthSeries(D.security, "HighVulns"), backgroundColor: COLORS.orange, stack: "s" },
          { label: "Medium", data: monthSeries(D.security, "MedVulns"), backgroundColor: COLORS.teal, stack: "s" },
          { label: "Low", data: monthSeries(D.security, "LowVulns"), backgroundColor: COLORS.grey, stack: "s" },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: TIP },
        scales: { x: { grid: { display: false }, stacked: true }, y: { grid: GRID, border: { display: false }, stacked: true } },
      },
      });
    });

    document.getElementById("sec-note").innerHTML =
      "<strong>" +
      monthLabel +
      "</strong> closed with workstation patch compliance at <strong>" +
      security.WkstationPatch +
      "</strong>, server compliance at <strong>" +
      security.ServerPatch +
      "</strong>, and a critical vulnerability backlog of <strong>" +
      security.CritVulns +
      "</strong>. Overdue remediation now sits at <strong>" +
      security.OverdueRemediation +
      "</strong> items with MFA coverage at <strong>" +
      security.MFACoverage +
      "</strong>.";
  }

  function buildAssetsPage() {
    var assets = byMonth(D.assets, activeMonth);
    var laptop = assets.find(function findAsset(asset) {
      return asset.AssetType === "Laptop";
    });
    var mobile = assets.find(function findAsset(asset) {
      return asset.AssetType === "Mobile";
    });
    var monitor = assets.find(function findAsset(asset) {
      return asset.AssetType === "Monitor";
    });
    var laptopSeries = D.assets.filter(function filterAssets(asset) {
      return asset.AssetType === "Laptop" && visibleMonths.indexOf(asset.Month) >= 0;
    });

    document.getElementById("asset-kpis").innerHTML =
      renderKpiCard(
        "asset-kpi-total-active-devices",
        "blue",
        "Total Active Devices",
        (laptop.ActiveDevices + mobile.ActiveDevices + monitor.ActiveDevices).toLocaleString(),
        "Laptop · Mobile · Monitor",
      ) +
      renderKpiCard("asset-kpi-laptops-in-lifecycle", "teal", "Laptops in Lifecycle", laptop.PctWithin, "▲ lifecycle coverage", "up") +
      renderKpiCard("asset-kpi-laptop-incidents", "orange", "Laptop Incidents", laptop.IncidentsLinked, "▼ linked hardware incidents", "up") +
      renderKpiCard("asset-kpi-stock-cover", "grey", "Stock Cover", laptop.StockOnHand + '<span class="u"> units</span>', "Ready to deploy");

    document.getElementById("asset-tiles").innerHTML = [laptop, mobile, monitor]
      .map(function renderAsset(asset) {
        var inLifecycle = pctNum(asset.PctWithin);
        var outLifecycle = pctNum(asset.PctOutside);
        return (
          '<div class="block mb-0"' +
          exportAttrs("asset-tile-" + slugify(asset.AssetType), asset.AssetType + " lifecycle tile") +
          '><div class="bh"><div><div class="bh-title">' +
          asset.AssetType +
          's</div><div class="bh-sub">' +
          asset.ActiveDevices +
          ' active devices</div></div><div class="bh-badge">Avg ' +
          asset.AvgAgeMths +
          ' mths</div></div><div class="bb"><div style="margin-bottom:14px"><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:0.8px;color:var(--text-3);margin-bottom:6px">Lifecycle health</div><div style="display:flex;gap:0;height:12px;border-radius:4px;overflow:hidden"><div style="width:' +
          inLifecycle +
          "%;background:" +
          (inLifecycle >= 85 ? COLORS.teal : inLifecycle >= 75 ? COLORS.blue : COLORS.orange) +
          '"></div><div style="width:' +
          outLifecycle +
          '%;background:var(--rule)"></div></div><div style="display:flex;justify-content:space-between;margin-top:5px;font-size:10px;color:var(--text-3)"><span style="color:var(--text-2);font-weight:700">' +
          asset.PctWithin +
          ' in lifecycle</span><span>' +
          asset.PctOutside +
          '</span></div></div><div style="display:grid;grid-template-columns:1fr 1fr;gap:10px"><div><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:0.8px;color:var(--text-3)">Stock on hand</div><div style="font-family:\'Arial Black\',Arial;font-size:18px;font-weight:700;color:var(--text)">' +
          asset.StockOnHand +
          '</div></div><div><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:0.8px;color:var(--text-3)">Incident links</div><div style="font-family:\'Arial Black\',Arial;font-size:18px;font-weight:700;color:' +
          (asset.IncidentsLinked > 10 ? COLORS.orange : COLORS.teal) +
          '">' +
          asset.IncidentsLinked +
          "</div></div></div></div></div>"
        );
      })
      .join("");

    registerChart("c-asset-trend", function createAssetTrendChart() {
      return new Chart(resetChartCanvas("c-asset-trend"), {
      type: "line",
      data: {
        labels: visibleLabels,
        datasets: [
          {
            label: "% Within Lifecycle",
            data: laptopSeries.map(function getValue(asset) {
              return pctNum(asset.PctWithin);
            }),
            borderColor: COLORS.blue,
            backgroundColor: "rgba(0,82,146,0.07)",
            fill: true,
            tension: 0.3,
            pointRadius: 4,
            borderWidth: 2,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: NO_LEGEND,
        scales: { x: { grid: { display: false } }, y: { grid: GRID, border: { display: false }, min: 70, max: 95, ticks: { callback: function callback(v) { return v + "%"; } } } },
      },
      });
    });

    registerChart("c-asset-spend", function createAssetSpendChart() {
      return new Chart(resetChartCanvas("c-asset-spend"), {
      type: "bar",
      data: {
        labels: visibleLabels,
        datasets: [
          {
            label: "Refresh Spend (£)",
            data: laptopSeries.map(function getValue(asset) {
              return asset.RefreshSpend;
            }),
            backgroundColor: COLORS.orange,
            borderRadius: 3,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: NO_LEGEND,
        scales: { x: { grid: { display: false } }, y: { grid: GRID, border: { display: false }, ticks: { callback: function callback(v) { return "£" + fmt(v); } } } },
      },
      });
    });
  }

  function buildChangePage() {
    var change = latest(D.change);
    var changePrev = previous(D.change);

    document.getElementById("chg-hero-num").innerHTML = change.SuccessRate.replace("%", "") + "<sup>%</sup>";
    document.getElementById("chg-benchmarks").innerHTML =
      '<div class="hb-item"><div class="hb-label">First month</div><div class="hb-val">' +
      D.change[0].SuccessRate +
      '</div></div><div class="hb-item"><div class="hb-label">Prior month</div><div class="hb-val">' +
      changePrev.SuccessRate +
      '</div></div><div class="hb-item"><div class="hb-label">Releases</div><div class="hb-val">' +
      change.ReleasesDeployed +
      "</div></div>";

    document.getElementById("chg-trend").innerHTML = D.change.filter(function filterChange(row) {
      return visibleMonths.indexOf(row.Month) >= 0;
    })
      .map(function renderTrend(row) {
        var value = pctNum(row.SuccessRate);
        return (
          '<div class="week-row"><div class="week-label">' +
          (D.meta.monthLabels[row.Month] || row.Month) +
          '</div><div class="week-track"><div class="week-fill" style="width:' +
          (((value - 89) / 11) * 100).toFixed(0) +
          "%;background:" +
          (value >= 95 ? "var(--teal)" : "var(--blue)") +
          '"></div></div><div class="week-val ' +
          (value >= 95 ? "hi" : "") +
          '">' +
          row.SuccessRate +
          "</div></div>"
        );
      })
      .join("");

    document.getElementById("chg-kpis").innerHTML =
      renderKpiCard("change-kpi-total-changes", "blue", "Total Changes", change.TotalChanges, "selected month") +
      renderKpiCard(
        "change-kpi-releases-deployed",
        "teal",
        "Releases Deployed",
        change.ReleasesDeployed,
        change.ReleasesDeployed >= changePrev.ReleasesDeployed ? "▲ release throughput" : "▼ release throughput",
        "up",
      ) +
      renderKpiCard(
        "change-kpi-failed-changes",
        "orange",
        "Failed Changes",
        change.FailedChanges,
        change.FailedChanges <= changePrev.FailedChanges ? "▼ improved" : "▲ worsened",
        "up",
      ) +
      renderKpiCard(
        "change-kpi-incidents",
        "green",
        "Changes → Incidents",
        change.ChangesIncidents,
        change.ChangesIncidents === 0 ? "no service impact" : "service impact recorded",
        "up",
      );

    registerChart("c-change-breakdown", function createChangeBreakdownChart() {
      return new Chart(resetChartCanvas("c-change-breakdown"), {
      type: "bar",
      data: {
        labels: visibleLabels,
        datasets: [
          { label: "Standard", data: monthSeries(D.change, "StandardChanges"), backgroundColor: COLORS.blue, borderRadius: 3, stack: "s" },
          { label: "Normal", data: monthSeries(D.change, "NormalChanges"), backgroundColor: COLORS.teal, borderRadius: 0, stack: "s" },
          { label: "Emergency", data: monthSeries(D.change, "EmergencyChanges"), backgroundColor: COLORS.orange, borderRadius: 3, stack: "s" },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: TIP },
        scales: { x: { grid: { display: false }, stacked: true }, y: { grid: GRID, border: { display: false }, stacked: true } },
      },
      });
    });
  }

  function buildDevelopmentPage() {
    var dev = latest(D.dev);
    var devPrev = previous(D.dev);

    document.getElementById("dev-kpis").innerHTML =
      renderKpiCard(
        "dev-kpi-backlog-end",
        "blue",
        "Backlog End",
        dev.BacklogEnd,
        dev.BacklogEnd <= devPrev.BacklogEnd ? "▼ reduced backlog" : "▲ higher backlog",
        "up",
      ) +
      renderKpiCard(
        "dev-kpi-tasks-closed",
        "teal",
        "Tasks Closed",
        dev.Closed,
        dev.Closed >= devPrev.Closed ? "▲ improved throughput" : "▼ lower throughput",
        "up",
      ) +
      renderKpiCard(
        "dev-kpi-blocked-items",
        "orange",
        "Blocked Items",
        dev.Blocked,
        dev.Blocked <= devPrev.Blocked ? "▼ fewer blockers" : "▲ more blockers",
        "up",
      ) +
      renderKpiCard(
        "dev-kpi-csat",
        "green",
        "Dev CSAT",
        dev.CSAT,
        pctNum(dev.CSAT) >= pctNum(devPrev.CSAT) ? "▲ sponsor sentiment improving" : "▼ sponsor sentiment lower",
        "up",
      );

    registerChart("c-dev-pipeline", function createDevelopmentPipelineChart() {
      return new Chart(resetChartCanvas("c-dev-pipeline"), {
      type: "bar",
      data: {
        labels: visibleLabels,
        datasets: [
          { label: "Opened", data: monthSeries(D.dev, "Opened"), backgroundColor: COLORS.blue, borderRadius: 3, barPercentage: 0.45, order: 2 },
          { label: "Closed", data: monthSeries(D.dev, "Closed"), backgroundColor: COLORS.orange, borderRadius: 3, barPercentage: 0.45, order: 3 },
          { label: "Backlog", data: monthSeries(D.dev, "BacklogEnd"), type: "line", borderColor: COLORS.teal, backgroundColor: "rgba(33,157,152,0.07)", fill: true, tension: 0.3, pointBackgroundColor: COLORS.teal, pointRadius: 5, borderWidth: 2.5, order: 1, yAxisID: "y1" },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: NO_LEGEND,
        scales: {
          x: { grid: { display: false } },
          y: { grid: GRID, border: { display: false } },
          y1: { grid: { display: false }, border: { display: false }, position: "right", ticks: { color: COLORS.teal } },
        },
      },
      });
    });

    registerChart("c-dev-mix", function createDevelopmentMixChart() {
      return new Chart(resetChartCanvas("c-dev-mix"), {
      type: "doughnut",
      data: {
        labels: ["Defects", "Enhancements", "Tech Debt", "BAU"],
        datasets: [
          {
            data: [dev.Defects, dev.Enhancements, dev.TechDebt, dev.BAU],
            backgroundColor: [COLORS.red, COLORS.blue, COLORS.teal, COLORS.grey],
            borderWidth: 0,
            hoverOffset: 4,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: true, position: "bottom", labels: { boxWidth: 10, font: { size: 10 }, color: "#4B5563", padding: 14 } },
          tooltip: TIP,
        },
      },
      });
    });

    document.getElementById("dev-note").innerHTML =
      "<strong>" +
      monthLabel +
      "</strong> closed with <strong>" +
      dev.Closed +
      "</strong> tasks delivered, <strong>" +
      dev.Opened +
      "</strong> opened, and a backlog of <strong>" +
      dev.BacklogEnd +
      "</strong>. The oldest open item stands at <strong>" +
      dev.OldestOpen +
      "</strong> days and blocked items closed at <strong>" +
      dev.Blocked +
      "</strong>.";
  }

  function buildProjectsPage() {
    var projects = byMonth(D.projects, activeMonth);
    var decisionCount = projects.filter(function filterProjects(project) {
      return project.DecisionNeeded === "Yes";
    }).length;
    var averageConfidence = projects.length
      ? Math.round(
          projects.reduce(function sumConfidence(total, project) {
            return total + parseInt(project.Confidence, 10);
          }, 0) / projects.length,
        )
      : 0;

    document.getElementById("prj-kpis").innerHTML =
      renderKpiCard("projects-kpi-active-projects", "green", "Active Projects", projects.length, "selected month") +
      renderKpiCard("projects-kpi-avg-confidence", "teal", "Avg Confidence", averageConfidence + '<span class="u">%</span>', "delivery confidence", "up") +
      renderKpiCard("projects-kpi-decisions-needed", "orange", "Decisions Needed", decisionCount, "require sign-off");

    document.getElementById("prj-list").innerHTML =
      '<div class="prj-list">' +
      projects
        .map(function renderProject(project) {
          var confidenceValue = parseInt(project.Confidence, 10);
          return (
            '<div class="prj-card"' +
            exportAttrs("project-card-" + slugify(project.ProjectName), project.ProjectName + " project card") +
            '><div><div class="prj-name">' +
            project.ProjectName +
            '</div><div class="prj-sponsor">Sponsor: ' +
            project.Sponsor +
            " · " +
            project.BudgetStatus +
            '</div></div><div style="display:flex;align-items:center;gap:8px;"><div class="rag-dot ' +
            ragColor(project.StatusRAG) +
            '"></div><div style="font-size:11px;color:var(--text-2)">' +
            project.StatusRAG +
            '</div></div><div><div class="prj-conf-label">Confidence</div><div class="prj-conf-bar"><div class="prj-conf-fill" style="width:' +
            project.Confidence +
            ";background:" +
            (confidenceValue >= 85 ? COLORS.teal : confidenceValue >= 70 ? COLORS.blue : COLORS.orange) +
            '"></div></div><div class="prj-conf-val">' +
            project.Confidence +
            '</div></div><div><div class="prj-cell" style="font-weight:700">' +
            project.MilestoneNext +
            '</div><div class="prj-milestone-date">' +
            project.MilestoneDate +
            '</div></div><div><div class="prj-decision ' +
            (project.DecisionNeeded === "Yes" ? "rb-yes" : "rb-no") +
            '">' +
            (project.DecisionNeeded === "Yes" ? "Decision needed" : "No action needed") +
            "</div></div></div>"
          );
        })
        .join("") +
      "</div>";

    document.getElementById("prj-note").innerHTML =
      "<strong>" +
      monthLabel +
      "</strong> shows <strong>" +
      projects.length +
      "</strong> active projects with average delivery confidence at <strong>" +
      averageConfidence +
      "%</strong>. " +
      decisionCount +
      " project(s) currently require sponsor or board direction.";
  }

  function buildRoadmapPage() {
    var quarters = Array.from(
      new Set(
        D.roadmap.map(function getQuarter(item) {
          return item.Quarter;
        }),
      ),
    );

    document.getElementById("rdm-grid").innerHTML = quarters
      .map(function renderQuarter(quarter) {
        var items = D.roadmap.filter(function filterRoadmap(item) {
          return item.Quarter === quarter;
        });
        return (
          '<div style="display:grid;grid-template-columns:120px 1fr;border-top:1px solid var(--rule);margin-bottom:0"' +
          exportAttrs("roadmap-quarter-" + slugify(quarter), quarter + " roadmap section") +
          '><div style="font-family:\'Arial Black\',Arial;font-size:12px;font-weight:700;color:var(--text-2);padding:16px 8px;border-right:1px solid var(--rule)">' +
          quarter +
          '</div><div style="padding:10px;display:flex;flex-direction:column;gap:7px">' +
          items
            .map(function renderItem(item) {
              return (
                '<div class="rdm-item ' +
                ragColor(item.StatusRAG) +
                '" style="' +
                (item.DecisionRequired === "Yes" ? "border-left-color:" + COLORS.blue : "") +
                '"><div><div class="rdm-lane">' +
                item.Lane +
                '</div></div><div><div class="rdm-name">' +
                item.Initiative +
                '</div><div style="font-size:10px;color:var(--text-3);margin-top:2px">' +
                item.Outcome +
                '</div></div><div class="rdm-owner">' +
                item.Owner +
                '</div><div class="rdm-dep">' +
                (item.Dependency || "—") +
                "</div><div>" +
                (item.DecisionRequired === "Yes"
                  ? '<span class="rb-yes" style="display:inline-block;padding:2px 7px;border-radius:3px;font-size:9px;font-weight:700;background:var(--blue-l);color:var(--blue)">Decision</span>'
                  : '<span style="font-size:9px;color:var(--text-3)">—</span>') +
                "</div></div>"
              );
            })
            .join("") +
          "</div></div>"
        );
      })
      .join("");
  }

  function buildBudgetPage() {
    var budgetRows = byMonth(D.budget, activeMonth);
    var totals = latest(D.budgetMonthlyTotals);

    document.getElementById("bud-kpis").innerHTML =
      renderKpiCard("budget-kpi-total-budget", "blue", "Total Budget", "£" + fmt(totals.Budget), monthLabel) +
      renderKpiCard(
        "budget-kpi-total-actual",
        "orange",
        "Total Actual",
        "£" + fmt(totals.Actual),
        totals.Actual <= totals.Budget ? "▲ under budget" : "▼ over budget",
        totals.Actual <= totals.Budget ? "up" : "dn",
      ) +
      renderKpiCard(
        "budget-kpi-variance",
        totals.Variance <= 0 ? "green" : "red",
        "Variance",
        (totals.Variance < 0 ? "" : "+") + '<span class="u">£</span>' + Math.abs(totals.Variance / 1000).toFixed(1) + '<span class="u">k</span>',
        totals.Variance <= 0 ? "favourable" : "adverse",
      ) +
      renderKpiCard("budget-kpi-forecast", "grey", "Forecast", "£" + fmt(totals.Forecast), "month-end forecast");

    document.getElementById("bud-table-wrap").innerHTML =
      '<table class="bud-table"><thead><tr><th>Budget Line</th><th>Budget</th><th>Actual</th><th>Forecast</th><th>Variance</th><th>Commentary</th></tr></thead><tbody>' +
      budgetRows
        .map(function renderBudget(row) {
          return (
            '<tr><td style="font-weight:700">' +
            row.BudgetLine +
            "</td><td>£" +
            row.Budget.toLocaleString() +
            "</td><td>£" +
            row.Actual.toLocaleString() +
            "</td><td>£" +
            row.Forecast.toLocaleString() +
            '</td><td><span class="bud-num ' +
            (row.Variance < 0 ? "bud-pos" : row.Variance > 500 ? "bud-neg" : "bud-neu") +
            '">' +
            (row.Variance < 0 ? "▼ £" + Math.abs(row.Variance).toLocaleString() : row.Variance === 0 ? "On plan" : "▲ £" + row.Variance.toLocaleString()) +
            '</span></td><td style="font-size:10px;color:var(--text-3)">' +
            (row.Commentary || "—") +
            "</td></tr>"
          );
        })
        .join("") +
      "</tbody></table>";

    registerChart("c-budget-trend", function createBudgetTrendChart() {
      return new Chart(resetChartCanvas("c-budget-trend"), {
      type: "line",
      data: {
        labels: visibleLabels,
        datasets: [
          {
            label: "Budget",
            data: monthSeries(D.budgetMonthlyTotals, "Budget"),
            borderColor: COLORS.grey,
            borderDash: [4, 4],
            backgroundColor: "transparent",
            tension: 0,
            pointRadius: 0,
            borderWidth: 1.5,
          },
          {
            label: "Actual",
            data: monthSeries(D.budgetMonthlyTotals, "Actual"),
            borderColor: COLORS.blue,
            backgroundColor: "rgba(0,82,146,0.07)",
            fill: true,
            tension: 0.3,
            pointRadius: 4,
            borderWidth: 2,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 10 }, color: "#4B5563" } }, tooltip: TIP },
        scales: { x: { grid: { display: false } }, y: { grid: GRID, border: { display: false }, ticks: { callback: function callback(v) { return "£" + fmt(v); } } } },
      },
      });
    });

    document.getElementById("bud-renewals").innerHTML = budgetRows
      .slice()
      .sort(function sortBudget(a, b) {
        return a.RenewalDue < b.RenewalDue ? -1 : 1;
      })
      .map(function renderRenewal(row) {
        var color = row.RenewalValue >= 200000 ? COLORS.red : row.RenewalValue >= 50000 ? COLORS.orange : COLORS.teal;
        return (
          '<div style="display:grid;grid-template-columns:1fr auto;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid var(--rule-l)"><div><div style="font-size:12px;font-weight:700;color:var(--text)">' +
          row.Vendor +
          '</div><div style="font-size:10px;color:var(--text-3);margin-top:2px">Due ' +
          row.RenewalDue +
          " · Owner: " +
          row.Owner +
          '</div></div><div style="text-align:right"><div style="font-family:\'Arial Black\',Arial;font-size:14px;font-weight:700;color:' +
          color +
          '">£' +
          row.RenewalValue.toLocaleString() +
          "</div></div></div>"
        );
      })
      .join("");
  }

  function buildRisksPage() {
    var risks = byMonth(D.risks, activeMonth);
    var decisionsNeeded = risks.filter(function filterRisks(risk) {
      return risk.DecisionRequired === "Yes";
    }).length;
    var amberRisks = risks.filter(function filterAmber(risk) {
      return risk.RAG === "Amber";
    }).length;

    document.getElementById("risk-badge").textContent = decisionsNeeded + " decisions required";
    document.getElementById("risk-kpis").innerHTML =
      renderKpiCard("risk-kpi-total-risks", "orange", "Total Risks", risks.length, "active items") +
      renderKpiCard("risk-kpi-decisions-needed", "red", "Decisions Needed", decisionsNeeded, "require board action") +
      renderKpiCard("risk-kpi-amber-risks", "grey", "Amber Risks", amberRisks + '<span class="u"> amber</span>', "current month");

    document.getElementById("risk-rows").innerHTML = risks
      .map(function renderRisk(risk) {
        return (
          '<div class="risk-row"><div><div class="risk-rag ' +
          ragColor(risk.RAG) +
          '"></div></div><div class="risk-title">' +
          risk.RiskIssue +
          '</div><div><span class="risk-badge rb-' +
          (risk.Impact === "High" ? "high" : "medium") +
          '">' +
          risk.Impact +
          '</span></div><div class="risk-cell">' +
          risk.Likelihood +
          '</div><div class="risk-cell" style="font-size:10px">' +
          risk.Owner +
          '</div><div><span class="risk-badge ' +
          (risk.DecisionRequired === "Yes" ? "rb-yes" : "rb-no") +
          '">' +
          (risk.DecisionRequired === "Yes" ? "Yes — needed" : "No") +
          "</span></div></div>"
        );
      })
      .join("");

    document.getElementById("risk-note").innerHTML =
      "<strong>" +
      monthLabel +
      "</strong> contains <strong>" +
      risks.length +
      "</strong> active risk items and <strong>" +
      decisionsNeeded +
      "</strong> board or sponsor decisions. The lowest-performing network site this month was <strong>" +
      latest(D.derivedNetwork).WorstOffice +
      "</strong> and the largest open delivery governance concern remains the development backlog age profile.";
  }

  updateStaticChrome();
  buildExecutivePage();
  buildAvailabilityPage();
  buildNetworkPage();
  buildSupportPage();
  buildSecurityPage();
  buildAssetsPage();
  buildChangePage();
  buildDevelopmentPage();
  buildProjectsPage();
  buildRoadmapPage();
  buildBudgetPage();
  buildRisksPage();

  window.__REPORT_READY = true;

  return {
    showPage: function showPageController(id) {
      showPage(id, null, { silent: true });
    },
    destroy: function destroy() {
      Object.keys(CHARTS).forEach(function destroyChart(id) {
        CHARTS[id].destroy();
        delete CHARTS[id];
      });
      if (options.attachGlobals !== false) {
        delete window.showPage;
        delete window.showMapTip;
        delete window.hideMapTip;
      }
    },
  };
}
