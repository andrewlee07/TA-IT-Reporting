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
  var INITIAL_PAGE_ID = options.initialPageId || "p-summary";
  var INITIAL_TAB_ID = options.initialTabId || null;
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
  var GANTT_DOMAIN_COLOURS = {
    Infrastructure: "#005292",
    "End-user computing": "#F57D00",
    "Security & compliance": "#219D98",
    "Applications & data": "#B45309",
    "Product / development": "#7C3AED",
    "Business transformation": "#A9A9AA",
  };
  var CHARTS = Object.create(null);
  var PAGE_TABS = (D.meta && D.meta.pageTabs) || {};
  var activeTabsByPage = Object.create(null);

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

  function getPageTabs(pageId) {
    return PAGE_TABS[pageId] || [];
  }

  function resolveTabId(pageId, tabId) {
    var tabs = getPageTabs(pageId);
    if (!tabs.length) {
      return null;
    }

    if (tabId && tabs.some(function matchesTab(tab) { return tab.id === tabId; })) {
      return tabId;
    }

    return tabs[0].id;
  }

  function getSlideId(pageId, tabId) {
    var resolvedTabId = resolveTabId(pageId, tabId);
    return resolvedTabId ? pageId + "-" + resolvedTabId : pageId;
  }

  function getActiveTabId(pageId) {
    return resolveTabId(pageId, activeTabsByPage[pageId] || (pageId === INITIAL_PAGE_ID ? INITIAL_TAB_ID : null));
  }

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

  function toMetricNumber(value) {
    if (value === null || typeof value === "undefined" || value === "") {
      return null;
    }

    if (typeof value === "number") {
      return Number.isFinite(value) ? value : null;
    }

    var normalized = parseFloat(String(value).replace(/[^0-9.-]+/g, ""));
    return Number.isFinite(normalized) ? normalized : null;
  }

  function numericMonthSeries(rows, metric) {
    return monthSeries(rows, metric)
      .map(function mapMetric(value) {
        return toMetricNumber(value);
      })
      .filter(function filterMetric(value) {
        return value !== null;
      });
  }

  function roundTo(value, decimals) {
    var precision = typeof decimals === "number" ? decimals : 1;
    var factor = Math.pow(10, precision);
    return Math.round(value * factor) / factor;
  }

  function findChartSetting(chartKey, page) {
    return (D.chartSettings || []).find(function findSetting(setting) {
      return setting.Month === activeMonth && setting.ChartKey === chartKey && (!page || setting.Page === page);
    }) || null;
  }

  function buildRollingAverageSeries(values, windowSize) {
    return values.map(function mapRolling(_, index) {
      var windowValues = values
        .slice(Math.max(0, index - windowSize + 1), index + 1)
        .filter(function filterValue(value) {
          return value !== null;
        });

      if (windowValues.length === 0) {
        return null;
      }

      var total = windowValues.reduce(function reduceTotal(sum, value) {
        return sum + value;
      }, 0);

      return roundTo(total / windowValues.length, 1);
    });
  }

  function buildCloseBalanceSeries(openedSeries, closedSeries) {
    return openedSeries.map(function mapBalance(openedValue, index) {
      var closedValue = closedSeries[index];

      if (openedValue === null || closedValue === null || openedValue <= 0) {
        return null;
      }

      return roundTo((closedValue / openedValue) * 100, 1);
    });
  }

  function overlayHealthColor(value, setting) {
    if (value === null || typeof value === "undefined") {
      return COLORS.grey;
    }

    if (value >= setting.HealthyMin) {
      return COLORS.teal;
    }

    if (value >= setting.AmberMin) {
      return COLORS.orange;
    }

    return COLORS.alert;
  }

  function buildSupportVolumeLegend(setting) {
    var legend = document.getElementById("support-vol-legend");

    if (!legend) {
      return;
    }

    var html =
      '<div class="legend-item"><div class="legend-swatch" style="background:' +
      COLORS.blue +
      '"></div>Opened</div>' +
      '<div class="legend-item"><div class="legend-swatch" style="background:' +
      COLORS.orange +
      '"></div>Closed</div>';

    if (setting && setting.OverlayEnabled === "Yes") {
      html +=
        '<div class="legend-item"><div class="legend-line health"></div>' +
        setting.RollingWindow +
        'M close balance %</div>';
    }

    legend.innerHTML = html;
  }

  function sparkSegmentColor(value, baseColor, cfg) {
    if (!cfg) {
      return baseColor;
    }

    if (cfg.dir === "up") {
      if (value >= cfg.good) {
        return COLORS.teal;
      }
      if (value >= cfg.warn) {
        return COLORS.orange;
      }
      return "#C0392B";
    }

    if (value <= cfg.good) {
      return COLORS.teal;
    }
    if (value <= cfg.warn) {
      return COLORS.orange;
    }
    return "#C0392B";
  }

  function renderSparkline(values, baseColor, cfg, width, height) {
    if (!values || values.length < 2) {
      return "";
    }

    var numericValues = values
      .map(function mapValue(value) {
        return toMetricNumber(value);
      })
      .filter(function filterValue(value) {
        return value !== null;
      });

    if (numericValues.length < 2) {
      return "";
    }

    var minValue = Math.min.apply(null, numericValues);
    var maxValue = Math.max.apply(null, numericValues);
    var range = maxValue - minValue || 1;
    var innerWidth = width || 84;
    var innerHeight = height || 30;
    var points = numericValues.map(function mapPoint(value, index) {
      return {
        x: numericValues.length === 1 ? innerWidth / 2 : (index / (numericValues.length - 1)) * innerWidth,
        y: innerHeight - ((value - minValue) / range) * (innerHeight - 4) - 2,
        value: value,
      };
    });

    var segments = "";
    for (var index = 0; index < points.length - 1; index += 1) {
      var point = points[index];
      var nextPoint = points[index + 1];
      var averageValue = (point.value + nextPoint.value) / 2;
      segments +=
        '<line x1="' +
        point.x.toFixed(1) +
        '" y1="' +
        point.y.toFixed(1) +
        '" x2="' +
        nextPoint.x.toFixed(1) +
        '" y2="' +
        nextPoint.y.toFixed(1) +
        '" stroke="' +
        sparkSegmentColor(averageValue, baseColor, cfg) +
        '" stroke-width="2.4" stroke-linecap="round"/>';
    }

    var dots = points
      .map(function mapDot(point) {
        return (
          '<circle cx="' +
          point.x.toFixed(1) +
          '" cy="' +
          point.y.toFixed(1) +
          '" r="1.65" fill="' +
          sparkSegmentColor(point.value, baseColor, cfg) +
          '"/>'
        );
      })
      .join("");

    return (
      '<div class="kc-spark" aria-hidden="true"><svg viewBox="0 0 ' +
      innerWidth +
      " " +
      innerHeight +
      '" preserveAspectRatio="none">' +
      segments +
      dots +
      "</svg></div>"
    );
  }

  var KPI_SPARK_CONFIG = {
    resolutionSLA: { dir: "up", good: 98.5, warn: 95 },
    csat: { dir: "up", good: 4.6, warn: 4.3 },
    critVulns: { dir: "down", good: 1, warn: 3 },
    changeSuccess: { dir: "up", good: 94, warn: 90 },
    devBacklog: { dir: "down", good: 22, warn: 30 },
    supportOpened: null,
    supportClosed: null,
    supportBacklog: { dir: "down", good: 25, warn: 35 },
    supportResolutionDays: { dir: "down", good: 1.4, warn: 1.7 },
    securityPatch: { dir: "up", good: 95, warn: 92 },
    securityMfa: { dir: "up", good: 99, warn: 97 },
    securityOverdue: { dir: "down", good: 18, warn: 28 },
    assetLifecycle: { dir: "up", good: 82, warn: 78 },
    assetIncidents: { dir: "down", good: 9, warn: 13 },
    changeReleases: { dir: "up", good: 7, warn: 5 },
    changeFailures: { dir: "down", good: 1, warn: 3 },
    changeIncidents: { dir: "down", good: 0, warn: 1 },
    devClosed: { dir: "up", good: 8, warn: 5 },
    devBlocked: { dir: "down", good: 4, warn: 6 },
    devCsat: { dir: "up", good: 4.6, warn: 4.2 },
  };

  function slugify(value) {
    return String(value)
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .replace(/-{2,}/g, "-");
  }

  function parseIsoDate(value) {
    if (!value) {
      return null;
    }

    var parts = String(value).split("-");

    if (parts.length !== 3) {
      return null;
    }

    return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  }

  function addDays(date, days) {
    var nextDate = new Date(date);
    nextDate.setDate(nextDate.getDate() + days);
    return nextDate;
  }

  function dayDiff(start, end) {
    return (end.getTime() - start.getTime()) / (24 * 60 * 60 * 1000);
  }

  function firstMondayOnOrAfter(monthValue) {
    var monthStart = parseIsoDate(monthValue + "-01");
    var dayOfWeek = monthStart.getDay();
    var daysToMonday = dayOfWeek === 1 ? 0 : dayOfWeek === 0 ? 1 : 8 - dayOfWeek;
    return addDays(monthStart, daysToMonday);
  }

  function escapeAttr(value) {
    return String(value).replace(/&/g, "&amp;").replace(/"/g, "&quot;");
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function exportAttrs(id, label) {
    return ' id="' + escapeAttr(id) + '" data-export-id="' + escapeAttr(id) + '" data-export-label="' + escapeAttr(label) + '"';
  }

  function renderKpiCard(id, tone, label, valueHtml, deltaHtml, deltaClass, sparkHtml) {
    return (
      '<div class="kc ' +
      tone +
      (sparkHtml ? " has-spark" : "") +
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
      "</div>" +
      (sparkHtml || "") +
      "</div>"
    );
  }

  function rebuildSlideHeaders() {
    document.querySelectorAll(".report-page").forEach(function rebuildHeader(page) {
      var header = page.querySelector(".ph");
      if (!header) {
        return;
      }

      var pageId = page.getAttribute("data-page-id") || page.id;
      var tabId = page.getAttribute("data-tab-id");
      var titleEl = header.querySelector(".ph-title");
      var subtitleEl = header.querySelector(".ph-sub");
      var periodLabelEl = header.querySelector(".ph-period-label");
      var periodValueEl = header.querySelector(".ph-period-val");

      if (!page.dataset.headerTitle && titleEl) {
        page.dataset.headerTitle = titleEl.textContent.trim();
      }
      if (!page.dataset.headerSubtitle && subtitleEl) {
        page.dataset.headerSubtitle = subtitleEl.textContent.trim();
      }
      if (!page.dataset.headerPeriodLabel && periodLabelEl) {
        page.dataset.headerPeriodLabel = periodLabelEl.textContent.trim();
      }
      if (!page.dataset.headerPeriodValue && periodValueEl) {
        page.dataset.headerPeriodValue = periodValueEl.textContent.trim();
      }
      if (!page.dataset.headerPeriodValueId && periodValueEl && periodValueEl.id) {
        page.dataset.headerPeriodValueId = periodValueEl.id;
      }

      var title = page.dataset.headerTitle || "";
      var subtitle = (subtitleEl ? subtitleEl.textContent : page.dataset.headerSubtitle) || "";
      var periodLabel = (periodLabelEl ? periodLabelEl.textContent : page.dataset.headerPeriodLabel) || "";
      var periodValue = (periodValueEl ? periodValueEl.textContent : page.dataset.headerPeriodValue) || "";
      var periodValueId = page.dataset.headerPeriodValueId || "";
      var activeTab = tabId
        ? getPageTabs(pageId).find(function matchesTab(tab) {
            return tab.id === tabId;
          }) || null
        : null;
      var accentHtml = activeTab
        ? '<span class="ph-title-sep">-</span><span class="ph-title-accent">' + escapeHtml(activeTab.label) + "</span>"
        : "";
      var metaLine = "";

      if (subtitle) {
        metaLine += '<span class="ph-sub">' + escapeHtml(subtitle) + "</span>";
      }

      if (periodLabel && periodValue) {
        if (metaLine) {
          metaLine += '<span class="ph-meta-divider">·</span>';
        }
        metaLine +=
          '<span class="ph-period-label">' +
          escapeHtml(periodLabel) +
          '</span><span class="ph-period-val"' +
          (periodValueId ? ' id="' + escapeAttr(periodValueId) + '"' : "") +
          ">" +
          escapeHtml(periodValue) +
          "</span>";
      }

      header.innerHTML =
        '<div class="ph-stack">' +
        '<div class="ph-title-line"><span class="ph-title">' +
        escapeHtml(title) +
        "</span>" +
        accentHtml +
        "</div>" +
        '<div class="ph-meta-line">' +
        metaLine +
        "</div></div>";
    });
  }

  function updateStaticChrome() {
    document.title = "TeacherActive — IT Reporting · " + monthLabel;

    document.querySelectorAll(".report-page").forEach(function updatePeriod(page) {
      var labelEl = page.querySelector(".ph-period-label");
      var valueEl = page.querySelector(".ph-period-val");
      if (labelEl && valueEl && labelEl.textContent.trim() === "Reporting Period") {
        valueEl.textContent = monthLabel;
      }
    });

    rebuildSlideHeaders();

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
        page.style.display = "flex";
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

    var initialTabId = getActiveTabId(INITIAL_PAGE_ID);
    var page = document.getElementById(getSlideId(INITIAL_PAGE_ID, initialTabId));
    if (page) {
      page.classList.add("active");
    }

    var nav = document.querySelector('.nav-link[data-page-id="' + INITIAL_PAGE_ID + '"]');
    if (nav) {
      nav.classList.add("active");
    }
  }

  function ensureTabSlot(pageId, tabId) {
    var slideId = getSlideId(pageId, tabId);
    var page = document.getElementById(slideId);
    if (!page) {
      return null;
    }

    var slot = page.querySelector(".slide-tabs-slot");
    if (slot) {
      return slot;
    }

    var createdSlot = nativeDocument.createElement("div");
    createdSlot.className = "slide-tabs-slot";
    createdSlot.setAttribute("data-page-id", pageId);
    var body = page.querySelector(".pb");
    if (body && body.parentNode) {
      body.parentNode.insertBefore(createdSlot, body);
    } else {
      page.appendChild(createdSlot);
    }
    return createdSlot;
  }

  function renderPageTabs(pageId) {
    var tabs = getPageTabs(pageId);
    if (!tabs.length) {
      return;
    }

    tabs.forEach(function renderTabSlot(tab) {
      var slot = ensureTabSlot(pageId, tab.id);
      if (!slot) {
        return;
      }

      slot.innerHTML =
        '<div class="slide-tabs" role="tablist" aria-label="' +
        pageId +
        ' tabs">' +
        tabs
          .map(function renderButton(entry) {
            var isActive = entry.id === tab.id;
            return (
              '<button class="slide-tab' +
              (isActive ? " active" : "") +
              '" role="tab" aria-selected="' +
              (isActive ? "true" : "false") +
              `" onclick='showPageTab(` +
              JSON.stringify(pageId) +
              "," +
              JSON.stringify(entry.id) +
              `)'>` +
              entry.label +
              "</button>"
            );
          })
          .join("") +
        "</div>";
    });
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
      case "p-summary":
        buildSummaryPage();
        break;
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
      case "p-gantt":
        buildGanttPage();
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

  function buildSummaryPage() {
    var summary = D.execSummary || {
      mode: "empty",
      contentHtml: "",
      excerpt: "",
      updatedAt: null,
      sourceReportId: null,
    };
    var stateBadge = document.getElementById("summary-state-badge");
    var updatedAt = document.getElementById("summary-updated-at");
    var content = document.getElementById("summary-content");
    var emptyState = document.getElementById("summary-empty-state");

    if (!stateBadge || !updatedAt || !content || !emptyState) {
      return;
    }

    var emptyTitle = emptyState.querySelector(".summary-empty-title");
    var emptyCopy = emptyState.querySelector(".summary-empty-copy");
    var badgeLabel = "Saved summary";
    var updatedLabel = "";
    var showContent = false;
    var showEmptyState = false;

    stateBadge.className = "summary-state-badge " + summary.mode;

    if (summary.mode === "loading") {
      badgeLabel = "Loading";
      showEmptyState = true;
      if (emptyTitle) {
        emptyTitle.textContent = "Loading exec summary";
      }
      if (emptyCopy) {
        emptyCopy.textContent = "Fetching the latest narrative for this report and reporting month.";
      }
    } else if (summary.mode === "empty") {
      badgeLabel = "No summary yet";
      showEmptyState = true;
      if (emptyTitle) {
        emptyTitle.textContent = "No exec summary has been written yet";
      }
      if (emptyCopy) {
        emptyCopy.textContent =
          "Use this page to add a concise leadership narrative for the selected month. This text is stored in the app, not the workbook, so you can keep the structured reporting model intact.";
      }
    } else if (summary.mode === "carried-forward") {
      badgeLabel = "Inherited draft";
      showContent = true;
      if (summary.updatedAt) {
        updatedLabel = "Inherited from a prior report · last updated " + new Date(summary.updatedAt).toLocaleString("en-GB", {
          day: "numeric",
          month: "short",
          year: "numeric",
          hour: "2-digit",
          minute: "2-digit",
        });
      }
    } else if (summary.mode === "demo-readonly") {
      badgeLabel = "Bundled example";
      showContent = true;
      if (summary.updatedAt) {
        updatedLabel = "Read-only demo narrative · refreshed " + new Date(summary.updatedAt).toLocaleDateString("en-GB", {
          day: "numeric",
          month: "short",
          year: "numeric",
        });
      }
    } else {
      badgeLabel = "Saved summary";
      showContent = true;
      if (summary.updatedAt) {
        updatedLabel = "Last updated " + new Date(summary.updatedAt).toLocaleString("en-GB", {
          day: "numeric",
          month: "short",
          year: "numeric",
          hour: "2-digit",
          minute: "2-digit",
        });
      }
    }

    stateBadge.textContent = badgeLabel;
    updatedAt.textContent = updatedLabel;
    content.innerHTML = showContent ? summary.contentHtml || "" : "";
    content.style.display = showContent ? "block" : "none";
    emptyState.classList.toggle("active", showEmptyState);
  }

  function showPage(id, el, runtimeOptions) {
    if (SHOW_ALL_PAGES) {
      return;
    }

    var nextTabId = resolveTabId(id, runtimeOptions && Object.prototype.hasOwnProperty.call(runtimeOptions, "tabId") ? runtimeOptions.tabId : getActiveTabId(id));
    activeTabsByPage[id] = nextTabId;

    document.querySelectorAll(".report-page").forEach(function deactivatePage(page) {
      page.classList.remove("active");
    });
    document.querySelectorAll(".nav-link").forEach(function clearNav(link) {
      link.classList.remove("active");
    });

    var page = document.getElementById(getSlideId(id, nextTabId));
    if (page) {
      page.classList.add("active");
    }

    var activeNav = el || document.querySelector('.nav-link[data-page-id="' + id + '"]');
    if (activeNav) {
      activeNav.classList.add("active");
    }

    rebuildVisiblePage(id);
    renderPageTabs(id);
    window.scrollTo(0, 0);

    if ((!runtimeOptions || runtimeOptions.silent !== true) && typeof options.onPageChange === "function") {
      options.onPageChange(id, nextTabId);
    }

  }

  if (options.attachGlobals !== false) {
    window.showPage = showPage;
    window.showPageTab = function showPageTab(pageId, tabId) {
      showPage(pageId, null, { tabId: tabId });
    };
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
    var execSupportSlaSpark = renderSparkline(numericMonthSeries(D.support, "ResolutionSLA"), COLORS.blue, KPI_SPARK_CONFIG.resolutionSLA);
    var execCsatSpark = renderSparkline(numericMonthSeries(D.support, "CSAT"), COLORS.teal, KPI_SPARK_CONFIG.csat);
    var execCriticalVulnSpark = renderSparkline(numericMonthSeries(D.security, "CritVulns"), COLORS.teal, KPI_SPARK_CONFIG.critVulns);
    var execChangeSpark = renderSparkline(numericMonthSeries(D.change, "SuccessRate"), COLORS.orange, KPI_SPARK_CONFIG.changeSuccess);
    var execBacklogSpark = renderSparkline(numericMonthSeries(D.dev, "BacklogEnd"), COLORS.grey, KPI_SPARK_CONFIG.devBacklog);

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
          execSupportSlaSpark,
        ) +
        renderKpiCard(
          "exec-kpi-user-csat",
          "teal",
          "User CSAT",
          support.CSAT,
          "▲ " + (pctNum(support.CSAT) - pctNum(supportPrev.CSAT)).toFixed(1) + " vs prior month",
          "up",
          execCsatSpark,
        ) +
        renderKpiCard(
          "exec-kpi-critical-vulns",
          "green",
          "Critical Vulns",
          security.CritVulns,
          security.CritVulns <= securityPrev.CritVulns ? "▼ reduced backlog" : "▲ increased backlog",
          "up",
          execCriticalVulnSpark,
        ) +
        renderKpiCard(
          "exec-kpi-change-success",
          "orange",
          "Change Success",
          change.SuccessRate.replace("%", "") + '<span class="u">%</span>',
          "▲ " + (pctNum(change.SuccessRate) - pctNum(changePrev.SuccessRate)).toFixed(1) + " pts vs prior month",
          "up",
          execChangeSpark,
        ) +
        renderKpiCard(
          "exec-kpi-dev-backlog",
          "grey",
          "Dev Backlog",
          dev.BacklogEnd,
          dev.BacklogEnd <= devPrev.BacklogEnd ? "▼ reduced backlog" : "▲ increased backlog",
          "up",
          execBacklogSpark,
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
    var networkMetric =
      D.derivedNetwork.find(function findMetric(item) {
        return item.Month === activeMonth;
      }) || latest(D.derivedNetwork);
    var averageAvailability = pctNum(networkMetric.Availability);
    var perfect = networkMetric.PerfectOffices;
    var below99 = networkMetric.Below99Offices;
    var below999 = networkMetric.Below99_9Offices;
    var netKpis = document.getElementById("net-kpis-map");
    var mapBadge = document.getElementById("net-map-badge");
    var dots = document.getElementById("office-dots");
    var officeList = document.getElementById("office-list");
    var detailNote = document.getElementById("network-detail-note");
    var sortedOffices = officeRows
      .slice()
      .sort(function sortOffice(a, b) {
        return pctNum(b.Availability) - pctNum(a.Availability);
      });
    var worstOffice = sortedOffices.length ? sortedOffices[sortedOffices.length - 1] : null;

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
      officeList.innerHTML = sortedOffices
        .map(function renderOffice(office) {
          var pct = pctNum(office.Availability);
          var color = pct === 100 ? COLORS.teal : pct >= 99.9 ? COLORS.orange : COLORS.alert;
          var label = pct === 100 ? "Perfect" : pct >= 99.9 ? "Good" : pct >= 99 ? "Minor" : "Impacted";
          return (
            '<div class="office-list-item">' +
            '<div class="office-list-meta">' +
            '<div class="office-list-name">' +
            '<span class="office-list-dot" style="background:' +
            color +
            ';"></span>' +
            "<strong>" +
            office.OfficeName +
            "</strong>" +
            "<span>" +
            office.Region +
            "</span>" +
            "</div>" +
            '<div class="office-list-track"><div class="office-list-fill" style="width:' +
            Math.max(0, ((pct - 97) / 3) * 100).toFixed(1) +
            "%;background:" +
            color +
            ';"></div></div>' +
            "</div>" +
            '<div class="office-list-value"><strong style="color:' +
            color +
            '">' +
            office.Availability +
            "</strong><span>" +
            label +
            "</span></div>" +
            "</div>"
          );
        })
        .join("");
    }

    if (detailNote) {
      detailNote.innerHTML =
        "<strong>" +
        monthLabel +
        "</strong> closed with estate-wide average availability of <strong>" +
        averageAvailability.toFixed(2) +
        "%</strong>. " +
        (below99 === 0
          ? "No offices fell below the 99% material-impact threshold, and "
          : "<strong>" + below99 + "</strong> offices fell below 99%, while ") +
        "<strong>" +
        below999 +
        "</strong> offices sat below 99.9%. " +
        (worstOffice
          ? "Lowest-performing site was <strong>" +
            worstOffice.OfficeName +
            "</strong> at <strong>" +
            worstOffice.Availability +
            "</strong>, so that office remains the priority focus for follow-up."
          : "All offices remained at or above the expected service range.");
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
    renderPageTabs("p-support");

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
      renderKpiCard(
        "support-kpi-opened",
        "blue",
        "Opened",
        support.Opened.toLocaleString(),
        "tickets this month",
        "",
        renderSparkline(numericMonthSeries(D.support, "Opened"), COLORS.blue, KPI_SPARK_CONFIG.supportOpened),
      ) +
      renderKpiCard(
        "support-kpi-closed",
        "orange",
        "Closed",
        support.Closed.toLocaleString(),
        support.Closed >= support.Opened ? "▲ net positive flow" : "▼ net negative flow",
        "up",
        renderSparkline(numericMonthSeries(D.support, "Closed"), COLORS.orange, KPI_SPARK_CONFIG.supportClosed),
      ) +
      renderKpiCard(
        "support-kpi-backlog",
        "teal",
        "Backlog End",
        support.Backlog,
        support.Backlog <= supportPrev.Backlog ? "▼ lower than prior month" : "▲ higher than prior month",
        "up",
        renderSparkline(numericMonthSeries(D.support, "Backlog"), COLORS.teal, KPI_SPARK_CONFIG.supportBacklog),
      ) +
      renderKpiCard(
        "support-kpi-avg-resolution",
        "green",
        "Avg Resolution",
        support.AvgResolution + '<span class="u"> days</span>',
        support.AvgResolution <= supportPrev.AvgResolution ? "▲ improved turnaround" : "▼ slower turnaround",
        "up",
        renderSparkline(numericMonthSeries(D.support, "AvgResolution"), COLORS.teal, KPI_SPARK_CONFIG.supportResolutionDays),
      ) +
      renderKpiCard(
        "support-kpi-major-incidents",
        "grey",
        "Major Incidents",
        support.MajorIncidents,
        support.MajorIncidents === 0 ? "clean month" : "incident activity recorded",
      );

    registerChart("c-support-vol", function createSupportVolumeChart() {
      var supportChartSetting = findChartSetting("support_ticket_volumes", "Support Operations");
      var overlayEnabled = Boolean(
        supportChartSetting &&
          supportChartSetting.OverlayEnabled === "Yes" &&
          supportChartSetting.OverlayMetric === "Close Balance %",
      );
      var rollingWindow = overlayEnabled ? Math.max(1, Number(supportChartSetting.RollingWindow) || 3) : 0;
      var openedSeries = monthSeries(D.support, "Opened").map(function mapOpened(value) {
        return toMetricNumber(value);
      });
      var closedSeries = monthSeries(D.support, "Closed").map(function mapClosed(value) {
        return toMetricNumber(value);
      });
      var closeBalanceSeries = buildCloseBalanceSeries(openedSeries, closedSeries);
      var rollingCloseBalanceSeries = overlayEnabled ? buildRollingAverageSeries(closeBalanceSeries, rollingWindow) : [];
      var overlayValues = rollingCloseBalanceSeries.filter(function filterValue(value) {
        return value !== null;
      });
      var overlayMin = overlayValues.length
        ? Math.min.apply(null, overlayValues.concat([Number(supportChartSetting.AmberMin) - 2]))
        : Number(supportChartSetting ? supportChartSetting.AmberMin : 95) - 2;
      var overlayMax = overlayValues.length
        ? Math.max.apply(null, overlayValues.concat([Number(supportChartSetting.HealthyMin) + 2]))
        : Number(supportChartSetting ? supportChartSetting.HealthyMin : 100) + 2;
      var supportTooltip = {
        backgroundColor: TIP.backgroundColor,
        borderColor: TIP.borderColor,
        borderWidth: TIP.borderWidth,
        titleColor: TIP.titleColor,
        bodyColor: TIP.bodyColor,
        padding: TIP.padding,
        cornerRadius: TIP.cornerRadius,
        callbacks: {
          label: function label(context) {
            if (context.dataset.yAxisID === "yBalance") {
              return context.dataset.label + ": " + Number(context.parsed.y).toFixed(1) + "%";
            }

            return context.dataset.label + ": " + fmt(context.parsed.y);
          },
          afterBody: function afterBody(items) {
            if (!items || items.length === 0) {
              return [];
            }

            var itemIndex = items[0].dataIndex;
            var opened = openedSeries[itemIndex];
            var closed = closedSeries[itemIndex];
            var closeBalance = closeBalanceSeries[itemIndex];
            var rollingBalance = overlayEnabled ? rollingCloseBalanceSeries[itemIndex] : null;
            var lines = [];

            if (opened !== null && closed !== null) {
              var netFlow = closed - opened;
              lines.push("Net flow: " + (netFlow > 0 ? "+" : "") + fmt(netFlow));
            }

            if (closeBalance !== null) {
              lines.push("Close balance: " + closeBalance.toFixed(1) + "%");
            }

            if (rollingBalance !== null) {
              lines.push(rollingWindow + "M rolling close balance: " + rollingBalance.toFixed(1) + "%");
            }

            return lines;
          },
        },
      };

      buildSupportVolumeLegend(supportChartSetting);

      var datasets = [
        { label: "Opened", data: openedSeries, backgroundColor: COLORS.blue, borderRadius: 3, barPercentage: 0.45, order: 2 },
        { label: "Closed", data: closedSeries, backgroundColor: COLORS.orange, borderRadius: 3, barPercentage: 0.45, order: 3 },
      ];

      if (overlayEnabled) {
        datasets.push({
          type: "line",
          label: rollingWindow + "M close balance %",
          data: rollingCloseBalanceSeries,
          yAxisID: "yBalance",
          borderWidth: 2,
          borderDash: [6, 4],
          tension: 0.32,
          spanGaps: true,
          pointRadius: 2.5,
          pointHoverRadius: 4,
          pointBorderWidth: 0,
          fill: false,
          order: 1,
          segment: {
            borderColor: function borderColor(context) {
              var startY = context.p0 && context.p0.parsed ? context.p0.parsed.y : null;
              var endY = context.p1 && context.p1.parsed ? context.p1.parsed.y : null;
              var average = startY === null || endY === null ? (startY !== null ? startY : endY) : (startY + endY) / 2;
              return overlayHealthColor(average, supportChartSetting);
            },
          },
          pointBackgroundColor: function pointBackgroundColor(context) {
            return overlayHealthColor(context.raw, supportChartSetting);
          },
          pointHoverBackgroundColor: function pointHoverBackgroundColor(context) {
            return overlayHealthColor(context.raw, supportChartSetting);
          },
        });
      }

      return new Chart(resetChartCanvas("c-support-vol"), {
        type: "bar",
        data: {
          labels: visibleLabels,
          datasets: datasets,
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false }, tooltip: supportTooltip },
          scales: {
            x: { grid: { display: false } },
            y: { grid: GRID, border: { display: false }, ticks: { callback: fmt } },
            yBalance: overlayEnabled
              ? {
                  position: "right",
                  grid: { display: false },
                  border: { display: false },
                  min: Math.floor(overlayMin),
                  max: Math.ceil(overlayMax),
                  ticks: {
                    callback: function balanceTick(value) {
                      return value + "%";
                    },
                    color: COLORS.teal,
                  },
                }
              : { display: false, grid: { display: false }, border: { display: false } },
          },
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
    var sortedCategories = Object.entries(categories).sort(function sortCategories(a, b) {
      return b[1] - a[1];
    });
    var criticalTickets = tickets.filter(function filterCritical(ticket) {
      return ticket.BusinessCritical === "Yes";
    }).length;
    var supportDetailNote = document.getElementById("support-detail-note");

    if (supportDetailNote) {
      var topCategory = sortedCategories.length ? sortedCategories[0][0] : support.TopCategory;
      var ageingPressure = tickets.filter(function filterAgeing(ticket) {
        return ticket.AgeDays >= 40;
      }).length;
      supportDetailNote.innerHTML =
        "<strong>" +
        topCategory +
        "</strong> remains the highest-volume support theme in the visible period. " +
        (criticalTickets > 0
          ? criticalTickets +
            " business-critical ticket" +
            (criticalTickets === 1 ? " is" : "s are") +
            " still open, with the oldest ticket sitting at <strong>" +
            oldestTicket +
            " days</strong>."
          : "There are no business-critical tickets currently open.") +
        " " +
        ageingPressure +
        " ticket" +
        (ageingPressure === 1 ? " is" : "s are") +
        " older than 40 days, which is the primary ageing pressure on the queue.";
    }

    document.getElementById("support-cats").innerHTML = sortedCategories
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
        renderSparkline(numericMonthSeries(D.security, "CritVulns"), COLORS.teal, KPI_SPARK_CONFIG.critVulns),
      ) +
      renderKpiCard(
        "sec-kpi-workstation-patch",
        "blue",
        "Workstation Patch",
        security.WkstationPatch.replace("%", "") + '<span class="u">%</span>',
        pctNum(security.WkstationPatch) >= pctNum(previousSecurity.WkstationPatch) ? "▲ improving" : "▼ lower than prior month",
        "up",
        renderSparkline(numericMonthSeries(D.security, "WkstationPatch"), COLORS.blue, KPI_SPARK_CONFIG.securityPatch),
      ) +
      renderKpiCard(
        "sec-kpi-mfa-coverage",
        "teal",
        "MFA Coverage",
        security.MFACoverage.replace("%", "") + '<span class="u">%</span>',
        "▲ near full coverage",
        "up",
        renderSparkline(numericMonthSeries(D.security, "MFACoverage"), COLORS.teal, KPI_SPARK_CONFIG.securityMfa),
      ) +
      renderKpiCard(
        "sec-kpi-overdue-remediation",
        "orange",
        "Overdue Remediation",
        security.OverdueRemediation,
        security.OverdueRemediation <= previousSecurity.OverdueRemediation ? "▼ reduced backlog" : "▲ higher than prior month",
        "up",
        renderSparkline(numericMonthSeries(D.security, "OverdueRemediation"), COLORS.orange, KPI_SPARK_CONFIG.securityOverdue),
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
      renderKpiCard(
        "asset-kpi-laptops-in-lifecycle",
        "teal",
        "Laptops in Lifecycle",
        laptop.PctWithin,
        "▲ lifecycle coverage",
        "up",
        renderSparkline(
          laptopSeries.map(function mapLifecycle(asset) {
            return asset.PctWithin;
          }),
          COLORS.teal,
          KPI_SPARK_CONFIG.assetLifecycle,
        ),
      ) +
      renderKpiCard(
        "asset-kpi-laptop-incidents",
        "orange",
        "Laptop Incidents",
        laptop.IncidentsLinked,
        "▼ linked hardware incidents",
        "up",
        renderSparkline(
          laptopSeries.map(function mapIncident(asset) {
            return asset.IncidentsLinked;
          }),
          COLORS.orange,
          KPI_SPARK_CONFIG.assetIncidents,
        ),
      ) +
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
        renderSparkline(numericMonthSeries(D.change, "ReleasesDeployed"), COLORS.teal, KPI_SPARK_CONFIG.changeReleases),
      ) +
      renderKpiCard(
        "change-kpi-failed-changes",
        "orange",
        "Failed Changes",
        change.FailedChanges,
        change.FailedChanges <= changePrev.FailedChanges ? "▼ improved" : "▲ worsened",
        "up",
        renderSparkline(numericMonthSeries(D.change, "FailedChanges"), COLORS.orange, KPI_SPARK_CONFIG.changeFailures),
      ) +
      renderKpiCard(
        "change-kpi-incidents",
        "green",
        "Changes → Incidents",
        change.ChangesIncidents,
        change.ChangesIncidents === 0 ? "no service impact" : "service impact recorded",
        "up",
        renderSparkline(numericMonthSeries(D.change, "ChangesIncidents"), COLORS.teal, KPI_SPARK_CONFIG.changeIncidents),
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
        renderSparkline(numericMonthSeries(D.dev, "BacklogEnd"), COLORS.blue, KPI_SPARK_CONFIG.devBacklog),
      ) +
      renderKpiCard(
        "dev-kpi-tasks-closed",
        "teal",
        "Tasks Closed",
        dev.Closed,
        dev.Closed >= devPrev.Closed ? "▲ improved throughput" : "▼ lower throughput",
        "up",
        renderSparkline(numericMonthSeries(D.dev, "Closed"), COLORS.teal, KPI_SPARK_CONFIG.devClosed),
      ) +
      renderKpiCard(
        "dev-kpi-blocked-items",
        "orange",
        "Blocked Items",
        dev.Blocked,
        dev.Blocked <= devPrev.Blocked ? "▼ fewer blockers" : "▲ more blockers",
        "up",
        renderSparkline(numericMonthSeries(D.dev, "Blocked"), COLORS.orange, KPI_SPARK_CONFIG.devBlocked),
      ) +
      renderKpiCard(
        "dev-kpi-csat",
        "green",
        "Dev CSAT",
        dev.CSAT,
        pctNum(dev.CSAT) >= pctNum(devPrev.CSAT) ? "▲ sponsor sentiment improving" : "▼ sponsor sentiment lower",
        "up",
        renderSparkline(numericMonthSeries(D.dev, "CSAT"), COLORS.teal, KPI_SPARK_CONFIG.devCsat),
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

  function buildGanttPage() {
    var chartBlock = document.getElementById("gantt-chart-block");
    var legend = chartBlock ? chartBlock.querySelector(".gantt-legend") : null;
    var wrap = chartBlock ? chartBlock.querySelector(".gantt-wrap") : null;
    var svg = document.getElementById("gantt-svg");
    var summary = document.getElementById("gantt-summary");
    var periodLabel = document.getElementById("gantt-period-label");
    var subtitle = document.getElementById("gantt-sub");
    var emptyState = document.getElementById("gantt-empty-state");
    var emptyCopy = document.getElementById("gantt-empty-copy");
    var tooltip = document.getElementById("gantt-tooltip");
    var openDemoLink = document.getElementById("gantt-open-demo-link");
    var uploadButton = document.getElementById("gantt-upload-btn");

    if (!svg || !summary || !periodLabel || !subtitle || !chartBlock) {
      return;
    }

    var workstreams = byMonth(D.ganttWorkstreams, activeMonth)
      .filter(function filterWorkstream(item) {
        return item.InScope === true;
      })
      .slice()
      .sort(function sortWorkstreams(left, right) {
        return left.DisplayOrder - right.DisplayOrder || String(left.WorkstreamName).localeCompare(String(right.WorkstreamName));
      });
    var milestones = byMonth(D.ganttMilestones, activeMonth).slice().sort(function sortMilestones(left, right) {
      return left.DisplayOrder - right.DisplayOrder || String(left.MilestoneDate).localeCompare(String(right.MilestoneDate));
    });
    var milestonesByWorkstream = milestones.reduce(function reduceMilestones(map, milestone) {
      var key = String(milestone.WorkstreamName);
      var values = map[key] || [];
      values.push(milestone);
      map[key] = values;
      return map;
    }, Object.create(null));

    var WEEKS = 12;
    var LEFT_W = 210;
    var ROW_H = 40;
    var ROW_PAD = 7;
    var BAR_H = ROW_H - ROW_PAD * 2;
    var HEADER_H = 46;
    var FOOTER_H = 8;
    var WEEK_W = 74;
    var CHART_W = LEFT_W + WEEKS * WEEK_W;
    var CHART_H = HEADER_H + workstreams.length * ROW_H + FOOTER_H;
    var CORNER_R = 6;

    var baseDate = firstMondayOnOrAfter(activeMonth);
    var windowEnd = addDays(baseDate, WEEKS * 7);
    var cutOffDate = parseIsoDate((D.meta.reportCutOffDates && D.meta.reportCutOffDates[activeMonth]) || "");
    var cutOffOffsetW = cutOffDate ? dayDiff(baseDate, cutOffDate) / 7 : null;
    var templateVersion = Number((D.meta && D.meta.templateVersion) || 0);
    var hasAnyGanttSourceData = Array.isArray(D.ganttWorkstreams) && D.ganttWorkstreams.length > 0;
    var hoverItems = Object.create(null);

    function clamp(value, min, max) {
      return Math.max(min, Math.min(max, value));
    }

    function hideGanttTooltip() {
      if (!tooltip) {
        return;
      }
      tooltip.classList.remove("active");
      tooltip.setAttribute("aria-hidden", "true");
    }

    function formatTooltipDate(date) {
      return date
        ? date.toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" })
        : "—";
    }

    function statusCopy(value) {
      if (value === "Green") {
        return "On track";
      }
      if (value === "Amber") {
        return "At risk";
      }
      if (value === "Red") {
        return "Off track";
      }
      return value || "Unknown";
    }

    function renderTooltipMarkup(meta) {
      var content = '<div class="gantt-tooltip-kicker">' + escapeHtml(meta.kicker) + "</div>";
      content += '<div class="gantt-tooltip-title">' + escapeHtml(meta.title) + "</div>";

      if (meta.sub) {
        content += '<div class="gantt-tooltip-sub">' + escapeHtml(meta.sub) + "</div>";
      }

      if (meta.rows && meta.rows.length) {
        content += '<div class="gantt-tooltip-grid">';
        meta.rows.forEach(function renderRow(row) {
          content += '<div class="gantt-tooltip-key">' + escapeHtml(row.label) + '</div><div class="gantt-tooltip-val">' + escapeHtml(row.value) + "</div>";
        });
        content += "</div>";
      }

      if (meta.body) {
        content += '<div class="gantt-tooltip-body">' + escapeHtml(meta.body) + "</div>";
      }

      return content;
    }

    function placeGanttTooltip(event, meta) {
      if (!tooltip) {
        return;
      }

      tooltip.style.setProperty("--tooltip-accent", meta.accent || COLORS.blue);
      tooltip.innerHTML = renderTooltipMarkup(meta);
      tooltip.classList.add("active");
      tooltip.setAttribute("aria-hidden", "false");

      var panelRect = chartBlock.getBoundingClientRect();
      var tooltipRect = tooltip.getBoundingClientRect();
      var left = event.clientX - panelRect.left + 16;
      var top = event.clientY - panelRect.top + 18;

      left = clamp(left, 10, Math.max(10, panelRect.width - tooltipRect.width - 10));
      top = clamp(top, 10, Math.max(10, panelRect.height - tooltipRect.height - 10));

      tooltip.style.left = left + "px";
      tooltip.style.top = top + "px";
    }

    function xForWeek(weekOffset) {
      return LEFT_W + weekOffset * WEEK_W;
    }

    var weekDates = Array.from({ length: WEEKS }, function mapWeek(_, index) {
      return addDays(baseDate, index * 7);
    });

    var monthGroups = {};
    weekDates.forEach(function groupMonth(date, index) {
      var label = date.toLocaleDateString("en-GB", { month: "short", year: "numeric" });
      if (!monthGroups[label]) {
        monthGroups[label] = { start: index, count: 0 };
      }
      monthGroups[label].count += 1;
    });

    var visibleWorkstreams = workstreams.filter(function filterVisibleWorkstream(item) {
      var startDate = parseIsoDate(item.StartDate);
      var endDateExclusive = addDays(parseIsoDate(item.EndDate), 1);
      return startDate < windowEnd && endDateExclusive > baseDate;
    });
    var visibleMilestones = milestones.filter(function filterVisibleMilestone(milestone) {
      var milestoneDate = parseIsoDate(milestone.MilestoneDate);
      var milestoneOffset = dayDiff(baseDate, milestoneDate) / 7;
      return milestoneOffset >= 0 && milestoneOffset <= WEEKS;
    });
    var completing = visibleWorkstreams.filter(function filterCompleting(item) {
      var endDate = parseIsoDate(item.EndDate);
      return endDate >= baseDate && endDate <= windowEnd;
    }).length;
    var onTrack = visibleWorkstreams.filter(function filterOnTrack(item) {
      return item.StatusRAG === "Green";
    }).length;
    var atRisk = visibleWorkstreams.filter(function filterAtRisk(item) {
      return item.StatusRAG === "Amber";
    }).length;

    periodLabel.textContent =
      baseDate.toLocaleDateString("en-GB", { day: "numeric", month: "short" }) +
      " – " +
      windowEnd.toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
    if (openDemoLink) {
      openDemoLink.setAttribute("href", "/?report=demo&month=" + activeMonth + "&page=p-gantt");
    }
    if (uploadButton) {
      uploadButton.onclick = function requestUpload() {
        if (window && typeof window.dispatchEvent === "function" && typeof window.CustomEvent === "function") {
          window.dispatchEvent(new window.CustomEvent("ta:request-upload"));
        }
      };
    }

    hideGanttTooltip();

    if (!workstreams.length && !milestones.length) {
      subtitle.textContent =
        templateVersion >= 3
          ? "No in-scope workstreams were provided for this reporting period."
          : "This workbook version does not include Portfolio Gantt inputs.";

      if (emptyCopy) {
        emptyCopy.textContent =
          templateVersion >= 3 && hasAnyGanttSourceData
            ? "This workbook does not include any in-scope Gantt workstreams for the selected month. Open the bundled demo for a populated example, or upload a refreshed v4 workbook with Gantt inputs."
            : "The active report was created from a legacy workbook that predates the Portfolio Gantt sheets. Open the bundled demo to see the page populated, or upload a v4 workbook to render this view from your own data.";
      }

      if (legend) {
        legend.style.display = "none";
      }
      if (wrap) {
        wrap.style.display = "none";
      }
      if (summary) {
        summary.style.display = "none";
        summary.innerHTML = "";
      }
      if (emptyState) {
        emptyState.classList.add("active");
      }
      svg.innerHTML = "";
      return;
    }

    if (legend) {
      legend.style.display = "";
    }
      if (wrap) {
        wrap.style.display = "";
        wrap.onscroll = hideGanttTooltip;
      }
      if (summary) {
        summary.style.display = "";
      }
    if (emptyState) {
      emptyState.classList.remove("active");
    }

    svg.setAttribute("width", String(CHART_W));
    svg.setAttribute("height", String(CHART_H));
    svg.setAttribute("viewBox", "0 0 " + CHART_W + " " + CHART_H);
    svg.style.minWidth = CHART_W + "px";

    var html =
      '<defs><clipPath id="gantt-clip"><rect x="' +
      LEFT_W +
      '" y="0" width="' +
      WEEKS * WEEK_W +
      '" height="' +
      CHART_H +
      '"/></clipPath><filter id="gantt-shadow" x="-5%" y="-10%" width="110%" height="130%"><feDropShadow dx="0" dy="1" stdDeviation="2" flood-color="rgba(0,0,0,0.10)"/></filter></defs>';
    html += '<rect width="' + CHART_W + '" height="' + CHART_H + '" fill="#FFFFFF"/>';

    workstreams.forEach(function renderRowBand(_, index) {
      var y = HEADER_H + index * ROW_H;
      html +=
        '<rect x="0" y="' +
        y +
        '" width="' +
        CHART_W +
        '" height="' +
        ROW_H +
        '" fill="' +
        (index % 2 === 0 ? "#FAFAFA" : "#FFFFFF") +
        '"/>';
    });

    for (var weekIndex = 0; weekIndex <= WEEKS; weekIndex += 1) {
      var x = xForWeek(weekIndex);
      var isMonthBoundary = weekIndex > 0 && weekDates[weekIndex] && weekDates[weekIndex - 1] && weekDates[weekIndex].getMonth() !== weekDates[weekIndex - 1].getMonth();
      html +=
        '<line x1="' +
        x +
        '" y1="' +
        (HEADER_H - 8) +
        '" x2="' +
        x +
        '" y2="' +
        (CHART_H - FOOTER_H) +
        '" stroke="' +
        (isMonthBoundary ? "#CBD5E1" : "#F0F2F5") +
        '" stroke-width="' +
        (isMonthBoundary ? 1.5 : 0.8) +
        '"/>';
    }

    html += '<line x1="' + LEFT_W + '" y1="0" x2="' + LEFT_W + '" y2="' + CHART_H + '" stroke="#E5E7EB" stroke-width="1"/>';
    html += '<rect x="0" y="0" width="' + CHART_W + '" height="' + HEADER_H + '" fill="#F8F9FA"/>';
    html += '<line x1="0" y1="' + HEADER_H + '" x2="' + CHART_W + '" y2="' + HEADER_H + '" stroke="#E5E7EB" stroke-width="1"/>';

    Object.entries(monthGroups).forEach(function renderMonthGroup(entry) {
      var label = entry[0];
      var group = entry[1];
      var x1 = xForWeek(group.start);
      var x2 = xForWeek(group.start + group.count);
      var centerX = (x1 + x2) / 2;
      html += '<text x="' + centerX + '" y="18" text-anchor="middle" class="gantt-month-label">' + label + "</text>";
      if (group.start > 0) {
        html += '<line x1="' + x1 + '" y1="0" x2="' + x1 + '" y2="' + HEADER_H + '" stroke="#E5E7EB" stroke-width="1"/>';
      }
    });

    weekDates.forEach(function renderWeekLabel(date, index) {
      var centerX = xForWeek(index) + WEEK_W / 2;
      var dayLabel = date.getDate() + " " + date.toLocaleDateString("en-GB", { month: "short" });
      html += '<text x="' + centerX + '" y="' + (HEADER_H - 22) + '" text-anchor="middle" style="font-size:8px;fill:#C0C8D4;font-family:Arial;">W' + (index + 1) + "</text>";
      html += '<text x="' + centerX + '" y="' + (HEADER_H - 10) + '" text-anchor="middle" class="gantt-week-label">' + dayLabel + "</text>";
    });

    if (cutOffOffsetW !== null && cutOffOffsetW >= 0 && cutOffOffsetW <= WEEKS) {
      var cutOffX = xForWeek(cutOffOffsetW);
      html +=
        '<g clip-path="url(#gantt-clip)"><line x1="' +
        cutOffX +
        '" y1="' +
        HEADER_H +
        '" x2="' +
        cutOffX +
        '" y2="' +
        (CHART_H - FOOTER_H) +
        '" stroke="#F57D00" stroke-width="2" opacity="0.9"/><polygon points="' +
        (cutOffX - 6) +
        "," +
        HEADER_H +
        " " +
        (cutOffX + 6) +
        "," +
        HEADER_H +
        " " +
        cutOffX +
        "," +
        (HEADER_H + 8) +
        '" fill="#F57D00"/><text x="' +
        cutOffX +
        '" y="' +
        (CHART_H - 2) +
        '" text-anchor="middle" style="font-size:8px;font-weight:700;fill:#F57D00;font-family:Arial;">CUT-OFF</text></g>';
    }

    workstreams.forEach(function renderWorkstream(item, index) {
      var y = HEADER_H + index * ROW_H;
      var centerY = y + ROW_H / 2;
      var barY = y + ROW_PAD;
      var domainColor = GANTT_DOMAIN_COLOURS[item.Domain] || COLORS.grey;
      var ragColorValue = item.StatusRAG === "Green" ? COLORS.teal : item.StatusRAG === "Amber" ? COLORS.orange : COLORS.alert;
      var startDateInclusive = parseIsoDate(item.StartDate);
      var endDateInclusive = parseIsoDate(item.EndDate);
      var progressDateInclusive = item.ProgressDate ? parseIsoDate(item.ProgressDate) : null;
      var startDate = startDateInclusive;
      var endDate = addDays(endDateInclusive, 1);
      var progressDate = progressDateInclusive ? addDays(progressDateInclusive, 1) : null;
      var startW = dayDiff(baseDate, startDateInclusive) / 7;
      var endW = dayDiff(baseDate, endDate) / 7;
      var durationW = endW - startW;
      var visibleStart = clamp(startW, 0, WEEKS);
      var visibleEnd = clamp(endW, 0, WEEKS);
      var startCapped = startW < 0;
      var endCapped = endW > WEEKS;
      var fullX = xForWeek(visibleStart);
      var fullX2 = xForWeek(visibleEnd);
      var fullW = fullX2 - fullX;
      var nameY = item.SponsorOwner ? centerY - 5 : centerY + 4;
      var rowHoverId = "gantt-workstream-" + index;

      hoverItems[rowHoverId] = {
        accent: domainColor,
        kicker: item.Domain + " · " + statusCopy(item.StatusRAG),
        title: item.WorkstreamName,
        sub: item.SponsorOwner || "",
        rows: [
          { label: "Start", value: formatTooltipDate(startDateInclusive) },
          { label: "End", value: formatTooltipDate(endDateInclusive) },
          { label: "Progress", value: progressDateInclusive ? formatTooltipDate(progressDateInclusive) : "No progress date" },
        ],
        body: item.Detail || "",
      };

      html += '<rect x="0" y="' + y + '" width="4" height="' + ROW_H + '" fill="' + domainColor + '"/>';
      html += '<circle cx="16" cy="' + centerY + '" r="4" fill="' + ragColorValue + '"/>';
      html += '<text x="28" y="' + (nameY + 1) + '" class="gantt-lane-label" style="font-size:10.5px;">' + item.WorkstreamName + "</text>";
      if (item.SponsorOwner) {
        html += '<text x="28" y="' + (centerY + 10) + '" class="gantt-lane-sub">' + item.SponsorOwner + "</text>";
      }

      if (visibleEnd > visibleStart) {
        var leftRadius = startCapped ? 0 : CORNER_R;
        html += '<g clip-path="url(#gantt-clip)">';
        html +=
          '<rect x="' +
          fullX +
          '" y="' +
          (barY + 2) +
          '" width="' +
          fullW +
          '" height="' +
          (BAR_H - 4) +
          '" rx="' +
          leftRadius +
          '" ry="' +
          leftRadius +
          '" fill="' +
          domainColor +
          '" opacity="0.12"/>';

        if (progressDate) {
          var progressEndW = dayDiff(baseDate, progressDate) / 7;
          var completeWidth = clamp(progressEndW - startW, 0, durationW);
          if (completeWidth > 0) {
            var progressEndX = xForWeek(clamp(startW + completeWidth, 0, WEEKS));
            var progressVisibleWidth = progressEndX - fullX;
            if (progressVisibleWidth > 0) {
              html +=
                '<rect x="' +
                fullX +
                '" y="' +
                (barY + 2) +
                '" width="' +
                progressVisibleWidth +
                '" height="' +
                (BAR_H - 4) +
                '" rx="' +
                leftRadius +
                '" ry="' +
                leftRadius +
                '" fill="' +
                domainColor +
                '" opacity="0.85"/>';

              if (completeWidth < durationW) {
                var completionPct = Math.round((completeWidth / durationW) * 100);
                var labelX = xForWeek(clamp(startW + completeWidth / 2, 0, WEEKS));
                if (labelX > LEFT_W + 4 && labelX < xForWeek(WEEKS) - 4) {
                  html +=
                    '<text x="' +
                    labelX +
                    '" y="' +
                    (centerY + 1) +
                    '" text-anchor="middle" dominant-baseline="middle" style="font-size:8.5px;font-weight:700;fill:white;font-family:Arial;pointer-events:none;">' +
                    completionPct +
                    "%</text>";
                }
              }
            }
          }
        }

        if (progressDate && dayDiff(startDate, progressDate) / 7 >= durationW) {
          var completionX = xForWeek(clamp(startW + durationW / 2, 0, WEEKS));
          if (completionX > LEFT_W + 4) {
            html +=
              '<text x="' +
              completionX +
              '" y="' +
              (centerY + 1) +
              '" text-anchor="middle" dominant-baseline="middle" style="font-size:8.5px;font-weight:700;fill:white;font-family:Arial;pointer-events:none;">✓ Complete</text>';
          }
        }

        html +=
          '<rect x="' +
          fullX +
          '" y="' +
          (barY + 2) +
          '" width="' +
          fullW +
          '" height="' +
          (BAR_H - 4) +
          '" rx="' +
          leftRadius +
          '" ry="' +
          leftRadius +
          '" fill="none" stroke="' +
          domainColor +
          '" stroke-width="1.5" opacity="0.5"/>';

        if (startCapped) {
          html +=
            '<polygon points="' +
            LEFT_W +
            "," +
            (barY + 6) +
            " " +
            (LEFT_W + 7) +
            "," +
            centerY +
            " " +
            LEFT_W +
            "," +
            (barY + BAR_H - 4) +
            '" fill="' +
            domainColor +
            '" opacity="0.7"/>';
        }

        if (endCapped) {
          var clipX = xForWeek(WEEKS);
          html +=
            '<polygon points="' +
            clipX +
            "," +
            (barY + 6) +
            " " +
            (clipX - 7) +
            "," +
            centerY +
            " " +
            clipX +
            "," +
            (barY + BAR_H - 4) +
            '" fill="' +
            domainColor +
            '" opacity="0.7"/>';
        }

        html += "</g>";
      }

      html +=
        '<rect class="gantt-hover-target" data-hover-id="' +
        rowHoverId +
        '" x="0" y="' +
        y +
        '" width="' +
        CHART_W +
        '" height="' +
        ROW_H +
        '" fill="#FFFFFF" fill-opacity="0.001"/>';

      (milestonesByWorkstream[item.WorkstreamName] || []).forEach(function renderMilestone(milestone, milestoneIndex) {
        var milestoneDate = parseIsoDate(milestone.MilestoneDate);
        var milestoneW = dayDiff(baseDate, milestoneDate) / 7;
        if (milestoneW < 0 || milestoneW > WEEKS) {
          return;
        }

        var milestoneX = xForWeek(milestoneW);
        var milestoneSize = 7;
        var milestoneHoverId = rowHoverId + "-milestone-" + milestoneIndex;
        var hitWidth = Math.min(180, 26 + milestone.MilestoneLabel.length * 5.4);
        hoverItems[milestoneHoverId] = {
          accent: domainColor,
          kicker: "Milestone · " + item.Domain,
          title: milestone.MilestoneLabel,
          sub: item.WorkstreamName + (item.SponsorOwner ? " · " + item.SponsorOwner : ""),
          rows: [
            { label: "Due", value: formatTooltipDate(milestoneDate) },
            { label: "Status", value: statusCopy(item.StatusRAG) },
            { label: "Window", value: formatTooltipDate(startDateInclusive) + " – " + formatTooltipDate(endDateInclusive) },
          ],
          body: item.Detail || "",
        };
        html +=
          '<g clip-path="url(#gantt-clip)"><rect x="' +
          (milestoneX - milestoneSize) +
          '" y="' +
          (centerY - milestoneSize) +
          '" width="' +
          milestoneSize * 2 +
          '" height="' +
          milestoneSize * 2 +
          '" transform="rotate(45 ' +
          milestoneX +
          " " +
          centerY +
          ')" fill="white" stroke="' +
          domainColor +
          '" stroke-width="2"/><line x1="' +
          milestoneX +
          '" y1="' +
          (centerY + milestoneSize + 3) +
          '" x2="' +
          milestoneX +
          '" y2="' +
          (y + ROW_H) +
          '" stroke="' +
          domainColor +
          '" stroke-width="1" stroke-dasharray="2,2" opacity="0.4"/><text x="' +
          (milestoneX + 10) +
          '" y="' +
          (centerY - 4) +
          '" style="font-size:8.5px;fill:#4B5563;font-family:Arial;font-weight:600;">' +
          milestone.MilestoneLabel +
          '</text><rect class="gantt-hover-target" data-hover-id="' +
          milestoneHoverId +
          '" x="' +
          (milestoneX - 14) +
          '" y="' +
          (centerY - 14) +
          '" width="' +
          hitWidth +
          '" height="28" fill="#FFFFFF" fill-opacity="0.001"/></g>';
      });

      html += '<line x1="0" y1="' + (y + ROW_H) + '" x2="' + CHART_W + '" y2="' + (y + ROW_H) + '" stroke="#F0F2F5" stroke-width="0.8"/>';
    });

    html += '<line x1="' + LEFT_W + '" y1="' + (HEADER_H - 1) + '" x2="' + CHART_W + '" y2="' + (HEADER_H - 1) + '" stroke="#E5E7EB" stroke-width="1"/>';
    svg.innerHTML = html;
    svg.onmouseleave = hideGanttTooltip;
    Array.from(svg.querySelectorAll(".gantt-hover-target")).forEach(function bindHover(node) {
      var hoverId = node.getAttribute("data-hover-id");
      var meta = hoverItems[hoverId];
      if (!meta) {
        return;
      }

      node.addEventListener("mouseenter", function onMouseEnter(event) {
        placeGanttTooltip(event, meta);
      });
      node.addEventListener("mousemove", function onMouseMove(event) {
        placeGanttTooltip(event, meta);
      });
      node.addEventListener("mouseleave", hideGanttTooltip);
    });

    summary.innerHTML =
      renderKpiCard("gantt-kpi-active-workstreams", "blue", "Active Workstreams", visibleWorkstreams.length, "in this 12-week window") +
      renderKpiCard("gantt-kpi-on-track", "teal", "On Track", onTrack, "Green RAG", "up") +
      renderKpiCard("gantt-kpi-at-risk", "orange", "At Risk", atRisk, "Amber RAG · monitoring") +
      renderKpiCard("gantt-kpi-milestones-due", "blue", "Milestones Due", visibleMilestones.length, "within this window");

    subtitle.textContent =
      visibleWorkstreams.length +
      " active workstreams · " +
      visibleMilestones.length +
      " milestones · " +
      completing +
      " completing this window";
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
  buildSummaryPage();
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
  buildGanttPage();
  buildBudgetPage();
  buildRisksPage();

  window.__REPORT_READY = true;

  return {
    showPage: function showPageController(id, tabId) {
      showPage(id, null, { silent: true, tabId: tabId });
    },
    destroy: function destroy() {
      Object.keys(CHARTS).forEach(function destroyChart(id) {
        CHARTS[id].destroy();
        delete CHARTS[id];
      });
      if (options.attachGlobals !== false) {
        delete window.showPage;
        delete window.showPageTab;
        delete window.showMapTip;
        delete window.hideMapTip;
      }
    },
  };
}
