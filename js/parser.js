/**
 * Client-side Excel (.xlsx) & CSV Parser + Exporter for Naue Opportunity Tracker.
 * Allows reviewers to drop an updated Excel spreadsheet directly in the browser
 * (works both statically and with server.py) and export enriched review reports.
 */
(function () {
  function detectRelevance(relText) {
    const t = (relText || "").trim();
    const u = t.toUpperCase();
    if (u.startsWith("VERY HIGH")) return { tier: "Very High", rank: 6 };
    if (u.startsWith("HIGH TO MODERATE") || u.startsWith("MEDIUM TO HIGH") || u.startsWith("MEDIUM-TO-HIGH")) {
      return { tier: "Medium-to-High", rank: 4 };
    }
    if (u.startsWith("HIGH")) return { tier: "High", rank: 5 };
    if (u.startsWith("MODERATE") || u.startsWith("MEDIUM")) return { tier: "Moderate", rank: 3 };
    if (u.startsWith("LOW TO MEDIUM") || u.startsWith("LOW-TO-MEDIUM")) return { tier: "Low-to-Medium", rank: 2 };
    if (u.startsWith("LOW")) return { tier: "Low", rank: 1 };
    if (u.startsWith("NONE")) return { tier: "None", rank: 0 };
    return { tier: "Moderate", rank: 3 };
  }

  function detectProducts(scopeText, relText) {
    const combined = `${scopeText || ""}\n${relText || ""}`;
    const list = [];
    const explicit = {};
    if (/Secutex[®\s]*Soft\s*Rock/i.test(combined)) {
      list.push("Secutex Soft Rock");
      explicit["Secutex Soft Rock"] = true;
    }
    const defs = [
      ["Carbofol", /Carbofol/i],
      ["Bentofix", /Bentofix/i],
      ["Secutex", /Secutex(?!\s*Soft\s*Rock)/i],
      ["Secugrid", /Secugrid/i],
      ["Combigrid", /Combigrid/i],
      ["Secudrain", /Secudrain/i],
      ["Secumat", /Secumat/i]
    ];
    defs.forEach(([id, rx]) => {
      if (rx.test(combined)) {
        list.push(id);
        explicit[id] = true;
      }
    });
    // Fallback inference from engineering keywords if brand names omitted
    if (list.length === 0) {
      if (/sanitary landfill|geomembrane|leachate|basal lin/i.test(combined)) {
        ["Carbofol", "Bentofix", "Secudrain", "Secutex"].forEach(p => { list.push(p); explicit[p] = false; });
      } else if (/geogrid|road.*rehab|trackbed/i.test(combined)) {
        ["Secugrid", "Secutex"].forEach(p => { list.push(p); explicit[p] = false; });
      }
    }
    return { products: list, explicitProductMap: explicit };
  }

  function parseCoordinatesFromText(coordText) {
    if (!coordText || /not specified|not explicitly stated/i.test(coordText) && !/\d+\.\d+/.test(coordText)) {
      return [];
    }
    const pins = [];
    // Decimal degrees e.g. 13.083992°N, 44.887034°E
    const decRx = /([A-Za-z0-9\-\.\s'()&]+?)?:\s*~?(\d{1,2}\.\d+)\s*°?\s*([NS])[,;\s]+~?(\d{1,3}\.\d+)\s*°?\s*([EW])/g;
    let m;
    while ((m = decRx.exec(coordText)) !== null) {
      let lat = parseFloat(m[2]) * (m[3] === "S" ? -1 : 1);
      let lng = parseFloat(m[4]) * (m[5] === "W" ? -1 : 1);
      pins.push({
        name: (m[1] || "Project Site").replace(/^[•\-\s]+/, "").trim() || "Project Site",
        lat,
        lng,
        precision: coordText.toLowerCase().includes("approx") ? "approximate" : "exact",
        note: "Extracted from document coordinates"
      });
    }
    return pins;
  }

  function parseScopeClient(scopeText) {
    if (!scopeText) return { lead: "", items: [], callout: "" };
    const lines = scopeText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    let lead = "";
    let callout = "";
    const items = [];

    lines.forEach(ln => {
      if (/^OPPORTUNITIES FOR NAUE/i.test(ln)) {
        callout = ln.replace(/^OPPORTUNITIES FOR NAUE\s*:?\s*/i, "");
        return;
      }
      if (/^WHAT IS HAPPENING:/i.test(ln)) {
        const cleaned = ln.replace(/^WHAT IS HAPPENING:\s*/i, "");
        const parts = cleaned.split(/(?=\(\d+\)\s+)/);
        if (parts.length > 1) {
          lead = parts[0].replace(/\s*(Activities|Projects)\s+happening\s+include:?\s*$/i, "").trim();
          parts.slice(1).forEach(p => {
            const m = p.match(/^\((\d+)\)\s*(.*)/);
            if (m) items.push({ num: m[1], title: "", body: m[2].replace(/;\s*$/, "").trim() });
          });
        } else {
          lead = cleaned;
        }
        return;
      }
      const mNum = ln.match(/^(\d+)\.\s*([^:]+):\s*(.+)$/);
      if (mNum) {
        items.push({ num: mNum[1], title: mNum[2].trim(), body: mNum[3].trim() });
        return;
      }
      const mBul = ln.match(/^[•\-]\s*([^:]+):\s*(.+)$/);
      if (mBul) {
        items.push({ num: "•", title: mBul[1].trim(), body: mBul[2].trim() });
        return;
      }
      if (!lead) {
        lead = ln.replace(/:$/, "");
      } else {
        items.push({ num: "•", title: "", body: ln.replace(/^[•\-]\s*/, "") });
      }
    });
    return { lead, items, callout };
  }

  function parseWorkbookBuffer(arrayBuffer, existingCatalog) {
    if (typeof XLSX === "undefined") {
      throw new Error("SheetJS library is not loaded; run via start_website.bat for Python openpyxl parsing.");
    }
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = wb.SheetNames.find(n => /tracker/i.test(n)) || wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    // Locate header row containing "Folder Name" or "Tender Title"
    let headerIdx = rows.findIndex(r => r.some(c => /tender title|project country/i.test(String(c))));
    if (headerIdx === -1) headerIdx = 3;

    // Match existing curated projects by title if available so we keep rich map pins
    const existingByTitle = {};
    if (window.NAUE_INITIAL_DATA && window.NAUE_INITIAL_DATA.projects) {
      window.NAUE_INITIAL_DATA.projects.forEach(p => {
        existingByTitle[p.title.trim().toLowerCase()] = p;
      });
    }

    const parsedProjects = [];
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || !r.some(cell => String(cell).trim() !== "")) continue;
      const folder = String(r[0] || "").trim();
      const countryRaw = String(r[1] || "").trim();
      const locationRaw = String(r[2] || "").trim();
      const coordsRaw = String(r[3] || "").trim();
      const refRaw = String(r[4] || "").trim();
      const title = String(r[5] || "").trim();
      const orgRaw = String(r[6] || "").trim();
      const scopeRaw = String(r[7] || "").trim();
      const timelineRaw = String(r[8] || "").trim();
      const relRaw = String(r[9] || "").trim();

      if (!title && !scopeRaw) continue;

      const matched = existingByTitle[title.toLowerCase()];
      if (matched) {
        // Keep curated enrichments while updating any edited raw text fields!
        const rel = detectRelevance(relRaw);
        const prod = detectProducts(scopeRaw, relRaw);
        const sc = parseScopeClient(scopeRaw);
        parsedProjects.push({
          ...matched,
          seq: parsedProjects.length + 1,
          id: `NAUE-${String(parsedProjects.length + 1).padStart(2, "0")}`,
          folder: folder || matched.folder,
          countryRaw: countryRaw || matched.countryRaw,
          locationRaw: locationRaw || matched.locationRaw,
          coordinatesRaw: coordsRaw || matched.coordinatesRaw,
          tenderRefRaw: refRaw || matched.tenderRefRaw,
          title: title || matched.title,
          procuringOrgRaw: orgRaw || matched.procuringOrgRaw,
          technicalScopeRaw: scopeRaw || matched.technicalScopeRaw,
          scopeLead: sc.lead || matched.scopeLead,
          scopeItems: sc.items.length ? sc.items : matched.scopeItems,
          timelineRaw: timelineRaw || matched.timelineRaw,
          relevanceRaw: relRaw || matched.relevanceRaw,
          relevanceTier: rel.tier,
          relevanceRank: rel.rank,
          products: prod.products.length ? prod.products : matched.products
        });
        continue;
      }

      const rel = detectRelevance(relRaw);
      const prod = detectProducts(scopeRaw, relRaw);
      const sc = parseScopeClient(scopeRaw);
      const tlLines = timelineRaw.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      const relLines = relRaw.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      const oppMatch = relRaw.match(/(\d+)\s*\+?\s*(?:Direct|High|Relevant|Core|Potential|Minor|Indirect|Major)?\s*Opportunities/i);
      const actionableCount = oppMatch ? parseInt(oppMatch[1], 10) : (rel.rank > 0 ? 1 : 0);

      const seq = parsedProjects.length + 1;
      parsedProjects.push({
        id: `NAUE-${String(seq).padStart(2, "0")}`,
        seq,
        excelRow: i + 1,
        folder: folder || "Imported",
        folderLabel: folder || "Imported Batch",
        shortName: title.length > 68 ? title.slice(0, 65) + "…" : title,
        title,
        countryRaw,
        countries: [countryRaw || "International"],
        locationRaw,
        coordinatesRaw: coordsRaw,
        mapLocations: parseCoordinatesFromText(coordsRaw),
        tenderRefRaw: refRaw,
        tenderCodes: refRaw ? [refRaw.split(/\r?\n/)[0].slice(0, 42)] : [],
        procuringOrgRaw: orgRaw,
        agencyShort: orgRaw.split(",")[0].slice(0, 45),
        donors: /World Bank/i.test(orgRaw) ? ["World Bank"] : /USAID/i.test(orgRaw) ? ["USAID"] : /KfW/i.test(orgRaw) ? ["KfW"] : ["International Donor"],
        sectors: /landfill|waste/i.test(title + scopeRaw)
          ? ["Sanitary Landfills & Waste"]
          : /road|bridge|highway/i.test(title + scopeRaw)
          ? ["Roads, Highways & Bridges"]
          : /canal|irrigat|dam/i.test(title + scopeRaw)
          ? ["Irrigation, Canals & Dams"]
          : ["Water, WWTP & Stormwater"],
        salesChannel: /supervision|design|feasibility|study/i.test(title + scopeRaw)
          ? "Consultant / BoQ Specification"
          : rel.rank === 0
          ? "Non-Civil (No Action)"
          : "Direct + Specification",
        technicalScopeRaw: scopeRaw,
        scopeLead: sc.lead,
        scopeItems: sc.items,
        scopeOppCallout: sc.callout,
        timelineRaw,
        timelineTag: (tlLines[0] || "Pending Execution").slice(0, 46),
        timelineHeadline: tlLines[0] || "",
        timelineMilestones: tlLines.slice(1).map(s => s.replace(/^[•\-]\s*/, "")),
        relevanceRaw: relRaw,
        relevanceTier: rel.tier,
        relevanceRank: rel.rank,
        relevanceHeadline: relLines[0] || rel.tier,
        whyItMatters: relLines.find(l => /why it matters/i.test(l))?.replace(/^.*why it matters:\s*/i, "") || "",
        relevanceBreakdown: relLines.slice(1).map(s => ({ tier: "info", text: s.replace(/^[•\-]\s*/, "") })),
        products: prod.products,
        explicitProductMap: prod.explicitProductMap,
        primaryProductSummary: "",
        budget: "See Scope",
        stats: {
          total: Math.max(actionableCount, 1),
          high: rel.rank >= 5 ? actionableCount : 0,
          medium: rel.rank >= 3 && rel.rank < 5 ? actionableCount : 0,
          low: rel.rank < 3 ? 1 : 0,
          actionable: actionableCount
        }
      });
    }

    return {
      documentTitle: "NAUE GEOSYNTHETICS — INTERNATIONAL DEVELOPMENT AID PROJECT OPPORTUNITY TRACKER",
      documentSubtitle: `Imported Workbook (${parsedProjects.length} Programs Parsed)`,
      sourceFile: "Imported Excel Workbook",
      generatedAt: new Date().toISOString().slice(0, 10),
      executionContextNote: "All documented programs are active in this tracker for review, consultant specification, and contractor pursuit.",
      productsCatalog: existingCatalog || [],
      projects: parsedProjects
    };
  }

  function exportCSV(projects, reviews) {
    const statusMap = {
      unreviewed: "Unreviewed",
      pursuing: "Pursuing (Priority)",
      consultant_spec: "Consultant Specification",
      contractor_bid: "Target Civil Contractors",
      monitoring: "Monitoring / Hold",
      not_applicable: "Not Applicable"
    };
    const headers = [
      "Project ID",
      "Review Status",
      "Starred",
      "Assigned Lead",
      "Relevance Tier",
      "Actionable Opps",
      "Total Sub-Tenders",
      "Sales Channel",
      "Applicable Naue Products",
      "Country",
      "Program Short Name",
      "Tender Title",
      "Tender Reference",
      "Procuring Agency",
      "Budget",
      "Timeline",
      "Reviewer Notes",
      "Source Batch"
    ];
    const esc = val => `"${String(val ?? "").replace(/"/g, '""')}"`;
    const lines = [headers.map(esc).join(",")];

    projects.forEach(p => {
      const r = reviews[p.id] || {};
      const entries = r.entries || r.comments || [];
      const notes = [
        r.quickNote || "",
        ...entries.map(
          c => `[${c.displayTime || c.date || ""}] ${c.author || "Reviewer"} (${c.verdictLabel || statusMap[c.verdict || r.status || "unreviewed"] || ""}): ${c.text || ""}`
        )
      ].filter(Boolean).join(" || ");

      lines.push([
        p.id,
        statusMap[r.status || "unreviewed"] || r.status,
        r.starred ? "Yes" : "No",
        r.assignee || "",
        p.relevanceTier,
        p.stats.actionable,
        p.stats.total,
        p.salesChannel,
        p.products.map(x => x + "®").join("; "),
        p.countryRaw,
        p.shortName,
        p.title,
        p.tenderRefRaw,
        p.procuringOrgRaw,
        p.budget,
        p.timelineTag,
        notes,
        p.folder
      ].map(esc).join(","));
    });

    const blob = new Blob(["\uFEFF" + lines.join("\r\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Naue_Opportunity_Tracker_Reviewed_${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  window.NaueParser = {
    parseWorkbookBuffer,
    exportCSV
  };
})();

