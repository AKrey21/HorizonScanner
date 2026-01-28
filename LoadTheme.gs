function loadThemeRules() {
  if (typeof repo_getActiveThemeRules_ === "function") {
    const rules = repo_getActiveThemeRules_() || [];
    return rules.map(rule => ({
      theme: rule.theme,
      poi: rule.poi,
      keywords: rule.keywords || []
    }));
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ThemeRules");
  if (!sheet) return [];

  const startRow =
    (typeof THEME_RULES_START_ROW !== "undefined" && Number(THEME_RULES_START_ROW) >= 1)
      ? Number(THEME_RULES_START_ROW)
      : 5;

  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];

  // Cols A:D => Theme, POI, Keywords, Active?
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
  const rules = [];

  values.forEach(row => {
    const theme = String(row[0] || "").trim();
    const poi = String(row[1] || "").trim();
    const keywordString = String(row[2] || "").trim();
    const isActive = row[3] === true;

    // Skip blank / inactive rows
    if (!isActive) return;
    if (!theme || !keywordString) return;

    const keywordList = keywordString
      .split(",")
      .map(k => String(k || "").trim().toLowerCase())
      .filter(Boolean);

    if (!keywordList.length) return;

    rules.push({
      theme,
      poi,
      keywords: keywordList
    });
  });

  return rules;
}
