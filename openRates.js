function buildCourseDictionary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Offered Appointments");
  const courseData = sheet.getRange("F2:F" + sheet.getLastRow()).getValues().flat();

  const results = courseData.map(cell => {
    if (!cell || typeof cell !== "string") return "";

    // Match valid course codes (2+ letters, 1+ digits, optional letter)
    let matches = cell.match(/\b[A-Z]{2,}\d{1,3}[A-Z]?\b/g);
    if (!matches) return "";

    // Hardcode exclusion of FA25
    matches = matches.filter(c => c !== "WI26");

    if (matches.length === 0) return "";

    // Count occurrences
    const counts = {};
    matches.forEach(c => counts[c] = (counts[c] || 0) + 1);

    // Convert to COURSE:COUNT string
    return Object.entries(counts)
      .map(([key, val]) => `${key}:${val}`)
      .join(",");
  });

  // Write to column O
  sheet.getRange(2, 15, results.length, 1).setValues(results.map(r => [r]));
}


function countAppointments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const countsSheet = ss.getSheetByName("Open % Per Course");
  const offeredSheet = ss.getSheetByName("Offered Appointments");

  const countsData = countsSheet.getDataRange().getValues();
  const offeredData = offeredSheet.getDataRange().getValues();

  const WEEK_COL = 8;     // I
  const MODE_COL = 9;     // J
  const COURSES_COL = 12; // M  (eligible courses)
  const DICT_COL = 14;    // O  (keys:values)

  for (let i = 2; i < countsData.length; i++) {
    const week = countsData[i][0];
    const course = countsData[i][1];
    if (!course) continue;

    // --- Determine token ---
    let token = course.replace("*", "").trim();
    let starMode = course.startsWith("*");

    let takenIP = 0;
    let takenOnline = 0;

    for (let j = 1; j < offeredData.length; j++) {
      const row = offeredData[j];
      const rowWeek = row[WEEK_COL];
      const mode = row[MODE_COL] ? row[MODE_COL].toString() : "";
      const eligibleString = row[COURSES_COL] ? row[COURSES_COL].toString() : "";
      const dictString = row[DICT_COL] ? row[DICT_COL].toString() : "";

      if (rowWeek !== week) continue;
      if (!eligibleString || !dictString) continue;

      // eligibility: column M contains token
      const eligible = eligibleString
        .split(",")
        .some(c => c.trim().includes(token));

      if (!eligible) continue;

      // parse keys:values
      const pairs = dictString.split(",").map(x => x.trim());
      for (const p of pairs) {
        const [key, val] = p.split(":").map(x => x && x.trim());
        if (!key) continue;
        const n = Number(val) || 0;

        // star rows → count only keys containing substring token
        if (starMode && !key.includes(token)) continue;

        if (mode === "In-Person") takenIP += n;
        else if (mode === "Online") takenOnline += n;
      }
    }

    // --- Only write to D and F ---
    countsSheet.getRange(i + 1, 4).setValue(takenIP);
    countsSheet.getRange(i + 1, 6).setValue(takenOnline);
  }
}








function calculateOpenRateSimplifiedLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("Open % Per Course");
  const offeredSheet = ss.getSheetByName("Offered Appointments");

  const targetData = targetSheet.getDataRange().getValues();
  const offeredData = offeredSheet.getDataRange().getValues();

  // 0-based indices in Offered Appointments
  const WEEK_COL = 8;       // I
  const MODE_COL = 9;       // J
  const COURSES_COL = 12;   // M
  const SLOTS_COL = 13;     // N  <-- number of slots for that offered appointment
  const DICT_COL = 14;      // O

  for (let i = 2; i < targetData.length; i++) {  // start at row 3
    const week = targetData[i][0];    // Column A
    const rawCourse = targetData[i][1];  // Column B
    if (!rawCourse) continue;

    // Detect star rows and normalize course name
    const isStar = rawCourse.toString().includes("*");
    const normalizedCourse = rawCourse.toString().replace(/^\*+/, "").trim();

    // By default take totals from target sheet (unchanged behavior)
    let totalOfferedInPerson = Number(targetData[i][2]) || 0; // Column C
    let totalOfferedOnline = Number(targetData[i][4]) || 0;   // Column E

    // For star rows, we'll recompute totalOffered from offeredData
    if (isStar) {
      totalOfferedInPerson = 0;
      totalOfferedOnline = 0;
    }

    let totalTakenInPerson = 0;
    let totalTakenOnline = 0;

    Logger.log(`\nRow ${i+1}: rawCourse="${rawCourse}", normalized="${normalizedCourse}", week=${week}, isStar=${isStar}`);
    Logger.log(`Starting totals (from target sheet unless star): IP=${totalOfferedInPerson}, Online=${totalOfferedOnline}`);

    for (let j = 1; j < offeredData.length; j++) {
      const row = offeredData[j];
      const offeredWeek = row[WEEK_COL];
      const offeredModeRaw = row[MODE_COL] ? row[MODE_COL].toString().trim() : "";
      const offeredMode = offeredModeRaw.toLowerCase();
      const offeredCoursesRaw = row[COURSES_COL] ? row[COURSES_COL].toString() : "";
      const dictString = row[DICT_COL] ? row[DICT_COL].toString() : "";
      const offeredSlots = Number(row[SLOTS_COL]) || 0;

      if (offeredWeek !== week) continue; // week must match
      if (!dictString && !isStar) continue; // if not a star row and no dict, skip (no taken info)
      if (!offeredCoursesRaw) continue;

      // Build array of course names in this offered row (strip any :count suffix)
      const offeredCourseNames = offeredCoursesRaw.split(",").map(s => {
        return s.split(":")[0].trim();
      }).filter(Boolean);

      // Decide whether this offered row is relevant:
      // - For non-star rows: require exact includes of the course (one of the offeredCourseNames equals the course)
      // - For star rows: require the offered courses list to include the normalizedCourse (exact or partial match)
      let rowMatches = false;
      if (isStar) {
        // match when any offered course name contains normalizedCourse OR equals normalizedCourse
        rowMatches = offeredCourseNames.some(cn => cn && cn.includes(normalizedCourse));
      } else {
        rowMatches = offeredCourseNames.some(cn => cn === rawCourse.toString());
      }
      if (!rowMatches) continue;

      // If this is a star row we want to include the offeredSlots into the totals per mode
      if (isStar) {
        if (offeredMode === "in-person" || offeredMode === "in-person".toLowerCase()) {
          totalOfferedInPerson += offeredSlots;
        } else if (offeredMode === "online") {
          totalOfferedOnline += offeredSlots;
        } else {
          // if mode label is slightly different, try simple checks:
          if (row[MODE_COL] === "In-Person" || /in-?person/i.test(row[MODE_COL])) totalOfferedInPerson += offeredSlots;
          else totalOfferedOnline += offeredSlots;
        }
      }

      // Parse dictString pairs and add taken counts for keys that match/contain the normalized course
      if (dictString) {
        const pairs = dictString.split(",").map(p => p.trim()).filter(Boolean);
        for (const pair of pairs) {
          const [keyRaw, valRaw] = pair.split(":").map(x => x && x.trim());
          if (!keyRaw) continue;
          const key = keyRaw;
          const val = Number(valRaw) || 0;

          // Matching logic for taken counts:
          // - For star rows: match if key contains normalizedCourse
          // - For non-star rows: match if key === course exactly
          let keyMatches = false;
          if (isStar) {
            keyMatches = key.includes(normalizedCourse);
          } else {
            keyMatches = (key === rawCourse.toString());
          }

          if (!keyMatches) continue;

          // Add to taken totals by mode
          if (offeredMode === "in-person" || offeredMode === "in-person".toLowerCase()) {
            totalTakenInPerson += val;
          } else if (offeredMode === "online") {
            totalTakenOnline += val;
          } else {
            // fallback checks
            if (row[MODE_COL] === "In-Person" || /in-?person/i.test(row[MODE_COL])) totalTakenInPerson += val;
            else totalTakenOnline += val;
          }

          Logger.log(`  Matched dict pair "${pair}" in offered row ${j+1} (mode=${offeredModeRaw}) → +${val}`);
        }
      }
    } // end offeredData loop

    // If totals were not recomputed for non-star rows, they remain as read from target sheet.
    // Compute open rates (guard against division by zero).
    const openRateInPerson = totalOfferedInPerson > 0 ? 1 - totalTakenInPerson / totalOfferedInPerson : 0;
    const openRateOnline = totalOfferedOnline > 0 ? 1 - totalTakenOnline / totalOfferedOnline : 0;

    // Write results to columns G (7) and H (8)
    targetSheet.getRange(i + 1, 7).setValue(openRateInPerson); // G
    targetSheet.getRange(i + 1, 8).setValue(openRateOnline);   // H

    Logger.log(`Row ${i+1} result: Offered IP=${totalOfferedInPerson}, Taken IP=${totalTakenInPerson}, OpenRate IP=${openRateInPerson}`);
    Logger.log(`Row ${i+1} result: Offered Online=${totalOfferedOnline}, Taken Online=${totalTakenOnline}, OpenRate Online=${openRateOnline}`);
  }
}







