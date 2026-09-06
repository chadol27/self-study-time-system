function teacherExtraDirectory(token) {
  return publicCall_(function () {
    requireConfig_();
    requireTeacher_(token);
    return withWriteLock_(function () {
      requireTeacher_(token);
      recoverExtraPending_();
      const directory = readStudentDirectorySheet_();
      return Object.keys(directory.students)
        .sort()
        .map(function (studentId) {
          return {
            studentId: studentId,
            name: directory.students[studentId].name,
          };
        });
    });
  });
}

function readExtraSheet_() {
  const sheet = spreadsheet_().getSheetByName(APP.SHEETS.EXTRAS);
  const blocks = [];
  const errors = [];
  if (!sheet || !sheet.getLastColumn())
    return { sheet: sheet, blocks: blocks, errors: errors };
  const width = sheet.getLastColumn();
  if (sheet.getMaxRows() < 2) {
    errors.push("미신청자 시트 헤더 2행이 없습니다.");
    return { sheet: sheet, blocks: blocks, errors: errors };
  }
  const range = sheet.getRange(1, 1, 2, width);
  const values = range.getValues();
  const formulas = range.getFormulas();
  let previous = "";
  for (let i = 0; i < width; i += 4) {
    const dates = values[0].slice(i, i + 4);
    const key =
      dates[0] instanceof Date && !isNaN(dates[0].getTime())
        ? dateKey_(dates[0])
        : "";
    if (
      !key ||
      dates.length !== 4 ||
      dates.some(function (date) {
        return (
          !(date instanceof Date) ||
          isNaN(date.getTime()) ||
          dateKey_(date) !== key
        );
      }) ||
      values[1].slice(i, i + 4).join("|") !== "학번|1교시|2교시|3교시" ||
      formulas.some(function (row) {
        return row.slice(i, i + 4).some(Boolean);
      })
    ) {
      errors.push("미신청자 날짜/교시 헤더 오류: " + (i + 1) + "열");
      continue;
    }
    if (previous && previous >= key)
      errors.push(
        "미신청자 날짜 중복 또는 정렬 오류: " + key + " (" + (i + 1) + "열)",
      );
    previous = key;
    blocks.push({ key: key, col: i + 1 });
  }
  return { sheet: sheet, blocks: blocks, errors: errors };
}

function readExtraRows_(sheet, block) {
  const rows = [];
  const errors = [];
  const invalidIds = new Set();
  let nextRow = 3;
  if (!sheet || !block || sheet.getLastRow() < 3)
    return {
      rows: rows,
      errors: errors,
      invalidIds: invalidIds,
      nextRow: nextRow,
    };
  const range = sheet.getRange(3, block.col, sheet.getLastRow() - 2, 4);
  const raw = range.getValues();
  const display = range.getDisplayValues();
  const formulas = range.getFormulas();
  const seen = new Set();
  raw.forEach(function (cells, index) {
    if (
      cells.every(function (v) {
        return v === "";
      }) &&
      !formulas[index].some(Boolean)
    )
      return;
    nextRow = index + 4;
    const id = display[index][0].trim();
    const label =
      "미신청자 " +
      block.key +
      " " +
      (index + 3) +
      "행 (" +
      (id || "학번 없음") +
      ")";
    let invalid = !id || formulas[index].some(Boolean);
    if (invalid) errors.push(label + ": 학번 누락 또는 수식 사용");
    cells.slice(1).forEach(function (value, p) {
      if (!["", "2", "3", 2, 3].includes(value)) {
        errors.push(label + ": " + (p + 1) + "교시 출결값 오류");
        invalid = true;
      }
    });
    if (seen.has(id)) {
      errors.push(label + ": 날짜 내 학번 중복");
      invalid = true;
    }
    seen.add(id);
    if (invalid) invalidIds.add(id);
    rows.push({
      studentId: id,
      row: index + 3,
      values: cells.slice(1).map(String),
    });
  });
  return {
    rows: rows.filter(function (row) {
      return !invalidIds.has(row.studentId);
    }),
    errors: errors,
    invalidIds: invalidIds,
    nextRow: nextRow,
  };
}

function extraDatePeriod_(date, period, write) {
  parseDateKey_(date);
  if (!Number.isInteger(period) || period < 1 || period > 3)
    throw userError_("교시가 올바르지 않습니다.", "INVALID_PERIOD");
  if (!isOperatingDate_(date))
    throw userError_(
      "미운영일은 조회하거나 변경할 수 없습니다.",
      "CLOSED_DATE",
    );
  if (date > addDays_(todayKey_(), 30))
    throw userError_("조회 가능한 날짜를 벗어났습니다.", "DATE_OUT_OF_RANGE");
  if (write && date !== todayKey_())
    throw userError_("오늘 날짜만 수정할 수 있습니다.", "DATE_READ_ONLY");
}

function extraMetadata_(date) {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const prefix = "EXTRA_" + date + "_";
  const pendingDates = new Set(
    Object.keys(all)
      .filter(function (key) {
        return (
          /^EXTRA_\d{4}-\d{2}-\d{2}_R_/.test(key) &&
          JSON.parse(all[key]).state === "pending"
        );
      })
      .map(function (key) {
        return key.slice(6, 16);
      }),
  );
  // Keep yesterday as well, so interrupted work can finish across midnight.
  const cutoff = addDays_(todayKey_(), -1);
  Object.keys(all)
    .filter(function (key) {
      return (
        /^EXTRA_\d{4}-\d{2}-\d{2}_/.test(key) &&
        key.slice(6, 16) < cutoff &&
        !pendingDates.has(key.slice(6, 16))
      );
    })
    .slice(0, 100)
    .forEach(function (key) {
      props.deleteProperty(key);
      delete all[key];
    });
  return { props: props, all: all, prefix: prefix };
}

function extraView_(date, period, metadata) {
  const source = readExtraSheet_();
  const block = source.blocks.find(function (item) {
    return item.key === date;
  });
  const data = readExtraRows_(source.sheet, block);
  const directory = readStudentDirectorySheet_();
  const versions = Object.create(null);
  const generations = Object.create(null);
  Object.keys(metadata.all)
    .filter(function (key) {
      return key.indexOf(metadata.prefix + "V_") === 0;
    })
    .forEach(function (key) {
      const record = JSON.parse(metadata.all[key]);
      generations[record.studentId] = record.generation;
      versions[record.studentId] = record.generation + ":missing";
    });
  const students = source.errors.length
    ? []
    : data.rows.map(function (row) {
        const version =
          (generations[row.studentId] || "0") +
          ":" +
          hashToken_(JSON.stringify([row.row, row.values]));
        versions[row.studentId] = version;
        const identity = Object.prototype.hasOwnProperty.call(
          directory.students,
          row.studentId,
        )
          ? directory.students[row.studentId]
          : null;
        return {
          studentId: row.studentId,
          name: identity ? identity.name : "학생 정보 확인 불가",
          values: row.values,
          version: version,
        };
      });
  const errors = source.errors.concat(data.errors);
  return {
    date: date,
    period: period,
    readOnly: date !== todayKey_(),
    students: students,
    versions: versions,
    errors: errors,
  };
}

function getTeacherExtras(token, date, period) {
  return publicCall_(function () {
    requireConfig_();
    requireTeacher_(token);
    return withWriteLock_(function () {
      requireTeacher_(token);
      recoverExtraPending_();
      date = String(date);
      period = Number(period);
      extraDatePeriod_(date, period, false);
      return extraView_(date, period, extraMetadata_(date));
    });
  });
}

function ensureExtraBlock_(source, date) {
  let sheet = source.sheet;
  if (!sheet) sheet = spreadsheet_().insertSheet(APP.SHEETS.EXTRAS);
  const existing = source.blocks.find(function (block) {
    return block.key === date;
  });
  if (existing) {
    sheet.getRange(1, existing.col, 1, 4).setNumberFormat("M/d");
    return { sheet: sheet, col: existing.col };
  }
  const later = source.blocks.find(function (block) {
    return block.key > date;
  });
  const col = later ? later.col : sheet.getLastColumn() + 1;
  if (later) sheet.insertColumnsBefore(col, 4);
  else if (sheet.getMaxColumns() < col + 3)
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      col + 3 - sheet.getMaxColumns(),
    );
  if (sheet.getMaxRows() < 3)
    sheet.insertRowsAfter(sheet.getMaxRows(), 3 - sheet.getMaxRows());
  const value = parseDateKey_(date);
  sheet.getRange(1, col, 2, 4).setValues([
    [value, value, value, value],
    ["학번", "1교시", "2교시", "3교시"],
  ]);
  sheet.getRange(1, col, 1, 4).setNumberFormat("M/d");
  sheet.setFrozenRows(2);
  return { sheet: sheet, col: col };
}

function extraAuditKey_(row) {
  return JSON.stringify(
    row.map(function (value) {
      return value instanceof Date ? value.getTime() : String(value);
    }),
  );
}

function recoverExtraPending_() {
  requireConfig_();
  // Finish already-authorized intents across midnight; new writes still require today.
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all)
    .filter(function (key) {
      return (
        /^EXTRA_\d{4}-\d{2}-\d{2}_R_/.test(key) &&
        JSON.parse(all[key]).state === "pending"
      );
    })
    .sort()
    .forEach(function (receiptKey) {
      const intent = JSON.parse(all[receiptKey]);
      const blocked = function () {
        throw userError_(
          "미신청자 미완료 작업과 시트 또는 기록이 일치하지 않습니다. 관리자 확인이 필요합니다.",
          "EXTRA_RECOVERY_REQUIRED",
        );
      };
      // Old/incomplete intents lack enough information for safe reconciliation.
      if (
        !intent.date ||
        !Array.isArray(intent.plans) ||
        !intent.plans.length ||
        !Number.isInteger(intent.period) ||
        intent.period < 1 ||
        intent.period > 3 ||
        !Number.isInteger(intent.auditStartRow) ||
        intent.auditStartRow < 2 ||
        !intent.startedAt ||
        isNaN(new Date(intent.startedAt).getTime())
      )
        blocked();
      const source = readExtraSheet_();
      const block = source.blocks.find(function (item) {
        return item.key === intent.date;
      });
      const log = spreadsheet_().getSheetByName(APP.SHEETS.LOG);
      if (
        source.errors.length ||
        !block ||
        !simpleHeadersValid_(log, APP.LOG_HEADERS) ||
        log.getLastRow() < intent.auditStartRow - 1
      )
        blocked();
      const data = readExtraRows_(source.sheet, block);
      /** @type {{[key: string]: string}} */
      const revisions = {};
      const audits = [];
      const writes = intent.plans.map(function (plan) {
        if (
          !Number.isInteger(plan.row) ||
          plan.row < 3 ||
          !plan.generation ||
          typeof plan.previousGeneration !== "string" ||
          !Array.isArray(plan.before) ||
          !Array.isArray(plan.after) ||
          plan.before.length !== 3 ||
          plan.after.length !== 3 ||
          typeof plan.beforeId !== "string" ||
          typeof plan.afterId !== "string"
        )
          blocked();
        if (
          data.invalidIds.has(plan.studentId) ||
          data.rows.some(function (row) {
            return row.studentId === plan.studentId && row.row !== plan.row;
          })
        )
          blocked();
        const range =
          plan.row <= source.sheet.getMaxRows()
            ? source.sheet.getRange(plan.row, block.col, 1, 4)
            : null;
        const current = range ? range.getValues()[0] : ["", "", "", ""];
        const id = range ? range.getDisplayValues()[0][0].trim() : "";
        if (
          (range && range.getFormulas()[0].some(Boolean)) ||
          (id !== plan.beforeId && id !== plan.afterId) ||
          current.slice(1).some(function (value, p) {
            return (
              !["", "2", "3", 2, 3].includes(value) ||
              (String(value) !== plan.before[p] &&
                String(value) !== plan.after[p])
            );
          })
        )
          blocked();
        const revisionKey =
          "EXTRA_" + intent.date + "_V_" + hashToken_(plan.studentId);
        const previous = props.getProperty(revisionKey);
        const revision = previous
          ? JSON.parse(previous)
          : { studentId: plan.studentId, generation: "" };
        if (
          revision.studentId !== plan.studentId ||
          (revision.generation !== plan.previousGeneration &&
            revision.generation !== plan.generation)
        )
          blocked();
        revisions[revisionKey] = JSON.stringify({
          studentId: plan.studentId,
          generation: plan.generation,
        });
        audits.push([
          new Date(intent.startedAt),
          "교사",
          "extra:" + plan.studentId,
          plan.studentId,
          parseDateKey_(intent.date),
          intent.period,
          plan.before[intent.period - 1],
          plan.after[intent.period - 1],
        ]);
        return { plan: plan, id: id, values: current.slice(1).map(String) };
      });
      const expected = new Map(
        audits.map(function (row) {
          return [extraAuditKey_(row), row];
        }),
      );
      if (log.getLastRow() >= intent.auditStartRow) {
        const range = log.getRange(
          intent.auditStartRow,
          1,
          log.getLastRow() - intent.auditStartRow + 1,
          8,
        );
        const formulas = range.getFormulas();
        range.getValues().forEach(function (row, index) {
          // Other services may append roster/student audits while an intent waits.
          if (String(row[2]).indexOf("extra:") !== 0) return;
          const key = extraAuditKey_(row);
          if (formulas[index].some(Boolean) || !expected.has(key)) blocked();
          expected.delete(key);
        });
      }
      // Validate the entire batch and its audit suffix before replaying any write.
      props.setProperties(revisions);
      writes.forEach(function (write) {
        const plan = write.plan;
        if (source.sheet.getMaxRows() < plan.row)
          source.sheet.insertRowsAfter(
            source.sheet.getMaxRows(),
            plan.row - source.sheet.getMaxRows(),
          );
        if (!plan.afterId) {
          if (write.id)
            source.sheet.getRange(plan.row, block.col, 1, 4).clearContent();
        } else {
          if (write.id !== plan.afterId)
            source.sheet
              .getRange(plan.row, block.col)
              .setNumberFormat("@")
              .setRichTextValue(
                SpreadsheetApp.newRichTextValue().setText(plan.afterId).build(),
              );
          if (write.values[intent.period - 1] !== plan.after[intent.period - 1])
            source.sheet
              .getRange(plan.row, block.col + intent.period)
              .setValue(
                plan.after[intent.period - 1] === ""
                  ? ""
                  : Number(plan.after[intent.period - 1]),
              );
        }
      });
      SpreadsheetApp.flush();
      const missing = Array.from(expected.values()).map(function (row) {
        const values = row.slice();
        if (String(values[3]).charAt(0) === "=") values[3] = "'" + values[3];
        return values;
      });
      appendAudits_(missing);
      SpreadsheetApp.flush();
      props.setProperty(
        receiptKey,
        JSON.stringify({
          payload: intent.payload,
          state: "done",
          changed: intent.plans.length,
          conflict: false,
          appliedVersions: intent.appliedVersions,
        }),
      );
    });
}

function teacherExtraChange(
  token,
  date,
  period,
  action,
  selections,
  requestId,
) {
  return publicCall_(function () {
    requireConfig_();
    requireTeacher_(token);
    date = String(date);
    period = Number(period);
    if (!["add", "present", "absent", "restore", "remove"].includes(action))
      throw userError_("변경 작업이 올바르지 않습니다.", "INVALID_ACTION");
    if (
      typeof requestId !== "string" ||
      !/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
        requestId,
      )
    )
      throw userError_(
        "요청 식별자가 올바르지 않습니다.",
        "INVALID_REQUEST_ID",
      );
    if (
      !Array.isArray(selections) ||
      !selections.length ||
      selections.length > 100 ||
      (action === "add" && selections.length !== 1) ||
      selections.some(function (item) {
        return (
          !item ||
          typeof item.studentId !== "string" ||
          !item.studentId.trim() ||
          item.studentId !== item.studentId.trim() ||
          item.studentId.length > 100 ||
          typeof item.version !== "string" ||
          item.version.length > 200
        );
      }) ||
      new Set(
        selections.map(function (item) {
          return item.studentId;
        }),
      ).size !== selections.length
    )
      throw userError_("선택한 학생을 확인해 주세요.", "INVALID_SELECTION");
    const selected = selections
      .map(function (item) {
        return { studentId: item.studentId, version: item.version };
      })
      .sort(function (a, b) {
        return a.studentId.localeCompare(b.studentId);
      });
    const payload = hashToken_(
      JSON.stringify([date, period, action, selected]),
    );
    return withWriteLock_(function () {
      requireConfig_();
      requireTeacher_(token);
      recoverExtraPending_();
      extraDatePeriod_(date, period, false);
      const metadata = extraMetadata_(date);
      const receiptKey = metadata.prefix + "R_" + requestId.toLowerCase();
      const previous = metadata.all[receiptKey]
        ? JSON.parse(metadata.all[receiptKey])
        : null;
      const view = extraView_(date, period, metadata);
      if (previous) {
        if (previous.payload !== payload)
          throw userError_(
            "같은 요청 식별자를 다른 작업에 사용할 수 없습니다.",
            "REQUEST_ID_REUSED",
          );
        if (previous.state === "done")
          return {
            changed: previous.changed,
            view: view,
            conflict: previous.conflict,
            appliedVersions: previous.appliedVersions,
          };
      }
      extraDatePeriod_(date, period, true);
      // Reserve capacity before any sheet writes; never discard today's receipts.
      if (
        Object.keys(metadata.all).reduce(function (size, key) {
          return size + key.length + metadata.all[key].length;
        }, 0) > 120000
      )
        throw userError_(
          "변경 기록 저장 공간을 확인해 주세요.",
          "EXTRA_METADATA_FULL",
        );
      const receipt = {
        payload: payload,
        state: "done",
        changed: 0,
        conflict: false,
        appliedVersions: Object.fromEntries(
          selected.map(function (item) {
            return [item.studentId, view.versions[item.studentId] || ""];
          }),
        ),
      };
      if (
        selected.some(function (item) {
          return item.version !== (view.versions[item.studentId] || "");
        })
      ) {
        receipt.conflict = true;
        metadata.props.setProperty(receiptKey, JSON.stringify(receipt));
        return { changed: 0, view: view, conflict: true };
      }
      const source = readExtraSheet_();
      if (source.errors.length)
        throw userError_(
          "미신청자 시트 헤더를 확인해 주세요.",
          "INVALID_HEADERS",
        );
      const block = source.blocks.find(function (item) {
        return item.key === date;
      });
      const data = readExtraRows_(source.sheet, block);
      const directory = readStudentDirectorySheet_();
      const log = spreadsheet_().getSheetByName(APP.SHEETS.LOG);
      if (!simpleHeadersValid_(log, APP.LOG_HEADERS))
        throw userError_("기록 시트 헤더를 확인해 주세요.", "INVALID_HEADERS");
      const plans = selected
        .map(function (item) {
          const row = data.rows.find(function (entry) {
            return entry.studentId === item.studentId;
          });
          if (
            data.invalidIds.has(item.studentId) ||
            (action !== "add" && !row) ||
            (action === "add" &&
              !Object.prototype.hasOwnProperty.call(
                directory.students,
                item.studentId,
              ))
          )
            throw userError_(
              "선택한 학생 정보를 다시 확인해 주세요: " + item.studentId,
              "INVALID_SELECTION",
            );
          if (action === "remove" && row.values.some(Boolean))
            throw userError_(
              "모든 교시를 되돌린 후 삭제해 주세요: " + item.studentId,
              "EXTRA_NOT_EMPTY",
            );
          const before = row ? row.values.slice() : ["", "", ""];
          const after = before.slice();
          after[period - 1] =
            action === "add" || action === "present"
              ? "2"
              : action === "absent"
                ? "3"
                : "";
          return {
            studentId: item.studentId,
            row: row ? row.row : data.nextRow,
            beforeId: row ? item.studentId : "",
            afterId: action === "remove" ? "" : item.studentId,
            before: before,
            after: after,
            generation: Utilities.getUuid(),
            previousGeneration: metadata.all[
              metadata.prefix + "V_" + hashToken_(item.studentId)
            ]
              ? JSON.parse(
                  metadata.all[
                    metadata.prefix + "V_" + hashToken_(item.studentId)
                  ],
                ).generation
              : "",
            changed:
              action === "remove" || before[period - 1] !== after[period - 1],
          };
        })
        .filter(function (plan) {
          return plan.changed;
        });
      receipt.changed = plans.length;
      plans.forEach(function (plan) {
        receipt.appliedVersions[plan.studentId] =
          plan.generation +
          ":" +
          (plan.afterId
            ? hashToken_(JSON.stringify([plan.row, plan.after]))
            : "missing");
      });
      if (!plans.length) {
        metadata.props.setProperty(receiptKey, JSON.stringify(receipt));
        return {
          changed: 0,
          view: view,
          conflict: false,
          appliedVersions: receipt.appliedVersions,
        };
      }
      // Headers are allocated first; no student cells change before the intent.
      const pending = JSON.stringify({
        payload: payload,
        state: "pending",
        date: date,
        action: action,
        period: period,
        startedAt: new Date(
          Math.floor(now_().getTime() / 1000) * 1000,
        ).toISOString(),
        auditStartRow: log.getLastRow() + 1,
        plans: plans,
        appliedVersions: receipt.appliedVersions,
      });
      if (Utilities.newBlob(pending).getBytes().length > 8000)
        throw userError_("학생을 나누어 선택해 주세요.", "INVALID_SELECTION");
      ensureExtraBlock_(source, date);
      SpreadsheetApp.flush();
      metadata.props.setProperty(receiptKey, pending);
      recoverExtraPending_();
      return {
        changed: receipt.changed,
        view: extraView_(date, period, extraMetadata_(date)),
        conflict: false,
        appliedVersions: receipt.appliedVersions,
      };
    });
  });
}
