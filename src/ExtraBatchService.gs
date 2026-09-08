function extraBatchSnapshot_(date) {
  const source = readExtraSheet_();
  const block = source.blocks.find(function (item) {
    return item.key === date;
  });
  return {
    source: source,
    block: block,
    data: readExtraRows_(source.sheet, block),
    directory: readStudentDirectorySheet_(),
  };
}

// All callers hold the script lock. Chunked intents avoid the per-property 9 KB limit.
function recoverExtraBatch_(
  receiptKey,
  intent,
  snapshot,
  metadata,
  preparedPlans,
  preparedLog,
) {
  const blocked = function () {
    throw userError_(
      "미신청자 미완료 작업과 시트 또는 기록이 일치하지 않습니다. 관리자 확인이 필요합니다.",
      "EXTRA_RECOVERY_REQUIRED",
    );
  };
  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(intent.date) ||
    !Number.isInteger(intent.period) ||
    intent.period < 1 ||
    intent.period > 3 ||
    !Number.isInteger(intent.chunks) ||
    intent.chunks < 1 ||
    !Number.isInteger(intent.auditStartRow) ||
    intent.auditStartRow < 2 ||
    !intent.startedAt ||
    isNaN(new Date(intent.startedAt).getTime())
  )
    blocked();
  metadata = metadata || extraMetadata_(intent.date);
  const plans =
    preparedPlans ||
    Array.from({ length: intent.chunks }, function (_, index) {
      const chunk = metadata.all[receiptKey + "_C_" + index];
      if (!chunk) blocked();
      return JSON.parse(chunk);
    }).flat();
  if (!plans.length || plans.length > 1000) blocked();
  snapshot = snapshot || extraBatchSnapshot_(intent.date);
  const source = snapshot.source;
  const data = snapshot.data;
  const block = snapshot.block;
  const log = preparedLog || spreadsheet_().getSheetByName(APP.SHEETS.LOG);
  const logLastRow = log ? log.getLastRow() : 0;
  if (
    source.errors.length ||
    data.errors.length ||
    !block ||
    (!preparedLog && !simpleHeadersValid_(log, APP.LOG_HEADERS)) ||
    logLastRow < intent.auditStartRow - 1
  )
    blocked();
  const raw = data.raw.map(function (row) {
    return row.slice();
  });
  const revisions = {};
  const audits = [];
  const seen = new Set();
  const positions = new Set();
  plans.forEach(function (plan) {
    if (
      !plan ||
      typeof plan.studentId !== "string" ||
      !plan.studentId ||
      seen.has(plan.studentId) ||
      positions.has(plan.row) ||
      !Number.isInteger(plan.row) ||
      plan.row < 3 ||
      !Array.isArray(plan.before) ||
      !Array.isArray(plan.after) ||
      plan.before.length !== 3 ||
      plan.after.length !== 3 ||
      !plan.before.concat(plan.after).every(function (v) {
        return ["", "2", "3"].includes(v);
      }) ||
      !["", plan.studentId].includes(plan.beforeId) ||
      !["", plan.studentId].includes(plan.afterId) ||
      (!plan.beforeId && !plan.afterId) ||
      (!plan.beforeId && plan.before.some(Boolean)) ||
      (!plan.afterId && plan.after.some(Boolean)) ||
      (plan.afterId &&
        plan.before.some(function (value, p) {
          return p !== intent.period - 1 && value !== plan.after[p];
        })) ||
      typeof plan.previousGeneration !== "string" ||
      typeof plan.generation !== "string" ||
      !plan.generation ||
      typeof plan.previousDeletion !== "string" ||
      typeof plan.deletion !== "string" ||
      !plan.deletion
    )
      blocked();
    seen.add(plan.studentId);
    positions.add(plan.row);
    const existing = data.rows.find(function (row) {
      return row.row === plan.row;
    });
    const id = existing ? existing.studentId : "";
    const values = existing ? existing.values : ["", "", ""];
    if (
      (id !== plan.beforeId && id !== plan.afterId) ||
      data.rows.some(function (row) {
        return row.studentId === plan.studentId && row.row !== plan.row;
      }) ||
      values.some(function (v, p) {
        return v !== plan.before[p] && v !== plan.after[p];
      })
    )
      blocked();
    const revisionKey = metadata.prefix + "V_" + hashToken_(plan.studentId);
    const deletionKey = metadata.prefix + "D_" + hashToken_(plan.studentId);
    const previous = metadata.all[revisionKey]
      ? JSON.parse(metadata.all[revisionKey])
      : null;
    const deletion = metadata.all[deletionKey]
      ? JSON.parse(metadata.all[deletionKey])
      : null;
    if (
      (previous && previous.studentId !== plan.studentId) ||
      ![plan.previousGeneration, plan.generation].includes(
        previous ? previous.generation : "",
      ) ||
      (deletion && deletion.studentId !== plan.studentId) ||
      ![plan.previousDeletion, plan.deletion].includes(
        deletion ? deletion.generation : "",
      )
    )
      blocked();
    revisions[revisionKey] = JSON.stringify({
      studentId: plan.studentId,
      generation: plan.generation,
    });
    revisions[deletionKey] = JSON.stringify({
      studentId: plan.studentId,
      generation: plan.deletion,
    });
    plan.before.forEach(function (v, p) {
      if (v === plan.after[p]) return;
      audits.push([
        new Date(intent.startedAt),
        "교사",
        "extra:" + plan.studentId,
        plan.studentId,
        parseDateKey_(intent.date),
        p + 1,
        v,
        plan.after[p],
      ]);
    });
    while (raw.length < plan.row - 2) raw.push(["", "", "", ""]);
    if (!plan.afterId) raw[plan.row - 3] = ["", "", "", ""];
    else {
      if (!plan.beforeId) raw[plan.row - 3][0] = plan.afterId;
      const value = plan.after[intent.period - 1];
      raw[plan.row - 3][intent.period] = value === "" ? "" : Number(value);
    }
  });
  const expected = new Map(
    audits.map(function (row) {
      return [extraAuditKey_(row), row];
    }),
  );
  if (logLastRow >= intent.auditStartRow) {
    const range = log.getRange(
      intent.auditStartRow,
      1,
      logLastRow - intent.auditStartRow + 1,
      8,
    );
    const formulas = range.getFormulas();
    range.getValues().forEach(function (row, index) {
      if (String(row[2]).indexOf("extra:") !== 0) return;
      const key = extraAuditKey_(row);
      if (formulas[index].some(Boolean) || !expected.has(key)) blocked();
      expected.delete(key);
    });
  }
  // Do not rewrite gap rows, existing IDs, or another period's raw values/formatting.
  const writes = [];
  plans
    .slice()
    .sort(function (a, b) {
      return a.row - b.row;
    })
    .forEach(function (plan) {
      const statusOnly = !!plan.beforeId && !!plan.afterId;
      const col = statusOnly ? block.col + intent.period : block.col;
      const cells = statusOnly
        ? [raw[plan.row - 3][intent.period]]
        : raw[plan.row - 3].slice();
      const current = data.raw[plan.row - 3] || ["", "", "", ""];
      // Completed ranges need no replay after a later range or audit append fails.
      if (
        cells.every(function (value, index) {
          return value === current[statusOnly ? intent.period : index];
        })
      )
        return;
      if (
        !statusOnly &&
        typeof cells[0] === "string" &&
        cells[0].charAt(0) === "="
      )
        cells[0] = "'" + cells[0];
      const previous = writes[writes.length - 1];
      if (
        previous &&
        previous.col === col &&
        previous.row + previous.values.length === plan.row
      )
        previous.values.push(cells);
      else writes.push({ row: plan.row, col: col, values: [cells] });
    });
  // Validate the entire intent and audit suffix before replaying any contiguous ranges.
  metadata.props.setProperties(revisions);
  Object.assign(metadata.all, revisions);
  const sheet = source.sheet;
  const lastRow = writes.reduce(function (last, write) {
    return Math.max(last, write.row + write.values.length - 1);
  }, 2);
  if (sheet.getMaxRows() < lastRow)
    sheet.insertRowsAfter(sheet.getMaxRows(), lastRow - sheet.getMaxRows());
  writes.forEach(function (write) {
    sheet
      .getRange(
        write.row,
        write.col,
        write.values.length,
        write.values[0].length,
      )
      .setValues(write.values);
  });
  SpreadsheetApp.flush();
  appendAudits_(
    Array.from(expected.values()).map(function (row) {
      const cells = row.slice();
      if (String(cells[3]).charAt(0) === "=") cells[3] = "'" + cells[3];
      return cells;
    }),
  );
  SpreadsheetApp.flush();
  const receipt = JSON.stringify({
    payload: intent.payload,
    state: "done",
    kind: "batch",
    changed: plans.length,
  });
  metadata.props.setProperty(receiptKey, receipt);
  metadata.all[receiptKey] = receipt;
  for (let i = 0; i < intent.chunks; i++) {
    metadata.props.deleteProperty(receiptKey + "_C_" + i);
    delete metadata.all[receiptKey + "_C_" + i];
  }
  data.raw = raw;
  const changedRows = new Set(
    plans.map(function (plan) {
      return plan.row;
    }),
  );
  data.rows = data.rows
    .filter(function (row) {
      return !changedRows.has(row.row);
    })
    .concat(
      plans
        .filter(function (plan) {
          return !!plan.afterId;
        })
        .map(function (plan) {
          return {
            studentId: plan.studentId,
            row: plan.row,
            values: plan.after.slice(),
          };
        }),
    );
  data.nextRow = data.rows.reduce(function (nextRow, row) {
    return Math.max(nextRow, row.row + 1);
  }, 3);
  return snapshot;
}

function teacherExtraSave(token, date, period, action, selections, requestId) {
  return publicCall_(function () {
    requireConfig_();
    requireTeacher_(token);
    if (
      typeof date !== "string" ||
      !Number.isInteger(period) ||
      !["save", "remove", "removeAll"].includes(action) ||
      typeof requestId !== "string" ||
      !/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
        requestId,
      )
    )
      throw userError_("변경 요청이 올바르지 않습니다.", "INVALID_REQUEST");
    parseDateKey_(date);
    if (
      !Array.isArray(selections) ||
      selections.length > 1000 ||
      (action === "remove" && selections.length !== 1) ||
      selections.some(function (item) {
        return (
          !item ||
          typeof item.studentId !== "string" ||
          !item.studentId.trim() ||
          item.studentId !== item.studentId.trim() ||
          item.studentId.length > 100 ||
          typeof item.generation !== "string" ||
          item.generation.length > 100 ||
          typeof item.isNew !== "boolean" ||
          typeof item.checked !== "boolean" ||
          (action !== "save" && item.isNew)
        );
      }) ||
      new Set(
        selections.map(function (item) {
          return item.studentId;
        }),
      ).size !== selections.length
    )
      throw userError_("미신청자 목록을 확인해 주세요.", "INVALID_SELECTION");
    const selected = selections
      .map(function (item) {
        return {
          studentId: item.studentId,
          generation: item.generation,
          isNew: item.isNew,
          checked: item.checked,
        };
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
      const metadata = extraMetadata_(date);
      const recovered = recoverExtraPending_(metadata);
      const receiptKey = metadata.prefix + "R_" + requestId.toLowerCase();
      const previous = metadata.all[receiptKey]
        ? JSON.parse(metadata.all[receiptKey])
        : null;
      if (previous && previous.payload !== payload)
        throw userError_(
          "같은 요청 식별자를 다른 작업에 사용할 수 없습니다.",
          "REQUEST_ID_REUSED",
        );
      if (previous && (previous.kind !== "batch" || previous.state !== "done"))
        throw userError_(
          "미신청자 요청 기록을 확인할 수 없습니다. 관리자 확인이 필요합니다.",
          "EXTRA_RECOVERY_REQUIRED",
        );
      // A completed retry may be acknowledged after midnight, but never applied again.
      extraDatePeriod_(date, period, !previous);
      const snapshot = recovered || extraBatchSnapshot_(date);
      const view = extraView_(date, period, metadata, snapshot);
      if (previous) return { changed: previous.changed, view: view };
      if (view.errors.length)
        throw userError_(
          "미신청자 시트 오류를 먼저 확인해 주세요.",
          "INVALID_HEADERS",
        );
      const rows = new Map(
        snapshot.data.rows.map(function (row) {
          return [row.studentId, row];
        }),
      );
      const stale = selected.some(function (item) {
        return (
          item.generation !== (view.generations[item.studentId] || "0") ||
          (!item.isNew && !rows.has(item.studentId)) ||
          (item.isNew &&
            !Object.prototype.hasOwnProperty.call(
              snapshot.directory.students,
              item.studentId,
            ))
        );
      });
      // Whole-date deletion must confirm exactly the membership seen by the teacher.
      if (
        stale ||
        (action === "removeAll" &&
          (rows.size !== selected.length ||
            selected.some(function (item) {
              return !rows.has(item.studentId);
            })))
      )
        throw userError_(
          "목록이 삭제되거나 변경되었습니다. 서버에서 다시 읽고 확인해 주세요.",
          "EXTRA_MEMBERSHIP_CHANGED",
        );
      const log = spreadsheet_().getSheetByName(APP.SHEETS.LOG);
      if (!simpleHeadersValid_(log, APP.LOG_HEADERS))
        throw userError_("기록 시트 헤더를 확인해 주세요.", "INVALID_HEADERS");
      let nextRow = snapshot.data.nextRow;
      const plans = selected
        .map(function (item) {
          const row = rows.get(item.studentId);
          const before = row ? row.values.slice() : ["", "", ""];
          const after = action === "save" ? before.slice() : ["", "", ""];
          if (action === "save") after[period - 1] = item.checked ? "2" : "3";
          const revisionKey =
            metadata.prefix + "V_" + hashToken_(item.studentId);
          const deletionKey =
            metadata.prefix + "D_" + hashToken_(item.studentId);
          const revision = metadata.all[revisionKey]
            ? JSON.parse(metadata.all[revisionKey])
            : null;
          const deletion = metadata.all[deletionKey]
            ? JSON.parse(metadata.all[deletionKey])
            : null;
          return {
            studentId: item.studentId,
            row: row ? row.row : nextRow++,
            beforeId: row ? item.studentId : "",
            afterId: action === "save" ? item.studentId : "",
            before: before,
            after: after,
            previousGeneration: revision ? revision.generation : "",
            generation: Utilities.getUuid(),
            previousDeletion: deletion ? deletion.generation : "",
            deletion:
              action === "save"
                ? view.generations[item.studentId] || "0"
                : Utilities.getUuid(),
          };
        })
        .filter(function (plan) {
          return (
            plan.beforeId !== plan.afterId ||
            plan.before.some(function (v, p) {
              return v !== plan.after[p];
            })
          );
        });
      const size = Utilities.newBlob(JSON.stringify(metadata.all)).getBytes()
        .length;
      if (size > 399000)
        throw userError_(
          "변경 기록 저장 공간을 확인해 주세요.",
          "EXTRA_METADATA_FULL",
        );
      if (!plans.length) {
        metadata.props.setProperty(
          receiptKey,
          JSON.stringify({
            payload: payload,
            state: "done",
            kind: "batch",
            changed: 0,
          }),
        );
        return { changed: 0, view: view };
      }
      const chunks = [];
      let chunk = [];
      plans.forEach(function (plan) {
        if (
          chunk.length &&
          Utilities.newBlob(JSON.stringify(chunk.concat([plan]))).getBytes()
            .length > 7000
        ) {
          chunks.push(JSON.stringify(chunk));
          chunk = [];
        }
        chunk.push(plan);
      });
      if (chunk.length) chunks.push(JSON.stringify(chunk));
      const intent = {
        kind: "batch",
        payload: payload,
        state: "pending",
        date: date,
        period: period,
        startedAt: new Date(
          Math.floor(now_().getTime() / 1000) * 1000,
        ).toISOString(),
        auditStartRow: log.getLastRow() + 1,
        chunks: chunks.length,
      };
      const pending = JSON.stringify(intent);
      if (
        size +
          chunks.reduce(function (total, text) {
            return total + Utilities.newBlob(text).getBytes().length;
          }, 0) +
          plans.length * 800 +
          pending.length >
        400000
      )
        throw userError_(
          "변경 기록 저장 공간을 확인해 주세요.",
          "EXTRA_METADATA_FULL",
        );
      if (!snapshot.block) {
        const block = ensureExtraBlock_(snapshot.source, date);
        snapshot.source.sheet = block.sheet;
        snapshot.block = { key: date, col: block.col };
        snapshot.source.blocks.push(snapshot.block);
        SpreadsheetApp.flush();
      }
      /** @type {{[key: string]: string}} */
      const properties = {};
      chunks.forEach(function (text, index) {
        properties[receiptKey + "_C_" + index] = text;
      });
      // Publish the pending marker only after all immutable chunks are durable.
      metadata.props.setProperties(properties);
      metadata.props.setProperty(receiptKey, pending);
      recoverExtraBatch_(receiptKey, intent, snapshot, metadata, plans, log);
      return {
        changed: plans.length,
        view: extraView_(date, period, metadata, snapshot),
      };
    });
  });
}
