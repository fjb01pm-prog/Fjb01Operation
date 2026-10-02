/**
 * ============================================================
 * FJB OPERATIONS CONTROL SYSTEM
 * AUTO DATABASE + DUMMY DATA + CLEAN WEB APP BACKEND
 * Google Apps Script / Code.gs
 * ============================================================
 *
 * CARA PAKAI:
 * 1. Buat / buka Google Spreadsheet.
 * 2. Extensions > Apps Script.
 * 3. Hapus isi Code.gs lalu paste seluruh script ini.
 * 4. Save.
 * 5. Jalankan fungsi: setupFJBSystem()
 * 6. Izinkan authorization.
 * 7. Kembali ke Spreadsheet -> semua database otomatis dibuat.
 *
 * AKUN DUMMY:
 * ADMIN      : 9000001 / admin123
 * GL         : 3102001 / gl12345
 * OPERATOR   : 3101021 / op12345
 * MECHANIC   : 4102101 / mech12345
 *
 * CATATAN:
 * - Password dibuat plaintext untuk DEMO sesuai kebutuhan saat ini.
 * - Untuk production sebaiknya diganti hash password.
 * - setupFJBSystem() NON-DESTRUCTIVE:
 *   sheet yang sudah berisi data tidak akan dihapus.
 * - resetAndSeedFJBSystem() akan rebuild sheet sistem FJB.
 * ============================================================
 */

const FJB_VERSION = '2.6.5';
const SESSION_TTL_HOURS = 8;
const FJB_DEFAULT_SPREADSHEET_ID = '16SztjFNvt28Hm9cfo_uw_0_5Nhd_LSdsTB9KLyPZPKo';
const FJB_SCRIPT_ID = '1_bQAhI52jUE0sahZ4mVeXS3dihMoffL8LmQPJY5KbGKV9lW0jJvP9_iM';
const FJB_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwFHFYH4trkXYzrLR6Fl0sLH6fF2gToTBJJb9BDcjGLPcPqqAXb5vnbooaYw5_-QrED/exec';

// Fast runtime object cache.
// Spreadsheet/Sheet objects are reused inside a warm Apps Script instance.
// Row values are NOT cached here, so manual sheet changes remain readable.
let FJB_DB_INSTANCE_ = null;
const FJB_SHEET_INSTANCE_CACHE_ = {};


/**
 * Global write lock.
 *
 * Semua write API kritikal menggunakan satu ScriptLock sehingga dua user
 * yang menulis pada waktu hampir bersamaan tidak berebut row / master state.
 */
function withWriteLock_(label, callback) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error(
      'WRITE_BUSY: Sistem sedang menyimpan data user lain. ' +
      'Silakan coba kembali beberapa detik.'
    );
  }

  try {
    const result = callback();

    // Pastikan buffered Spreadsheet writes benar-benar dikirim.
    SpreadsheetApp.flush();

    return result;

  } catch (err) {
    throw new Error(
      (label ? label + ': ' : '') +
      (err && err.message ? err.message : String(err))
    );

  } finally {
    lock.releaseLock();
  }
}


/**
 * Normalisasi value untuk read-back verification.
 */
function normalizeVerifyValue_(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  if (value instanceof Date) {
    // Google Sheet serial date may round sub-second precision.
    // Verify at second precision, which is sufficient for operational timestamps.
    return String(Math.floor(value.getTime() / 1000));
  }

  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }

  if (typeof value === 'number') {
    return String(Number(value));
  }

  return String(value);
}


function verifyValuesEquivalent_(expected, actual) {
  return normalizeVerifyValue_(expected) ===
    normalizeVerifyValue_(actual);
}


/**
 * Verifikasi append dengan membaca kembali primary key kolom pertama.
 */
function verifyAppendRange_(sheetName, startRow, sourceRows) {
  if (!sourceRows || !sourceRows.length) return true;

  const sheet = getSheet_(sheetName);
  const count = sourceRows.length;

  const keyValues = sheet
    .getRange(startRow, 1, count, 1)
    .getValues();

  const expectedFirst = sourceRows[0][0];
  const expectedLast = sourceRows[count - 1][0];

  const actualFirst = keyValues[0][0];
  const actualLast = keyValues[count - 1][0];

  if (
    !verifyValuesEquivalent_(expectedFirst, actualFirst) ||
    !verifyValuesEquivalent_(expectedLast, actualLast)
  ) {
    throw new Error(
      'WRITE_VERIFY_FAILED: ' + sheetName +
      ' append tidak dapat diverifikasi.'
    );
  }

  return true;
}


/**
 * Verifikasi update row dengan membaca kembali field yang diubah.
 */
function verifyRowPatch_(sheetName, rowNumber, patch) {
  const sheet = getSheet_(sheetName);
  const headers = FJB_SCHEMA[sheetName];

  const row = sheet
    .getRange(rowNumber, 1, 1, headers.length)
    .getValues()[0];

  Object.keys(patch || {}).forEach(function(field) {
    const idx = headers.indexOf(field);

    if (idx < 0) return;

    if (!verifyValuesEquivalent_(patch[field], row[idx])) {
      throw new Error(
        'WRITE_VERIFY_FAILED: ' +
        sheetName + ' row ' + rowNumber +
        ' field ' + field +
        ' gagal diverifikasi.'
      );
    }
  });

  return true;
}

function clearDbObjectCache_() {
  FJB_DB_INSTANCE_ = null;
  Object.keys(FJB_SHEET_INSTANCE_CACHE_).forEach(function(k) {
    delete FJB_SHEET_INSTANCE_CACHE_[k];
  });
}


const FJB_SHEETS = {
  CONFIG: '00_CONFIG',
  USERS: '01_USERS',
  PERSONNEL: '02_PERSONNEL',
  UNITS: '03_UNITS',
  UNIT_ALIAS: '04_UNIT_ALIAS',
  OWNER: '05_OWNER',
  PRODUCT: '06_PRODUCT',
  LOCATION: '07_LOCATION',
  OPTIONS: '08_OPTIONS',
  P2H_MASTER: '09_P2H_MASTER',
  ROSTER: '10_ROSTER',
  WASHING: '11_WASHING',
  HAULING: '12_HAULING',
  P2H: '13_P2H',
  P2H_DETAIL: '14_P2H_DETAIL',
  MAINTENANCE: '15_MAINTENANCE',
  UNIT_STATUS_HISTORY: '16_UNIT_STATUS_HISTORY',
  HISTORY: '17_HISTORY',
  SESSIONS: '18_SESSIONS',
  FUEL_USAGE: '19_FUEL_USAGE',
  HM_OPERATION: '20_HM_OPERATION'
};

const FJB_SCHEMA = {
  [FJB_SHEETS.CONFIG]: [
    'key','value','description'
  ],

  [FJB_SHEETS.USERS]: [
    'user_id','nik','password','role_override','status',
    'must_change_password','last_login_at','created_at','updated_at'
  ],

  [FJB_SHEETS.PERSONNEL]: [
    'nik','name','category','position','assigned_unit',
    'team','status','phone','join_date','created_at','updated_at'
  ],

  [FJB_SHEETS.UNITS]: [
    'unit_id','unit_code','unit_type','owner','status',
    'operational_status','hm_km','availability_pct','achievement_pct',
    'tonase_today','ritase_today','assigned_nik','assigned_name',
    'location','active','updated_at'
  ],

  [FJB_SHEETS.UNIT_ALIAS]: [
    'alias_code','canonical_unit','active','note'
  ],

  [FJB_SHEETS.OWNER]: [
    'owner_code','owner_name','active'
  ],

  [FJB_SHEETS.PRODUCT]: [
    'product_code','product_name','active'
  ],

  [FJB_SHEETS.LOCATION]: [
    'location_code','location_name','active'
  ],

  [FJB_SHEETS.OPTIONS]: [
    'option_group','option_code','option_label','sort_order','active'
  ],

  [FJB_SHEETS.P2H_MASTER]: [
    'item_id','item_name','sort_order','active','required'
  ],

  [FJB_SHEETS.ROSTER]: [
    'roster_id','month','date','nik','name','category','position',
    'assigned_unit','roster_status','source',
    'created_by_nik','created_by_name','created_at',
    'updated_by_nik','updated_by_name','updated_at'
  ],

  [FJB_SHEETS.WASHING]: [
    'washing_id','month','plan_date','actual_date','reschedule_date',
    'unit','pic_nik','pic_name','status','note',
    'created_by_nik','created_by_name','created_at',
    'updated_by_nik','updated_by_name','updated_at'
  ],

  [FJB_SHEETS.HAULING]: [
    'transaction_id','date','shift','time','hauler','loader',
    'gross','tare','net_ton','product_seam','remark',
    'input_by_nik','input_by_name','created_at','updated_at',
    'distance','coal_product','jam_ritase','ritase'
  ],

  [FJB_SHEETS.P2H]: [
    'p2h_id','date','shift','time','unit','hm_km',
    'total_items','ok_count','finding_count','notes',
    'input_by_nik','input_by_name','created_at','updated_at'
  ],

  [FJB_SHEETS.P2H_DETAIL]: [
    'detail_id','p2h_id','item_id','item_name',
    'checked','status','note'
  ],

  [FJB_SHEETS.MAINTENANCE]: [
    'maintenance_id','date','unit','type','hm_km',
    'start_time','finish_time','duration_min','result',
    'problem','action','part_material',
    'mechanic_nik','mechanic_name','created_at','updated_at'
  ],

  [FJB_SHEETS.UNIT_STATUS_HISTORY]: [
    'status_id','timestamp','unit',
    'old_status','new_status',
    'old_operational_status','new_operational_status',
    'reason','updated_by_nik','updated_by_name'
  ],

  [FJB_SHEETS.HISTORY]: [
    'log_id','timestamp','user_name','nik','role',
    'module','activity','entity','summary','payload_json'
  ],

  [FJB_SHEETS.SESSIONS]: [
    'token','nik','role','name','position','assigned_unit',
    'created_at','expires_at','last_seen_at','active'
  ],

  [FJB_SHEETS.FUEL_USAGE]: [
    'fuel_id','date','shift','fuel_source','entity_used',
    'unit_code','hm_km','total_liter','fill_time',
    'dedicated','unit_day','location','source',
    'input_by_nik','input_by_name','created_at','updated_at'
  ],

  [FJB_SHEETS.HM_OPERATION]: [
    'hm_id','date','shift','nik','operator_name','unit',
    'hm_start','hm_end','total_hm','note','source',
    'created_by_nik','created_by_name','created_at','updated_at'
  ]
};


/* ============================================================
 * FAST READ CACHE — V2.5
 * ============================================================
 *
 * Google Sheets getValues() is the main latency source.
 * Read-only APIs use CacheService and a per-sheet revision.
 * Writes invalidate only the affected sheet.
 *
 * Cache values are chunked because Apps Script CacheService has
 * a per-value size limit.
 */

const FJB_FAST_CACHE_TTL_SEC = 900;
const FJB_FAST_CACHE_CHUNK = 70000;


function fastCacheRevisionKey_(sheetName) {
  return 'FJB_FC_REV_' + sheetName;
}


function getFastCacheRevision_(sheetName) {
  return (
    PropertiesService
      .getScriptProperties()
      .getProperty(
        fastCacheRevisionKey_(sheetName)
      ) ||
    '0'
  );
}


function getFastViewStamp_() {
  return (
    PropertiesService
      .getScriptProperties()
      .getProperty(
        'FJB_VIEW_STAMP'
      ) ||
    '0'
  );
}


function bumpFastCacheRevision_(
  sheetName,
  forceViewStamp
) {
  if (!sheetName) return;

  const props =
    PropertiesService
      .getScriptProperties();

  props.setProperty(
    fastCacheRevisionKey_(
      sheetName
    ),
    String(Date.now()) +
      '-' +
      Utilities
        .getUuid()
        .substring(0,8)
  );

  /*
   * Runtime audit/session changes should not constantly invalidate
   * the user's browser operational cache.
   *
   * Manual database edits call this function with forceViewStamp=true,
   * so even USERS / HISTORY manual edits become visible on next login.
   */
  const noViewStamp = {};
  noViewStamp[
    FJB_SHEETS.SESSIONS
  ] = true;

  noViewStamp[
    FJB_SHEETS.HISTORY
  ] = true;

  noViewStamp[
    FJB_SHEETS.USERS
  ] = true;

  noViewStamp[
    FJB_SHEETS.UNIT_STATUS_HISTORY
  ] = true;

  if (
    forceViewStamp ||
    !noViewStamp[sheetName]
  ) {
    touchFastViewStamp_(
      'SHEET:' + sheetName
    );
  }
}


function fastCacheCanUse_(sheetName) {
  /*
   * Never cache authentication/session source tables.
   */
  return (
    sheetName !== FJB_SHEETS.SESSIONS &&
    sheetName !== FJB_SHEETS.USERS
  );
}


function fastCacheKey_(
  sheetName,
  scope
) {
  return [
    'FJBFC',
    sheetName,
    getFastCacheRevision_(
      sheetName
    ),
    scope
  ].join(':');
}


function cachePutLargeString_(
  baseKey,
  text,
  ttl
) {
  try {
    const cache =
      CacheService
        .getScriptCache();

    const chunks = [];

    for (
      let i = 0;
      i < text.length;
      i += FJB_FAST_CACHE_CHUNK
    ) {
      chunks.push(
        text.substring(
          i,
          i +
          FJB_FAST_CACHE_CHUNK
        )
      );
    }

    const values = {};

    values[
      baseKey + ':manifest'
    ] = JSON.stringify({
      count:chunks.length
    });

    chunks.forEach(
      function(chunk, i) {
        values[
          baseKey + ':c' + i
        ] = chunk;
      }
    );

    cache.putAll(
      values,
      Number(
        ttl ||
        FJB_FAST_CACHE_TTL_SEC
      )
    );

    return true;

  } catch (err) {
    console.warn(
      'FAST_CACHE_PUT_WARNING: ' +
      (
        err && err.message
          ? err.message
          : String(err)
      )
    );

    return false;
  }
}


function cacheGetLargeString_(
  baseKey
) {
  try {
    const cache =
      CacheService
        .getScriptCache();

    const manifestRaw =
      cache.get(
        baseKey + ':manifest'
      );

    if (!manifestRaw) {
      return '';
    }

    const manifest =
      JSON.parse(
        manifestRaw
      );

    const count =
      Number(
        manifest.count || 0
      );

    if (!count) {
      return '';
    }

    const keys = [];

    for (
      let i = 0;
      i < count;
      i++
    ) {
      keys.push(
        baseKey + ':c' + i
      );
    }

    const values =
      cache.getAll(keys);

    let out = '';

    for (
      let i = 0;
      i < count;
      i++
    ) {
      const value =
        values[
          baseKey + ':c' + i
        ];

      if (
        value === undefined ||
        value === null
      ) {
        return '';
      }

      out += value;
    }

    return out;

  } catch (err) {
    console.warn(
      'FAST_CACHE_GET_WARNING: ' +
      (
        err && err.message
          ? err.message
          : String(err)
      )
    );

    return '';
  }
}


function cacheRowsForRead_(
  rows
) {
  return (rows || [])
    .map(function(x) {
      const out =
        cleanObject_(x);

      if (
        x &&
        x._row !== undefined
      ) {
        out._row = x._row;
      }

      return out;
    });
}


function readObjectsFast_(
  sheetName
) {
  if (
    !fastCacheCanUse_(
      sheetName
    )
  ) {
    return readObjects_(
      sheetName
    );
  }

  const key =
    fastCacheKey_(
      sheetName,
      'ALL'
    );

  const cached =
    cacheGetLargeString_(
      key
    );

  if (cached) {
    try {
      return JSON.parse(
        cached
      );
    } catch (err) {
      // fall through
    }
  }

  const rows =
    cacheRowsForRead_(
      readObjects_(
        sheetName
      )
    );

  cachePutLargeString_(
    key,
    JSON.stringify(rows),
    FJB_FAST_CACHE_TTL_SEC
  );

  return rows;
}


function readObjectsTailFast_(
  sheetName,
  maxRows
) {
  if (
    !fastCacheCanUse_(
      sheetName
    )
  ) {
    return readObjectsTail_(
      sheetName,
      maxRows
    );
  }

  const limit =
    Math.max(
      1,
      Number(
        maxRows || 500
      )
    );

  const key =
    fastCacheKey_(
      sheetName,
      'TAIL-' + limit
    );

  const cached =
    cacheGetLargeString_(
      key
    );

  if (cached) {
    try {
      return JSON.parse(
        cached
      );
    } catch (err) {
      // fall through
    }
  }

  const rows =
    cacheRowsForRead_(
      readObjectsTail_(
        sheetName,
        limit
      )
    );

  cachePutLargeString_(
    key,
    JSON.stringify(rows),
    FJB_FAST_CACHE_TTL_SEC
  );

  return rows;
}


function apiGetCacheStampJson(
  token
) {
  requireSession_(token);

  return JSON.stringify({
    ok:true,
    stamp:
      getFastViewStamp_()
  });
}



/* ============================================================
 * DATABASE DIRECT-EDIT DETECTOR — V2.5.1
 * ============================================================
 *
 * IMPORTANT:
 * Project ini standalone, jadi trigger dipasang ke Spreadsheet
 * database secara installable (bukan simple onEdit).
 *
 * Run ONCE dari Apps Script editor:
 *   installFJBDatabaseChangeTriggers()
 */


function touchFastViewStamp_(reason) {
  const value =
    String(Date.now()) +
    '-' +
    Utilities
      .getUuid()
      .substring(0,8);

  PropertiesService
    .getScriptProperties()
    .setProperties(
      {
        FJB_VIEW_STAMP:value,
        FJB_VIEW_STAMP_REASON:
          String(reason || ''),
        FJB_VIEW_STAMP_AT:
          new Date().toISOString()
      },
      false
    );

  return value;
}


function isFJBDatabaseSheet_(
  sheetName
) {
  if (!sheetName) {
    return false;
  }

  return Object.keys(
    FJB_SCHEMA
  ).indexOf(
    String(sheetName)
  ) >= 0;
}


function fjbDatabaseSheetNames_() {
  return Object.keys(
    FJB_SCHEMA
  );
}


function invalidateAllFJBDatabaseCaches_(
  reason
) {
  const props =
    PropertiesService
      .getScriptProperties();

  const values = {};
  const now =
    String(Date.now());

  fjbDatabaseSheetNames_()
    .forEach(function(sheetName) {
      /*
       * Session rows are temporary and do not belong to
       * user-facing business data.
       */
      if (
        sheetName ===
        FJB_SHEETS.SESSIONS
      ) {
        return;
      }

      values[
        fastCacheRevisionKey_(
          sheetName
        )
      ] =
        now +
        '-' +
        Utilities
          .getUuid()
          .substring(0,8);
    });

  values.FJB_VIEW_STAMP =
    now +
    '-' +
    Utilities
      .getUuid()
      .substring(0,8);

  values.FJB_VIEW_STAMP_REASON =
    String(
      reason ||
      'DATABASE_CHANGE'
    );

  values.FJB_VIEW_STAMP_AT =
    new Date()
      .toISOString();

  props.setProperties(
    values,
    false
  );

  return {
    ok:true,
    invalidated:
      Object.keys(values)
        .filter(function(k) {
          return k.indexOf(
            'FJB_FC_REV_'
          ) === 0;
        }).length,
    stamp:
      values.FJB_VIEW_STAMP
  };
}


/**
 * Installable spreadsheet ON EDIT handler.
 *
 * Direct cell edits / paste ranges:
 * only the edited sheet cache is invalidated.
 */
function handleFJBDatabaseEdit(e) {
  try {
    if (
      !e ||
      !e.range
    ) {
      return;
    }

    const sheet =
      e.range.getSheet();

    const sheetName =
      sheet.getName();

    if (
      !isFJBDatabaseSheet_(
        sheetName
      )
    ) {
      return;
    }

    if (
      sheetName ===
      FJB_SHEETS.SESSIONS
    ) {
      return;
    }

    bumpFastCacheRevision_(
      sheetName,
      true
    );

    console.log(
      'FJB_DB_EDIT_DETECTED: ' +
      sheetName +
      '!' +
      e.range.getA1Notation()
    );

  } catch (err) {
    console.warn(
      'FJB_DB_EDIT_TRIGGER_WARNING: ' +
      (
        err && err.message
          ? err.message
          : String(err)
      )
    );
  }
}


/**
 * Installable spreadsheet ON CHANGE handler.
 *
 * Used for structural changes such as:
 * INSERT_ROW, REMOVE_ROW, INSERT_COLUMN, REMOVE_COLUMN,
 * INSERT_GRID, REMOVE_GRID, FORMAT, OTHER.
 *
 * EDIT is ignored because handleFJBDatabaseEdit() already handles it.
 */
function handleFJBDatabaseChange(e) {
  try {
    const changeType =
      String(
        e && e.changeType
          ? e.changeType
          : 'UNKNOWN'
      )
      .toUpperCase();

    if (
      changeType === 'EDIT'
    ) {
      return;
    }

    const result =
      invalidateAllFJBDatabaseCaches_(
        'STRUCTURE:' +
        changeType
      );

    console.log(
      'FJB_DB_CHANGE_DETECTED: ' +
      changeType +
      ' | invalidated=' +
      result.invalidated
    );

  } catch (err) {
    console.warn(
      'FJB_DB_CHANGE_TRIGGER_WARNING: ' +
      (
        err && err.message
          ? err.message
          : String(err)
      )
    );
  }
}


function removeFJBDatabaseChangeTriggers() {
  const handlerNames = {
    handleFJBDatabaseEdit:true,
    handleFJBDatabaseChange:true
  };

  let removed = 0;

  ScriptApp
    .getProjectTriggers()
    .forEach(function(trigger) {
      const handler =
        trigger
          .getHandlerFunction();

      if (
        handlerNames[handler]
      ) {
        ScriptApp
          .deleteTrigger(
            trigger
          );

        removed++;
      }
    });

  return {
    ok:true,
    removed:removed
  };
}


/**
 * Run ONCE manually from Apps Script editor after deploying V2.5.1.
 *
 * This function:
 * 1. removes old duplicate FJB DB triggers;
 * 2. attaches ON EDIT to the actual database spreadsheet;
 * 3. attaches ON CHANGE to the same spreadsheet;
 * 4. invalidates all current caches once.
 */
function installFJBDatabaseChangeTriggers() {
  const db =
    getDb_();

  const spreadsheetId =
    db.getId();

  removeFJBDatabaseChangeTriggers();

  const editTrigger =
    ScriptApp
      .newTrigger(
        'handleFJBDatabaseEdit'
      )
      .forSpreadsheet(
        spreadsheetId
      )
      .onEdit()
      .create();

  const changeTrigger =
    ScriptApp
      .newTrigger(
        'handleFJBDatabaseChange'
      )
      .forSpreadsheet(
        spreadsheetId
      )
      .onChange()
      .create();

  const invalidate =
    invalidateAllFJBDatabaseCaches_(
      'TRIGGER_INSTALL'
    );

  const result = {
    ok:true,

    version:
      FJB_VERSION,

    database_name:
      db.getName(),

    spreadsheet_id:
      spreadsheetId,

    edit_trigger_id:
      editTrigger.getUniqueId(),

    change_trigger_id:
      changeTrigger.getUniqueId(),

    cache_stamp:
      invalidate.stamp
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


/**
 * Diagnostic: run directly from Apps Script editor.
 */
function checkFJBDatabaseChangeTriggers() {
  const db =
    getDb_();

  const triggers =
    ScriptApp
      .getProjectTriggers()
      .map(function(trigger) {
        return {
          handler:
            trigger
              .getHandlerFunction(),

          source:
            String(
              trigger.getTriggerSource()
            ),

          event_type:
            String(
              trigger.getEventType()
            ),

          id:
            trigger.getUniqueId()
        };
      })
      .filter(function(x) {
        return (
          x.handler ===
            'handleFJBDatabaseEdit' ||
          x.handler ===
            'handleFJBDatabaseChange'
        );
      });

  const result = {
    ok:
      triggers.length >= 2,

    version:
      FJB_VERSION,

    database_name:
      db.getName(),

    spreadsheet_id:
      db.getId(),

    view_stamp:
      getFastViewStamp_(),

    view_stamp_reason:
      PropertiesService
        .getScriptProperties()
        .getProperty(
          'FJB_VIEW_STAMP_REASON'
        ) ||
      '',

    view_stamp_at:
      PropertiesService
        .getScriptProperties()
        .getProperty(
          'FJB_VIEW_STAMP_AT'
        ) ||
      '',

    triggers:
      triggers
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


/**
 * Manual emergency cache invalidation.
 *
 * Useful after a large external import if the change trigger
 * has not run yet.
 */
function forceRefreshFJBCache() {
  const result =
    invalidateAllFJBDatabaseCaches_(
      'MANUAL_FORCE_REFRESH'
    );

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}




const ROLE_MENUS = {
  ADMIN: [
    'dashboard','input','daily','roster','units','hauling','fuel','hm','washing',
    'maintenance','manpower','reports','history','master',
    'p2hmaster','users','settings'
  ],
  GL: [
    'dashboard','daily','roster','units','hauling','fuel','hm','washing',
    'maintenance','manpower','reports','history','master'
  ],
  OPERATOR: [
    'dashboard','roster','fuel','hm','washing'
  ],
  MECHANIC: [
    'dashboard','roster'
  ]
};



/* ============================================================
 * STANDALONE / BOUND SCRIPT SAFE UI HELPERS
 * ============================================================
 */

/**
 * Coba ambil Spreadsheet UI.
 * Pada standalone Apps Script / web app context, getUi() bisa tidak tersedia.
 */
function getSpreadsheetUiSafe_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (err) {
    return null;
  }
}


/**
 * Alert aman:
 * - bound spreadsheet -> popup;
 * - standalone/editor/web app -> Execution Log.
 */
function safeUiAlert_(title, message) {
  const ui = getSpreadsheetUiSafe_();
  const text = String(message || '');

  if (ui) {
    try {
      ui.alert(
        String(title || 'FJB SYSTEM'),
        text,
        ui.ButtonSet.OK
      );

      return {
        shown_in_ui:true,
        logged:false
      };

    } catch (err) {
      // fallback ke log
    }
  }

  console.log(
    '\n========== ' +
    String(title || 'FJB SYSTEM') +
    ' ==========\n' +
    text +
    '\n========================================'
  );

  return {
    shown_in_ui:false,
    logged:true
  };
}


/**
 * Toast aman untuk spreadsheet object.
 */
function safeToast_(ss, message, title, seconds) {
  try {
    if (ss && typeof ss.toast === 'function') {
      ss.toast(
        String(message || ''),
        String(title || 'FJB SYSTEM'),
        Number(seconds || 5)
      );
      return true;
    }
  } catch (err) {}

  console.log(
    '[' + String(title || 'FJB SYSTEM') + '] ' +
    String(message || '')
  );

  return false;
}


/* ============================================================
 * 01. MENU & SETUP
 * ============================================================
 */

function onOpen() {
  const ui = getSpreadsheetUiSafe_();

  // Standalone Apps Script tidak memiliki Spreadsheet UI.
  if (!ui) {
    console.log(
      'FJB SYSTEM: onOpen menu dilewati karena project bukan bound Spreadsheet UI.'
    );
    return;
  }

  ui.createMenu('FJB SYSTEM')
    .addItem('1. Setup / Repair Database', 'setupFJBSystem')
    .addItem('2. Seed Dummy Jika Kosong', 'seedFJBDummyData')
    .addItem('3. Repair Roster & Washing', 'repairRosterWashingDatabase')
    .addItem('4. Sync Personnel ↔ Unit', 'repairPersonnelUnitLinks')
    .addSeparator()
    .addItem('Status Database', 'showFJBDatabaseStatus')
    .addItem('Run System Self-Test', 'runFJBSelfTest')
    .addItem('Show Write Safety Status', 'showWriteSafetyStatus')
    .addSeparator()
    .addItem('RESET & Rebuild Dummy', 'resetAndSeedFJBSystem')
    .addToUi();
}


/**
 * Fungsi utama.
 * Jalankan SEKALI setelah script dipaste.
 */
function setupFJBSystem() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDb_();

    PropertiesService.getScriptProperties()
      .setProperty('FJB_SPREADSHEET_ID', ss.getId());

    PropertiesService.getScriptProperties()
      .setProperty('FJB_TIMEZONE', ss.getSpreadsheetTimeZone() || 'Asia/Makassar');

    Object.keys(FJB_SCHEMA).forEach(function(sheetName) {
      ensureSheet_(ss, sheetName, FJB_SCHEMA[sheetName]);
    });

    ensureOperationalExtensionV260_(true);
    seedFJBDummyData_(ss);
    repairRosterWashingDatabase_(ss);
    repairPersonnelUnitLinks_();
    formatAllSheets_(ss);

    appendHistorySystem_(
      'System',
      'SETUP',
      'DATABASE',
      'Database FJB berhasil dibuat / diperbaiki',
      { version: FJB_VERSION }
    );

    SpreadsheetApp.flush();

    safeToast_(
      ss,
      'Database FJB siap. Dummy data sudah dibuat.',
      'FJB SYSTEM',
      8
    );

    return getDatabaseStatus_();

  } finally {
    lock.releaseLock();
  }
}


/**
 * Alias jika user ingin nama fungsi yang jelas.
 */
function autoSetupFJBDatabase() {
  return setupFJBSystem();
}


function seedFJBDummyData() {
  const ss = getDb_();
  seedFJBDummyData_(ss);
  formatAllSheets_(ss);

  safeToast_(
    ss,
    'Dummy data ditambahkan hanya pada sheet yang masih kosong.',
    'FJB SYSTEM',
    6
  );

  return getDatabaseStatus_();
}


/**
 * RESET hanya sheet yang dimiliki sistem FJB.
 * Sheet lain milik user tidak dihapus.
 */
function resetAndSeedFJBSystem() {
  clearDbObjectCache_();

  const ui = getSpreadsheetUiSafe_();

  // Demi keamanan: standalone tidak boleh reset hanya karena fungsi ter-run.
  if (!ui) {
    const message =
      'RESET DIBATALKAN.\n\n' +
      'Project Apps Script ini tidak memiliki Spreadsheet UI, sehingga ' +
      'konfirmasi YES/NO tidak dapat ditampilkan.\n\n' +
      'Jika benar-benar ingin reset database, buka Apps Script dari ' +
      'Google Sheet (Extensions > Apps Script), lalu jalankan kembali.';

    console.log(message);
    throw new Error(
      'RESET_REQUIRES_BOUND_SPREADSHEET_UI'
    );
  }

  const answer = ui.alert(
    'RESET DATABASE FJB',
    'Semua data pada sheet database FJB akan dihapus lalu dibuat ulang ' +
    'dengan dummy data. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (answer !== ui.Button.YES) {
    return {
      ok:false,
      cancelled:true
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDb_();

    Object.keys(FJB_SCHEMA).forEach(function(sheetName) {
      const sh = ss.getSheetByName(sheetName);

      if (sh) {
        ss.deleteSheet(sh);
      }
    });

    clearDbObjectCache_();

    const freshDb = getDb_();

    Object.keys(FJB_SCHEMA).forEach(function(sheetName) {
      ensureSheet_(
        freshDb,
        sheetName,
        FJB_SCHEMA[sheetName]
      );
    });

    seedFJBDummyData_(freshDb);
    repairRosterWashingDatabase_(freshDb);
    repairPersonnelUnitLinks_();
    formatAllSheets_(freshDb);

    appendHistorySystem_(
      'System',
      'RESET',
      'DATABASE',
      'Database FJB di-reset dan dummy data dibuat ulang',
      { version:FJB_VERSION }
    );

    SpreadsheetApp.flush();

    safeToast_(
      freshDb,
      'Reset selesai. Database dan dummy data sudah dibuat ulang.',
      'FJB SYSTEM',
      8
    );

    return getDatabaseStatus_();

  } finally {
    lock.releaseLock();
  }
}



function showWriteSafetyStatus() {
  const text =
    'FJB WRITE SAFETY V1.9.2\n\n' +
    '✓ Global write-lock aktif\n' +
    '✓ Append read-back verification aktif\n' +
    '✓ Update read-back verification aktif\n' +
    '✓ Delete verification aktif\n' +
    '✓ Roster batch lock aktif\n' +
    '✓ P2H transaction lock aktif\n\n' +
    'Database tetap Google Sheet.';

  console.log(text);

  safeUiAlert_(
    'FJB WRITE SAFETY',
    text
  );

  return {
    ok:true,
    version:FJB_VERSION,
    write_lock:true,
    append_verify:true,
    update_verify:true,
    delete_verify:true
  };
}


function showFJBDatabaseStatus() {
  const status = getDatabaseStatus_();

  const lines = status.sheets.map(function(x) {
    return x.name + ' = ' + x.rows + ' data row';
  });

  const text =
    'Spreadsheet ID:\n' +
    status.spreadsheetId +
    '\n\n' +
    lines.join('\n');

  console.log(
    'FJB DATABASE STATUS\n' + text
  );

  safeUiAlert_(
    'FJB DATABASE STATUS',
    text
  );

  return status;
}


/* ============================================================
 * 02. DATABASE CORE
 * ============================================================
 */

function getDb_() {
  if (FJB_DB_INSTANCE_) {
    return FJB_DB_INSTANCE_;
  }

  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('FJB_SPREADSHEET_ID') || FJB_DEFAULT_SPREADSHEET_ID;

  if (id) {
    try {
      FJB_DB_INSTANCE_ = SpreadsheetApp.openById(id);
      return FJB_DB_INSTANCE_;
    } catch (openErr) {
      // Fallback ke active spreadsheet jika openById gagal
    }
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    id = active.getId();
    props.setProperty('FJB_SPREADSHEET_ID', id);
    FJB_DB_INSTANCE_ = active;
    return FJB_DB_INSTANCE_;
  }

  throw new Error(
    'Database belum terhubung. Pastikan FJB_DEFAULT_SPREADSHEET_ID atau setupFJBSystem() sudah dijalankan.'
  );
}


function getTimezone_() {
  return PropertiesService.getScriptProperties()
    .getProperty('FJB_TIMEZONE') || 'Asia/Makassar';
}


function ensureSheet_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      headers.length - sheet.getMaxColumns()
    );
  }

  const existingHeaders = sheet
    .getRange(1, 1, 1, headers.length)
    .getValues()[0];

  const needHeader =
    sheet.getLastRow() === 0 ||
    existingHeaders.join('|') !== headers.join('|');

  if (needHeader && sheet.getLastRow() <= 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.setFrozenRows(1);

  return sheet;
}

function ensureSheetSchemaExtended_(ss, sheetName, expectedHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  let created = false;
  const added = [];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    created = true;
  }

  const currentLastCol = Math.max(1, sheet.getLastColumn());
  let currentHeaders = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, currentLastCol).getValues()[0]
        .map(function(x) { return String(x).trim(); })
    : [];

  if (!currentHeaders.some(Boolean)) {
    if (sheet.getMaxColumns() < expectedHeaders.length) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(),
        expectedHeaders.length - sheet.getMaxColumns()
      );
    }
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    currentHeaders = expectedHeaders.slice();
    expectedHeaders.forEach(function(h) { added.push(h); });
  } else {
    expectedHeaders.forEach(function(header) {
      if (currentHeaders.indexOf(header) >= 0) return;
      const col = currentHeaders.length + 1;
      if (sheet.getMaxColumns() < col) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
      }
      sheet.getRange(1, col).setValue(header);
      currentHeaders.push(header);
      added.push(header);
    });
  }

  sheet.setFrozenRows(1);

  if (created || added.length) {
    sheet.getRange(1, 1, 1, currentHeaders.length)
      .setBackground('#10283A')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 28);
    bumpFastCacheRevision_(sheetName, true);
  }

  return {created:created, added:added};
}


/* ============================================================
 * HM DATABASE SYNC — V2.6.4
 * ============================================================
 *
 * The ONLY Web App HM database is:
 *   20_HM_OPERATION
 *
 * Legacy sheets are migration sources only.
 */


function hmHeaderKeyV264_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/[\/\-]+/g, ' ')
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}


function hmFindHeaderV264_(headers, aliases) {
  const normalized =
    headers.map(
      hmHeaderKeyV264_
    );

  for (
    let i = 0;
    i < aliases.length;
    i++
  ) {
    const idx =
      normalized.indexOf(
        aliases[i]
      );

    if (idx >= 0) {
      return idx;
    }
  }

  return -1;
}


function hmLegacyMapV264_(headers) {
  const map = {
    date:
      hmFindHeaderV264_(
        headers,
        ['TANGGAL','DATE']
      ),

    shift:
      hmFindHeaderV264_(
        headers,
        ['SHIFT']
      ),

    nik:
      hmFindHeaderV264_(
        headers,
        ['NIK','NRP']
      ),

    operator:
      hmFindHeaderV264_(
        headers,
        [
          'NAMA_OPERATOR',
          'OPERATOR',
          'NAMA'
        ]
      ),

    unit:
      hmFindHeaderV264_(
        headers,
        [
          'NO_UNIT',
          'UNIT',
          'UNIT_CODE'
        ]
      ),

    hmStart:
      hmFindHeaderV264_(
        headers,
        [
          'HM_AWAL',
          'HM_START'
        ]
      ),

    hmEnd:
      hmFindHeaderV264_(
        headers,
        [
          'HM_AKHIR',
          'HM_END'
        ]
      ),

    total:
      hmFindHeaderV264_(
        headers,
        ['TOTAL_HM']
      ),

    note:
      hmFindHeaderV264_(
        headers,
        [
          'CATATAN',
          'NOTE',
          'NOTES',
          'REMARK'
        ]
      )
  };

  map.isLegacy =
    map.date >= 0 &&
    map.operator >= 0 &&
    map.unit >= 0 &&
    map.hmStart >= 0 &&
    map.hmEnd >= 0;

  return map;
}


function hmCanonicalHeaderV264_(headers) {
  const expected =
    FJB_SCHEMA[
      FJB_SHEETS.HM_OPERATION
    ];

  if (
    headers.length <
    expected.length
  ) {
    return false;
  }

  for (
    let i = 0;
    i < expected.length;
    i++
  ) {
    if (
      String(headers[i] || '')
        .trim() !==
      expected[i]
    ) {
      return false;
    }
  }

  return true;
}


function hmShiftV264_(value) {
  const s =
    String(value || '')
      .trim()
      .toUpperCase();

  if (
    s === 'N' ||
    s === 'NIGHT'
  ) {
    return 'NIGHT';
  }

  if (
    s === 'D' ||
    s === 'DAY'
  ) {
    return 'DAY';
  }

  return '';
}


function hmPersonnelLookupV264_() {
  const byNik = {};
  const byName = {};

  readObjects_(
    FJB_SHEETS.PERSONNEL
  ).forEach(function(p) {
    const nik =
      String(p.nik || '')
        .trim();

    if (nik) {
      byNik[nik] = p;
    }

    const nameKey =
      String(p.name || '')
        .trim()
        .toUpperCase()
        .replace(
          /[^A-Z0-9]/g,
          ''
        );

    if (nameKey) {
      byName[nameKey] = p;
    }
  });

  return {
    byNik:byNik,
    byName:byName
  };
}


function hmResolvePersonV264_(
  nik,
  name,
  lookup
) {
  const nikKey =
    String(nik || '')
      .trim();

  if (
    nikKey &&
    lookup.byNik[
      nikKey
    ]
  ) {
    return lookup.byNik[
      nikKey
    ];
  }

  const nameKey =
    String(name || '')
      .trim()
      .toUpperCase()
      .replace(
        /[^A-Z0-9]/g,
        ''
      );

  return (
    nameKey &&
    lookup.byName[
      nameKey
    ]
  )
    ? lookup.byName[
        nameKey
      ]
    : null;
}


function hmLegacyRowV264_(
  row,
  map,
  lookup,
  sourceName
) {
  const dateKey =
    toIsoDate_(
      row[map.date]
    );

  const operatorName =
    String(
      row[map.operator] || ''
    )
      .trim();

  const unit =
    String(
      row[map.unit] || ''
    )
      .trim()
      .toUpperCase();

  const hmStart =
    Number(
      row[map.hmStart]
    );

  const hmEnd =
    Number(
      row[map.hmEnd]
    );

  if (
    !dateKey ||
    !isValidIsoDate_(
      dateKey
    ) ||
    !unit ||
    !Number.isFinite(
      hmStart
    ) ||
    !Number.isFinite(
      hmEnd
    ) ||
    hmEnd < hmStart
  ) {
    return null;
  }

  const rawNik =
    map.nik >= 0
      ? String(
          row[map.nik] || ''
        )
      : '';

  const person =
    hmResolvePersonV264_(
      rawNik,
      operatorName,
      lookup
    );

  const shift =
    (
      map.shift >= 0
        ? hmShiftV264_(
            row[map.shift]
          )
        : ''
    ) ||
    'DAY';

  let note =
    map.note >= 0
      ? String(
          row[map.note] || ''
        )
      : '';

  if (
    map.shift < 0
  ) {
    note +=
      (
        note
          ? ' | '
          : ''
      ) +
      'Legacy: shift default DAY';
  }

  const total =
    Number(
      (
        hmEnd -
        hmStart
      ).toFixed(2)
    );

  return {
    hm_id:
      generateId_('HM'),

    date:
      parseIsoDate_(
        dateKey
      ),

    shift:
      shift,

    nik:
      person
        ? String(
            person.nik
          )
        : rawNik,

    operator_name:
      person
        ? person.name
        : operatorName,

    unit:
      unit,

    hm_start:
      hmStart,

    hm_end:
      hmEnd,

    total_hm:
      total,

    note:
      note,

    source:
      'LEGACY_MIGRATION:' +
      sourceName,

    created_by_nik:
      'SYSTEM',

    created_by_name:
      'HM MIGRATION',

    created_at:
      new Date(),

    updated_at:''
  };
}


function hmShiftedLegacyCanonicalV264_(obj) {
  /*
   * Detect old 7-column data pasted directly under canonical headers:
   *
   * hm_id          <- NO
   * date           <- TANGGAL
   * shift          <- NAMA OPERATOR
   * nik            <- NO UNIT
   * operator_name  <- HM AWAL
   * unit           <- HM AKHIR
   * hm_start       <- TOTAL HM
   */
  const dateKey =
    toIsoDate_(
      obj.date
    );

  const validShift =
    hmShiftV264_(
      obj.shift
    );

  return (
    !!dateKey &&
    !validShift &&
    !!String(
      obj.nik || ''
    ).trim() &&
    Number.isFinite(
      Number(
        obj.operator_name
      )
    ) &&
    Number.isFinite(
      Number(
        obj.unit
      )
    ) &&
    (
      obj.hm_end === '' ||
      obj.hm_end === null ||
      obj.hm_end === undefined
    )
  );
}


function hmNormalizeCanonicalRowV264_(
  obj,
  lookup
) {
  if (
    hmShiftedLegacyCanonicalV264_(
      obj
    )
  ) {
    const operatorName =
      String(
        obj.shift || ''
      ).trim();

    const unit =
      String(
        obj.nik || ''
      )
        .trim()
        .toUpperCase();

    const hmStart =
      Number(
        obj.operator_name
      );

    const hmEnd =
      Number(
        obj.unit
      );

    const person =
      hmResolvePersonV264_(
        '',
        operatorName,
        lookup
      );

    return {
      hm_id:
        generateId_('HM'),

      date:
        parseIsoDate_(
          toIsoDate_(
            obj.date
          )
        ),

      shift:
        'DAY',

      nik:
        person
          ? String(
              person.nik
            )
          : '',

      operator_name:
        person
          ? person.name
          : operatorName,

      unit:
        unit,

      hm_start:
        hmStart,

      hm_end:
        hmEnd,

      total_hm:
        Number(
          (
            hmEnd -
            hmStart
          ).toFixed(2)
        ),

      note:
        'Legacy direct-paste normalized; shift default DAY',

      source:
        'LEGACY_DIRECT',

      created_by_nik:
        'SYSTEM',

      created_by_name:
        'HM NORMALIZER',

      created_at:
        new Date(),

      updated_at:''
    };
  }

  const dateKey =
    toIsoDate_(
      obj.date
    );

  const unit =
    String(
      obj.unit || ''
    )
      .trim()
      .toUpperCase();

  const hmStart =
    Number(
      obj.hm_start
    );

  const hmEnd =
    Number(
      obj.hm_end
    );

  if (
    !dateKey ||
    !isValidIsoDate_(
      dateKey
    ) ||
    !unit ||
    !Number.isFinite(
      hmStart
    ) ||
    !Number.isFinite(
      hmEnd
    ) ||
    hmEnd < hmStart
  ) {
    return obj;
  }

  const person =
    hmResolvePersonV264_(
      obj.nik,
      obj.operator_name,
      lookup
    );

  return {
    hm_id:
      String(
        obj.hm_id || ''
      ).trim() ||
      generateId_('HM'),

    date:
      parseIsoDate_(
        dateKey
      ),

    shift:
      hmShiftV264_(
        obj.shift
      ) ||
      'DAY',

    nik:
      person
        ? String(
            person.nik
          )
        : String(
            obj.nik || ''
          ),

    operator_name:
      person
        ? person.name
        : String(
            obj.operator_name || ''
          ),

    unit:
      unit,

    hm_start:
      hmStart,

    hm_end:
      hmEnd,

    total_hm:
      Number(
        (
          hmEnd -
          hmStart
        ).toFixed(2)
      ),

    note:
      String(
        obj.note || ''
      ),

    source:
      String(
        obj.source ||
        'DATABASE'
      ),

    created_by_nik:
      String(
        obj.created_by_nik || ''
      ),

    created_by_name:
      String(
        obj.created_by_name || ''
      ),

    created_at:
      obj.created_at ||
      new Date(),

    updated_at:
      obj.updated_at || ''
  };
}


function hmRewriteCanonicalV264_(
  sheet,
  records
) {
  const headers =
    FJB_SCHEMA[
      FJB_SHEETS.HM_OPERATION
    ];

  if (
    sheet.getMaxColumns() <
    headers.length
  ) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      headers.length -
      sheet.getMaxColumns()
    );
  }

  sheet.clearContents();

  sheet.getRange(
    1,
    1,
    1,
    headers.length
  ).setValues([
    headers
  ]);

  if (
    records &&
    records.length
  ) {
    const values =
      records.map(
        function(obj) {
          return headers.map(
            function(h) {
              return obj[h] ===
                undefined
                  ? ''
                  : obj[h];
            }
          );
        }
      );

    sheet.getRange(
      2,
      1,
      values.length,
      headers.length
    ).setValues(
      values
    );
  }

  sheet.setFrozenRows(1);

  sheet.getRange(
    1,
    1,
    1,
    headers.length
  )
  .setBackground('#10283A')
  .setFontColor('#FFFFFF')
  .setFontWeight('bold')
  .setHorizontalAlignment(
    'center'
  );

  SpreadsheetApp.flush();

  bumpFastCacheRevision_(
    FJB_SHEETS.HM_OPERATION,
    true
  );
}



function findLegacyHMHeaderV265_(
  sheet
) {
  if (
    !sheet ||
    sheet.getLastRow() < 2 ||
    sheet.getLastColumn() < 5
  ) {
    return null;
  }

  const maxHeaderRow =
    Math.min(
      10,
      sheet.getLastRow()
    );

  const width =
    Math.min(
      30,
      sheet.getLastColumn()
    );

  const values =
    sheet.getRange(
      1,
      1,
      maxHeaderRow,
      width
    ).getValues();

  for (
    let r = 0;
    r < values.length;
    r++
  ) {
    const map =
      hmLegacyMapV264_(
        values[r]
      );

    if (
      map.isLegacy
    ) {
      return {
        header_row:r + 1,
        headers:values[r],
        map:map,
        width:width
      };
    }
  }

  return null;
}


function hmLegacyDateV265_(
  value
) {
  if (
    value instanceof Date &&
    !isNaN(
      value.getTime()
    )
  ) {
    return Utilities.formatDate(
      value,
      getTimezone_(),
      'yyyy-MM-dd'
    );
  }

  let direct =
    toIsoDate_(
      value
    );

  if (
    direct &&
    isValidIsoDate_(
      direct
    )
  ) {
    return direct;
  }

  const s =
    String(value || '')
      .trim();

  /*
   * Handles examples such as:
   * 30-Aug-26 / 1-Sep-26 / 30-Aug-2026
   */
  const m =
    s.match(
      /^(\d{1,2})[-\/ ]([A-Za-z]{3,9})[-\/ ](\d{2}|\d{4})$/
    );

  if (!m) {
    return '';
  }

  const months = {
    JAN:1,
    FEB:2,
    MAR:3,
    APR:4,
    MAY:5,
    JUN:6,
    JUL:7,
    AUG:8,
    SEP:9,
    OCT:10,
    NOV:11,
    DEC:12
  };

  const monKey =
    m[2]
      .substring(0,3)
      .toUpperCase();

  const month =
    months[monKey];

  if (!month) {
    return '';
  }

  let year =
    Number(
      m[3]
    );

  if (
    year < 100
  ) {
    year +=
      year >= 70
        ? 1900
        : 2000;
  }

  const iso =
    [
      String(year)
        .padStart(4,'0'),
      String(month)
        .padStart(2,'0'),
      String(
        Number(m[1])
      )
        .padStart(2,'0')
    ].join('-');

  return isValidIsoDate_(
    iso
  )
    ? iso
    : '';
}


function hmLegacyRowV265_(
  row,
  map,
  lookup,
  sourceName
) {
  const dateKey =
    hmLegacyDateV265_(
      row[map.date]
    );

  const operatorName =
    String(
      row[map.operator] || ''
    )
      .trim();

  const unit =
    String(
      row[map.unit] || ''
    )
      .trim()
      .toUpperCase();

  const hmStart =
    Number(
      row[map.hmStart]
    );

  const hmEnd =
    Number(
      row[map.hmEnd]
    );

  if (
    !dateKey ||
    !unit ||
    !Number.isFinite(
      hmStart
    ) ||
    !Number.isFinite(
      hmEnd
    ) ||
    hmEnd < hmStart
  ) {
    return null;
  }

  const rawNik =
    map.nik >= 0
      ? String(
          row[map.nik] || ''
        )
      : '';

  const person =
    hmResolvePersonV264_(
      rawNik,
      operatorName,
      lookup
    );

  const shift =
    (
      map.shift >= 0
        ? hmShiftV264_(
            row[map.shift]
          )
        : ''
    ) ||
    'DAY';

  let note =
    map.note >= 0
      ? String(
          row[map.note] || ''
        )
      : '';

  if (
    map.shift < 0
  ) {
    note +=
      (
        note
          ? ' | '
          : ''
      ) +
      'Legacy: shift default DAY';
  }

  const totalHM =
    Number(
      (
        hmEnd -
        hmStart
      ).toFixed(2)
    );

  return {
    hm_id:
      generateId_(
        'HM'
      ),

    date:
      parseIsoDate_(
        dateKey
      ),

    shift:
      shift,

    nik:
      person
        ? String(
            person.nik
          )
        : rawNik,

    operator_name:
      person
        ? person.name
        : operatorName,

    unit:
      unit,

    hm_start:
      hmStart,

    hm_end:
      hmEnd,

    total_hm:
      totalHM,

    note:
      note,

    source:
      'LEGACY_MIGRATION:' +
      sourceName,

    created_by_nik:
      'SYSTEM',

    created_by_name:
      'HM MIGRATION',

    created_at:
      new Date(),

    updated_at:''
  };
}


function scanLegacyHMSourcesV264_() {
  const db =
    getDb_();

  return db
    .getSheets()
    .filter(function(sheet) {
      if (
        sheet.getName() ===
        FJB_SHEETS.HM_OPERATION
      ) {
        return false;
      }

      return !!findLegacyHMHeaderV265_(
        sheet
      );
    })
    .map(function(sheet) {
      const found =
        findLegacyHMHeaderV265_(
          sheet
        );

      return {
        name:
          sheet.getName(),

        header_row:
          found
            ? found.header_row
            : 0,

        rows:
          found
            ? Math.max(
                0,
                sheet.getLastRow() -
                found.header_row
              )
            : 0
      };
    });
}


function hmImportLegacySheetV264_(
  sourceSheet
) {
  const found =
    findLegacyHMHeaderV265_(
      sourceSheet
    );

  if (
    !found
  ) {
    return {
      ok:false,
      source:
        sourceSheet
          ? sourceSheet.getName()
          : '',
      imported:0,
      skipped:0
    };
  }

  const startRow =
    found.header_row + 1;

  const count =
    sourceSheet.getLastRow() -
    found.header_row;

  if (
    count <= 0
  ) {
    return {
      ok:true,
      source:
        sourceSheet.getName(),
      imported:0,
      skipped:0
    };
  }

  const values =
    sourceSheet.getRange(
      startRow,
      1,
      count,
      found.width
    ).getValues();

  const lookup =
    hmPersonnelLookupV264_();

  const existing =
    readObjects_(
      FJB_SHEETS.HM_OPERATION
    );

  const keys = {};

  existing.forEach(function(x) {
    const d =
      toIsoDate_(
        x.date
      );

    if (!d) return;

    /*
     * Legacy source often has no shift, so dedupe also uses HM values.
     */
    keys[
      d +
      '|' +
      String(
        x.unit || ''
      )
        .trim()
        .toUpperCase() +
      '|' +
      Number(
        x.hm_start || 0
      ) +
      '|' +
      Number(
        x.hm_end || 0
      )
    ] = true;
  });

  const newRows = [];
  let skipped = 0;

  values.forEach(function(row) {
    const rec =
      hmLegacyRowV265_(
        row,
        found.map,
        lookup,
        sourceSheet.getName()
      );

    if (!rec) {
      /*
       * Completely empty rows are ignored rather than counted as errors.
       */
      if (
        row.some(
          function(v) {
            return String(v || '')
              .trim() !== '';
          }
        )
      ) {
        skipped++;
      }

      return;
    }

    const key =
      toIsoDate_(
        rec.date
      ) +
      '|' +
      rec.unit +
      '|' +
      rec.hm_start +
      '|' +
      rec.hm_end;

    if (
      keys[key]
    ) {
      return;
    }

    keys[key] =
      true;

    newRows.push(
      rec
    );
  });

  if (
    newRows.length
  ) {
    appendObjects_(
      FJB_SHEETS.HM_OPERATION,
      newRows
    );
  }

  return {
    ok:true,
    source:
      sourceSheet.getName(),
    header_row:
      found.header_row,
    imported:
      newRows.length,
    skipped:
      skipped
  };
}


function ensureHMOperationDatabaseV264_(
  force
) {
  const db =
    getDb_();

  let sheet =
    db.getSheetByName(
      FJB_SHEETS.HM_OPERATION
    );

  if (!sheet) {
    ensureSheetSchemaExtended_(
      db,
      FJB_SHEETS.HM_OPERATION,
      FJB_SCHEMA[
        FJB_SHEETS.HM_OPERATION
      ]
    );

    sheet =
      db.getSheetByName(
        FJB_SHEETS.HM_OPERATION
      );
  }

  const width =
    Math.max(
      1,
      sheet.getLastColumn()
    );

  const headers =
    sheet.getLastRow() >= 1
      ? sheet.getRange(
          1,
          1,
          1,
          width
        ).getValues()[0]
      : [];

  const canonical =
    hmCanonicalHeaderV264_(
      headers
    );

  const legacyMap =
    hmLegacyMapV264_(
      headers
    );

  const lookup =
    hmPersonnelLookupV264_();

  let normalized = 0;
  let migratedExternal = 0;
  let migratedFrom = '';

  if (
    legacyMap.isLegacy &&
    !canonical
  ) {
    const values =
      sheet.getLastRow() >= 2
        ? sheet.getRange(
            2,
            1,
            sheet.getLastRow() -
            1,
            width
          ).getValues()
        : [];

    const records =
      values
        .map(function(row) {
          return hmLegacyRowV264_(
            row,
            legacyMap,
            lookup,
            sheet.getName()
          );
        })
        .filter(Boolean);

    hmRewriteCanonicalV264_(
      sheet,
      records
    );

    normalized =
      records.length;
  }
  else if (canonical) {
    const props =
      PropertiesService
        .getScriptProperties();

    const needsVersion =
      props.getProperty(
        'FJB_HM_SCHEMA_VERSION'
      ) !== '2.6.4';

    if (
      force ||
      needsVersion
    ) {
      const rows =
        readObjects_(
          FJB_SHEETS.HM_OPERATION
        );

      const ids = {};
      let changed = false;

      const normalizedRows =
        rows.map(function(obj) {
          const wasShifted =
            hmShiftedLegacyCanonicalV264_(
              obj
            );

          let rec =
            hmNormalizeCanonicalRowV264_(
              obj,
              lookup
            );

          let id =
            String(
              rec.hm_id || ''
            )
            .trim();

          if (
            !id ||
            ids[id]
          ) {
            rec.hm_id =
              generateId_('HM');

            id =
              rec.hm_id;

            changed =
              true;
          }

          ids[id] =
            true;

          if (
            wasShifted
          ) {
            normalized++;
            changed = true;
          }

          return rec;
        });

      if (
        changed ||
        force ||
        needsVersion
      ) {
        hmRewriteCanonicalV264_(
          sheet,
          normalizedRows
        );
      }
    }
  }
  else if (
    sheet.getLastRow() <= 1
  ) {
    hmRewriteCanonicalV264_(
      sheet,
      []
    );
  }

  /*
   * If canonical database is empty:
   * - normal page load auto-imports when there is one clear source;
   * - explicit Sync DB HM (force=true) imports every detected legacy
   *   source and deduplicates rows into the ONE canonical database.
   */
  if (
    sheet.getLastRow() <= 1 ||
    force
  ) {
    const candidates =
      scanLegacyHMSourcesV264_();

    const shouldImport =
      force
        ? candidates
        : (
            candidates.length === 1
              ? candidates
              : []
          );

    const importedNames = [];

    shouldImport.forEach(
      function(candidate) {
        const source =
          db.getSheetByName(
            candidate.name
          );

        if (!source) {
          return;
        }

        const result =
          hmImportLegacySheetV264_(
            source
          );

        if (
          result.imported
        ) {
          migratedExternal +=
            Number(
              result.imported
            );

          importedNames.push(
            result.source
          );
        }
      }
    );

    migratedFrom =
      importedNames.join(
        ', '
      );
  }

  PropertiesService
    .getScriptProperties()
    .setProperty(
      'FJB_HM_SCHEMA_VERSION',
      '2.6.4'
    );

  const candidates =
    scanLegacyHMSourcesV264_();

  return {
    ok:true,

    database_sheet:
      FJB_SHEETS.HM_OPERATION,

    database_rows:
      Math.max(
        0,
        sheet.getLastRow() -
        1
      ),

    normalized:
      normalized,

    migrated_external:
      migratedExternal,

    migrated_from:
      migratedFrom,

    legacy_candidates:
      candidates
  };
}


/**
 * Run manually from Apps Script editor if desired.
 */
function repairFJBHMOperationV264() {
  const result =
    ensureHMOperationDatabaseV264_(
      true
    );

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


/**
 * Read-only diagnostic.
 */
function diagnoseFJBHMOperationV264() {
  const db =
    getDb_();

  const sheet =
    db.getSheetByName(
      FJB_SHEETS.HM_OPERATION
    );

  const headers =
    sheet &&
    sheet.getLastRow() >= 1
      ? sheet.getRange(
          1,
          1,
          1,
          Math.max(
            1,
            sheet.getLastColumn()
          )
        ).getValues()[0]
      : [];

  const rawRows =
    sheet
      ? Math.max(
          0,
          sheet.getLastRow() -
          1
        )
      : 0;

  const parsedRows =
    sheet
      ? readObjects_(
          FJB_SHEETS.HM_OPERATION
        )
      : [];

  const validRows =
    parsedRows.filter(
      function(x) {
        const d =
          toIsoDate_(
            x.date
          );

        return (
          !!d &&
          !!String(
            x.unit || ''
          ).trim() &&
          Number.isFinite(
            Number(
              x.hm_start
            )
          ) &&
          Number.isFinite(
            Number(
              x.hm_end
            )
          )
        );
      }
    );

  const dates =
    validRows
      .map(function(x) {
        return toIsoDate_(
          x.date
        );
      })
      .filter(Boolean)
      .sort();

  const candidates =
    scanLegacyHMSourcesV264_();

  const result = {
    ok:true,

    version:
      FJB_VERSION,

    canonical_database:
      FJB_SHEETS.HM_OPERATION,

    hm_database_definitions:
      1,

    sheet_exists:
      !!sheet,

    canonical_raw_rows:
      rawRows,

    canonical_valid_rows:
      validRows.length,

    canonical_header:
      hmCanonicalHeaderV264_(
        headers
      ),

    legacy_header_on_canonical:
      hmLegacyMapV264_(
        headers
      ).isLegacy,

    earliest_date:
      dates.length
        ? dates[0]
        : '',

    latest_date:
      dates.length
        ? dates[
            dates.length - 1
          ]
        : '',

    legacy_candidate_sheets:
      candidates,

    recommendation:
      validRows.length
        ? 'Canonical HM database contains valid rows.'
        : candidates.length
          ? 'Run repairFJBHMOperationV264 or click Sync DB HM.'
          : 'No HM legacy source was detected in this spreadsheet.',

    explanation:
      'Web App membaca hanya 20_HM_OPERATION. Legacy sheets are migration sources only.'
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


/**
 * Web App ADMIN repair endpoint.
 */
function apiRepairHMDatabaseV264(
  token
) {
  requireSession_(
    token,
    ['ADMIN']
  );

  return ensureHMOperationDatabaseV264_(
    true
  );
}


function ensureOperationalExtensionV260_(force) {
  const props =
    PropertiesService
      .getScriptProperties();

  const targetVersion =
    '2.6.4';

  const db =
    getDb_();

  const result = [];

  [
    FJB_SHEETS.HAULING,
    FJB_SHEETS.FUEL_USAGE
  ].forEach(function(name) {
    const r =
      ensureSheetSchemaExtended_(
        db,
        name,
        FJB_SCHEMA[name]
      );

    result.push({
      name:name,
      created:r.created,
      added:r.added
    });
  });

  const hm =
    ensureHMOperationDatabaseV264_(
      !!force
    );

  result.push({
    name:
      FJB_SHEETS.HM_OPERATION,
    created:false,
    added:[],
    normalized:
      hm.normalized,
    migrated_external:
      hm.migrated_external,
    migrated_from:
      hm.migrated_from
  });

  props.setProperty(
    'FJB_OPERATIONAL_SCHEMA_VERSION',
    targetVersion
  );

  SpreadsheetApp.flush();

  return {
    ok:true,
    changed:true,
    version:
      targetVersion,
    database:
      db.getName(),
    sheets:
      result
  };
}

function upgradeFJBDatabaseV260() {
  const result = ensureOperationalExtensionV260_(true);
  console.log(JSON.stringify(result, null, 2));
  return result;
}



function getSheet_(sheetName) {
  if (FJB_SHEET_INSTANCE_CACHE_[sheetName]) {
    return FJB_SHEET_INSTANCE_CACHE_[sheetName];
  }

  const sheet = getDb_().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      'Sheet ' + sheetName +
      ' tidak ditemukan. Jalankan setupFJBSystem().'
    );
  }

  FJB_SHEET_INSTANCE_CACHE_[sheetName] = sheet;
  return sheet;
}


function getHeaders_(sheetName) {
  const sheet = getSheet_(sheetName);
  const lastCol = sheet.getLastColumn();

  if (!lastCol) return [];

  return sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(x) { return String(x).trim(); });
}


function readObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const schema = FJB_SCHEMA[sheetName];
  const lastCol = schema ? schema.length : sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return [];

  const values = sheet
    .getRange(1, 1, lastRow, lastCol)
    .getValues();

  const headers = values[0].map(function(x) {
    return String(x).trim();
  });

  return values.slice(1)
    .map(function(row, index) {
      const hasValue = row.some(function(v) {
        return v !== '' && v !== null;
      });

      if (!hasValue) return null;

      const obj = { _row:index + 2 };

      headers.forEach(function(h, i) {
        obj[h] = row[i];
      });

      return obj;
    })
    .filter(Boolean);
}


/**
 * Baca hanya N row terakhir untuk tabel transaksi besar.
 * Header tetap dibaca dari row 1.
 */
function readObjectsTail_(sheetName, maxRows) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const schema = FJB_SCHEMA[sheetName];
  const lastCol = schema ? schema.length : sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return [];

  const headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(function(x) {
      return String(x).trim();
    });

  const count = Math.min(
    Math.max(1, Number(maxRows || 500)),
    lastRow - 1
  );

  const startRow = lastRow - count + 1;

  return sheet
    .getRange(startRow, 1, count, lastCol)
    .getValues()
    .map(function(row, index) {
      const hasValue = row.some(function(v) {
        return v !== '' && v !== null;
      });

      if (!hasValue) return null;

      const obj = {
        _row:startRow + index
      };

      headers.forEach(function(h, i) {
        obj[h] = row[i];
      });

      return obj;
    })
    .filter(Boolean);
}


function objectToRow_(sheetName, obj) {
  return FJB_SCHEMA[sheetName].map(function(h) {
    return obj[h] === undefined ? '' : obj[h];
  });
}


function appendObject_(sheetName, obj) {
  const result = appendObjects_(sheetName, [obj]);
  return result ? result.startRow : null;
}


function appendObjects_(sheetName, objects) {
  if (!objects || !objects.length) {
    return {
      startRow:null,
      count:0
    };
  }

  const sheet = getSheet_(sheetName);

  const rows = objects.map(function(obj) {
    return objectToRow_(sheetName, obj);
  });

  const startRow = sheet.getLastRow() + 1;

  sheet.getRange(
    startRow,
    1,
    rows.length,
    rows[0].length
  ).setValues(rows);

  verifyAppendRange_(
    sheetName,
    startRow,
    rows
  );

  bumpFastCacheRevision_(
    sheetName
  );

  return {
    startRow:startRow,
    count:rows.length
  };
}


function updateRowObject_(sheetName, rowNumber, patch) {
  const sheet = getSheet_(sheetName);
  const headers = FJB_SCHEMA[sheetName];

  const current = sheet
    .getRange(
      rowNumber,
      1,
      1,
      headers.length
    )
    .getValues()[0];

  headers.forEach(function(h, i) {
    if (patch[h] !== undefined) {
      current[i] = patch[h];
    }
  });

  sheet.getRange(
    rowNumber,
    1,
    1,
    headers.length
  ).setValues([current]);

  verifyRowPatch_(
    sheetName,
    rowNumber,
    patch
  );

  bumpFastCacheRevision_(
    sheetName
  );

  return rowNumber;
}


function deleteRowByNumber_(sheetName, rowNumber) {
  const sheet = getSheet_(sheetName);
  const before = sheet.getLastRow();

  sheet.deleteRow(rowNumber);

  const after = sheet.getLastRow();

  if (
    before > 1 &&
    after !== before - 1
  ) {
    throw new Error(
      'WRITE_VERIFY_FAILED: Delete pada ' +
      sheetName +
      ' tidak dapat diverifikasi.'
    );
  }

  bumpFastCacheRevision_(
    sheetName
  );

  return true;
}


function findOne_(sheetName, column, value) {
  const target = String(value).trim();

  return readObjects_(sheetName).find(function(x) {
    return String(x[column]).trim() === target;
  }) || null;
}


function seedIfEmpty_(sheetName, rows) {
  const sheet = getSheet_(sheetName);

  if (sheet.getLastRow() <= 1 && rows && rows.length) {
    appendObjects_(sheetName, rows);
  }
}


function getDatabaseStatus_() {
  const ss = getDb_();

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    version: FJB_VERSION,
    sheets: Object.keys(FJB_SCHEMA).map(function(name) {
      const sh = ss.getSheetByName(name);

      return {
        name: name,
        rows: sh ? Math.max(0, sh.getLastRow() - 1) : 0
      };
    })
  };
}



/**
 * Repair NON-DESTRUCTIVE untuk database existing.
 * - Normalisasi month roster/washing menjadi yyyy-MM.
 * - Tidak menghapus transaksi.
 * - Mengisi dummy roster/washing yang hilang hanya untuk database dummy.
 */
function repairRosterWashingDatabase() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDb_();
    const result = repairRosterWashingDatabase_(ss);

    formatAllSheets_(ss);
    SpreadsheetApp.flush();

    safeToast_(
      ss,
      'Repair selesai: Roster ' + result.rosterNormalized +
      ' row, Washing ' + result.washingNormalized + ' row.',
      'FJB SYSTEM',
      8
    );

    return result;

  } finally {
    lock.releaseLock();
  }
}


function repairRosterWashingDatabase_(ss) {
  const result = {
    ok:true,
    rosterNormalized:0,
    washingNormalized:0,
    rosterDummyAdded:0,
    washingDummyAdded:0
  };

  result.rosterNormalized = normalizeMonthColumn_(
    FJB_SHEETS.ROSTER,
    'month',
    'date'
  );

  result.washingNormalized = normalizeMonthColumn_(
    FJB_SHEETS.WASHING,
    'month',
    'plan_date'
  );

  // Repair dummy database only when known dummy users exist.
  const hasDummyOperator = !!findOne_(
    FJB_SHEETS.PERSONNEL,
    'nik',
    '3101021'
  );

  if (hasDummyOperator) {
    result.rosterDummyAdded = ensureDummyRosterCoverage_('2026-08');
    result.washingDummyAdded = ensureDummyWashingCoverage_('2026-08');
  }

  return result;
}


function normalizeMonthColumn_(sheetName, monthField, dateField) {
  const sheet = getSheet_(sheetName);
  const headers = FJB_SCHEMA[sheetName];

  const monthCol = headers.indexOf(monthField) + 1;
  const dateCol = headers.indexOf(dateField) + 1;

  if (monthCol < 1 || dateCol < 1) {
    return 0;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }

  // Jadikan kolom month Plain Text agar Google Sheet tidak auto-convert.
  sheet.getRange(2, monthCol, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setNumberFormat('@');

  const values = sheet.getRange(
    2,
    1,
    lastRow - 1,
    headers.length
  ).getValues();

  const output = [];
  let changed = 0;

  values.forEach(function(row) {
    const rawMonth = row[monthCol - 1];
    const rawDate = row[dateCol - 1];

    let monthKey = normalizeMonthKey_(rawMonth);

    if (!monthKey) {
      monthKey = monthKeyFromValue_(rawDate);
    }

    // Date field is authoritative if available.
    const fromDate = monthKeyFromValue_(rawDate);
    if (fromDate) {
      monthKey = fromDate;
    }

    output.push([monthKey]);

    if (String(rawMonth || '') !== String(monthKey || '')) {
      changed++;
    }
  });

  if (output.length) {
    sheet.getRange(2, monthCol, output.length, 1)
      .setValues(output)
      .setNumberFormat('@');
  }

  return changed;
}


function normalizeMonthKey_(value) {
  if (!value) {
    return '';
  }

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      getTimezone_(),
      'yyyy-MM'
    );
  }

  const s = String(value).trim();

  let m = s.match(/^(\d{4})-(\d{2})$/);
  if (m) {
    return m[1] + '-' + m[2];
  }

  m = s.match(/^(\d{4})-(\d{2})-\d{2}/);
  if (m) {
    return m[1] + '-' + m[2];
  }

  m = s.match(/^(\d{2})\/(\d{4})$/);
  if (m) {
    return m[2] + '-' + m[1];
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(
      d,
      getTimezone_(),
      'yyyy-MM'
    );
  }

  return '';
}


function monthKeyFromValue_(value) {
  if (!value) {
    return '';
  }

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      getTimezone_(),
      'yyyy-MM'
    );
  }

  const iso = toIsoDate_(value);
  return iso ? iso.substring(0, 7) : '';
}


function rowMonthMatches_(row, requestedMonth, dateField) {
  const target = normalizeMonthKey_(requestedMonth);

  if (!target) {
    return true;
  }

  const monthFromDate = monthKeyFromValue_(row[dateField]);
  if (monthFromDate) {
    return monthFromDate === target;
  }

  return normalizeMonthKey_(row.month) === target;
}


function ensureDummyRosterCoverage_(month) {
  const targetMonth = normalizeMonthKey_(month);
  if (!targetMonth) {
    return 0;
  }

  const personnel = readObjects_(FJB_SHEETS.PERSONNEL)
    .filter(function(p) {
      return String(p.status).toUpperCase() === 'ACTIVE' &&
        String(p.category).toUpperCase() !== 'ADMIN';
    });

  const existing = readObjects_(FJB_SHEETS.ROSTER);

  const existingKeys = {};
  existing.forEach(function(r) {
    const dateKey = toIsoDate_(r.date);
    if (!dateKey) return;

    existingKeys[
      String(r.nik) + '|' + dateKey
    ] = true;
  });

  const parts = targetMonth.split('-');
  const year = Number(parts[0]);
  const monthIndex = Number(parts[1]) - 1;
  const days = new Date(year, monthIndex + 1, 0).getDate();

  const driverPattern = [
    'D','D','D','D','OFF','OFF',
    'N','N','N','N','OFF','OFF'
  ];

  const mechanicPattern = [
    'D','D','N','N','OFF','OFF'
  ];

  const now = new Date();
  const rows = [];

  personnel.forEach(function(p, pIndex) {
    for (let d = 1; d <= days; d++) {
      const dayKey = String(d).padStart(2, '0');
      const dateKey = targetMonth + '-' + dayKey;
      const uniqueKey = String(p.nik) + '|' + dateKey;

      if (existingKeys[uniqueKey]) {
        continue;
      }

      let status = 'D';
      const category = String(p.category).toUpperCase();

      if (category.indexOf('GL') >= 0) {
        const dow = new Date(year, monthIndex, d).getDay();
        status = dow === 0 ? 'OFF' : 'D';

      } else if (
        category.indexOf('MECHANIC') >= 0 ||
        category.indexOf('MEKANIK') >= 0
      ) {
        status = mechanicPattern[
          (d - 1 + pIndex) % mechanicPattern.length
        ];

      } else {
        status = driverPattern[
          (d - 1 + pIndex * 2) % driverPattern.length
        ];
      }

      rows.push({
        roster_id:'ROS-' +
          targetMonth.replace('-', '') + '-' +
          p.nik + '-' + dayKey,
        month:targetMonth,
        date:new Date(year, monthIndex, d),
        nik:p.nik,
        name:p.name,
        category:p.category,
        position:p.position,
        assigned_unit:p.assigned_unit,
        roster_status:status,
        source:'DUMMY_REPAIR',
        created_by_nik:'SYSTEM',
        created_by_name:'SYSTEM',
        created_at:now,
        updated_by_nik:'',
        updated_by_name:'',
        updated_at:''
      });

      existingKeys[uniqueKey] = true;
    }
  });

  if (rows.length) {
    appendObjects_(FJB_SHEETS.ROSTER, rows);
  }

  return rows.length;
}


function ensureDummyWashingCoverage_(month) {
  const targetMonth = normalizeMonthKey_(month);
  if (!targetMonth) {
    return 0;
  }

  const existing = readObjects_(FJB_SHEETS.WASHING);

  const unitSet = {};
  existing.forEach(function(r) {
    if (rowMonthMatches_(r, targetMonth, 'plan_date')) {
      unitSet[String(r.unit)] = true;
    }
  });

  const personnel = readObjects_(FJB_SHEETS.PERSONNEL)
    .filter(function(p) {
      const category = String(p.category).toUpperCase();

      return String(p.status).toUpperCase() === 'ACTIVE' &&
        (
          category.indexOf('DRIVER') >= 0 ||
          category.indexOf('OPERATOR') >= 0
        ) &&
        String(p.assigned_unit || '').indexOf('DT-') === 0;
    });

  const parts = targetMonth.split('-');
  const year = Number(parts[0]);
  const monthIndex = Number(parts[1]) - 1;
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();

  const now = new Date();
  const rows = [];

  personnel.forEach(function(p, i) {
    const unit = String(p.assigned_unit);

    if (unitSet[unit]) {
      return;
    }

    // 2 plan washing per unit pada database dummy.
    const d1 = 3 + ((i * 2) % 16);
    const d2 = Math.min(daysInMonth, d1 + 17);

    rows.push({
      washing_id:'WASH-' + Utilities.getUuid(),
      month:targetMonth,
      plan_date:new Date(year, monthIndex, d1),
      actual_date:new Date(year, monthIndex, d1),
      reschedule_date:'',
      unit:unit,
      pic_nik:p.nik,
      pic_name:p.name,
      status:'DONE',
      note:'Dummy repair',
      created_by_nik:'SYSTEM',
      created_by_name:'SYSTEM',
      created_at:now,
      updated_by_nik:'',
      updated_by_name:'',
      updated_at:''
    });

    rows.push({
      washing_id:'WASH-' + Utilities.getUuid(),
      month:targetMonth,
      plan_date:new Date(year, monthIndex, d2),
      actual_date:'',
      reschedule_date:'',
      unit:unit,
      pic_nik:p.nik,
      pic_name:p.name,
      status:'PLAN',
      note:'Dummy repair',
      created_by_nik:'SYSTEM',
      created_by_name:'SYSTEM',
      created_at:now,
      updated_by_nik:'',
      updated_by_name:'',
      updated_at:''
    });

    unitSet[unit] = true;
  });

  if (rows.length) {
    appendObjects_(FJB_SHEETS.WASHING, rows);
  }

  return rows.length;
}


/* ============================================================
 * 03. FORMAT DATABASE
 * ============================================================
 */

function formatAllSheets_(ss) {
  const headerColor = '#10283A';
  const headerText = '#FFFFFF';

  Object.keys(FJB_SCHEMA).forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const headers = FJB_SCHEMA[sheetName];
    const lastRow = Math.max(1, sheet.getLastRow());

    const header = sheet.getRange(1, 1, 1, headers.length);

    header
      .setBackground(headerColor)
      .setFontColor(headerText)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    sheet.setRowHeight(1, 28);
    sheet.setFrozenRows(1);

    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, headers.length)
        .setVerticalAlignment('middle');
    }

    try {
      if (sheet.getFilter()) {
        sheet.getFilter().remove();
      }

      if (lastRow >= 2) {
        sheet.getRange(1, 1, lastRow, headers.length)
          .createFilter();
      }
    } catch (err) {
      // Filter bukan fitur kritikal.
    }

    try {
      sheet.autoResizeColumns(1, headers.length);

      for (let c = 1; c <= headers.length; c++) {
        const width = sheet.getColumnWidth(c);

        if (width < 85) sheet.setColumnWidth(c, 85);
        if (width > 260) sheet.setColumnWidth(c, 260);
      }
    } catch (err) {
      // Auto resize best effort.
    }
  });

  applyDateFormats_();
}


function applyDateFormats_() {
  const configs = [
    [FJB_SHEETS.USERS, ['last_login_at','created_at','updated_at']],
    [FJB_SHEETS.PERSONNEL, ['join_date','created_at','updated_at']],
    [FJB_SHEETS.UNITS, ['updated_at']],
    [FJB_SHEETS.ROSTER, ['date','created_at','updated_at']],
    [FJB_SHEETS.WASHING, [
      'plan_date','actual_date','reschedule_date',
      'created_at','updated_at'
    ]],
    [FJB_SHEETS.HAULING, ['date','created_at','updated_at']],
    [FJB_SHEETS.P2H, ['date','created_at','updated_at']],
    [FJB_SHEETS.MAINTENANCE, ['date','created_at','updated_at']],
    [FJB_SHEETS.UNIT_STATUS_HISTORY, ['timestamp']],
    [FJB_SHEETS.HISTORY, ['timestamp']],
    [FJB_SHEETS.SESSIONS, [
      'created_at','expires_at','last_seen_at'
    ]]
  ];

  configs.forEach(function(item) {
    const sheetName = item[0];
    const cols = item[1];
    const sheet = getSheet_(sheetName);
    const headers = FJB_SCHEMA[sheetName];

    cols.forEach(function(colName) {
      const index = headers.indexOf(colName) + 1;
      if (index < 1) return;

      if (colName === 'date' ||
          colName === 'join_date' ||
          colName.indexOf('_date') >= 0) {
        sheet.getRange(2, index, Math.max(1, sheet.getMaxRows() - 1), 1)
          .setNumberFormat('dd/MM/yyyy');
      } else {
        sheet.getRange(2, index, Math.max(1, sheet.getMaxRows() - 1), 1)
          .setNumberFormat('dd/MM/yyyy HH:mm:ss');
      }
    });
  });

  formatNumericColumns_();
}


function formatNumericColumns_() {
  const numeric = {
    [FJB_SHEETS.UNITS]: [
      'hm_km','availability_pct','achievement_pct',
      'tonase_today','ritase_today'
    ],
    [FJB_SHEETS.HAULING]: [
      'gross','tare','net_ton'
    ],
    [FJB_SHEETS.P2H]: [
      'hm_km','total_items','ok_count','finding_count'
    ],
    [FJB_SHEETS.MAINTENANCE]: [
      'hm_km','duration_min'
    ]
  };

  Object.keys(numeric).forEach(function(sheetName) {
    const sheet = getSheet_(sheetName);
    const headers = FJB_SCHEMA[sheetName];

    numeric[sheetName].forEach(function(col) {
      const index = headers.indexOf(col) + 1;
      if (index < 1) return;

      sheet.getRange(
        2,
        index,
        Math.max(1, sheet.getMaxRows() - 1),
        1
      ).setNumberFormat('#,##0.00');
    });
  });
}


/* ============================================================
 * 04. DUMMY DATA
 * ============================================================
 */

function seedFJBDummyData_(ss) {
  seedConfig_();
  seedPersonnel_();
  seedUsers_();
  seedOwners_();
  seedProducts_();
  seedLocations_();
  seedOptions_();
  seedUnits_();
  seedUnitAliases_();
  seedP2HMaster_();
  seedRoster_();
  seedWashing_();
  seedHauling_();
  seedP2H_();
  seedMaintenance_();
  seedUnitStatusHistory_();
  seedHistory_();
}


function seedConfig_() {
  const rows = [
    {
      key: 'APP_NAME',
      value: 'FJB Operations Control',
      description: 'Nama aplikasi'
    },
    {
      key: 'COMPANY_NAME',
      value: 'PT. FORTUNA JAYA BERSAUDARA',
      description: 'Nama perusahaan'
    },
    {
      key: 'COMPANY_SHORT',
      value: 'FJB',
      description: 'Nama singkat'
    },
    {
      key: 'TAGLINE',
      value: 'Excellent Service',
      description: 'Tagline perusahaan'
    },
    {
      key: 'APP_VERSION',
      value: FJB_VERSION,
      description: 'Versi backend'
    },
    {
      key: 'ROSTER_DEFAULT_MONTH',
      value: '2026-08',
      description: 'Dummy roster default'
    },
    {
      key: 'SESSION_HOURS',
      value: SESSION_TTL_HOURS,
      description: 'Masa aktif login'
    },
    {
      key: 'PASSWORD_MODE',
      value: 'PLAIN_DEMO',
      description: 'Mode password demo'
    }
  ];

  seedIfEmpty_(FJB_SHEETS.CONFIG, rows);
}


function seedPersonnel_() {
  const now = new Date();

  const rows = [
    {
      nik:'9000001', name:'ADMIN FJB',
      category:'ADMIN', position:'ADMIN',
      assigned_unit:'', team:'ADMINISTRATION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'3102001', name:'BUDI SANTOSO',
      category:'GL / PENGAWAS', position:'GL OPERASIONAL',
      assigned_unit:'', team:'OPERATION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },

    {
      nik:'3101021', name:'AHMAD RIZKY',
      category:'DRIVER DT', position:'DRIVER DT',
      assigned_unit:'DT-001', team:'PRODUCTION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'3101022', name:'BUDI SANTOSO DRIVER',
      category:'DRIVER DT', position:'DRIVER DT',
      assigned_unit:'DT-002', team:'PRODUCTION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'3101023', name:'DEDI ARYANTO',
      category:'DRIVER DT', position:'DRIVER DT',
      assigned_unit:'DT-004', team:'PRODUCTION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'3101024', name:'FAJAR HIDAYAT',
      category:'DRIVER DT', position:'DRIVER DT',
      assigned_unit:'DT-005', team:'PRODUCTION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'3101025', name:'HENDRA',
      category:'DRIVER DT', position:'DRIVER DT',
      assigned_unit:'DT-006', team:'PRODUCTION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'3101026', name:'IMAM SYAFII',
      category:'DRIVER DT', position:'DRIVER DT',
      assigned_unit:'DT-007', team:'PRODUCTION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'3101027', name:'JOKO SUSILO',
      category:'DRIVER DT', position:'DRIVER DT',
      assigned_unit:'DT-008', team:'PRODUCTION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'3101028', name:'MULYADI',
      category:'DRIVER DT', position:'DRIVER DT',
      assigned_unit:'DT-009', team:'PRODUCTION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'3101029', name:'RIZAL',
      category:'DRIVER DT', position:'DRIVER DT',
      assigned_unit:'DT-010', team:'PRODUCTION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'3101030', name:'SUPARDI',
      category:'DRIVER DT', position:'DRIVER DT',
      assigned_unit:'DT-014', team:'PRODUCTION',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },

    {
      nik:'4102101', name:'SUMARNO',
      category:'MECHANIC', position:'MECHANIC WORKSHOP',
      assigned_unit:'WORKSHOP', team:'MAINTENANCE',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'4102102', name:'AGUNG SETIAWAN',
      category:'MECHANIC', position:'MECHANIC FIELD',
      assigned_unit:'FIELD', team:'MAINTENANCE',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'4102103', name:'RIDWAN',
      category:'MECHANIC', position:'MECHANIC FIELD',
      assigned_unit:'FIELD', team:'MAINTENANCE',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    },
    {
      nik:'4102104', name:'BAYU',
      category:'MECHANIC', position:'MECHANIC WORKSHOP',
      assigned_unit:'WORKSHOP', team:'MAINTENANCE',
      status:'ACTIVE', phone:'', join_date:new Date(2025,0,1),
      created_at:now, updated_at:now
    }
  ];

  seedIfEmpty_(FJB_SHEETS.PERSONNEL, rows);
}


function seedUsers_() {
  const now = new Date();

  const personnel = readObjects_(FJB_SHEETS.PERSONNEL);

  const rows = personnel.map(function(p) {
    let password = 'op12345';
    let roleOverride = '';

    if (p.category === 'ADMIN') {
      password = 'admin123';
      roleOverride = 'ADMIN';
    } else if (p.category.indexOf('GL') >= 0) {
      password = 'gl12345';
      roleOverride = 'GL';
    } else if (p.category.indexOf('MECHANIC') >= 0) {
      password = 'mech12345';
    }

    return {
      user_id: 'USR-' + p.nik,
      nik: p.nik,
      password: password,
      role_override: roleOverride,
      status: 'ACTIVE',
      must_change_password: false,
      last_login_at: '',
      created_at: now,
      updated_at: now
    };
  });

  seedIfEmpty_(FJB_SHEETS.USERS, rows);
}


function seedOwners_() {
  seedIfEmpty_(FJB_SHEETS.OWNER, [
    {owner_code:'FJB', owner_name:'PT. Fortuna Jaya Bersaudara', active:true},
    {owner_code:'PPA', owner_name:'PT. Putra Perkasa Abadi', active:true},
    {owner_code:'RENTAL', owner_name:'Rental / Vendor', active:true}
  ]);
}


function seedProducts_() {
  seedIfEmpty_(FJB_SHEETS.PRODUCT, [
    {product_code:'SEAM-A', product_name:'Coal Seam A', active:true},
    {product_code:'SEAM-B', product_name:'Coal Seam B', active:true},
    {product_code:'OB', product_name:'Overburden', active:true},
    {product_code:'BASE', product_name:'Base Course', active:true}
  ]);
}


function seedLocations_() {
  seedIfEmpty_(FJB_SHEETS.LOCATION, [
    {location_code:'PIT-A', location_name:'Pit A', active:true},
    {location_code:'PIT-B', location_name:'Pit B', active:true},
    {location_code:'ROM', location_name:'ROM', active:true},
    {location_code:'DISPOSAL', location_name:'Disposal', active:true},
    {location_code:'WORKSHOP', location_name:'Workshop', active:true}
  ]);
}


function seedOptions_() {
  const rows = [];

  function push(group, code, label, sort) {
    rows.push({
      option_group: group,
      option_code: code,
      option_label: label,
      sort_order: sort,
      active: true
    });
  }

  push('SHIFT','DAY','Day Shift',1);
  push('SHIFT','NIGHT','Night Shift',2);

  push('ROSTER_STATUS','D','Day',1);
  push('ROSTER_STATUS','N','Night',2);
  push('ROSTER_STATUS','OFF','Off',3);
  push('ROSTER_STATUS','CT','Cuti',4);
  push('ROSTER_STATUS','SK','Sakit',5);

  push('UNIT_STATUS','READY','Ready',1);
  push('UNIT_STATUS','BD','Breakdown',2);
  push('UNIT_STATUS','SERVICE','Service',3);
  push('UNIT_STATUS','COMMISSIONING','Commissioning',4);

  push('WASHING_STATUS','PLAN','Plan',1);
  push('WASHING_STATUS','DONE','Done',2);
  push('WASHING_STATUS','RESCHEDULE','Reschedule',3);

  push('MAINT_TYPE','BREAKDOWN','Breakdown',1);
  push('MAINT_TYPE','SERVICE','Service',2);
  push('MAINT_TYPE','INSPECTION','Inspection',3);
  push('MAINT_TYPE','COMMISSIONING','Commissioning',4);

  push('MAINT_RESULT','NORMAL','Normal',1);
  push('MAINT_RESULT','MONITORING','Monitoring',2);
  push('MAINT_RESULT','NOT READY','Not Ready',3);

  seedIfEmpty_(FJB_SHEETS.OPTIONS, rows);
}


function seedUnits_() {
  const now = new Date();

  const rows = [
    unitRow_('DT-001','DUMP TRUCK','FJB','READY','RUNNING',
      15240,96,96,1248,85,'3101021','AHMAD RIZKY','PIT-A',now),

    unitRow_('DT-002','DUMP TRUCK','FJB','READY','RUNNING',
      14880,94,92,1110,78,'3101022','BUDI SANTOSO DRIVER','PIT-A',now),

    unitRow_('DT-004','DUMP TRUCK','FJB','READY','RUNNING',
      16020,95,98,1195,82,'3101023','DEDI ARYANTO','PIT-A',now),

    unitRow_('DT-005','DUMP TRUCK','FJB','READY','RUNNING',
      14410,97,103,1325,90,'3101024','FAJAR HIDAYAT','PIT-A',now),

    unitRow_('DT-006','DUMP TRUCK','FJB','READY','RUNNING',
      13980,91,88,980,70,'3101025','HENDRA','PIT-B',now),

    unitRow_('DT-007','DUMP TRUCK','FJB','SERVICE','SERVICE PM',
      15100,78,0,0,0,'3101026','IMAM SYAFII','WORKSHOP',now),

    unitRow_('DT-008','DUMP TRUCK','FJB','READY','RUNNING',
      14550,92,91,1105,77,'3101027','JOKO SUSILO','PIT-B',now),

    unitRow_('DT-009','DUMP TRUCK','FJB','READY','RUNNING',
      14690,95,100,1260,86,'3101028','MULYADI','PIT-B',now),

    unitRow_('DT-010','DUMP TRUCK','FJB','READY','RUNNING',
      15005,90,89,1030,72,'3101029','RIZAL','PIT-B',now),

    unitRow_('DT-014','DUMP TRUCK','FJB','BD',
      'HYDRAULIC LEAKAGE - 6H20M',
      15580,62,0,0,0,'3101030','SUPARDI','WORKSHOP',now),

    unitRow_('EX-001','EXCAVATOR','FJB','READY','LOADING',
      12250,97,99,0,0,'','','PIT-A',now),

    unitRow_('EX-002','EXCAVATOR','FJB','READY','LOADING',
      11820,95,96,0,0,'','','PIT-B',now)
  ];

  seedIfEmpty_(FJB_SHEETS.UNITS, rows);
}


function unitRow_(
  code,type,owner,status,op,hm,avail,ach,ton,rit,
  nik,name,location,updated
) {
  return {
    unit_id:'UNIT-' + code,
    unit_code:code,
    unit_type:type,
    owner:owner,
    status:status,
    operational_status:op,
    hm_km:hm,
    availability_pct:avail,
    achievement_pct:ach,
    tonase_today:ton,
    ritase_today:rit,
    assigned_nik:nik,
    assigned_name:name,
    location:location,
    active:true,
    updated_at:updated
  };
}


function seedUnitAliases_() {
  seedIfEmpty_(FJB_SHEETS.UNIT_ALIAS, [
    {alias_code:'DT001', canonical_unit:'DT-001', active:true, note:'Alias tanpa dash'},
    {alias_code:'DT002', canonical_unit:'DT-002', active:true, note:'Alias tanpa dash'},
    {alias_code:'DT004', canonical_unit:'DT-004', active:true, note:'Alias tanpa dash'},
    {alias_code:'EX001', canonical_unit:'EX-001', active:true, note:'Alias tanpa dash'}
  ]);
}


function seedP2HMaster_() {
  const items = [
    'Ban / Tyre',
    'Lampu',
    'Rem',
    'Klakson',
    'Wiper',
    'Oli Mesin',
    'Air Radiator',
    'Body / Kaca',
    'APAR',
    'Alarm Mundur'
  ];

  const rows = items.map(function(name, i) {
    return {
      item_id:'P2H-' + String(i + 1).padStart(3,'0'),
      item_name:name,
      sort_order:i + 1,
      active:true,
      required:true
    };
  });

  seedIfEmpty_(FJB_SHEETS.P2H_MASTER, rows);
}


function seedRoster_() {
  const sheet = getSheet_(FJB_SHEETS.ROSTER);

  if (sheet.getLastRow() > 1) return;

  const personnel = readObjects_(FJB_SHEETS.PERSONNEL)
    .filter(function(p) {
      return p.category !== 'ADMIN';
    });

  const now = new Date();
  const year = 2026;
  const monthIndex = 7; // Agustus
  const monthKey = '2026-08';
  const days = new Date(year, monthIndex + 1, 0).getDate();

  const driverPattern = [
    'D','D','D','D','OFF','OFF',
    'N','N','N','N','OFF','OFF'
  ];

  const mechanicPattern = [
    'D','D','N','N','OFF','OFF'
  ];

  const rows = [];

  personnel.forEach(function(p, pIndex) {
    for (let d = 1; d <= days; d++) {
      let status = 'D';

      if (p.category.indexOf('GL') >= 0) {
        const day = new Date(year, monthIndex, d).getDay();
        status = day === 0 ? 'OFF' : 'D';

      } else if (p.category.indexOf('MECHANIC') >= 0) {
        status = mechanicPattern[
          (d - 1 + pIndex) % mechanicPattern.length
        ];

      } else {
        status = driverPattern[
          (d - 1 + pIndex * 2) % driverPattern.length
        ];
      }

      // Sedikit dummy cuti / sakit.
      if (p.nik === '3101025' && d === 17) status = 'CT';
      if (p.nik === '3101028' && d === 24) status = 'SK';

      const date = new Date(year, monthIndex, d);
      const dayKey = String(d).padStart(2,'0');

      rows.push({
        roster_id:'ROS-' + monthKey.replace('-','') + '-' + p.nik + '-' + dayKey,
        month:monthKey,
        date:date,
        nik:p.nik,
        name:p.name,
        category:p.category,
        position:p.position,
        assigned_unit:p.assigned_unit,
        roster_status:status,
        source:'DUMMY_SETUP',
        created_by_nik:'SYSTEM',
        created_by_name:'SYSTEM',
        created_at:now,
        updated_by_nik:'',
        updated_by_name:'',
        updated_at:''
      });
    }
  });

  appendObjects_(FJB_SHEETS.ROSTER, rows);
}


function seedWashing_() {
  const sheet = getSheet_(FJB_SHEETS.WASHING);
  if (sheet.getLastRow() > 1) return;

  const now = new Date();

  const configs = [
    ['DT-001','3101021','AHMAD RIZKY',5,'DONE',5,''],
    ['DT-001','3101021','AHMAD RIZKY',22,'PLAN','', ''],

    ['DT-002','3101022','BUDI SANTOSO DRIVER',7,'DONE',7,''],
    ['DT-002','3101022','BUDI SANTOSO DRIVER',24,'PLAN','', ''],

    ['DT-004','3101023','DEDI ARYANTO',8,'DONE',8,''],
    ['DT-004','3101023','DEDI ARYANTO',25,'PLAN','', ''],

    ['DT-005','3101024','FAJAR HIDAYAT',10,'DONE',10,''],
    ['DT-005','3101024','FAJAR HIDAYAT',27,'PLAN','', ''],

    ['DT-006','3101025','HENDRA',12,'RESCHEDULE','',14],
    ['DT-006','3101025','HENDRA',29,'PLAN','', ''],

    ['DT-007','3101026','IMAM SYAFII',14,'DONE',14,''],
    ['DT-008','3101027','JOKO SUSILO',16,'DONE',16,''],
    ['DT-009','3101028','MULYADI',18,'PLAN','', ''],
    ['DT-010','3101029','RIZAL',20,'PLAN','', ''],
    ['DT-014','3101030','SUPARDI',21,'PLAN','', '']
  ];

  const rows = configs.map(function(x, i) {
    const planDay = x[3];
    const status = x[4];
    const actualDay = x[5];
    const rescheduleDay = x[6];

    return {
      washing_id:'WASH-202608-' + String(i + 1).padStart(3,'0'),
      month:'2026-08',
      plan_date:new Date(2026,7,planDay),
      actual_date:actualDay ? new Date(2026,7,actualDay) : '',
      reschedule_date:rescheduleDay ? new Date(2026,7,rescheduleDay) : '',
      unit:x[0],
      pic_nik:x[1],
      pic_name:x[2],
      status:status,
      note:status === 'RESCHEDULE'
        ? 'Reschedule karena unit masih operasi'
        : '',
      created_by_nik:'9000001',
      created_by_name:'ADMIN FJB',
      created_at:now,
      updated_by_nik:'',
      updated_by_name:'',
      updated_at:''
    };
  });

  appendObjects_(FJB_SHEETS.WASHING, rows);
}


function seedHauling_() {
  const sheet = getSheet_(FJB_SHEETS.HAULING);
  if (sheet.getLastRow() > 1) return;

  const now = new Date();

  const dummies = [
    ['2026-08-28','DAY','07:15','DT-001','EX-001',83500,41000,'SEAM-A','','3101021','AHMAD RIZKY'],
    ['2026-08-28','DAY','08:02','DT-002','EX-001',82400,40800,'SEAM-A','','3101022','BUDI SANTOSO DRIVER'],
    ['2026-08-28','DAY','09:10','DT-004','EX-001',84200,41200,'SEAM-A','','3101023','DEDI ARYANTO'],
    ['2026-08-28','NIGHT','20:05','DT-005','EX-002',85100,41400,'SEAM-B','','3101024','FAJAR HIDAYAT'],
    ['2026-08-28','NIGHT','21:20','DT-006','EX-002',81900,40600,'SEAM-B','','3101025','HENDRA'],

    ['2026-08-29','DAY','07:12','DT-001','EX-001',84800,41100,'SEAM-A','','3101021','AHMAD RIZKY'],
    ['2026-08-29','DAY','08:44','DT-002','EX-001',83300,40900,'SEAM-A','','3101022','BUDI SANTOSO DRIVER'],
    ['2026-08-29','DAY','09:32','DT-004','EX-001',85600,41300,'SEAM-A','','3101023','DEDI ARYANTO'],
    ['2026-08-29','DAY','10:17','DT-005','EX-001',86200,41700,'SEAM-A','','3101024','FAJAR HIDAYAT'],
    ['2026-08-29','NIGHT','19:45','DT-009','EX-002',84700,41000,'SEAM-B','','3101028','MULYADI']
  ];

  const rows = dummies.map(function(x, i) {
    const gross = Number(x[5]);
    const tare = Number(x[6]);

    return {
      transaction_id:'HAL-202608-' + String(i + 1).padStart(4,'0'),
      date:parseIsoDate_(x[0]),
      shift:x[1],
      time:x[2],
      hauler:x[3],
      loader:x[4],
      gross:gross,
      tare:tare,
      net_ton:(gross - tare) / 1000,
      product_seam:x[7],
      remark:x[8],
      input_by_nik:x[9],
      input_by_name:x[10],
      created_at:now,
      updated_at:''
    };
  });

  appendObjects_(FJB_SHEETS.HAULING, rows);
}


function seedP2H_() {
  const headerSheet = getSheet_(FJB_SHEETS.P2H);
  if (headerSheet.getLastRow() > 1) return;

  const now = new Date();
  const master = readObjects_(FJB_SHEETS.P2H_MASTER);

  const headers = [
    {
      p2h_id:'P2H-20260829-001',
      date:new Date(2026,7,29),
      shift:'DAY',
      time:'06:55',
      unit:'DT-001',
      hm_km:15240,
      notes:'',
      input_by_nik:'3101021',
      input_by_name:'AHMAD RIZKY'
    },
    {
      p2h_id:'P2H-20260829-002',
      date:new Date(2026,7,29),
      shift:'DAY',
      time:'06:58',
      unit:'DT-002',
      hm_km:14880,
      notes:'Lampu kerja kanan redup',
      input_by_nik:'3101022',
      input_by_name:'BUDI SANTOSO DRIVER'
    },
    {
      p2h_id:'P2H-20260829-003',
      date:new Date(2026,7,29),
      shift:'NIGHT',
      time:'18:45',
      unit:'DT-009',
      hm_km:14690,
      notes:'',
      input_by_nik:'3101028',
      input_by_name:'MULYADI'
    }
  ];

  const p2hRows = [];
  const detailRows = [];

  headers.forEach(function(h, hIndex) {
    let findings = 0;

    master.forEach(function(item, itemIndex) {
      const isFinding =
        hIndex === 1 &&
        item.item_name === 'Lampu';

      if (isFinding) findings++;

      detailRows.push({
        detail_id:h.p2h_id + '-' + item.item_id,
        p2h_id:h.p2h_id,
        item_id:item.item_id,
        item_name:item.item_name,
        checked:!isFinding,
        status:isFinding ? 'TEMUAN' : 'OK',
        note:isFinding ? 'Lampu kerja kanan redup' : ''
      });
    });

    p2hRows.push({
      p2h_id:h.p2h_id,
      date:h.date,
      shift:h.shift,
      time:h.time,
      unit:h.unit,
      hm_km:h.hm_km,
      total_items:master.length,
      ok_count:master.length - findings,
      finding_count:findings,
      notes:h.notes,
      input_by_nik:h.input_by_nik,
      input_by_name:h.input_by_name,
      created_at:now,
      updated_at:''
    });
  });

  appendObjects_(FJB_SHEETS.P2H, p2hRows);
  appendObjects_(FJB_SHEETS.P2H_DETAIL, detailRows);
}


function seedMaintenance_() {
  const now = new Date();

  const rows = [
    {
      maintenance_id:'MNT-20260829-001',
      date:new Date(2026,7,29),
      unit:'DT-014',
      type:'BREAKDOWN',
      hm_km:15580,
      start_time:'08:00',
      finish_time:'14:20',
      duration_min:380,
      result:'MONITORING',
      problem:'Hydraulic leakage',
      action:'Replace hydraulic hose dan cleaning area',
      part_material:'Hydraulic hose 1 pcs',
      mechanic_nik:'4102101',
      mechanic_name:'SUMARNO',
      created_at:now,
      updated_at:''
    },
    {
      maintenance_id:'MNT-20260829-002',
      date:new Date(2026,7,29),
      unit:'DT-007',
      type:'SERVICE',
      hm_km:15100,
      start_time:'09:00',
      finish_time:'11:30',
      duration_min:150,
      result:'NORMAL',
      problem:'Periodic maintenance',
      action:'Service PM dan inspection',
      part_material:'Filter set',
      mechanic_nik:'4102102',
      mechanic_name:'AGUNG SETIAWAN',
      created_at:now,
      updated_at:''
    },
    {
      maintenance_id:'MNT-20260828-001',
      date:new Date(2026,7,28),
      unit:'DT-006',
      type:'INSPECTION',
      hm_km:13980,
      start_time:'15:00',
      finish_time:'15:40',
      duration_min:40,
      result:'NORMAL',
      problem:'Inspection tyre',
      action:'Pressure check dan visual inspection',
      part_material:'',
      mechanic_nik:'4102103',
      mechanic_name:'RIDWAN',
      created_at:now,
      updated_at:''
    }
  ];

  seedIfEmpty_(FJB_SHEETS.MAINTENANCE, rows);
}


function seedUnitStatusHistory_() {
  const now = new Date();

  seedIfEmpty_(FJB_SHEETS.UNIT_STATUS_HISTORY, [
    {
      status_id:'UST-001',
      timestamp:now,
      unit:'DT-014',
      old_status:'READY',
      new_status:'BD',
      old_operational_status:'RUNNING',
      new_operational_status:'HYDRAULIC LEAKAGE - 6H20M',
      reason:'Hydraulic leakage',
      updated_by_nik:'3102001',
      updated_by_name:'BUDI SANTOSO'
    },
    {
      status_id:'UST-002',
      timestamp:now,
      unit:'DT-007',
      old_status:'READY',
      new_status:'SERVICE',
      old_operational_status:'RUNNING',
      new_operational_status:'SERVICE PM',
      reason:'Scheduled PM',
      updated_by_nik:'3102001',
      updated_by_name:'BUDI SANTOSO'
    }
  ]);
}


function seedHistory_() {
  const sheet = getSheet_(FJB_SHEETS.HISTORY);
  if (sheet.getLastRow() > 1) return;

  const now = new Date();

  const rows = [
    historyRow_(
      now,'AHMAD RIZKY','3101021','OPERATOR',
      'Hauling','CREATE','HAL-202608-0006',
      'Input hauling DT-001 43.70 T',
      {
        unit:'DT-001',
        shift:'DAY',
        net_ton:43.7
      }
    ),
    historyRow_(
      now,'AHMAD RIZKY','3101021','OPERATOR',
      'P2H','CREATE','DT-001',
      'P2H DT-001 disimpan',
      {
        unit:'DT-001',
        result:'OK'
      }
    ),
    historyRow_(
      now,'SUMARNO','4102101','MECHANIC',
      'Maintenance','CREATE','DT-014',
      'Breakdown hydraulic leakage',
      {
        result:'MONITORING',
        duration_min:380
      }
    ),
    historyRow_(
      now,'BUDI SANTOSO','3102001','GL',
      'Roster','UPDATE','AHMAD RIZKY',
      'Assign DT-001 dan update roster',
      {
        nik:'3101021',
        unit:'DT-001'
      }
    ),
    historyRow_(
      now,'ADMIN FJB','9000001','ADMIN',
      'Washing','UPDATE','DT-001',
      'Jadwal washing DT-001 diperbarui',
      {
        month:'2026-08'
      }
    )
  ];

  seedIfEmpty_(FJB_SHEETS.HISTORY, rows);
}


function historyRow_(
  timestamp,userName,nik,role,module,activity,entity,summary,payload
) {
  return {
    log_id:'HIST-' + Utilities.getUuid(),
    timestamp:timestamp,
    user_name:userName,
    nik:nik,
    role:role,
    module:module,
    activity:activity,
    entity:entity,
    summary:summary,
    payload_json:JSON.stringify(payload || {})
  };
}


/* ============================================================
 * 05. AUTH & ROLE
 * ============================================================
 */

function resolveRole_(user, personnel) {
  const override = String(user.role_override || '')
    .trim()
    .toUpperCase();

  if (override === 'ADMIN' || override === 'GL') {
    return override;
  }

  if (override === 'OPERATOR' || override === 'MECHANIC') {
    return override;
  }

  const category = String(
    personnel ? personnel.category : ''
  ).trim().toUpperCase();

  if (
    category.indexOf('MECHANIC') >= 0 ||
    category.indexOf('MEKANIK') >= 0
  ) {
    return 'MECHANIC';
  }

  if (
    category.indexOf('DRIVER') >= 0 ||
    category.indexOf('OPERATOR') >= 0
  ) {
    return 'OPERATOR';
  }

  return null;
}



/**
 * Jalankan dari Apps Script editor bila login bermasalah.
 *
 * Contoh:
 * diagnoseFJBLogin('9000001')
 *
 * Tidak menampilkan password.
 */

/**
 * Bisa langsung dipilih dari dropdown Apps Script tanpa parameter.
 * Default mengecek akun ADMIN dummy/utama 9000001.
 */
function diagnoseFJBLoginAdmin() {
  return diagnoseFJBLogin('9000001');
}


/**
 * Diagnosis Fast Prime untuk memastikan session response lengkap.
 * Jalankan setelah login web app bila dibutuhkan:
 * diagnoseFJBFastPrimeByNik('9000001')
 *
 * Fungsi ini tidak membuat session baru; hanya memeriksa mapping role/menu.
 */
function diagnoseFJBFastPrimeByNik(nik) {
  nik = String(nik || '9000001').trim();

  const user = findOne_(
    FJB_SHEETS.USERS,
    'nik',
    nik
  );

  const person = findOne_(
    FJB_SHEETS.PERSONNEL,
    'nik',
    nik
  );

  const role =
    user && person
      ? resolveRole_(user, person)
      : null;

  const result = {
    ok:!!(
      user &&
      person &&
      role &&
      ROLE_MENUS[role]
    ),
    nik:nik,
    user_exists:!!user,
    personnel_exists:!!person,
    role:role || '',
    menus:
      role && ROLE_MENUS[role]
        ? ROLE_MENUS[role]
        : [],
    assigned_unit:
      person
        ? String(person.assigned_unit || '')
        : ''
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


function diagnoseFJBLogin(nik) {
  nik = String(nik || '').trim();

  const result = {
    ok:true,
    version:FJB_VERSION,
    nik:nik,
    database_id:'',
    database_name:'',
    user_exists:false,
    user_status:'',
    personnel_exists:false,
    personnel_category:'',
    personnel_position:'',
    assigned_unit:'',
    resolved_role:'',
    users_rows:0,
    personnel_rows:0,
    sessions_rows:0,
    notes:[]
  };

  try {
    const db = getDb_();

    result.database_id = db.getId();
    result.database_name = db.getName();

    const users = readObjects_(
      FJB_SHEETS.USERS
    );

    const personnel = readObjects_(
      FJB_SHEETS.PERSONNEL
    );

    result.users_rows = users.length;
    result.personnel_rows =
      personnel.length;

    result.sessions_rows =
      Math.max(
        0,
        getSheet_(
          FJB_SHEETS.SESSIONS
        ).getLastRow() - 1
      );

    const user = users.find(
      function(x) {
        return String(x.nik) === nik;
      }
    );

    const person = personnel.find(
      function(x) {
        return String(x.nik) === nik;
      }
    );

    result.user_exists = !!user;
    result.user_status =
      user ? String(user.status) : '';

    result.personnel_exists =
      !!person;

    if (person) {
      result.personnel_category =
        String(person.category || '');

      result.personnel_position =
        String(person.position || '');

      result.assigned_unit =
        String(
          person.assigned_unit || ''
        );
    }

    if (user && person) {
      result.resolved_role =
        resolveRole_(user, person) || '';
    }

    if (!user) {
      result.notes.push(
        'NIK tidak ditemukan di 01_USERS.'
      );
    }

    if (
      user &&
      String(user.status)
        .toUpperCase() !== 'ACTIVE'
    ) {
      result.notes.push(
        'User status bukan ACTIVE.'
      );
    }

    if (!person) {
      result.notes.push(
        'NIK tidak ditemukan di 02_PERSONNEL.'
      );
    }

    if (
      user &&
      person &&
      !result.resolved_role
    ) {
      result.notes.push(
        'Role tidak dapat ditentukan.'
      );
    }

    console.log(
      JSON.stringify(
        result,
        null,
        2
      )
    );

    return result;

  } catch (err) {
    result.ok = false;
    result.error =
      err && err.message
        ? err.message
        : String(err);

    console.log(
      JSON.stringify(
        result,
        null,
        2
      )
    );

    return result;
  }
}


function apiLogin(nik, password) {
  nik = String(nik || '').trim();
  password = String(password || '');

  if (!nik || !password) {
    return {
      ok:false,
      message:
        'NIK / NRP dan password wajib diisi.'
    };
  }

  const user = findOne_(
    FJB_SHEETS.USERS,
    'nik',
    nik
  );

  if (!user) {
    return {
      ok:false,
      message:
        'NIK / NRP tidak terdaftar.'
    };
  }

  if (
    String(user.status)
      .trim()
      .toUpperCase() !== 'ACTIVE'
  ) {
    return {
      ok:false,
      message:'User tidak aktif.'
    };
  }

  if (
    String(user.password) !==
    password
  ) {
    return {
      ok:false,
      message:'Password salah.'
    };
  }

  const personnel = findOne_(
    FJB_SHEETS.PERSONNEL,
    'nik',
    nik
  );

  if (!personnel) {
    return {
      ok:false,
      message:
        'Master Personnel untuk NIK ' +
        nik +
        ' tidak ditemukan.'
    };
  }

  const role =
    resolveRole_(
      user,
      personnel
    );

  if (
    !role ||
    !ROLE_MENUS[role]
  ) {
    return {
      ok:false,
      message:
        'Kategori manpower belum memiliki akses aplikasi.'
    };
  }

  const now = new Date();

  const expires =
    new Date(
      now.getTime() +
      SESSION_TTL_HOURS *
      60 * 60 * 1000
    );

  const token =
    Utilities.getUuid();

  const sessionObject = {
    token:token,
    nik:nik,
    role:role,
    name:personnel.name,
    position:personnel.position,
    assigned_unit:
      personnel.assigned_unit,
    created_at:now,
    expires_at:expires,
    last_seen_at:now,
    active:true
  };

  /*
   * Only session creation is critical for login.
   */
  withWriteLock_(
    'CREATE_LOGIN_SESSION',
    function() {
      appendObject_(
        FJB_SHEETS.SESSIONS,
        sessionObject
      );
    }
  );

  cacheSession_({
    token:token,
    nik:nik,
    role:role,
    name:personnel.name,
    position:
      personnel.position,
    assigned_unit:
      personnel.assigned_unit,
    expires_at:expires,
    active:true
  });

  return {
    ok:true,
    token:token,

    user:{
      nik:nik,
      name:personnel.name,
      role:role,
      category:
        personnel.category,
      position:
        personnel.position,
      assigned_unit:
        personnel.assigned_unit,
      must_change_password:
        toBool_(
          user.must_change_password
        )
    },

    menus:
      ROLE_MENUS[role],

    expires_at:
      expires.toISOString(),

    audit_deferred:true,

    cache_stamp:
      getFastViewStamp_()
  };
}


function apiLogout(token) {
  let session = null;

  try {
    session = requireSession_(token);
  } catch (err) {
    // Session invalid/expired is effectively already logged out.
    try {
      CacheService.getScriptCache()
        .remove(
          'FJB_SESSION_' +
          String(token || '')
        );
    } catch (_) {}

    return {
      ok:true,
      already_logged_out:true
    };
  }

  // Mark session inactive best effort.
  try {
    withWriteLock_(
      'LOGOUT_SESSION',
      function() {
        const rows = readObjects_(
          FJB_SHEETS.SESSIONS
        );

        const found = rows.find(
          function(x) {
            return String(x.token) ===
              String(token);
          }
        );

        if (found) {
          updateRowObject_(
            FJB_SHEETS.SESSIONS,
            found._row,
            {
              active:false,
              last_seen_at:new Date()
            }
          );
        }

        appendHistory_(
          session,
          'Auth',
          'LOGOUT',
          session.nik,
          session.role + ' logout',
          {}
        );
      }
    );

  } catch (err) {
    console.warn(
      'LOGOUT_AUDIT_WARNING: ' +
      (err && err.message
        ? err.message
        : String(err))
    );
  }

  try {
    CacheService.getScriptCache()
      .remove(
        'FJB_SESSION_' + token
      );
  } catch (_) {}

  return { ok:true };
}


function requireSession_(token, allowedRoles) {
  token = String(token || '').trim();

  if (!token) {
    throw new Error('SESSION_REQUIRED');
  }

  let session = getCachedSession_(token);

  if (!session) {
    const row = findOne_(
      FJB_SHEETS.SESSIONS,
      'token',
      token
    );

    if (!row) {
      throw new Error('SESSION_INVALID');
    }

    session = {
      token:row.token,
      nik:String(row.nik),
      role:String(row.role),
      name:String(row.name),
      position:String(row.position),
      assigned_unit:String(row.assigned_unit || ''),
      expires_at:row.expires_at,
      active:toBool_(row.active),
      _row:row._row
    };
  }

  if (!session.active) {
    throw new Error('SESSION_INACTIVE');
  }

  const expires = new Date(session.expires_at);

  if (expires.getTime() < Date.now()) {
    if (session._row) {
      updateRowObject_(FJB_SHEETS.SESSIONS, session._row, {
        active:false,
        last_seen_at:new Date()
      });
    }

    CacheService.getScriptCache()
      .remove('FJB_SESSION_' + token);

    throw new Error('SESSION_EXPIRED');
  }

  if (
    allowedRoles &&
    allowedRoles.length &&
    allowedRoles.indexOf(session.role) < 0
  ) {
    throw new Error('ACCESS_DENIED');
  }

  cacheSession_(session);

  return session;
}


function cacheSession_(session) {
  try {
    CacheService.getScriptCache().put(
      'FJB_SESSION_' + session.token,
      JSON.stringify(session),
      600
    );
  } catch (err) {}
}


function getCachedSession_(token) {
  try {
    const raw = CacheService.getScriptCache()
      .get('FJB_SESSION_' + token);

    return raw ? JSON.parse(raw) : null;
  } catch (err) {
    return null;
  }
}



/* ============================================================
 * 05B. SESSION / PERSONNEL SYNC HELPERS
 * ============================================================
 */

function syncActiveSessionsForNik_(nik, patch) {
  nik = String(nik || '').trim();
  if (!nik) return 0;

  const rows = readObjects_(FJB_SHEETS.SESSIONS)
    .filter(function(x) {
      return String(x.nik) === nik && toBool_(x.active);
    });

  let count = 0;

  rows.forEach(function(x) {
    const update = {
      last_seen_at:new Date()
    };

    ['name','position','assigned_unit','role'].forEach(function(k) {
      if (patch && patch[k] !== undefined) {
        update[k] = patch[k];
      }
    });

    updateRowObject_(FJB_SHEETS.SESSIONS, x._row, update);
    CacheService.getScriptCache().remove('FJB_SESSION_' + x.token);
    count++;
  });

  return count;
}


function invalidateSessionsForNik_(nik, exceptToken) {
  nik = String(nik || '').trim();
  if (!nik) return 0;

  const rows = readObjects_(FJB_SHEETS.SESSIONS)
    .filter(function(x) {
      return String(x.nik) === nik &&
        toBool_(x.active) &&
        (!exceptToken || String(x.token) !== String(exceptToken));
    });

  rows.forEach(function(x) {
    updateRowObject_(FJB_SHEETS.SESSIONS, x._row, {
      active:false,
      last_seen_at:new Date()
    });
    CacheService.getScriptCache().remove('FJB_SESSION_' + x.token);
  });

  return rows.length;
}

/* ============================================================
 * 06. BOOTSTRAP / MASTER
 * ============================================================
 */


/**
 * Stable small bootstrap for web login.
 *
 * Return STRING JSON, bukan object Apps Script besar.
 * Ini menghindari masalah serialisasi payload besar pada google.script.run.
 */
function apiBootstrapJson(token) {
  const data = apiBootstrap(token);

  return JSON.stringify({
    ok:true,
    server_time:data.server_time,
    app:data.app,
    session:data.session,
    personnel_self:data.personnel_self,
    units:data.units,
    products:data.products,
    locations:data.locations,
    options:data.options,
    p2h_master:data.p2h_master,
    db_info:data.db_info || {}
  });
}


/**
 * Lightweight ping setelah login.
 */
function apiSessionJson(token) {
  const session = requireSession_(token);

  return JSON.stringify({
    ok:true,
    session:{
      nik:String(session.nik || ''),
      name:String(session.name || ''),
      role:String(session.role || ''),
      position:String(session.position || ''),
      assigned_unit:String(session.assigned_unit || ''),
      menus:ROLE_MENUS[String(session.role || '')] || []
    }
  });
}


function apiBootstrap(token) {
  const session = requireSession_(token);

  const self = findOne_(
    FJB_SHEETS.PERSONNEL,
    'nik',
    session.nik
  );

  return {
    ok:true,
    server_time:new Date().toISOString(),
    app:{
      name:getConfigValue_('APP_NAME'),
      company:getConfigValue_('COMPANY_NAME'),
      version:FJB_VERSION
    },
    session:{
      nik:session.nik,
      name:session.name,
      role:session.role,
      position:session.position,
      assigned_unit:session.assigned_unit,
      menus:ROLE_MENUS[session.role]
    },
    personnel_self:self ? cleanObject_(self) : null,
    units:readObjects_(FJB_SHEETS.UNITS)
      .filter(function(x) { return toBool_(x.active); })
      .map(cleanObject_),
    products:readObjects_(FJB_SHEETS.PRODUCT)
      .filter(function(x) { return toBool_(x.active); })
      .map(cleanObject_),
    locations:readObjects_(FJB_SHEETS.LOCATION)
      .filter(function(x) { return toBool_(x.active); })
      .map(cleanObject_),
    options:readObjects_(FJB_SHEETS.OPTIONS)
      .filter(function(x) { return toBool_(x.active); })
      .map(cleanObject_),
    p2h_master:readObjects_(FJB_SHEETS.P2H_MASTER)
      .filter(function(x) { return toBool_(x.active); })
      .sort(function(a,b) {
        return Number(a.sort_order) - Number(b.sort_order);
      })
      .map(cleanObject_),

    db_info:{
      roster_rows:Math.max(
        0,
        getSheet_(FJB_SHEETS.ROSTER).getLastRow() - 1
      ),
      washing_rows:Math.max(
        0,
        getSheet_(FJB_SHEETS.WASHING).getLastRow() - 1
      ),
      personnel_rows:Math.max(
        0,
        getSheet_(FJB_SHEETS.PERSONNEL).getLastRow() - 1
      ),
      unit_rows:Math.max(
        0,
        getSheet_(FJB_SHEETS.UNITS).getLastRow() - 1
      )
    }
  };
}


function getConfigValue_(key) {
  const row = findOne_(FJB_SHEETS.CONFIG, 'key', key);
  return row ? row.value : '';
}


/* ============================================================
 * 07. ROSTER API
 * ============================================================
 */

function apiGetRoster(token, month) {
  const session = requireSession_(token);

  month = normalizeMonthKey_(month) ||
    Utilities.formatDate(new Date(), getTimezone_(), 'yyyy-MM');

  const personnelMap = {};
  readObjectsFast_(FJB_SHEETS.PERSONNEL).forEach(function(p) {
    personnelMap[String(p.nik)] = p;
  });

  function getRows() {
    let rows = readObjectsFast_(FJB_SHEETS.ROSTER)
      .filter(function(x) {
        return rowMonthMatches_(x, month, 'date');
      });

    if (session.role === 'OPERATOR' || session.role === 'MECHANIC') {
      rows = rows.filter(function(x) {
        return String(x.nik) === String(session.nik);
      });
    }

    // Personnel master is authoritative for identity/current batangan.
    rows.forEach(function(x) {
      const p = personnelMap[String(x.nik)];
      if (!p) return;
      x.name = p.name;
      x.category = p.category;
      x.position = p.position;
      x.assigned_unit = p.assigned_unit;
    });

    return rows;
  }

  let rows = getRows();

  if (!rows.length && month === '2026-08' &&
      findOne_(FJB_SHEETS.PERSONNEL, 'nik', '3101021')) {
    ensureDummyRosterCoverage_(month);
    rows = getRows();
  }

  rows.sort(function(a, b) {
    const an = String(a.name || '');
    const bn = String(b.name || '');
    if (an !== bn) return an.localeCompare(bn);
    return new Date(a.date) - new Date(b.date);
  });

  return {
    ok:true,
    month:month,
    role:session.role,
    nik:session.nik,
    assigned_unit:session.assigned_unit,
    count:rows.length,
    rows:rows.map(cleanObject_)
  };
}


function apiSaveRosterBatch(token, changes) {
  const session = requireSession_(token, ['ADMIN','GL']);

  if (!Array.isArray(changes) || !changes.length) {
    throw new Error('Tidak ada data roster untuk disimpan.');
  }

  const allowedStatus = ['D','N','OFF','CT','SK'];
  const maxBatch = 10000;

  if (changes.length > maxBatch) {
    throw new Error('Maksimal ' + maxBatch + ' perubahan dalam satu batch.');
  }

  // Deduplicate: last value wins for the same NIK/date.
  const normalizedMap = {};

  changes.forEach(function(ch) {
    const nik = String(ch.nik || '').trim();
    const dateKey = toIsoDate_(ch.date);
    const status = String(ch.roster_status || ch.status || '')
      .trim().toUpperCase();

    if (!nik || !dateKey) {
      throw new Error('NIK dan tanggal roster wajib diisi.');
    }

    if (!isValidIsoDate_(dateKey)) {
      throw new Error(
        'Tanggal roster tidak valid: ' + dateKey
      );
    }

    if (allowedStatus.indexOf(status) < 0) {
      throw new Error('Status roster tidak valid: ' + status);
    }

    normalizedMap[nik + '|' + dateKey] = {
      nik:nik,
      dateKey:dateKey,
      status:status
    };
  });

  const normalized = Object.keys(normalizedMap).map(function(k) {
    return normalizedMap[k];
  });

  const personnelMap = {};
  readObjects_(FJB_SHEETS.PERSONNEL).forEach(function(p) {
    personnelMap[String(p.nik)] = p;
  });

  normalized.forEach(function(ch) {
    if (!personnelMap[ch.nik]) {
      throw new Error('Personnel tidak ditemukan: ' + ch.nik);
    }
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getSheet_(FJB_SHEETS.ROSTER);
    const headers = FJB_SCHEMA[FJB_SHEETS.ROSTER];
    const existing = readObjects_(FJB_SHEETS.ROSTER);
    const existingMap = {};

    existing.forEach(function(x) {
      const d = toIsoDate_(x.date);
      if (d) existingMap[String(x.nik) + '|' + d] = x;
    });

    const now = new Date();
    const newRows = [];
    const historyRows = [];
    let inserted = 0;
    let updated = 0;

    normalized.forEach(function(ch) {
      const person = personnelMap[ch.nik];
      const key = ch.nik + '|' + ch.dateKey;
      const old = existingMap[key];

      if (old) {
        const before = String(old.roster_status || '');

        old.name = person.name;
        old.category = person.category;
        old.position = person.position;
        old.assigned_unit = person.assigned_unit;
        old.roster_status = ch.status;
        old.source = 'WEB_APP';
        old.updated_by_nik = session.nik;
        old.updated_by_name = session.name;
        old.updated_at = now;
        updated++;

        historyRows.push(historyRow_(
          now, session.name, session.nik, session.role,
          'Roster','UPDATE',person.name,
          ch.dateKey + ' • ' + before + ' → ' + ch.status,
          {
            nik:ch.nik,
            date:ch.dateKey,
            before:before,
            after:ch.status,
            assigned_unit:person.assigned_unit
          }
        ));
      } else {
        const month = ch.dateKey.substring(0,7);
        const obj = {
          roster_id:'ROS-' + month.replace('-','') + '-' + ch.nik + '-' + ch.dateKey.substring(8,10),
          month:month,
          date:parseIsoDate_(ch.dateKey),
          nik:ch.nik,
          name:person.name,
          category:person.category,
          position:person.position,
          assigned_unit:person.assigned_unit,
          roster_status:ch.status,
          source:'WEB_APP',
          created_by_nik:session.nik,
          created_by_name:session.name,
          created_at:now,
          updated_by_nik:'',
          updated_by_name:'',
          updated_at:''
        };

        newRows.push(obj);
        existingMap[key] = obj;
        inserted++;

        historyRows.push(historyRow_(
          now, session.name, session.nik, session.role,
          'Roster','CREATE',person.name,
          ch.dateKey + ' • ' + ch.status,
          {
            nik:ch.nik,
            date:ch.dateKey,
            status:ch.status,
            assigned_unit:person.assigned_unit
          }
        ));
      }
    });

    // One write for all existing rows instead of setValue per cell/row.
    if (existing.length) {
      const values = existing.map(function(obj) {
        return headers.map(function(h) {
          return obj[h] === undefined ? '' : obj[h];
        });
      });
      sheet.getRange(2, 1, values.length, headers.length).setValues(values);
    }

    if (newRows.length) appendObjects_(FJB_SHEETS.ROSTER, newRows);
    if (historyRows.length) appendObjects_(FJB_SHEETS.HISTORY, historyRows);

    bumpFastCacheRevision_(
      FJB_SHEETS.ROSTER
    );

    return {
      ok:true,
      inserted:inserted,
      updated:updated,
      total:inserted + updated,
      requested:changes.length,
      deduplicated:normalized.length
    };
  } finally {
    lock.releaseLock();
  }
}



/* ============================================================
 * 07B. FUEL USAGE + HM OPERATION
 * ============================================================ */

function excelSerialFromIsoDate_(dateKey) {
  if (!isValidIsoDate_(dateKey)) return 0;
  const p = String(dateKey).split('-').map(Number);
  return Math.floor((Date.UTC(p[0], p[1]-1, p[2]) - Date.UTC(1899,11,30)) / 86400000);
}

function makeFuelUnitDay_(dateKey, unitCode, shift) {
  return String(excelSerialFromIsoDate_(dateKey)) +
    String(unitCode || '') +
    (String(shift).toUpperCase() === 'NIGHT' ? 'N' : 'D');
}

function apiSaveFuelUsage(token, data) {
  ensureOperationalExtensionV260_(false);
  return withWriteLock_('apiSaveFuelUsage', function() {
    const session = requireSession_(token, ['ADMIN','OPERATOR']);
    data = data || {};
    const dateKey = toIsoDate_(data.date || new Date());
    const shift = String(data.shift || '').trim().toUpperCase();
    const fuelSource = String(data.fuel_source || '').trim().toUpperCase();
    const entityUsed = String(data.entity_used || '').trim();
    const unitCode = String(data.unit_code || '').trim().toUpperCase();
    const hmKm = Number(data.hm_km || 0);
    const totalLiter = Number(data.total_liter || 0);
    const fillTime = String(data.fill_time || formatTime_(new Date())).trim();
    const dedicated = String(data.dedicated || '').trim();
    const location = String(data.location || '').trim();

    if (!isValidIsoDate_(dateKey)) throw new Error('Tanggal Fuel tidak valid.');
    if (['DAY','NIGHT'].indexOf(shift) < 0) throw new Error('Shift Fuel harus DAY atau NIGHT.');
    if (!fuelSource || !unitCode || totalLiter <= 0) {
      throw new Error('Fuel Source, Unit Code dan Total Liter wajib diisi.');
    }
    if (!findOne_(FJB_SHEETS.UNITS, 'unit_code', unitCode)) {
      throw new Error('Unit tidak ditemukan: ' + unitCode);
    }

    const now = new Date();
    const rec = {
      fuel_id:generateId_('FUEL'),
      date:parseIsoDate_(dateKey),
      shift:shift,
      fuel_source:fuelSource,
      entity_used:entityUsed,
      unit_code:unitCode,
      hm_km:hmKm,
      total_liter:totalLiter,
      fill_time:fillTime,
      dedicated:dedicated,
      unit_day:makeFuelUnitDay_(dateKey, unitCode, shift),
      location:location,
      source:'WEB_APP',
      input_by_nik:session.nik,
      input_by_name:session.name,
      created_at:now,
      updated_at:''
    };
    appendObject_(FJB_SHEETS.FUEL_USAGE, rec);
    appendHistory_(session,'Fuel','CREATE',rec.fuel_id,
      unitCode + ' • ' + totalLiter + ' L', cleanObject_(rec));
    return {ok:true, data:cleanObject_(rec)};
  });
}

function apiGetFuelUsage(token, filters) {
  ensureOperationalExtensionV260_(false);
  const session = requireSession_(token, ['ADMIN','GL','OPERATOR']);
  filters = filters || {};
  let rows = readObjectsTailFast_(FJB_SHEETS.FUEL_USAGE, 5000);
  if (session.role === 'OPERATOR') {
    rows = rows.filter(function(x) {
      return String(x.input_by_nik) === String(session.nik);
    });
  }
  rows = filterCommon_(rows, filters, {
    date:'date', shift:'shift', unit:'unit_code', nik:'input_by_nik'
  });
  rows.sort(function(a,b) {
    const ad=toIsoDate_(a.date), bd=toIsoDate_(b.date);
    if (ad !== bd) return bd.localeCompare(ad);
    return String(b.fill_time||'').localeCompare(String(a.fill_time||''));
  });
  return {ok:true, count:rows.length, rows:rows.map(cleanObject_)};
}

function hmShiftOrder_(shift) {
  return String(shift).toUpperCase() === 'NIGHT' ? 2 : 1;
}

function getHMStartInfo_(unitCode, dateKey, shift) {
  const targetKey = dateKey + '|' + hmShiftOrder_(shift);
  const rows = readObjectsTailFast_(FJB_SHEETS.HM_OPERATION, 5000)
    .filter(function(x) { return String(x.unit||'') === String(unitCode||''); });
  let duplicate=null, previous=null, previousKey='';
  rows.forEach(function(x) {
    const d=toIsoDate_(x.date); if (!d) return;
    const k=d+'|'+hmShiftOrder_(x.shift);
    if (d===dateKey && String(x.shift||'').toUpperCase()===String(shift||'').toUpperCase()) duplicate=x;
    if (k<targetKey && (!previous || k>previousKey)) { previous=x; previousKey=k; }
  });
  let hmStart=previous ? Number(previous.hm_end||0) : 0;
  let source=previous ? 'LAST_HM_OPERATION' : 'MASTER_UNIT';
  if (!previous) {
    const unit=findOne_(FJB_SHEETS.UNITS,'unit_code',unitCode);
    hmStart=Number(unit ? unit.hm_km||0 : 0);
  }
  return {
    duplicate:duplicate ? cleanObject_(duplicate) : null,
    hm_start:hmStart,
    source:source,
    previous:previous ? cleanObject_(previous) : null
  };
}

function apiGetHMStart(token, unitCode, dateKey, shift) {
  ensureHMOperationDatabaseV264_(false);
  requireSession_(token, ['ADMIN','GL','OPERATOR']);
  unitCode=String(unitCode||'').trim().toUpperCase();
  dateKey=toIsoDate_(dateKey||new Date());
  shift=String(shift||'DAY').trim().toUpperCase();
  if (!unitCode || !isValidIsoDate_(dateKey)) throw new Error('Unit / tanggal HM tidak valid.');
  const info=getHMStartInfo_(unitCode,dateKey,shift);
  return {ok:true,unit:unitCode,date:dateKey,shift:shift,hm_start:info.hm_start,source:info.source,duplicate:info.duplicate};
}

function apiSaveHMOperation(token, data) {
  ensureHMOperationDatabaseV264_(false);
  return withWriteLock_('apiSaveHMOperation', function() {
    const session=requireSession_(token,['ADMIN','OPERATOR']);
    data=data||{};
    const dateKey=toIsoDate_(data.date||new Date());
    const shift=String(data.shift||'DAY').trim().toUpperCase();
    const unitCode=String(data.unit||session.assigned_unit||'').trim().toUpperCase();
    if (session.role==='OPERATOR' && session.assigned_unit && unitCode!==String(session.assigned_unit)) {
      throw new Error('Unit HM harus sesuai batangan aktif: '+session.assigned_unit);
    }
    if (!isValidIsoDate_(dateKey)||!unitCode) throw new Error('Tanggal dan Unit HM wajib diisi.');
    if (['DAY','NIGHT'].indexOf(shift)<0) throw new Error('Shift HM harus DAY atau NIGHT.');

    const startInfo=getHMStartInfo_(unitCode,dateKey,shift);
    if (startInfo.duplicate) {
      throw new Error('Data HM '+unitCode+' • '+dateKey+' • '+shift+' sudah ada.');
    }
    const hmStart=(data.hm_start!==undefined&&data.hm_start!=='') ? Number(data.hm_start) : Number(startInfo.hm_start||0);
    const hmEnd=Number(data.hm_end);
    if (!Number.isFinite(hmStart)||!Number.isFinite(hmEnd)||hmEnd<hmStart) {
      throw new Error('HM Akhir harus sama atau lebih besar dari HM Awal.');
    }
    const totalHM=Number((hmEnd-hmStart).toFixed(2));
    const operatorNik=session.role==='OPERATOR' ? String(session.nik) : String(data.operator_nik||data.nik||session.nik);
    const person=findOne_(FJB_SHEETS.PERSONNEL,'nik',operatorNik);
    if (!person) throw new Error('Operator HM tidak ditemukan: '+operatorNik);

    const now=new Date();
    const rec={
      hm_id:generateId_('HM'),date:parseIsoDate_(dateKey),shift:shift,
      nik:operatorNik,operator_name:person.name,unit:unitCode,
      hm_start:hmStart,hm_end:hmEnd,total_hm:totalHM,note:String(data.note||''),
      source:'WEB_APP',created_by_nik:session.nik,created_by_name:session.name,
      created_at:now,updated_at:''
    };
    appendObject_(FJB_SHEETS.HM_OPERATION,rec);

    const unit=findOne_(FJB_SHEETS.UNITS,'unit_code',unitCode);
    if (unit && Number(unit.hm_km||0)<=hmEnd) {
      updateRowObject_(FJB_SHEETS.UNITS,unit._row,{hm_km:hmEnd,updated_at:now});
    }
    appendHistory_(session,'HM Operation','CREATE',unitCode,
      dateKey+' • '+shift+' • '+totalHM+' HM',cleanObject_(rec));
    return {ok:true,warning:totalHM>24?'TOTAL_HM_OVER_24':'',data:cleanObject_(rec)};
  });
}

function apiGetHMOperation(token, filters) {
  const repair =
    ensureHMOperationDatabaseV264_(
      false
    );

  const session =
    requireSession_(
      token,
      [
        'ADMIN',
        'GL',
        'OPERATOR'
      ]
    );

  filters =
    filters || {};

  let rows =
    readObjectsTailFast_(
      FJB_SHEETS.HM_OPERATION,
      5000
    );

  if (
    session.role ===
    'OPERATOR'
  ) {
    rows =
      rows.filter(
        function(x) {
          return String(
            x.nik || ''
          ) ===
          String(
            session.nik
          );
        }
      );
  }

  rows =
    filterCommon_(
      rows,
      filters,
      {
        date:'date',
        shift:'shift',
        unit:'unit',
        nik:'nik'
      }
    );

  rows.sort(function(a,b) {
    const ad =
      toIsoDate_(
        a.date
      );

    const bd =
      toIsoDate_(
        b.date
      );

    if (
      ad !== bd
    ) {
      return bd.localeCompare(
        ad
      );
    }

    return (
      hmShiftOrder_(
        b.shift
      ) -
      hmShiftOrder_(
        a.shift
      )
    );
  });

  const sheet =
    getSheet_(
      FJB_SHEETS.HM_OPERATION
    );

  return {
    ok:true,

    database_sheet:
      FJB_SHEETS.HM_OPERATION,

    database_rows:
      Math.max(
        0,
        sheet.getLastRow() -
        1
      ),

    normalized:
      repair.normalized || 0,

    migrated_external:
      repair.migrated_external || 0,

    migrated_from:
      repair.migrated_from || '',

    legacy_candidates:
      repair.legacy_candidates || [],

    count:
      rows.length,

    rows:
      rows.map(
        cleanObject_
      )
  };
}

function apiSaveHMOperationBatch(token, changes) {
  ensureHMOperationDatabaseV264_(false);
  const session=requireSession_(token,['ADMIN']);
  if(!Array.isArray(changes)||!changes.length)throw new Error('Tidak ada data HM untuk diimport.');
  if(changes.length>3000)throw new Error('Maksimal 3000 row HM per import.');

  const people={}; readObjects_(FJB_SHEETS.PERSONNEL).forEach(function(p){people[String(p.nik)]=p;});
  const units={}; readObjects_(FJB_SHEETS.UNITS).forEach(function(u){units[String(u.unit_code).toUpperCase()]=u;});
  const existing=readObjects_(FJB_SHEETS.HM_OPERATION), existingMap={};
  existing.forEach(function(x){const d=toIsoDate_(x.date);if(d)existingMap[d+'|'+String(x.shift).toUpperCase()+'|'+String(x.unit).toUpperCase()]=x;});
  const now=new Date(), newRows=[], historyRows=[]; let inserted=0,updated=0;

  changes.forEach(function(ch){
    const dateKey=toIsoDate_(ch.date),shift=String(ch.shift||'DAY').toUpperCase(),unitCode=String(ch.unit||'').trim().toUpperCase();
    const hmStart=Number(ch.hm_start),hmEnd=Number(ch.hm_end);
    if(!isValidIsoDate_(dateKey)||!unitCode||!Number.isFinite(hmStart)||!Number.isFinite(hmEnd)||hmEnd<hmStart)throw new Error('Data import HM tidak valid pada '+unitCode+' / '+dateKey);
    if(!units[unitCode])throw new Error('Unit HM tidak ditemukan: '+unitCode);
    let nik=String(ch.nik||ch.operator_nik||''),person=nik?people[nik]:null;
    if(!person&&ch.operator_name){
      const target=String(ch.operator_name).trim().toUpperCase();
      Object.keys(people).some(function(k){if(String(people[k].name).trim().toUpperCase()===target){nik=k;person=people[k];return true;}return false;});
    }
    if(!person)throw new Error('Operator HM tidak ditemukan: '+(ch.operator_name||nik||'-'));
    const totalHM=Number((hmEnd-hmStart).toFixed(2));
    const key=dateKey+'|'+shift+'|'+unitCode,old=existingMap[key];
    if(old){
      old.nik=nik;old.operator_name=person.name;old.hm_start=hmStart;old.hm_end=hmEnd;old.total_hm=totalHM;old.note=String(ch.note||'');old.source='IMPORT';old.updated_at=now;updated++;
    }else{
      const rec={hm_id:generateId_('HM'),date:parseIsoDate_(dateKey),shift:shift,nik:nik,operator_name:person.name,unit:unitCode,hm_start:hmStart,hm_end:hmEnd,total_hm:totalHM,note:String(ch.note||''),source:'IMPORT',created_by_nik:session.nik,created_by_name:session.name,created_at:now,updated_at:''};
      newRows.push(rec);existingMap[key]=rec;inserted++;
    }
    historyRows.push(historyRow_(now,session.name,session.nik,session.role,'HM Operation',old?'UPDATE':'CREATE',unitCode,dateKey+' • '+shift+' • '+totalHM+' HM',{date:dateKey,shift:shift,unit:unitCode,nik:nik,operator_name:person.name,hm_start:hmStart,hm_end:hmEnd,total_hm:totalHM,source:'IMPORT'}));
  });

  const sheet=getSheet_(FJB_SHEETS.HM_OPERATION),headers=FJB_SCHEMA[FJB_SHEETS.HM_OPERATION];
  if(existing.length&&updated){
    const values=existing.map(function(obj){return headers.map(function(h){return obj[h]===undefined?'':obj[h];});});
    sheet.getRange(2,1,values.length,headers.length).setValues(values);
    bumpFastCacheRevision_(FJB_SHEETS.HM_OPERATION);
  }
  if(newRows.length)appendObjects_(FJB_SHEETS.HM_OPERATION,newRows);
  if(historyRows.length)appendObjects_(FJB_SHEETS.HISTORY,historyRows);
  return {ok:true,inserted:inserted,updated:updated,total:inserted+updated};
}


/* ============================================================
 * 07C. INPUT HISTORY ASSIST — V2.6.1
 * ============================================================ */

function sortRecentInputHistory_(rows, dateField, timeField) {
  return (rows || []).sort(function(a,b) {
    const ad = toIsoDate_(a[dateField]);
    const bd = toIsoDate_(b[dateField]);

    if (ad !== bd) {
      return bd.localeCompare(ad);
    }

    return String(b[timeField] || '')
      .localeCompare(String(a[timeField] || ''));
  });
}


function apiGetInputHistoryJson(token) {
  ensureOperationalExtensionV260_(false);

  const session = requireSession_(token, [
    'ADMIN','GL','OPERATOR','MECHANIC'
  ]);

  let hauling = readObjectsTailFast_(
    FJB_SHEETS.HAULING,
    350
  );

  let fuel = readObjectsTailFast_(
    FJB_SHEETS.FUEL_USAGE,
    350
  );

  let hm = readObjectsTailFast_(
    FJB_SHEETS.HM_OPERATION,
    350
  );

  let maintenance = readObjectsTailFast_(
    FJB_SHEETS.MAINTENANCE,
    250
  );

  /*
   * Operator uses unit history, not only their own previous input.
   * This keeps field patterns synchronized when the driver changes.
   */
  if (session.role === 'OPERATOR') {
    const assigned = String(session.assigned_unit || '');

    if (assigned) {
      hauling = hauling.filter(function(x) {
        return String(x.hauler || '') === assigned;
      });

      fuel = fuel.filter(function(x) {
        return String(x.unit_code || '') === assigned;
      });

      hm = hm.filter(function(x) {
        return String(x.unit || '') === assigned;
      });
    }
  }

  if (session.role === 'MECHANIC') {
    maintenance = maintenance.filter(function(x) {
      return String(x.mechanic_nik || '') === String(session.nik);
    });
  }

  hauling = sortRecentInputHistory_(hauling,'date','time').slice(0,80);
  fuel = sortRecentInputHistory_(fuel,'date','fill_time').slice(0,80);
  hm = sortRecentInputHistory_(hm,'date','created_at').slice(0,80);
  maintenance = sortRecentInputHistory_(maintenance,'date','start_time').slice(0,50);

  return JSON.stringify({
    ok:true,
    cache_stamp:getFastViewStamp_(),
    hauling:hauling.map(cleanObject_),
    fuel:fuel.map(cleanObject_),
    hm:hm.map(cleanObject_),
    maintenance:maintenance.map(cleanObject_)
  });
}


/* ============================================================
 * 08. WASHING API
 * ============================================================
 */

function apiGetWashing(token, month) {
  const session = requireSession_(
    token,
    ['ADMIN','GL','OPERATOR']
  );

  month = normalizeMonthKey_(month) ||
    Utilities.formatDate(
      new Date(),
      getTimezone_(),
      'yyyy-MM'
    );

  function getRows() {
    let rows = readObjectsFast_(
      FJB_SHEETS.WASHING
    ).filter(function(x) {
      return rowMonthMatches_(
        x,
        month,
        'plan_date'
      );
    });

    if (session.role === 'OPERATOR') {
      rows = rows.filter(function(x) {
        return String(x.unit) ===
          String(
            session.assigned_unit || ''
          );
      });
    }

    return rows;
  }

  let rows = getRows();

  // Self-heal dummy washing schedule.
  if (
    !rows.length &&
    month === '2026-08' &&
    findOne_(
      FJB_SHEETS.PERSONNEL,
      'nik',
      '3101021'
    )
  ) {
    ensureDummyWashingCoverage_(month);
    rows = getRows();
  }

  rows.sort(function(a, b) {
    const ua = String(a.unit || '');
    const ub = String(b.unit || '');

    if (ua !== ub) {
      return ua.localeCompare(ub);
    }

    return new Date(a.plan_date) -
      new Date(b.plan_date);
  });

  return {
    ok:true,
    month:month,
    role:session.role,
    nik:session.nik,
    assigned_unit:session.assigned_unit,
    count:rows.length,
    rows:rows.map(cleanObject_)
  };
}


function apiSaveWashing(token, data) {
  return withWriteLock_('apiSaveWashing', function() {
      const session = requireSession_(token, ['ADMIN']);

      data = data || {};

      const unit = String(data.unit || '').trim();
      const planDate = toIsoDate_(data.plan_date);
      const status = String(data.status || 'PLAN')
        .trim()
        .toUpperCase();

      if (!unit || !planDate) {
        throw new Error('Unit dan Plan Date wajib diisi.');
      }

      if (
        ['PLAN','DONE','RESCHEDULE'].indexOf(status) < 0
      ) {
        throw new Error('Status washing tidak valid.');
      }

      const unitRow = findOne_(
        FJB_SHEETS.UNITS,
        'unit_code',
        unit
      );

      if (!unitRow) {
        throw new Error('Unit tidak ditemukan: ' + unit);
      }

      const personnel = unitRow.assigned_nik
        ? findOne_(
            FJB_SHEETS.PERSONNEL,
            'nik',
            String(unitRow.assigned_nik)
          )
        : null;

      const now = new Date();

      const rec = {
        washing_id:'WASH-' + Utilities.getUuid(),
        month:planDate.substring(0,7),
        plan_date:parseIsoDate_(planDate),
        actual_date:(data.actual_date || status === 'DONE')
          ? parseIsoDate_(toIsoDate_(data.actual_date || planDate))
          : '',
        reschedule_date:(data.reschedule_date || status === 'RESCHEDULE')
          ? parseIsoDate_(toIsoDate_(data.reschedule_date || planDate))
          : '',
        unit:unit,
        pic_nik:personnel ? personnel.nik : '',
        pic_name:personnel ? personnel.name : '',
        status:status,
        note:String(data.note || ''),
        created_by_nik:session.nik,
        created_by_name:session.name,
        created_at:now,
        updated_by_nik:'',
        updated_by_name:'',
        updated_at:''
      };

      appendObject_(FJB_SHEETS.WASHING, rec);

      appendHistory_(
        session,
        'Washing',
        'CREATE',
        unit,
        planDate + ' • ' + status,
        cleanObject_(rec)
      );

      return {
        ok:true,
        data:cleanObject_(rec)
      };

  });
}


/**
 * Batch import Washing.
 *
 * UPSERT key = UNIT + PLAN_DATE
 * - existing -> UPDATE
 * - missing  -> INSERT
 * - rows/months not supplied -> untouched
 *
 * Supports import from matrix and long format on the client.
 */
function apiSaveWashingBatch(token, changes) {
  const session = requireSession_(
    token,
    ['ADMIN']
  );

  if (
    !Array.isArray(changes) ||
    !changes.length
  ) {
    throw new Error(
      'Tidak ada data Washing untuk disimpan.'
    );
  }

  const maxBatch = 3000;

  if (changes.length > maxBatch) {
    throw new Error(
      'Maksimal ' +
      maxBatch +
      ' perubahan Washing per upload.'
    );
  }

  const allowedStatus = [
    'PLAN',
    'DONE',
    'RESCHEDULE'
  ];

  /*
   * Last value wins for duplicate UNIT + PLAN_DATE.
   */
  const normalizedMap = {};

  changes.forEach(function(ch) {
    const unit =
      String(ch.unit || '')
        .trim()
        .toUpperCase();

    const planDate =
      toIsoDate_(ch.plan_date);

    const status =
      String(ch.status || 'PLAN')
        .trim()
        .toUpperCase();

    if (
      !unit ||
      !planDate
    ) {
      throw new Error(
        'Unit dan Plan Date Washing wajib diisi.'
      );
    }

    if (
      !isValidIsoDate_(planDate)
    ) {
      throw new Error(
        'Tanggal Washing tidak valid: ' +
        planDate
      );
    }

    if (
      allowedStatus.indexOf(status) < 0
    ) {
      throw new Error(
        'Status Washing tidak valid: ' +
        status
      );
    }

    let actualDate = '';
    let rescheduleDate = '';

    if (status === 'DONE') {
      actualDate =
        ch.actual_date
          ? toIsoDate_(ch.actual_date)
          : planDate;

      if (
        !isValidIsoDate_(actualDate)
      ) {
        throw new Error(
          'Actual Date Washing tidak valid: ' +
          actualDate
        );
      }
    }

    if (status === 'RESCHEDULE') {
      /*
       * Matrix R defaults to same displayed date,
       * so calendar immediately displays R.
       */
      rescheduleDate =
        ch.reschedule_date
          ? toIsoDate_(
              ch.reschedule_date
            )
          : planDate;

      if (
        !isValidIsoDate_(
          rescheduleDate
        )
      ) {
        throw new Error(
          'Reschedule Date Washing tidak valid: ' +
          rescheduleDate
        );
      }
    }

    normalizedMap[
      unit + '|' + planDate
    ] = {
      unit:unit,
      planDate:planDate,
      status:status,
      actualDate:actualDate,
      rescheduleDate:
        rescheduleDate,
      note:String(
        ch.note || ''
      )
    };
  });

  const normalized =
    Object.keys(normalizedMap)
      .map(function(k) {
        return normalizedMap[k];
      });

  /*
   * Read masters ONCE.
   */
  const unitMap = {};

  readObjects_(
    FJB_SHEETS.UNITS
  ).forEach(function(u) {
    unitMap[
      String(u.unit_code)
        .trim()
        .toUpperCase()
    ] = u;
  });

  const personnelMap = {};

  readObjects_(
    FJB_SHEETS.PERSONNEL
  ).forEach(function(p) {
    personnelMap[
      String(p.nik)
    ] = p;
  });

  normalized.forEach(function(ch) {
    if (!unitMap[ch.unit]) {
      throw new Error(
        'Unit tidak ditemukan: ' +
        ch.unit
      );
    }
  });

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const sheet =
      getSheet_(
        FJB_SHEETS.WASHING
      );

    const headers =
      FJB_SCHEMA[
        FJB_SHEETS.WASHING
      ];

    const existing =
      readObjects_(
        FJB_SHEETS.WASHING
      );

    const existingMap = {};

    existing.forEach(function(x) {
      const d =
        toIsoDate_(
          x.plan_date
        );

      if (!d) return;

      existingMap[
        String(x.unit)
          .trim()
          .toUpperCase() +
        '|' + d
      ] = x;
    });

    const now = new Date();

    const newRows = [];
    const savedRows = [];
    const historyRows = [];

    let inserted = 0;
    let updated = 0;

    normalized.forEach(function(ch) {
      const unitRow =
        unitMap[ch.unit];

      const pic =
        unitRow &&
        unitRow.assigned_nik
          ? personnelMap[
              String(
                unitRow.assigned_nik
              )
            ]
          : null;

      const key =
        ch.unit +
        '|' +
        ch.planDate;

      const old =
        existingMap[key];

      if (old) {
        const before =
          String(
            old.status || ''
          );

        old.month =
          ch.planDate.substring(0,7);

        old.plan_date =
          parseIsoDate_(
            ch.planDate
          );

        old.actual_date =
          ch.actualDate
            ? parseIsoDate_(
                ch.actualDate
              )
            : '';

        old.reschedule_date =
          ch.rescheduleDate
            ? parseIsoDate_(
                ch.rescheduleDate
              )
            : '';

        old.unit =
          ch.unit;

        old.pic_nik =
          pic
            ? pic.nik
            : '';

        old.pic_name =
          pic
            ? pic.name
            : '';

        old.status =
          ch.status;

        old.note =
          ch.note;

        old.updated_by_nik =
          session.nik;

        old.updated_by_name =
          session.name;

        old.updated_at =
          now;

        updated++;

        savedRows.push(
          old
        );

        historyRows.push(
          historyRow_(
            now,
            session.name,
            session.nik,
            session.role,
            'Washing',
            'UPDATE',
            ch.unit,
            ch.planDate +
              ' • ' +
              before +
              ' → ' +
              ch.status,
            {
              unit:ch.unit,
              plan_date:
                ch.planDate,
              before:before,
              after:
                ch.status,
              actual_date:
                ch.actualDate,
              reschedule_date:
                ch.rescheduleDate,
              note:ch.note
            }
          )
        );
      }
      else {
        const rec = {
          washing_id:
            'WASH-' +
            Utilities.getUuid(),

          month:
            ch.planDate.substring(
              0,
              7
            ),

          plan_date:
            parseIsoDate_(
              ch.planDate
            ),

          actual_date:
            ch.actualDate
              ? parseIsoDate_(
                  ch.actualDate
                )
              : '',

          reschedule_date:
            ch.rescheduleDate
              ? parseIsoDate_(
                  ch.rescheduleDate
                )
              : '',

          unit:
            ch.unit,

          pic_nik:
            pic
              ? pic.nik
              : '',

          pic_name:
            pic
              ? pic.name
              : '',

          status:
            ch.status,

          note:
            ch.note,

          created_by_nik:
            session.nik,

          created_by_name:
            session.name,

          created_at:
            now,

          updated_by_nik:'',
          updated_by_name:'',
          updated_at:''
        };

        newRows.push(rec);

        existingMap[key] =
          rec;

        savedRows.push(
          rec
        );

        inserted++;

        historyRows.push(
          historyRow_(
            now,
            session.name,
            session.nik,
            session.role,
            'Washing',
            'CREATE',
            ch.unit,
            ch.planDate +
              ' • ' +
              ch.status,
            {
              unit:ch.unit,
              plan_date:
                ch.planDate,
              status:
                ch.status,
              actual_date:
                ch.actualDate,
              reschedule_date:
                ch.rescheduleDate,
              note:ch.note
            }
          )
        );
      }
    });

    /*
     * One write for all existing records.
     * This is significantly faster than setValue per row.
     */
    if (
      existing.length &&
      updated
    ) {
      const values =
        existing.map(function(obj) {
          return headers.map(
            function(h) {
              return obj[h] ===
                undefined
                  ? ''
                  : obj[h];
            }
          );
        });

      sheet.getRange(
        2,
        1,
        values.length,
        headers.length
      ).setValues(values);
    }

    if (newRows.length) {
      appendObjects_(
        FJB_SHEETS.WASHING,
        newRows
      );
    }

    if (historyRows.length) {
      appendObjects_(
        FJB_SHEETS.HISTORY,
        historyRows
      );
    }

    bumpFastCacheRevision_(
      FJB_SHEETS.WASHING
    );

    return {
      ok:true,
      inserted:inserted,
      updated:updated,
      total:
        inserted + updated,
      requested:
        changes.length,
      deduplicated:
        normalized.length,
      rows:
        savedRows.map(
          cleanObject_
        )
    };

  } finally {
    lock.releaseLock();
  }
}



/* ============================================================
 * 09. HAULING API
 * ============================================================
 */

function apiSaveHauling(token, data) {
  ensureOperationalExtensionV260_(false);

  return withWriteLock_('apiSaveHauling', function() {
    const session = requireSession_(token, ['ADMIN','OPERATOR']);
    data = data || {};

    const dateKey = toIsoDate_(data.date || new Date());
    const shift = String(data.shift || '').toUpperCase();
    const hauler = String(data.hauler || data.unit || '').trim();
    const loader = String(data.loader || '').trim();
    const gross = Number(data.gross || 0);
    const tare = Number(data.tare || 0);
    const distance = Number(data.distance || 0);
    const coalProduct = String(
      data.coal_product || data.product_seam || data.product || ''
    ).trim();
    const jamRitase = String(data.jam_ritase || '').trim();
    const ritase = Math.max(1, Math.round(Number(data.ritase || 1)));

    if (!dateKey || !shift || !hauler || !loader) {
      throw new Error('Tanggal, Shift, Hauler dan Loader wajib diisi.');
    }
    if (['DAY','NIGHT'].indexOf(shift) < 0) {
      throw new Error('Shift harus DAY atau NIGHT.');
    }
    if (gross <= 0 || tare < 0 || gross <= tare) {
      throw new Error('Gross / Tare tidak valid.');
    }
    if (distance < 0) {
      throw new Error('Distance tidak boleh negatif.');
    }
    if (session.role === 'OPERATOR' && session.assigned_unit &&
        hauler !== String(session.assigned_unit)) {
      throw new Error('Hauler harus sesuai batangan aktif Anda: ' + session.assigned_unit);
    }
    if (!findOne_(FJB_SHEETS.UNITS, 'unit_code', hauler)) {
      throw new Error('Hauler tidak ditemukan pada Master Unit: ' + hauler);
    }
    if (!findOne_(FJB_SHEETS.UNITS, 'unit_code', loader)) {
      throw new Error('Loader tidak ditemukan pada Master Unit: ' + loader);
    }

    const now = new Date();
    const rec = {
      transaction_id:generateId_('HAL'),
      date:parseIsoDate_(dateKey),
      shift:shift,
      time:String(data.time || formatTime_(now)),
      hauler:hauler,
      loader:loader,
      gross:gross,
      tare:tare,
      net_ton:(gross - tare) / 1000,
      product_seam:coalProduct,
      remark:String(data.remark || ''),
      input_by_nik:session.nik,
      input_by_name:session.name,
      created_at:now,
      updated_at:'',
      distance:distance,
      coal_product:coalProduct,
      jam_ritase:jamRitase,
      ritase:ritase
    };

    appendObject_(FJB_SHEETS.HAULING, rec);
    appendHistory_(
      session,'Hauling','CREATE',rec.transaction_id,
      hauler + ' • ' + rec.net_ton.toFixed(2) + ' T • ' + ritase + ' rit',
      cleanObject_(rec)
    );
    return {ok:true, data:cleanObject_(rec)};
  });
}


function apiGetHauling(token, filters) {
  const session = requireSession_(token, [
    'ADMIN','GL','OPERATOR'
  ]);

  filters = filters || {};

  let allRows = readObjectsTailFast_(
    FJB_SHEETS.HAULING,
    5000
  );

  if (session.role === 'OPERATOR') {
    allRows = allRows.filter(function(x) {
      return String(x.input_by_nik) ===
        String(session.nik);
    });
  }

  let rows = filterCommon_(
    allRows,
    filters,
    {
      date:'date',
      shift:'shift',
      unit:'hauler',
      nik:'input_by_nik'
    }
  );

  const requestedFrom =
    filters.from
      ? toIsoDate_(filters.from)
      : '';

  const requestedTo =
    filters.to
      ? toIsoDate_(filters.to)
      : '';

  let effectiveDate = '';
  let fallbackUsed = false;

  if (
    rows.length &&
    requestedFrom &&
    requestedFrom === requestedTo
  ) {
    effectiveDate = requestedFrom;
  }

  /*
   * Operational page fallback:
   * if exact requested day has no hauling yet,
   * use latest available hauling date <= requested date.
   *
   * Reports do NOT use this unless latest_if_empty=true.
   */
  if (
    !rows.length &&
    toBool_(filters.latest_if_empty) &&
    requestedFrom &&
    requestedFrom === requestedTo
  ) {
    const candidates = filterCommon_(
      allRows,
      {
        to:requestedTo,
        shift:filters.shift || 'ALL',
        unit:filters.unit || 'ALL',
        nik:filters.nik || 'ALL'
      },
      {
        date:'date',
        shift:'shift',
        unit:'hauler',
        nik:'input_by_nik'
      }
    );

    const dates = candidates
      .map(function(x) {
        return toIsoDate_(x.date);
      })
      .filter(Boolean)
      .sort();

    if (dates.length) {
      effectiveDate = dates[dates.length - 1];

      rows = candidates.filter(function(x) {
        return toIsoDate_(x.date) ===
          effectiveDate;
      });

      fallbackUsed =
        effectiveDate !== requestedFrom;
    }
  }

  if (!effectiveDate && rows.length) {
    const dates = rows
      .map(function(x) {
        return toIsoDate_(x.date);
      })
      .filter(Boolean)
      .sort();

    effectiveDate =
      dates.length
        ? dates[dates.length - 1]
        : '';
  }

  rows.sort(function(a,b) {
    const ad = toIsoDate_(a.date);
    const bd = toIsoDate_(b.date);

    if (ad !== bd) {
      return bd.localeCompare(ad);
    }

    return String(a.time || '')
      .localeCompare(
        String(b.time || '')
      );
  });

  return {
    ok:true,
    requested_date:
      requestedFrom === requestedTo
        ? requestedFrom
        : '',
    effective_date:effectiveDate,
    fallback_used:fallbackUsed,
    count:rows.length,
    rows:rows.map(cleanObject_)
  };
}



/**
 * Jalankan langsung dari Apps Script editor.
 * Menampilkan tanggal hauling terakhir dan total tonasenya.
 */
function diagnoseFJBHaulingData() {
  const rows = readObjectsTail_(
    FJB_SHEETS.HAULING,
    5000
  );

  const dateMap = {};

  rows.forEach(function(x) {
    const d = toIsoDate_(x.date);

    if (!d) return;

    if (!dateMap[d]) {
      dateMap[d] = {
        ritase:0,
        tonase:0
      };
    }

    dateMap[d].ritase += 1;
    dateMap[d].tonase +=
      Number(x.net_ton || 0);
  });

  const dates =
    Object.keys(dateMap).sort();

  const latest =
    dates.length
      ? dates[dates.length - 1]
      : '';

  const result = {
    version:FJB_VERSION,
    total_rows:rows.length,
    latest_date:latest,
    latest_ritase:
      latest
        ? dateMap[latest].ritase
        : 0,
    latest_tonase:
      latest
        ? dateMap[latest].tonase
        : 0,
    last_7_dates:
      dates.slice(-7).map(function(d) {
        return {
          date:d,
          ritase:dateMap[d].ritase,
          tonase:dateMap[d].tonase
        };
      })
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}



/* ============================================================
 * 10. P2H API
 * ============================================================
 */

function apiSaveP2H(token, data) {
  const session = requireSession_(token, ['ADMIN','OPERATOR']);

  data = data || {};

  const dateKey = toIsoDate_(data.date || new Date());
  const shift = String(data.shift || '').toUpperCase();
  const unit = String(
    data.unit || session.assigned_unit || ''
  ).trim();

  const hmKm = Number(data.hm_km || data.hm || 0);

  if (!dateKey || !shift || !unit) {
    throw new Error('Tanggal, Shift dan Unit wajib diisi.');
  }

  if (!Array.isArray(data.items) || !data.items.length) {
    throw new Error('Checklist P2H tidak boleh kosong.');
  }

  if (
    session.role === 'OPERATOR' &&
    session.assigned_unit &&
    unit !== String(session.assigned_unit)
  ) {
    throw new Error(
      'Unit P2H harus sesuai batangan aktif Anda: ' +
      session.assigned_unit
    );
  }

  if (!findOne_(FJB_SHEETS.UNITS, 'unit_code', unit)) {
    throw new Error('Unit P2H tidak ditemukan pada Master Unit: ' + unit);
  }

  const activeMaster = readObjects_(FJB_SHEETS.P2H_MASTER)
    .filter(function(x) { return toBool_(x.active); });

  const allowedItemIds = {};
  const requiredItemIds = {};

  activeMaster.forEach(function(x) {
    allowedItemIds[String(x.item_id)] = true;
    if (toBool_(x.required)) requiredItemIds[String(x.item_id)] = true;
  });

  const receivedIds = {};
  data.items.forEach(function(item) {
    const id = String(item.item_id || item.id || '');
    if (!id || !allowedItemIds[id]) {
      throw new Error('Item P2H tidak valid / sudah tidak aktif: ' + id);
    }
    receivedIds[id] = true;
  });

  Object.keys(requiredItemIds).forEach(function(id) {
    if (!receivedIds[id]) {
      throw new Error('Checklist P2H wajib belum lengkap: ' + id);
    }
  });

  const p2hId = generateId_('P2H');
  const now = new Date();

  let okCount = 0;
  let findingCount = 0;

  const detailRows = data.items.map(function(item, index) {
    const checked = toBool_(item.checked);
    const status = checked ? 'OK' : 'TEMUAN';

    if (checked) okCount++;
    else findingCount++;

    return {
      detail_id:p2hId + '-' + String(index + 1).padStart(3,'0'),
      p2h_id:p2hId,
      item_id:String(item.item_id || item.id || ''),
      item_name:String(item.item_name || item.name || ''),
      checked:checked,
      status:status,
      note:String(item.note || '')
    };
  });

  const header = {
    p2h_id:p2hId,
    date:parseIsoDate_(dateKey),
    shift:shift,
    time:String(data.time || formatTime_(now)),
    unit:unit,
    hm_km:hmKm,
    total_items:detailRows.length,
    ok_count:okCount,
    finding_count:findingCount,
    notes:String(data.notes || data.remark || ''),
    input_by_nik:session.nik,
    input_by_name:session.name,
    created_at:now,
    updated_at:''
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    appendObject_(FJB_SHEETS.P2H, header);
    appendObjects_(FJB_SHEETS.P2H_DETAIL, detailRows);

    appendHistory_(
      session,
      'P2H',
      'CREATE',
      unit,
      okCount + ' OK • ' + findingCount + ' Temuan',
      {
        header:cleanObject_(header),
        items:detailRows.map(cleanObject_)
      }
    );

  } finally {
    lock.releaseLock();
  }

  return {
    ok:true,
    data:cleanObject_(header),
    details:detailRows.map(cleanObject_)
  };
}


function apiGetP2H(token, filters) {
  const session = requireSession_(token, [
    'ADMIN','GL','OPERATOR'
  ]);

  filters = filters || {};

  let rows = readObjects_(FJB_SHEETS.P2H);

  if (session.role === 'OPERATOR') {
    rows = rows.filter(function(x) {
      return String(x.input_by_nik) === String(session.nik);
    });
  }

  rows = filterCommon_(rows, filters, {
    date:'date',
    shift:'shift',
    unit:'unit',
    nik:'input_by_nik'
  });

  return {
    ok:true,
    rows:rows.map(cleanObject_)
  };
}


function apiGetP2HDetail(token, p2hId) {
  requireSession_(token, [
    'ADMIN','GL','OPERATOR'
  ]);

  const rows = readObjects_(FJB_SHEETS.P2H_DETAIL)
    .filter(function(x) {
      return String(x.p2h_id) === String(p2hId);
    });

  return {
    ok:true,
    rows:rows.map(cleanObject_)
  };
}


/* ============================================================
 * 11. MAINTENANCE API
 * ============================================================
 */

function apiSaveMaintenance(token, data) {
  return withWriteLock_('apiSaveMaintenance', function() {
      const session = requireSession_(token, [
        'ADMIN','MECHANIC'
      ]);

      data = data || {};

      const dateKey = toIsoDate_(data.date || new Date());
      const unit = String(data.unit || '').trim();
      const type = String(data.type || '').toUpperCase();
      const result = String(data.result || '').toUpperCase();

      if (!dateKey || !unit || !type || !result) {
        throw new Error(
          'Tanggal, Unit, Jenis dan Hasil wajib diisi.'
        );
      }

      if (['BREAKDOWN','SERVICE','INSPECTION','COMMISSIONING'].indexOf(type) < 0) {
        throw new Error('Jenis maintenance tidak valid.');
      }

      if (['NORMAL','MONITORING','NOT READY'].indexOf(result) < 0) {
        throw new Error('Hasil maintenance tidak valid.');
      }

      if (!findOne_(FJB_SHEETS.UNITS, 'unit_code', unit)) {
        throw new Error('Unit tidak ditemukan: ' + unit);
      }

      let mechanicNik = session.nik;
      let mechanicName = session.name;

      if (session.role === 'ADMIN') {
        mechanicNik = String(data.mechanic_nik || '').trim();
        if (!mechanicNik) {
          throw new Error('Mechanic wajib dipilih untuk input ADMIN.');
        }

        const mechanic = findOne_(FJB_SHEETS.PERSONNEL, 'nik', mechanicNik);
        if (!mechanic ||
            String(mechanic.category).toUpperCase().indexOf('MECHANIC') < 0) {
          throw new Error('Personnel mechanic tidak valid: ' + mechanicNik);
        }

        mechanicName = mechanic.name;
      }

      const start = String(data.start_time || data.start || '');
      const finish = String(data.finish_time || data.finish || '');

      const duration = durationMinutes_(start, finish);

      const now = new Date();

      const rec = {
        maintenance_id:generateId_('MNT'),
        date:parseIsoDate_(dateKey),
        unit:unit,
        type:type,
        hm_km:Number(data.hm_km || data.hm || 0),
        start_time:start,
        finish_time:finish,
        duration_min:duration,
        result:result,
        problem:String(data.problem || ''),
        action:String(data.action || ''),
        part_material:String(
          data.part_material || data.part || ''
        ),
        mechanic_nik:mechanicNik,
        mechanic_name:mechanicName,
        created_at:now,
        updated_at:''
      };

      appendObject_(FJB_SHEETS.MAINTENANCE, rec);

      appendHistory_(
        session,
        'Maintenance',
        'CREATE',
        unit,
        type + ' • ' + result,
        cleanObject_(rec)
      );

      return {
        ok:true,
        data:cleanObject_(rec)
      };

  });
}


function apiGetMaintenance(token, filters) {
  const session = requireSession_(token, [
    'ADMIN','GL','MECHANIC'
  ]);

  filters = filters || {};

  let rows = readObjectsTailFast_(FJB_SHEETS.MAINTENANCE, 2000);

  if (session.role === 'MECHANIC') {
    rows = rows.filter(function(x) {
      return String(x.mechanic_nik) === String(session.nik);
    });
  }

  rows = filterCommon_(rows, filters, {
    date:'date',
    unit:'unit',
    nik:'mechanic_nik'
  });

  return {
    ok:true,
    rows:rows.map(cleanObject_)
  };
}



function apiCloseMaintenance(token, maintenanceId, result) {
  return withWriteLock_('apiCloseMaintenance', function() {
      const session = requireSession_(token, ['ADMIN','MECHANIC']);
      const row = findOne_(FJB_SHEETS.MAINTENANCE, 'maintenance_id', String(maintenanceId));
      if (!row) throw new Error('Maintenance event tidak ditemukan.');

      if (session.role === 'MECHANIC' && String(row.mechanic_nik) !== String(session.nik)) {
        throw new Error('Anda hanya dapat menutup pekerjaan mechanic Anda sendiri.');
      }

      const finalResult = String(result || 'NORMAL').toUpperCase();
      if (['NORMAL','MONITORING','NOT READY'].indexOf(finalResult) < 0) {
        throw new Error('Hasil maintenance tidak valid.');
      }

      const finish = formatTime_(new Date()).substring(0,5);
      const duration = durationMinutes_(String(row.start_time || ''), finish);
      const now = new Date();

      updateRowObject_(FJB_SHEETS.MAINTENANCE, row._row, {
        finish_time:finish,
        duration_min:duration,
        result:finalResult,
        updated_at:now
      });

      appendHistory_(
        session,
        'Maintenance',
        'UPDATE',
        String(row.unit),
        'Close ' + String(row.maintenance_id) + ' • ' + finalResult,
        {
          maintenance_id:String(row.maintenance_id),
          finish_time:finish,
          duration_min:duration,
          result:finalResult
        }
      );

      return { ok:true };

  });
}


/* ============================================================
 * 12. UNIT STATUS API
 * ============================================================
 */

function apiUpdateUnitStatus(token, data) {
  return withWriteLock_('apiUpdateUnitStatus', function() {
      const session = requireSession_(token, ['ADMIN','GL']);

      data = data || {};

      const unitCode = String(data.unit || '').trim();
      const newStatus = String(data.status || '')
        .trim()
        .toUpperCase();

      const newOperational = String(
        data.operational_status || data.operational || ''
      ).trim();

      const reason = String(data.reason || '').trim();

      const unit = findOne_(
        FJB_SHEETS.UNITS,
        'unit_code',
        unitCode
      );

      if (!unit) {
        throw new Error('Unit tidak ditemukan: ' + unitCode);
      }

      const oldStatus = String(unit.status);
      const oldOperational = String(unit.operational_status || '');

      const now = new Date();

      updateRowObject_(FJB_SHEETS.UNITS, unit._row, {
        status:newStatus,
        operational_status:newOperational,
        updated_at:now
      });

      const log = {
        status_id:generateId_('UST'),
        timestamp:now,
        unit:unitCode,
        old_status:oldStatus,
        new_status:newStatus,
        old_operational_status:oldOperational,
        new_operational_status:newOperational,
        reason:reason,
        updated_by_nik:session.nik,
        updated_by_name:session.name
      };

      appendObject_(FJB_SHEETS.UNIT_STATUS_HISTORY, log);

      appendHistory_(
        session,
        'Unit',
        'UPDATE',
        unitCode,
        oldStatus + ' → ' + newStatus + ' • ' + reason,
        cleanObject_(log)
      );

      return {
        ok:true,
        data:cleanObject_(log)
      };

  });
}


/* ============================================================
 * 13. HISTORY API
 * ============================================================
 */


/* ============================================================
 * V2.6.3 FUEL HISTORY SOURCE SYNC
 * ============================================================
 *
 * 19_FUEL_USAGE = source of truth for Fuel transactions.
 * 17_HISTORY    = audit trail of Web App actions.
 *
 * Direct Spreadsheet imports/edits do not create audit rows.
 * Therefore Input History synthesizes Fuel entries from
 * 19_FUEL_USAGE and suppresses duplicate Fuel audit entries.
 */

function normalizeHistoryTimePart_(value) {
  if (!value) return '00:00:00';

  const clean = serializeValue_(value, 'time');
  const s = String(clean || '').trim();

  if (/^\d{2}:\d{2}:\d{2}$/.test(s)) {
    return s;
  }

  if (/^\d{2}:\d{2}$/.test(s)) {
    return s + ':00';
  }

  return '00:00:00';
}


function fuelHistoryTimestamp_(row) {
  const dateKey = toIsoDate_(row.date);

  if (!dateKey) {
    const created = serializeValue_(
      row.created_at,
      'created_at'
    );

    return String(created || '');
  }

  return dateKey + 'T' +
    normalizeHistoryTimePart_(
      row.fill_time
    );
}


function inferDatabaseHistoryRole_(row) {
  const name = String(
    row.input_by_name ||
    row.created_by_name ||
    ''
  ).toUpperCase();

  if (name.indexOf('ADMIN') >= 0) {
    return 'ADMIN';
  }

  const nik = String(
    row.input_by_nik ||
    row.created_by_nik ||
    ''
  );

  if (nik) {
    return 'OPERATOR';
  }

  return 'SYSTEM';
}


function fuelUsageToHistoryRow_(row) {
  const payload = cleanObject_(row);
  const sourceRow = Number(row._row || 0);
  const fuelId = String(row.fuel_id || '').trim();

  const syntheticKey = sourceRow
    ? String(sourceRow)
    : (fuelId ? ('ID_' + fuelId) : Utilities.getUuid().substring(0,8));

  return {
    log_id:'DB-FUEL-' + syntheticKey,
    timestamp:fuelHistoryTimestamp_(row),
    user_name:String(
      row.input_by_name ||
      'DATABASE'
    ),
    nik:String(
      row.input_by_nik ||
      ''
    ),
    role:inferDatabaseHistoryRole_(row),
    module:'Fuel',
    activity:String(row.updated_at || '').trim()
      ? 'UPDATE'
      : 'CREATE',
    entity:String(
      row.unit_code ||
      row.fuel_id ||
      '-'
    ),
    summary:
      String(row.unit_code || '-') +
      ' • ' +
      Number(row.total_liter || 0) +
      ' L' +
      (
        row.fuel_source
          ? ' • ' + String(row.fuel_source)
          : ''
      ),
    payload_json:JSON.stringify(payload),
    payload:payload,
    source_kind:'DATABASE',
    source_sheet:FJB_SHEETS.FUEL_USAGE,
    source_row:sourceRow,
    source_id:String(row.fuel_id || ''),
    can_mutate_history:true
  };
}


function historyRowForClient_(row) {
  const out = cleanObject_(row);

  if (
    row.payload &&
    typeof row.payload === 'object'
  ) {
    out.payload = row.payload;
  }
  else {
    try {
      out.payload = JSON.parse(
        String(row.payload_json || '{}')
      );
    }
    catch (err) {
      out.payload = {};
    }
  }

  return out;
}


function getFuelHistoryRows_(from, to) {
  return readObjectsTailFast_(
    FJB_SHEETS.FUEL_USAGE,
    8000
  )
  .filter(function(row) {
    const d = toIsoDate_(row.date);

    if (!d) return false;
    if (from && d < from) return false;
    if (to && d > to) return false;

    return true;
  })
  .map(fuelUsageToHistoryRow_);
}


/**
 * Diagnostic only. Does not modify database.
 */
function diagnoseFJBFuelHistoryV263() {
  ensureOperationalExtensionV260_(false);

  const fuelRows = readObjects_(
    FJB_SHEETS.FUEL_USAGE
  );

  const auditRows = readObjects_(
    FJB_SHEETS.HISTORY
  ).filter(function(x) {
    return String(x.module || '')
      .toUpperCase() === 'FUEL';
  });

  const idCounts = {};

  fuelRows.forEach(function(row) {
    const id = String(row.fuel_id || '').trim();
    const key = id || '(BLANK)';

    idCounts[key] =
      Number(idCounts[key] || 0) + 1;
  });

  const duplicateGroups = Object.keys(idCounts)
    .filter(function(id) {
      return id === '(BLANK)' || idCounts[id] > 1;
    })
    .map(function(id) {
      return {
        fuel_id:id,
        count:idCounts[id]
      };
    })
    .sort(function(a,b) {
      return b.count - a.count;
    });

  const result = {
    ok:true,
    version:FJB_VERSION,
    fuel_database_sheet:FJB_SHEETS.FUEL_USAGE,
    fuel_rows:fuelRows.length,
    unique_fuel_ids:Object.keys(idCounts).filter(function(x) {
      return x !== '(BLANK)';
    }).length,
    duplicate_id_groups:duplicateGroups.length,
    duplicate_examples:duplicateGroups.slice(0,10),
    fuel_audit_rows:auditRows.length,
    explanation:
      '19_FUEL_USAGE adalah database transaksi; 17_HISTORY adalah audit log.'
  };

  console.log(JSON.stringify(result,null,2));

  return result;
}


/**
 * Safe repair for manually pasted/imported Fuel rows.
 * Keeps the first occurrence of an ID and rewrites only blank/duplicate IDs.
 * No transaction data is deleted.
 */
function repairDuplicateFuelIdsV263() {
  ensureOperationalExtensionV260_(false);

  const sessionLock = LockService.getScriptLock();
  sessionLock.waitLock(30000);

  try {
    const sheet = getSheet_(FJB_SHEETS.FUEL_USAGE);
    const rows = readObjects_(FJB_SHEETS.FUEL_USAGE);

    if (!rows.length) {
      return {
        ok:true,
        total_rows:0,
        repaired:0
      };
    }

    const headers = FJB_SCHEMA[FJB_SHEETS.FUEL_USAGE];
    const idIndex = headers.indexOf('fuel_id');

    if (idIndex < 0) {
      throw new Error('Kolom fuel_id tidak ditemukan.');
    }

    const seen = {};
    const ids = [];
    let repaired = 0;

    rows.forEach(function(row) {
      let id = String(row.fuel_id || '').trim();

      if (!id || seen[id]) {
        id = generateId_('FUEL');
        repaired++;
      }

      seen[id] = true;
      ids.push([id]);
    });

    if (repaired) {
      sheet.getRange(
        2,
        idIndex + 1,
        ids.length,
        1
      ).setValues(ids);

      SpreadsheetApp.flush();

      bumpFastCacheRevision_(
        FJB_SHEETS.FUEL_USAGE,
        true
      );
    }

    const result = {
      ok:true,
      total_rows:rows.length,
      repaired:repaired,
      unchanged:rows.length - repaired,
      sheet:FJB_SHEETS.FUEL_USAGE
    };

    console.log(JSON.stringify(result,null,2));

    return result;
  }
  finally {
    sessionLock.releaseLock();
  }
}


function apiGetHistory(token, filters) {
  requireSession_(token, ['ADMIN','GL']);

  filters = filters || {};

  const q = String(filters.q || '')
    .trim()
    .toLowerCase();

  const module = String(filters.module || '')
    .trim();

  const activity = String(filters.activity || '')
    .trim()
    .toUpperCase();

  const role = String(filters.role || '')
    .trim()
    .toUpperCase();

  const from = filters.from
    ? toIsoDate_(filters.from)
    : '';

  const to = filters.to
    ? toIsoDate_(filters.to)
    : '';

  /*
   * 17_HISTORY remains the audit source for non-Fuel modules.
   * Fuel audit rows are suppressed because 19_FUEL_USAGE is now
   * the transaction source of truth shown in Input History.
   */
  let rows = readObjectsTailFast_(
    FJB_SHEETS.HISTORY,
    5000
  )
  .filter(function(x) {
    return String(x.module || '')
      .toUpperCase() !== 'FUEL';
  })
  .map(function(x) {
    const out = cleanObject_(x);

    out.source_kind = 'AUDIT';
    out.source_sheet = FJB_SHEETS.HISTORY;
    out.can_mutate_history = true;

    try {
      out.payload = JSON.parse(
        String(x.payload_json || '{}')
      );
    }
    catch (err) {
      out.payload = {};
    }

    return out;
  });

  /*
   * Only read Fuel transactions when the selected module can include Fuel.
   */
  if (
    !module ||
    module.toUpperCase() === 'FUEL'
  ) {
    rows = rows.concat(
      getFuelHistoryRows_(from,to)
    );
  }

  rows = rows.filter(function(x) {
    if (
      module &&
      String(x.module) !== module
    ) {
      return false;
    }

    if (
      activity &&
      String(x.activity)
        .toUpperCase() !== activity
    ) {
      return false;
    }

    if (
      role &&
      String(x.role)
        .toUpperCase() !== role
    ) {
      return false;
    }

    const d = toIsoDate_(x.timestamp);

    if (from && d && d < from) {
      return false;
    }

    if (to && d && d > to) {
      return false;
    }

    if (q) {
      const haystack = [
        x.log_id,
        x.user_name,
        x.nik,
        x.role,
        x.module,
        x.activity,
        x.entity,
        x.summary,
        x.payload_json,
        x.source_id
      ]
        .join(' ')
        .toLowerCase();

      if (haystack.indexOf(q) < 0) {
        return false;
      }
    }

    return true;
  });

  rows.sort(function(a,b) {
    return String(b.timestamp || '')
      .localeCompare(
        String(a.timestamp || '')
      );
  });

  const limit = Math.min(
    5000,
    Math.max(
      1,
      Number(filters.limit || 1500)
    )
  );

  return {
    ok:true,
    unified:true,
    from:from,
    to:to,
    count:Math.min(rows.length,limit),
    rows:rows
      .slice(0,limit)
      .map(historyRowForClient_)
  };
}


function findFuelUsageRow_(key, fallbackData) {
  const sheet = getSheet_(FJB_SHEETS.FUEL_USAGE);
  const headers = FJB_SCHEMA[FJB_SHEETS.FUEL_USAGE];
  const lastRow = sheet.getLastRow();
  const rawKey = String(key || '').replace('DB-FUEL-', '').trim();

  fallbackData = fallbackData || {};

  // 1. Try by source_row if provided
  if (fallbackData.source_row) {
    const sr = Number(fallbackData.source_row);
    if (!isNaN(sr) && sr >= 2 && sr <= lastRow) {
      const vals = sheet.getRange(sr, 1, 1, headers.length).getValues()[0];
      const obj = { _row: sr };
      headers.forEach(function(h, i) { obj[h] = vals[i]; });
      if (obj.unit_code || obj.total_liter || obj.fuel_id) {
        return obj;
      }
    }
  }

  // 2. Try by numeric row number
  const rowNum = Number(rawKey);
  if (!isNaN(rowNum) && rowNum >= 2 && rowNum <= lastRow) {
    const vals = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    const obj = { _row: rowNum };
    headers.forEach(function(h, i) { obj[h] = vals[i]; });
    if (obj.unit_code || obj.total_liter || obj.fuel_id) {
      return obj;
    }
  }

  // 3. Scan sheet for matching fuel_id or _row
  const all = readObjects_(FJB_SHEETS.FUEL_USAGE);
  const targetFuelId = String(fallbackData.fuel_id || rawKey).replace('ID_', '').trim();
  const byId = all.find(function(x) {
    return (targetFuelId && String(x.fuel_id || '').trim() === targetFuelId) ||
           String(x.fuel_id || '').trim() === rawKey ||
           String(x._row || '') === rawKey ||
           ('DB-FUEL-' + x._row) === key;
  });
  if (byId) return byId;

  // 4. Attribute fallback by unit_code, date, and total_liter
  const unitCode = String(fallbackData.unit_code || fallbackData.unit || fallbackData.entity || '').trim().toUpperCase();
  const dateVal = fallbackData.date ? toIsoDate_(fallbackData.date) : '';
  const literVal = fallbackData.total_liter !== undefined ? Number(fallbackData.total_liter) : (fallbackData.liter !== undefined ? Number(fallbackData.liter) : null);

  if (unitCode || dateVal) {
    const byAttrs = all.find(function(x) {
      const uMatch = unitCode ? String(x.unit_code || '').trim().toUpperCase() === unitCode : true;
      const dMatch = dateVal ? toIsoDate_(x.date) === dateVal : true;
      const lMatch = (literVal !== null && !isNaN(literVal)) ? Math.abs(Number(x.total_liter || 0) - literVal) < 0.05 : true;
      return uMatch && dMatch && lMatch;
    });
    if (byAttrs) return byAttrs;
  }

  return null;
}


function apiGetHistoryDetail(token, logId) {
  requireSession_(token, ['ADMIN','GL']);

  const id = String(logId || '').trim();

  if (id.indexOf('DB-FUEL-') === 0) {
    const foundFuel = findFuelUsageRow_(id);
    if (foundFuel) {
      return {
        ok: true,
        data: historyRowForClient_(fuelUsageToHistoryRow_(foundFuel))
      };
    }
    throw new Error('Transaksi Fuel tidak ditemukan: ' + id);
  }

  let row = findOne_(FJB_SHEETS.HISTORY, 'log_id', id);
  if (!row) {
    row = findOne_(FJB_SHEETS.HISTORY, 'entity', id);
  }

  if (!row) {
    throw new Error('History tidak ditemukan: ' + id);
  }

  const data = historyRowForClient_(row);
  data.source_kind = 'AUDIT';
  data.source_sheet = FJB_SHEETS.HISTORY;
  data.can_mutate_history = true;

  return {
    ok: true,
    data: data
  };
}


function apiEditHistory(token, logId, patch) {
  return withWriteLock_('apiEditHistory', function() {
    const session = requireSession_(token, ['ADMIN','GL']);
    const id = String(logId || '').trim();
    patch = patch || {};

    const isFuel = id.indexOf('DB-FUEL-') === 0;

    // 1. Direct DB-FUEL mutation
    if (isFuel) {
      const existing = findFuelUsageRow_(id, patch);
      if (!existing || !existing._row) {
        throw new Error('Transaksi Fuel tidak ditemukan: ' + id);
      }
      const rowNumber = existing._row;
      const before = cleanObject_(existing);

      const d = patch.date ? parseIsoDate_(patch.date) : (existing.date ? new Date(existing.date) : new Date());
      const dateKey = toIsoDate_(d);
      const shift = String(patch.shift !== undefined ? patch.shift : existing.shift).trim().toUpperCase();
      const unitCode = String(patch.unit_code !== undefined ? patch.unit_code : (patch.unit || existing.unit_code)).trim().toUpperCase();
      const totalLiter = Number(patch.total_liter !== undefined ? patch.total_liter : (patch.liter || existing.total_liter || 0));
      const hmKm = Number(patch.hm_km !== undefined ? patch.hm_km : (patch.hm || existing.hm_km || 0));
      const fuelSource = String(patch.fuel_source !== undefined ? patch.fuel_source : (existing.fuel_source || '')).trim().toUpperCase();
      const entityUsed = String(patch.entity_used !== undefined ? patch.entity_used : (existing.entity_used || '')).trim();
      const fillTime = String(patch.fill_time !== undefined ? patch.fill_time : (existing.fill_time || '')).trim();
      const dedicated = String(patch.dedicated !== undefined ? patch.dedicated : (existing.dedicated || '')).trim();
      const location = String(patch.location !== undefined ? patch.location : (existing.location || '')).trim();
      const unitDay = makeFuelUnitDay_(dateKey, unitCode, shift);
      const now = new Date();

      const updatedObj = {
        date: d,
        shift: shift,
        fuel_source: fuelSource,
        entity_used: entityUsed,
        unit_code: unitCode,
        hm_km: hmKm,
        total_liter: totalLiter,
        fill_time: fillTime,
        dedicated: dedicated,
        unit_day: unitDay,
        location: location,
        updated_at: now
      };

      updateRowObject_(FJB_SHEETS.FUEL_USAGE, rowNumber, updatedObj);

      bumpFastCacheRevision_(FJB_SHEETS.FUEL_USAGE, true);
      touchFastViewStamp_('FUEL_EDIT');

      appendHistory_(
        session,
        'Fuel',
        'UPDATE',
        unitCode,
        unitCode + ' • ' + totalLiter + ' L' + (fuelSource ? ' • ' + fuelSource : ''),
        {
          target_log_id: id,
          fuel_id: String(existing.fuel_id || ''),
          before: before,
          after: cleanObject_(updatedObj)
        }
      );

      return {
        ok: true,
        data: historyRowForClient_(
          fuelUsageToHistoryRow_(Object.assign({}, existing, updatedObj, { _row: rowNumber }))
        )
      };
    }

    // 2. Audit record in 17_HISTORY
    const row = findOne_(FJB_SHEETS.HISTORY, 'log_id', id);
    if (!row) {
      throw new Error('History tidak ditemukan: ' + id);
    }

    const before = cleanObject_(row);
    let payload = {};
    try {
      payload = JSON.parse(String(row.payload_json || '{}'));
    } catch (_) {
      payload = {};
    }

    const mod = String(row.module || '').trim().toUpperCase();
    const now = new Date();

    // Cascading update to operational sheets
    if (mod === 'FUEL') {
      const fuelId = String(payload.fuel_id || row.entity || '').trim();
      let fuelRow = fuelId ? findOne_(FJB_SHEETS.FUEL_USAGE, 'fuel_id', fuelId) : null;
      if (!fuelRow && patch.source_row) {
        const sr = Number(patch.source_row);
        if (sr >= 2) fuelRow = { _row: sr };
      }
      if (fuelRow) {
        const d = patch.date ? parseIsoDate_(patch.date) : (fuelRow.date ? new Date(fuelRow.date) : new Date());
        const dateKey = toIsoDate_(d);
        const shift = String(patch.shift !== undefined ? patch.shift : (fuelRow.shift || payload.shift || 'DAY')).trim().toUpperCase();
        const unitCode = String(patch.unit_code !== undefined ? patch.unit_code : (patch.unit || fuelRow.unit_code || payload.unit_code || row.entity)).trim().toUpperCase();
        const totalLiter = Number(patch.total_liter !== undefined ? patch.total_liter : (fuelRow.total_liter || payload.total_liter || 0));
        const hmKm = Number(patch.hm_km !== undefined ? patch.hm_km : (fuelRow.hm_km || payload.hm_km || 0));
        const fuelSource = String(patch.fuel_source !== undefined ? patch.fuel_source : (fuelRow.fuel_source || payload.fuel_source || '')).trim().toUpperCase();
        const entityUsed = String(patch.entity_used !== undefined ? patch.entity_used : (fuelRow.entity_used || payload.entity_used || '')).trim();
        const fillTime = String(patch.fill_time !== undefined ? patch.fill_time : (fuelRow.fill_time || payload.fill_time || '')).trim();
        const location = String(patch.location !== undefined ? patch.location : (fuelRow.location || payload.location || '')).trim();

        updateRowObject_(FJB_SHEETS.FUEL_USAGE, fuelRow._row, {
          date: d,
          shift: shift,
          unit_code: unitCode,
          total_liter: totalLiter,
          hm_km: hmKm,
          fuel_source: fuelSource,
          entity_used: entityUsed,
          fill_time: fillTime,
          location: location,
          unit_day: makeFuelUnitDay_(dateKey, unitCode, shift),
          updated_at: now
        });
        bumpFastCacheRevision_(FJB_SHEETS.FUEL_USAGE, true);
      }
    } else if (mod === 'HAULING') {
      const txId = String(payload.transaction_id || row.entity || '').trim();
      const haulRow = txId ? findOne_(FJB_SHEETS.HAULING, 'transaction_id', txId) : null;
      if (haulRow) {
        const gross = patch.gross !== undefined ? Number(patch.gross) : Number(haulRow.gross || 0);
        const tare = patch.tare !== undefined ? Number(patch.tare) : Number(haulRow.tare || 0);
        const netTon = patch.net_ton !== undefined ? Number(patch.net_ton) : (gross - tare) / 1000;
        const coalProduct = String(patch.coal_product !== undefined ? patch.coal_product : (patch.product_seam || haulRow.coal_product || haulRow.product_seam || '')).trim().toUpperCase();

        const patchHaul = {
          date: patch.date ? parseIsoDate_(patch.date) : haulRow.date,
          shift: patch.shift ? String(patch.shift).trim().toUpperCase() : haulRow.shift,
          time: patch.time ? String(patch.time).trim() : haulRow.time,
          hauler: patch.hauler ? String(patch.hauler).trim().toUpperCase() : haulRow.hauler,
          loader: patch.loader ? String(patch.loader).trim().toUpperCase() : haulRow.loader,
          gross: gross,
          tare: tare,
          net_ton: netTon,
          product_seam: coalProduct,
          coal_product: coalProduct,
          distance: patch.distance !== undefined ? Number(patch.distance) : haulRow.distance,
          jam_ritase: patch.jam_ritase !== undefined ? String(patch.jam_ritase) : haulRow.jam_ritase,
          ritase: patch.ritase !== undefined ? Number(patch.ritase) : haulRow.ritase,
          remark: patch.remark !== undefined ? String(patch.remark) : haulRow.remark,
          updated_at: now
        };
        updateRowObject_(FJB_SHEETS.HAULING, haulRow._row, patchHaul);
        bumpFastCacheRevision_(FJB_SHEETS.HAULING, true);
      }
    } else if (mod === 'HM OPERATION' || mod === 'HM') {
      const hmId = String(payload.hm_id || '').trim();
      let hmRow = hmId ? findOne_(FJB_SHEETS.HM_OPERATION, 'hm_id', hmId) : null;
      if (!hmRow && payload.date && payload.unit && payload.shift) {
        hmRow = readObjects_(FJB_SHEETS.HM_OPERATION).find(function(x) {
          return toIsoDate_(x.date) === toIsoDate_(payload.date) &&
            String(x.unit).toUpperCase() === String(payload.unit).toUpperCase() &&
            String(x.shift).toUpperCase() === String(payload.shift).toUpperCase();
        });
      }
      if (hmRow) {
        const hmStart = patch.hm_start !== undefined ? Number(patch.hm_start) : Number(hmRow.hm_start || 0);
        const hmEnd = patch.hm_end !== undefined ? Number(patch.hm_end) : Number(hmRow.hm_end || 0);
        const totalHM = patch.total_hm !== undefined ? Number(patch.total_hm) : (hmEnd - hmStart);

        updateRowObject_(FJB_SHEETS.HM_OPERATION, hmRow._row, {
          date: patch.date ? parseIsoDate_(patch.date) : hmRow.date,
          shift: patch.shift ? String(patch.shift).trim().toUpperCase() : hmRow.shift,
          unit: patch.unit ? String(patch.unit).trim().toUpperCase() : hmRow.unit,
          nik: patch.nik !== undefined ? String(patch.nik).trim() : hmRow.nik,
          operator_name: patch.operator_name !== undefined ? String(patch.operator_name).trim() : hmRow.operator_name,
          hm_start: hmStart,
          hm_end: hmEnd,
          total_hm: totalHM,
          note: patch.note !== undefined ? String(patch.note) : hmRow.note,
          updated_at: now
        });
        bumpFastCacheRevision_(FJB_SHEETS.HM_OPERATION, true);
      }
    } else if (mod === 'MAINTENANCE') {
      const mId = String(payload.maintenance_id || '').trim();
      const mRow = mId ? findOne_(FJB_SHEETS.MAINTENANCE, 'maintenance_id', mId) : null;
      if (mRow) {
        const start = patch.start_time !== undefined ? String(patch.start_time).trim() : String(mRow.start_time || '');
        const finish = patch.finish_time !== undefined ? String(patch.finish_time).trim() : String(mRow.finish_time || '');
        const duration = patch.duration_min !== undefined ? Number(patch.duration_min) : durationMinutes_(start, finish);

        updateRowObject_(FJB_SHEETS.MAINTENANCE, mRow._row, {
          date: patch.date ? parseIsoDate_(patch.date) : mRow.date,
          unit: patch.unit ? String(patch.unit).trim().toUpperCase() : mRow.unit,
          type: patch.type ? String(patch.type).trim().toUpperCase() : mRow.type,
          hm_km: patch.hm_km !== undefined ? Number(patch.hm_km) : mRow.hm_km,
          start_time: start,
          finish_time: finish,
          duration_min: duration,
          result: patch.result ? String(patch.result).trim().toUpperCase() : mRow.result,
          problem: patch.problem !== undefined ? String(patch.problem) : mRow.problem,
          action: patch.action !== undefined ? String(patch.action) : mRow.action,
          part_material: patch.part_material !== undefined ? String(patch.part_material) : mRow.part_material,
          mechanic_name: patch.mechanic_name !== undefined ? String(patch.mechanic_name) : mRow.mechanic_name,
          updated_at: now
        });
        bumpFastCacheRevision_(FJB_SHEETS.MAINTENANCE, true);
      }
    } else if (mod === 'WASHING') {
      const wId = String(payload.washing_id || '').trim();
      const wRow = wId ? findOne_(FJB_SHEETS.WASHING, 'washing_id', wId) : null;
      if (wRow) {
        updateRowObject_(FJB_SHEETS.WASHING, wRow._row, {
          plan_date: patch.plan_date ? parseIsoDate_(patch.plan_date) : wRow.plan_date,
          actual_date: patch.actual_date ? parseIsoDate_(patch.actual_date) : wRow.actual_date,
          reschedule_date: patch.reschedule_date ? parseIsoDate_(patch.reschedule_date) : wRow.reschedule_date,
          unit: patch.unit ? String(patch.unit).trim().toUpperCase() : wRow.unit,
          pic_name: patch.pic_name !== undefined ? String(patch.pic_name) : wRow.pic_name,
          status: patch.status ? String(patch.status).trim().toUpperCase() : wRow.status,
          note: patch.note !== undefined ? String(patch.note) : wRow.note,
          updated_at: now
        });
        bumpFastCacheRevision_(FJB_SHEETS.WASHING, true);
      }
    } else if (mod === 'P2H') {
      const p2hId = String(payload.p2h_id || (payload.header && payload.header.p2h_id) || '').trim();
      const p2hRow = p2hId ? findOne_(FJB_SHEETS.P2H, 'p2h_id', p2hId) : null;
      if (p2hRow) {
        updateRowObject_(FJB_SHEETS.P2H, p2hRow._row, {
          date: patch.date ? parseIsoDate_(patch.date) : p2hRow.date,
          shift: patch.shift ? String(patch.shift).trim().toUpperCase() : p2hRow.shift,
          time: patch.time ? String(patch.time).trim() : p2hRow.time,
          unit: patch.unit ? String(patch.unit).trim().toUpperCase() : p2hRow.unit,
          hm_km: patch.hm_km !== undefined ? Number(patch.hm_km) : p2hRow.hm_km,
          notes: patch.notes !== undefined ? String(patch.notes) : p2hRow.notes,
          updated_at: now
        });
        bumpFastCacheRevision_(FJB_SHEETS.P2H, true);
      }
    } else if (mod === 'ROSTER') {
      const rosterId = String(payload.roster_id || '').trim();
      const rRow = rosterId ? findOne_(FJB_SHEETS.ROSTER, 'roster_id', rosterId) : null;
      if (rRow) {
        updateRowObject_(FJB_SHEETS.ROSTER, rRow._row, {
          assigned_unit: patch.assigned_unit !== undefined ? String(patch.assigned_unit).trim().toUpperCase() : rRow.assigned_unit,
          roster_status: patch.roster_status !== undefined ? String(patch.roster_status).trim().toUpperCase() : (patch.status || rRow.roster_status),
          updated_at: now
        });
        bumpFastCacheRevision_(FJB_SHEETS.ROSTER, true);
      }
    }

    // Merge patch into history payload
    const mergedPayload = Object.assign({}, payload, patch.payload !== undefined ? patch.payload : patch);
    delete mergedPayload._row;

    // Calculate new Entity and Summary
    let newEntity = patch.entity !== undefined ? String(patch.entity) : String(row.entity);
    if (patch.hauler) newEntity = String(patch.hauler);
    else if (patch.unit_code) newEntity = String(patch.unit_code);
    else if (patch.unit) newEntity = String(patch.unit);

    let newSummary = patch.summary !== undefined ? String(patch.summary) : String(row.summary);
    if (mod === 'HAULING' && (patch.gross || patch.hauler || patch.net_ton)) {
      const nTon = patch.net_ton !== undefined ? Number(patch.net_ton).toFixed(2) : (Number(patch.gross||0) - Number(patch.tare||0))/1000;
      newSummary = (patch.hauler || newEntity) + ' • ' + nTon + ' T • ' + (patch.ritase || 1) + ' rit';
    } else if (mod === 'FUEL' && (patch.total_liter || patch.unit_code)) {
      newSummary = (patch.unit_code || newEntity) + ' • ' + (patch.total_liter || 0) + ' L' + (patch.fuel_source ? ' • ' + patch.fuel_source : '');
    } else if ((mod === 'HM OPERATION' || mod === 'HM') && (patch.total_hm || patch.unit)) {
      newSummary = (patch.date || toIsoDate_(new Date())) + ' • ' + (patch.shift || 'DAY') + ' • ' + (patch.total_hm || 0) + ' HM';
    }

    updateRowObject_(FJB_SHEETS.HISTORY, row._row, {
      entity: newEntity,
      summary: newSummary,
      payload_json: JSON.stringify(mergedPayload)
    });

    bumpFastCacheRevision_(FJB_SHEETS.HISTORY, true);
    touchFastViewStamp_('HISTORY_EDIT');

    appendHistory_(
      session,
      row.module,
      'UPDATE',
      newEntity,
      'Edit: ' + newSummary,
      {
        target_log_id: id,
        before: before,
        after: {
          entity: newEntity,
          summary: newSummary,
          payload: mergedPayload
        }
      }
    );

    const updatedClientRow = historyRowForClient_(
      Object.assign({}, row, {
        entity: newEntity,
        summary: newSummary,
        payload_json: JSON.stringify(mergedPayload),
        payload: mergedPayload
      })
    );

    return {
      ok: true,
      data: updatedClientRow
    };
  });
}


function apiDeleteHistory(token, logId) {
  return withWriteLock_('apiDeleteHistory', function() {
    const session = requireSession_(token, ['ADMIN','GL']);
    const id = String(logId || '').trim();

    const isFuel = id.indexOf('DB-FUEL-') === 0;

    // 1. Direct DB-FUEL deletion
    if (isFuel) {
      const existing = findFuelUsageRow_(id);
      if (!existing || !existing._row) {
        throw new Error('Transaksi Fuel tidak ditemukan: ' + id);
      }
      const rowNumber = existing._row;
      const deleted = cleanObject_(existing);

      deleteRowByNumber_(FJB_SHEETS.FUEL_USAGE, rowNumber);

      bumpFastCacheRevision_(FJB_SHEETS.FUEL_USAGE, true);
      touchFastViewStamp_('FUEL_DELETE');

      appendHistory_(
        session,
        'Fuel',
        'DELETE',
        String(existing.unit_code || existing.fuel_id || '-'),
        'Hapus transaksi Fuel ' + String(existing.unit_code || '') + ' (' + Number(existing.total_liter || 0) + ' L)',
        { deleted_record: deleted }
      );

      return { ok: true, deleted_id: id };
    }

    // 2. Audit record in 17_HISTORY
    const row = findOne_(FJB_SHEETS.HISTORY, 'log_id', id);
    if (!row) {
      throw new Error('History tidak ditemukan: ' + id);
    }

    const deleted = cleanObject_(row);
    let payload = {};
    try {
      payload = JSON.parse(String(row.payload_json || '{}'));
    } catch (_) {
      payload = {};
    }

    const mod = String(row.module || '').trim().toUpperCase();

    // Cascading delete to operational sheets
    if (mod === 'FUEL') {
      const fuelId = String(payload.fuel_id || row.entity || '').trim();
      if (fuelId) {
        const fuelRow = findOne_(FJB_SHEETS.FUEL_USAGE, 'fuel_id', fuelId);
        if (fuelRow) {
          deleteRowByNumber_(FJB_SHEETS.FUEL_USAGE, fuelRow._row);
          bumpFastCacheRevision_(FJB_SHEETS.FUEL_USAGE, true);
        }
      }
    } else if (mod === 'HAULING') {
      const txId = String(payload.transaction_id || row.entity || '').trim();
      if (txId) {
        const haulRow = findOne_(FJB_SHEETS.HAULING, 'transaction_id', txId);
        if (haulRow) {
          deleteRowByNumber_(FJB_SHEETS.HAULING, haulRow._row);
          bumpFastCacheRevision_(FJB_SHEETS.HAULING, true);
        }
      }
    } else if (mod === 'HM OPERATION' || mod === 'HM') {
      const hmId = String(payload.hm_id || '').trim();
      let hmRow = hmId ? findOne_(FJB_SHEETS.HM_OPERATION, 'hm_id', hmId) : null;
      if (!hmRow && payload.date && payload.unit && payload.shift) {
        hmRow = readObjects_(FJB_SHEETS.HM_OPERATION).find(function(x) {
          return toIsoDate_(x.date) === toIsoDate_(payload.date) &&
            String(x.unit).toUpperCase() === String(payload.unit).toUpperCase() &&
            String(x.shift).toUpperCase() === String(payload.shift).toUpperCase();
        });
      }
      if (hmRow) {
        deleteRowByNumber_(FJB_SHEETS.HM_OPERATION, hmRow._row);
        bumpFastCacheRevision_(FJB_SHEETS.HM_OPERATION, true);
      }
    } else if (mod === 'MAINTENANCE') {
      const mId = String(payload.maintenance_id || '').trim();
      if (mId) {
        const mRow = findOne_(FJB_SHEETS.MAINTENANCE, 'maintenance_id', mId);
        if (mRow) {
          deleteRowByNumber_(FJB_SHEETS.MAINTENANCE, mRow._row);
          bumpFastCacheRevision_(FJB_SHEETS.MAINTENANCE, true);
        }
      }
    } else if (mod === 'WASHING') {
      const wId = String(payload.washing_id || '').trim();
      if (wId) {
        const wRow = findOne_(FJB_SHEETS.WASHING, 'washing_id', wId);
        if (wRow) {
          deleteRowByNumber_(FJB_SHEETS.WASHING, wRow._row);
          bumpFastCacheRevision_(FJB_SHEETS.WASHING, true);
        }
      }
    } else if (mod === 'P2H') {
      const p2hId = String(payload.p2h_id || (payload.header && payload.header.p2h_id) || '').trim();
      if (p2hId) {
        const p2hRow = findOne_(FJB_SHEETS.P2H, 'p2h_id', p2hId);
        if (p2hRow) {
          deleteRowByNumber_(FJB_SHEETS.P2H, p2hRow._row);
        }
        const details = readObjects_(FJB_SHEETS.P2H_DETAIL).filter(function(x) {
          return String(x.p2h_id) === p2hId;
        });
        details.sort(function(a, b) { return b._row - a._row; }).forEach(function(d) {
          deleteRowByNumber_(FJB_SHEETS.P2H_DETAIL, d._row);
        });
        bumpFastCacheRevision_(FJB_SHEETS.P2H, true);
      }
    }

    // Delete history row from 17_HISTORY
    deleteRowByNumber_(FJB_SHEETS.HISTORY, row._row);
    bumpFastCacheRevision_(FJB_SHEETS.HISTORY, true);
    touchFastViewStamp_('HISTORY_DELETE');

    appendHistory_(
      session,
      row.module,
      'DELETE',
      row.entity,
      'Hapus data ' + row.module + ' ' + row.entity,
      { deleted_record: deleted }
    );

    return { ok: true, deleted_id: id };
  });
}


function appendHistory_(
  session,
  module,
  activity,
  entity,
  summary,
  payload
) {
  appendObject_(FJB_SHEETS.HISTORY, {
    log_id:'HIST-' + Utilities.getUuid(),
    timestamp:new Date(),
    user_name:session && session.name
      ? session.name
      : 'SYSTEM',
    nik:session && session.nik
      ? session.nik
      : 'SYSTEM',
    role:session && session.role
      ? session.role
      : 'SYSTEM',
    module:String(module || ''),
    activity:String(activity || '').toUpperCase(),
    entity:String(entity || ''),
    summary:String(summary || ''),
    payload_json:JSON.stringify(payload || {})
  });
}


function appendHistorySystem_(
  module,
  activity,
  entity,
  summary,
  payload
) {
  appendHistory_(
    {
      nik:'SYSTEM',
      name:'SYSTEM',
      role:'SYSTEM'
    },
    module,
    activity,
    entity,
    summary,
    payload
  );
}


/* ============================================================
 * 14. MASTER DATA API
 * ============================================================
 */

function apiGetMasterData(token, master) {
  const session = requireSession_(token, ['ADMIN','GL']);

  const map = {
    personnel:FJB_SHEETS.PERSONNEL,
    units:FJB_SHEETS.UNITS,
    unit_alias:FJB_SHEETS.UNIT_ALIAS,
    owner:FJB_SHEETS.OWNER,
    product:FJB_SHEETS.PRODUCT,
    location:FJB_SHEETS.LOCATION,
    options:FJB_SHEETS.OPTIONS,
    p2h_master:FJB_SHEETS.P2H_MASTER
  };

  const sheetName = map[String(master || '').toLowerCase()];

  if (!sheetName) {
    throw new Error('Master data tidak valid.');
  }

  return {
    ok:true,
    readonly:session.role !== 'ADMIN',
    rows:readObjectsFast_(sheetName).map(cleanObject_)
  };
}


function apiSaveP2HMaster(token, data) {
  return withWriteLock_('apiSaveP2HMaster', function() {
      const session = requireSession_(token, ['ADMIN']);

      data = data || {};

      const itemId = String(
        data.item_id || ''
      ).trim();

      const itemName = String(
        data.item_name || ''
      ).trim();

      if (!itemName) {
        throw new Error('Nama item P2H wajib diisi.');
      }

      const nowPayload = {
        item_id:itemId || ('P2H-' + Utilities.getUuid().substring(0,8)),
        item_name:itemName,
        sort_order:Number(data.sort_order || 999),
        active:data.active === undefined
          ? true
          : toBool_(data.active),
        required:data.required === undefined
          ? true
          : toBool_(data.required)
      };

      if (itemId) {
        const existing = findOne_(
          FJB_SHEETS.P2H_MASTER,
          'item_id',
          itemId
        );

        if (existing) {
          updateRowObject_(
            FJB_SHEETS.P2H_MASTER,
            existing._row,
            nowPayload
          );

          appendHistory_(
            session,
            'P2H Master',
            'UPDATE',
            itemId,
            itemName,
            nowPayload
          );

          return { ok:true, mode:'UPDATE' };
        }
      }

      appendObject_(
        FJB_SHEETS.P2H_MASTER,
        nowPayload
      );

      appendHistory_(
        session,
        'P2H Master',
        'CREATE',
        nowPayload.item_id,
        itemName,
        nowPayload
      );

      return { ok:true, mode:'CREATE' };

  });
}


/* ============================================================
 * 15. USER MANAGEMENT API
 * ============================================================
 */

function apiListUsers(token) {
  requireSession_(token, ['ADMIN']);

  const personnelMap = {};

  readObjects_(FJB_SHEETS.PERSONNEL)
    .forEach(function(p) {
      personnelMap[String(p.nik)] = p;
    });

  return {
    ok:true,
    rows:readObjects_(FJB_SHEETS.USERS)
      .map(function(u) {
        const p = personnelMap[String(u.nik)] || {};

        return {
          user_id:u.user_id,
          nik:String(u.nik),
          name:p.name || '',
          category:p.category || '',
          position:p.position || '',
          role:resolveRole_(u, p),
          role_override:String(u.role_override || ''),
          status:u.status,
          must_change_password:toBool_(u.must_change_password),
          last_login_at:serializeValue_(u.last_login_at)
        };
      })
  };
}


function apiResetPassword(token, targetNik, newPassword, mustChange) {
  return withWriteLock_('apiResetPassword', function() {
      const session = requireSession_(token, ['ADMIN']);

      const user = findOne_(
        FJB_SHEETS.USERS,
        'nik',
        String(targetNik)
      );

      if (!user) {
        throw new Error('User tidak ditemukan.');
      }

      if (!newPassword || String(newPassword).length < 6) {
        throw new Error('Password minimal 6 karakter.');
      }

      updateRowObject_(FJB_SHEETS.USERS, user._row, {
        password:String(newPassword),
        must_change_password:toBool_(mustChange),
        updated_at:new Date()
      });

      invalidateSessionsForNik_(String(targetNik), String(targetNik) === session.nik ? token : '');

      appendHistory_(
        session,
        'Users',
        'UPDATE',
        String(targetNik),
        'Reset password user',
        {
          nik:String(targetNik),
          must_change_password:toBool_(mustChange)
        }
      );

      return { ok:true };

  });
}



/* ============================================================
 * 15B. PASSWORD & GENERIC MASTER MANAGEMENT
 * ============================================================
 */

function apiChangePassword(token, oldPassword, newPassword) {
  return withWriteLock_('apiChangePassword', function() {
      const session = requireSession_(token);

      const user = findOne_(
        FJB_SHEETS.USERS,
        'nik',
        String(session.nik)
      );

      if (!user) {
        throw new Error('User tidak ditemukan.');
      }

      if (String(user.password) !== String(oldPassword || '')) {
        throw new Error('Password lama salah.');
      }

      if (!newPassword || String(newPassword).length < 6) {
        throw new Error('Password baru minimal 6 karakter.');
      }

      updateRowObject_(FJB_SHEETS.USERS, user._row, {
        password:String(newPassword),
        must_change_password:false,
        updated_at:new Date()
      });

      appendHistory_(
        session,
        'Users',
        'UPDATE',
        session.nik,
        'User mengganti password',
        { nik:session.nik }
      );

      return { ok:true };

  });
}


function getMasterConfig_(master) {
  const key = String(master || '').toLowerCase();

  const map = {
    personnel:{
      sheet:FJB_SHEETS.PERSONNEL,
      key:'nik',
      allowed:[
        'nik','name','category','position','assigned_unit',
        'team','status','phone','join_date'
      ]
    },

    units:{
      sheet:FJB_SHEETS.UNITS,
      key:'unit_code',
      allowed:[
        'unit_id','unit_code','unit_type','owner','status',
        'operational_status','hm_km','availability_pct',
        'achievement_pct','tonase_today','ritase_today',
        'assigned_nik','assigned_name','location','active'
      ]
    },

    unit_alias:{
      sheet:FJB_SHEETS.UNIT_ALIAS,
      key:'alias_code',
      allowed:[
        'alias_code','canonical_unit','active','note'
      ]
    },

    owner:{
      sheet:FJB_SHEETS.OWNER,
      key:'owner_code',
      allowed:[
        'owner_code','owner_name','active'
      ]
    },

    product:{
      sheet:FJB_SHEETS.PRODUCT,
      key:'product_code',
      allowed:[
        'product_code','product_name','active'
      ]
    },

    location:{
      sheet:FJB_SHEETS.LOCATION,
      key:'location_code',
      allowed:[
        'location_code','location_name','active'
      ]
    }
  };

  return map[key] || null;
}


function apiSaveMasterData(token, master, data) {
  return withWriteLock_('apiSaveMasterData', function() {
      const session = requireSession_(token, ['ADMIN']);

      const cfg = getMasterConfig_(master);
      if (!cfg) {
        throw new Error('Master data tidak didukung untuk edit.');
      }

      data = data || {};

      const keyValue = String(data[cfg.key] || '').trim();
      if (!keyValue) {
        throw new Error(cfg.key + ' wajib diisi.');
      }

      const patch = {};

      cfg.allowed.forEach(function(field) {
        if (data[field] !== undefined) {
          patch[field] = data[field];
        }
      });

      const now = new Date();

      if (cfg.sheet === FJB_SHEETS.PERSONNEL) {
        patch.updated_at = now;

        if (patch.join_date) {
          patch.join_date = parseIsoDate_(toIsoDate_(patch.join_date));
        }

        if (!patch.status) {
          patch.status = 'ACTIVE';
        }
      }

      if (cfg.sheet === FJB_SHEETS.UNITS) {
        patch.updated_at = now;

        if (!patch.unit_id) {
          patch.unit_id = 'UNIT-' + keyValue;
        }

        if (patch.active === undefined) {
          patch.active = true;
        }
      }

      if (
        cfg.sheet === FJB_SHEETS.UNIT_ALIAS ||
        cfg.sheet === FJB_SHEETS.OWNER ||
        cfg.sheet === FJB_SHEETS.PRODUCT ||
        cfg.sheet === FJB_SHEETS.LOCATION
      ) {
        if (patch.active === undefined) {
          patch.active = true;
        }
      }

      const existing = findOne_(
        cfg.sheet,
        cfg.key,
        keyValue
      );

      if (existing) {
        updateRowObject_(cfg.sheet, existing._row, patch);

        if (cfg.sheet === FJB_SHEETS.PERSONNEL) {
          syncActiveSessionsForNik_(keyValue, {
            name:patch.name !== undefined ? patch.name : existing.name,
            position:patch.position !== undefined ? patch.position : existing.position,
            assigned_unit:patch.assigned_unit !== undefined ? patch.assigned_unit : existing.assigned_unit
          });

          if (patch.assigned_unit !== undefined &&
              String(patch.assigned_unit) !== String(existing.assigned_unit || '')) {
            applyPersonnelAssignment_(keyValue, String(patch.assigned_unit || ''), null);
          }
        }

        appendHistory_(
          session,
          'Master Data',
          'UPDATE',
          keyValue,
          String(master) + ' diperbarui',
          patch
        );

        return {
          ok:true,
          mode:'UPDATE',
          data:patch
        };
      }

      if (cfg.sheet === FJB_SHEETS.PERSONNEL) {
        patch.created_at = now;
        patch.updated_at = now;
      }

      appendObject_(cfg.sheet, patch);

      appendHistory_(
        session,
        'Master Data',
        'CREATE',
        keyValue,
        String(master) + ' dibuat',
        patch
      );

      return {
        ok:true,
        mode:'CREATE',
        data:patch
      };

  });
}


function apiSaveMasterBatch(token, master, items) {
  return withWriteLock_('apiSaveMasterBatch', function() {
    const session = requireSession_(token, ['ADMIN']);

    const cfg = getMasterConfig_(master);
    if (!cfg) {
      throw new Error('Master data tidak didukung untuk batch edit.');
    }

    if (!Array.isArray(items) || !items.length) {
      throw new Error('Tidak ada data master untuk disimpan.');
    }

    if (items.length > 500) {
      throw new Error('Maksimal 500 baris master data per batch save.');
    }

    const existingRows = readObjects_(cfg.sheet);
    const existingMap = {};
    existingRows.forEach(function(r) {
      const k = String(r[cfg.key] || '').trim().toUpperCase();
      if (k) existingMap[k] = r;
    });

    const now = new Date();
    const newRows = [];
    let updatedCount = 0;
    let insertedCount = 0;
    const historySummaries = [];

    items.forEach(function(item, idx) {
      if (!item || typeof item !== 'object') return;
      const keyValue = String(item[cfg.key] || '').trim();
      if (!keyValue) {
        throw new Error('Baris ke-' + (idx + 1) + ': ' + cfg.key + ' wajib diisi.');
      }

      const patch = {};
      cfg.allowed.forEach(function(field) {
        if (item[field] !== undefined) {
          patch[field] = item[field];
        }
      });

      // Type conversion & module-specific defaults
      if (cfg.sheet === FJB_SHEETS.PERSONNEL) {
        patch.updated_at = now;
        if (patch.join_date) {
          patch.join_date = parseIsoDate_(toIsoDate_(patch.join_date));
        }
        if (!patch.status) {
          patch.status = 'ACTIVE';
        }
      }

      if (cfg.sheet === FJB_SHEETS.UNITS) {
        patch.updated_at = now;
        if (!patch.unit_id) {
          patch.unit_id = 'UNIT-' + keyValue;
        }
        if (patch.active === undefined) {
          patch.active = true;
        }
        if (patch.hm_km !== undefined && patch.hm_km !== '') {
          patch.hm_km = Number(patch.hm_km);
        }
        if (patch.availability_pct !== undefined && patch.availability_pct !== '') {
          patch.availability_pct = Number(patch.availability_pct);
        }
        if (patch.achievement_pct !== undefined && patch.achievement_pct !== '') {
          patch.achievement_pct = Number(patch.achievement_pct);
        }
      }

      if (
        cfg.sheet === FJB_SHEETS.UNIT_ALIAS ||
        cfg.sheet === FJB_SHEETS.OWNER ||
        cfg.sheet === FJB_SHEETS.PRODUCT ||
        cfg.sheet === FJB_SHEETS.LOCATION
      ) {
        if (patch.active === undefined) {
          patch.active = true;
        }
      }

      const lookupKey = keyValue.toUpperCase();
      const existing = existingMap[lookupKey];

      if (existing) {
        updateRowObject_(cfg.sheet, existing._row, patch);
        updatedCount++;

        if (cfg.sheet === FJB_SHEETS.PERSONNEL) {
          syncActiveSessionsForNik_(keyValue, {
            name: patch.name !== undefined ? patch.name : existing.name,
            position: patch.position !== undefined ? patch.position : existing.position,
            assigned_unit: patch.assigned_unit !== undefined ? patch.assigned_unit : existing.assigned_unit
          });
          if (patch.assigned_unit !== undefined && String(patch.assigned_unit) !== String(existing.assigned_unit || '')) {
            applyPersonnelAssignment_(keyValue, String(patch.assigned_unit || ''), null);
          }
        }
      } else {
        if (cfg.sheet === FJB_SHEETS.PERSONNEL) {
          patch.created_at = now;
          patch.updated_at = now;
        }
        newRows.push(patch);
        existingMap[lookupKey] = patch;
        insertedCount++;
      }

      if (historySummaries.length < 5) {
        historySummaries.push(keyValue);
      }
    });

    if (newRows.length) {
      appendObjects_(cfg.sheet, newRows);
    }

    bumpFastCacheRevision_(cfg.sheet, true);

    const summaryText = updatedCount + ' diupdate, ' + insertedCount + ' dibuat (' + historySummaries.join(', ') + (items.length > 5 ? '...' : '') + ')';
    appendHistory_(
      session,
      'Master Data',
      'UPDATE',
      String(master).toUpperCase(),
      'Simpan Masal ' + master + ': ' + summaryText,
      { count: items.length, updated: updatedCount, inserted: insertedCount }
    );

    return {
      ok: true,
      master: master,
      updated: updatedCount,
      inserted: insertedCount,
      total: updatedCount + insertedCount
    };
  });
}


function apiDeactivateMasterData(token, master, keyValue) {
  return withWriteLock_('apiDeactivateMasterData', function() {
      const session = requireSession_(token, ['ADMIN']);

      const cfg = getMasterConfig_(master);
      if (!cfg) {
        throw new Error('Master data tidak didukung.');
      }

      const existing = findOne_(
        cfg.sheet,
        cfg.key,
        String(keyValue)
      );

      if (!existing) {
        throw new Error('Data tidak ditemukan.');
      }

      let patch = {};

      if (cfg.sheet === FJB_SHEETS.PERSONNEL) {
        patch = {
          status:'INACTIVE',
          updated_at:new Date()
        };
      } else if (cfg.sheet === FJB_SHEETS.UNITS) {
        patch = {
          active:false,
          updated_at:new Date()
        };
      } else {
        patch = { active:false };
      }

      updateRowObject_(cfg.sheet, existing._row, patch);

      appendHistory_(
        session,
        'Master Data',
        'DELETE',
        String(keyValue),
        String(master) + ' dinonaktifkan',
        {
          key:String(keyValue),
          patch:patch
        }
      );

      return { ok:true };

  });
}


function apiSaveUserAccount(token, data) {
  return withWriteLock_('apiSaveUserAccount', function() {
      const session = requireSession_(token, ['ADMIN']);

      data = data || {};

      const nik = String(data.nik || '').trim();
      const password = String(data.password || '').trim();
      const roleOverride = String(
        data.role_override || ''
      ).trim().toUpperCase();

      if (!nik) {
        throw new Error('NIK / NRP wajib diisi.');
      }

      let person = findOne_(
        FJB_SHEETS.PERSONNEL,
        'nik',
        nik
      );

      const now = new Date();

      if (!person) {
        const personName = String(data.name || '').trim() || ('User ' + nik);
        const personCategory = String(
          data.category ||
          (roleOverride === 'MECHANIC' ? 'MECHANIC' :
           roleOverride === 'GL' ? 'GL / PENGAWAS' :
           roleOverride === 'ADMIN' ? 'ADMIN' : 'DRIVER DT')
        );
        const personPosition = String(data.position || roleOverride || personCategory || 'STAFF');

        person = {
          nik: nik,
          name: personName,
          category: personCategory,
          position: personPosition,
          assigned_unit: '',
          team: '',
          status: 'ACTIVE',
          phone: String(data.phone || ''),
          join_date: now,
          created_at: now,
          updated_at: now
        };

        appendObject_(FJB_SHEETS.PERSONNEL, person);
        bumpFastCacheRevision_(FJB_SHEETS.PERSONNEL, true);

        appendHistory_(
          session,
          'Master Data',
          'CREATE',
          nik,
          'Personnel otomatis dibuat saat penambahan user',
          { nik: nik, name: personName, category: personCategory }
        );
      } else if (data.name && String(data.name).trim() && String(data.name).trim() !== String(person.name || '')) {
        updateRowObject_(FJB_SHEETS.PERSONNEL, person._row, {
          name: String(data.name).trim(),
          updated_at: now
        });
        bumpFastCacheRevision_(FJB_SHEETS.PERSONNEL, true);
      }

      if (
        roleOverride &&
        ['ADMIN','GL','OPERATOR','MECHANIC'].indexOf(roleOverride) < 0
      ) {
        throw new Error('Role override tidak valid.');
      }

      const existing = findOne_(
        FJB_SHEETS.USERS,
        'nik',
        nik
      );

      if (existing) {
        const patch = {
          role_override:roleOverride,
          status:String(data.status || existing.status || 'ACTIVE'),
          must_change_password:data.must_change_password === undefined
            ? toBool_(existing.must_change_password)
            : toBool_(data.must_change_password),
          updated_at:now
        };

        if (password) {
          if (password.length < 6) {
            throw new Error('Password minimal 6 karakter.');
          }
          patch.password = password;
        }

        updateRowObject_(
          FJB_SHEETS.USERS,
          existing._row,
          patch
        );

        // Role/status/password changes must apply immediately to active sessions.
        invalidateSessionsForNik_(nik, nik === session.nik ? token : '');

        appendHistory_(
          session,
          'Users',
          'UPDATE',
          nik,
          'User account diperbarui',
          {
            role_override:roleOverride,
            status:patch.status
          }
        );

        return { ok:true, mode:'UPDATE' };
      }

      if (!password || password.length < 6) {
        throw new Error(
          'Password user baru minimal 6 karakter.'
        );
      }

      appendObject_(FJB_SHEETS.USERS, {
        user_id:'USR-' + nik,
        nik:nik,
        password:password,
        role_override:roleOverride,
        status:String(data.status || 'ACTIVE'),
        must_change_password:toBool_(data.must_change_password),
        last_login_at:'',
        created_at:now,
        updated_at:now
      });

      appendHistory_(
        session,
        'Users',
        'CREATE',
        nik,
        'User account dibuat',
        {
          role_override:roleOverride,
          status:String(data.status || 'ACTIVE')
        }
      );

      return { ok:true, mode:'CREATE' };

  });
}



/* ============================================================
 * 15C. ROSTER / WASHING CONTEXT & DIAGNOSTICS
 * ============================================================
 */

/**
 * Satu endpoint untuk mengecek hubungan:
 * Login NIK -> Personnel -> Assigned Unit -> Roster -> Washing.
 *
 * Berguna untuk troubleshooting dan UI.
 */
function apiGetRosterWashingContext(token, month) {
  const session = requireSession_(token);

  const monthKey = normalizeMonthKey_(month) ||
    Utilities.formatDate(
      new Date(),
      getTimezone_(),
      'yyyy-MM'
    );

  const person = findOne_(
    FJB_SHEETS.PERSONNEL,
    'nik',
    String(session.nik)
  );

  let roster = readObjects_(FJB_SHEETS.ROSTER)
    .filter(function(x) {
      return rowMonthMatches_(
        x,
        monthKey,
        'date'
      );
    });

  let washing = readObjects_(FJB_SHEETS.WASHING)
    .filter(function(x) {
      return rowMonthMatches_(
        x,
        monthKey,
        'plan_date'
      );
    });

  if (
    session.role === 'OPERATOR' ||
    session.role === 'MECHANIC'
  ) {
    roster = roster.filter(function(x) {
      return String(x.nik) ===
        String(session.nik);
    });
  }

  if (session.role === 'OPERATOR') {
    washing = washing.filter(function(x) {
      return String(x.unit) ===
        String(session.assigned_unit || '');
    });
  }

  return {
    ok:true,
    month:monthKey,

    session:{
      nik:session.nik,
      name:session.name,
      role:session.role,
      position:session.position,
      assigned_unit:session.assigned_unit
    },

    personnel:person
      ? cleanObject_(person)
      : null,

    roster_count:roster.length,
    washing_count:washing.length,

    roster:roster.map(cleanObject_),
    washing:washing.map(cleanObject_),

    database:{
      roster_total:Math.max(
        0,
        getSheet_(FJB_SHEETS.ROSTER)
          .getLastRow() - 1
      ),
      washing_total:Math.max(
        0,
        getSheet_(FJB_SHEETS.WASHING)
          .getLastRow() - 1
      )
    }
  };
}


/**
 * Cek dan sinkronkan Assigned Unit pada master unit
 * berdasarkan Personnel.
 * Tidak menghapus data existing.
 */
function repairPersonnelUnitLinks() {
  const sessionUser = {
    nik:'SYSTEM',
    name:'SYSTEM',
    role:'SYSTEM'
  };

  const result = repairPersonnelUnitLinks_();

  appendHistory_(
    sessionUser,
    'System',
    'UPDATE',
    'PERSONNEL_UNIT_LINK',
    'Sinkronisasi personnel ke assigned unit',
    result
  );

  SpreadsheetApp.getActiveSpreadsheet()
    .toast(
      'Sinkronisasi Personnel ↔ Unit selesai.',
      'FJB SYSTEM',
      6
    );

  return result;
}


function repairPersonnelUnitLinks_() {
  const people = readObjects_(
    FJB_SHEETS.PERSONNEL
  );

  const units = readObjects_(
    FJB_SHEETS.UNITS
  );

  const unitMap = {};

  units.forEach(function(u) {
    unitMap[
      String(u.unit_code)
    ] = u;
  });

  let updated = 0;
  let missingUnit = 0;

  people.forEach(function(p) {
    const assigned = String(
      p.assigned_unit || ''
    ).trim();

    if (!assigned ||
        assigned.indexOf('DT-') !== 0) {
      return;
    }

    const unit = unitMap[assigned];

    if (!unit) {
      missingUnit++;
      return;
    }

    const needUpdate =
      String(unit.assigned_nik || '') !==
        String(p.nik) ||
      String(unit.assigned_name || '') !==
        String(p.name);

    if (!needUpdate) {
      return;
    }

    updateRowObject_(
      FJB_SHEETS.UNITS,
      unit._row,
      {
        assigned_nik:p.nik,
        assigned_name:p.name,
        updated_at:new Date()
      }
    );

    updated++;
  });

  return {
    ok:true,
    unit_links_updated:updated,
    assigned_unit_not_found:missingUnit
  };
}



/* ============================================================
 * 15D. FAST SPA PRIME API
 * ============================================================
 */

/**
 * Memuat seluruh data yang umum dipakai UI dalam SATU request.
 *
 * Tujuan:
 * - pindah menu tidak memanggil server lagi;
 * - browser hanya render data dari memory;
 * - request Google Sheet dilakukan saat login, manual refresh,
 *   save, atau ganti bulan.
 */
function apiFastPrime(token, month) {
  const session = requireSession_(token);

  const monthKey = normalizeMonthKey_(month) ||
    Utilities.formatDate(
      new Date(),
      getTimezone_(),
      'yyyy-MM'
    );

  const now = new Date();
  const role = String(session.role || '').trim().toUpperCase();

  if (!role || !ROLE_MENUS[role]) {
    throw new Error(
      'FAST_PRIME_ROLE_INVALID: Session role tidak memiliki menu.'
    );
  }

  const sessionPayload = {
    nik:String(session.nik || ''),
    name:String(session.name || ''),
    role:role,
    position:String(session.position || ''),
    assigned_unit:String(session.assigned_unit || ''),
    menus:ROLE_MENUS[role].slice()
  };

  // Master kecil: dibaca satu kali.
  const configRows = readObjects_(FJB_SHEETS.CONFIG);
  const personnelAll = readObjects_(FJB_SHEETS.PERSONNEL);
  const unitsAll = readObjects_(FJB_SHEETS.UNITS)
    .filter(function(x) {
      return toBool_(x.active);
    });

  const products = readObjects_(FJB_SHEETS.PRODUCT)
    .filter(function(x) {
      return toBool_(x.active);
    });

  const locations = readObjects_(FJB_SHEETS.LOCATION)
    .filter(function(x) {
      return toBool_(x.active);
    });

  const options = readObjects_(FJB_SHEETS.OPTIONS)
    .filter(function(x) {
      return toBool_(x.active);
    });

  const p2hMaster = readObjects_(FJB_SHEETS.P2H_MASTER)
    .filter(function(x) {
      return toBool_(x.active);
    })
    .sort(function(a,b) {
      return Number(a.sort_order) - Number(b.sort_order);
    });

  const self = personnelAll.find(function(p) {
    return String(p.nik) === String(session.nik);
  }) || null;

  // Personnel visible berdasarkan role.
  let personnel = personnelAll.filter(function(p) {
    return String(p.status).toUpperCase() === 'ACTIVE';
  });

  if (role === 'OPERATOR' || role === 'MECHANIC') {
    personnel = self ? [self] : [];
  }

  // Roster current month.
  let roster = readObjects_(FJB_SHEETS.ROSTER)
    .filter(function(x) {
      return rowMonthMatches_(x, monthKey, 'date');
    });

  if (role === 'OPERATOR' || role === 'MECHANIC') {
    roster = roster.filter(function(x) {
      return String(x.nik) === String(session.nik);
    });
  }

  // Washing current month.
  let washing = [];
  if (role !== 'MECHANIC') {
    washing = readObjects_(FJB_SHEETS.WASHING)
      .filter(function(x) {
        return rowMonthMatches_(x, monthKey, 'plan_date');
      });

    if (role === 'OPERATOR') {
      washing = washing.filter(function(x) {
        return String(x.unit) === String(session.assigned_unit || '');
      });
    }
  }

  // Transaksi: hanya tail, agar tetap cepat walaupun database membesar.
  let hauling = [];
  if (role === 'ADMIN' || role === 'GL' || role === 'OPERATOR') {
    hauling = readObjectsTail_(FJB_SHEETS.HAULING, 2500);

    if (role === 'OPERATOR') {
      hauling = hauling.filter(function(x) {
        return String(x.input_by_nik) === String(session.nik);
      });
    }
  }

  let p2h = [];
  if (role === 'ADMIN' || role === 'GL' || role === 'OPERATOR') {
    p2h = readObjectsTail_(FJB_SHEETS.P2H, 1800);

    if (role === 'OPERATOR') {
      p2h = p2h.filter(function(x) {
        return String(x.input_by_nik) === String(session.nik);
      });
    }
  }

  let maintenance = [];
  if (role === 'ADMIN' || role === 'GL' || role === 'MECHANIC') {
    maintenance = readObjectsTail_(FJB_SHEETS.MAINTENANCE, 1200);

    if (role === 'MECHANIC') {
      maintenance = maintenance.filter(function(x) {
        return String(x.mechanic_nik) === String(session.nik);
      });
    }
  }

  // History preload secukupnya. Detail history tetap API khusus.
  let history = [];
  if (role === 'ADMIN' || role === 'GL') {
    history = readObjectsTail_(FJB_SHEETS.HISTORY, 500)
      .sort(function(a,b) {
        return new Date(b.timestamp) - new Date(a.timestamp);
      });
  }

  // Users only for ADMIN.
  let users = [];
  if (role === 'ADMIN') {
    const pMap = {};

    personnelAll.forEach(function(p) {
      pMap[String(p.nik)] = p;
    });

    users = readObjects_(FJB_SHEETS.USERS)
      .map(function(u) {
        const p = pMap[String(u.nik)] || {};

        return {
          user_id:u.user_id,
          nik:String(u.nik),
          name:p.name || '',
          category:p.category || '',
          position:p.position || '',
          role:resolveRole_(u, p),
          role_override:String(u.role_override || ''),
          status:u.status,
          must_change_password:toBool_(u.must_change_password),
          last_login_at:serializeValue_(u.last_login_at)
        };
      });
  }

  const settings = {};
  configRows.forEach(function(r) {
    settings[String(r.key)] = r.value;
  });

  const app = {
    name:settings.APP_NAME || 'FJB Operations Control',
    company:settings.COMPANY_NAME || 'PT. FORTUNA JAYA BERSAUDARA',
    version:FJB_VERSION
  };

  return {
    ok:true,
    server_time:now.toISOString(),
    month:monthKey,
    app:app,

    session:sessionPayload,

    personnel_self:self ? cleanObject_(self) : null,
    personnel:personnel.map(cleanObject_),
    units:unitsAll.map(cleanObject_),
    products:products.map(cleanObject_),
    locations:locations.map(cleanObject_),
    options:options.map(cleanObject_),
    p2h_master:p2hMaster.map(cleanObject_),

    roster:roster.map(cleanObject_),
    washing:washing.map(cleanObject_),
    hauling:hauling.map(cleanObject_),
    p2h:p2h.map(cleanObject_),
    maintenance:maintenance.map(cleanObject_),
    history:history.map(cleanObject_),
    users:users,

    masters:{
      unit_alias:(
        role === 'ADMIN' || role === 'GL'
          ? readObjects_(FJB_SHEETS.UNIT_ALIAS).map(cleanObject_)
          : []
      ),
      owner:(
        role === 'ADMIN' || role === 'GL'
          ? readObjects_(FJB_SHEETS.OWNER).map(cleanObject_)
          : []
      )
    },

    settings:settings,

    db_info:{
      roster_rows:Math.max(
        0,
        getSheet_(FJB_SHEETS.ROSTER).getLastRow() - 1
      ),
      washing_rows:Math.max(
        0,
        getSheet_(FJB_SHEETS.WASHING).getLastRow() - 1
      )
    }
  };
}



/**
 * Performance diagnostic.
 * Jalankan dari Apps Script untuk mengukur prime response actual.
 */

/**
 * Read-only verification that Safe Write components are available.
 */
function apiWriteSafetyStatus(token) {
  const session = requireSession_(token);

  return {
    ok:true,
    version:FJB_VERSION,
    nik:session.nik,
    role:session.role,
    write_lock:true,
    append_readback_verification:true,
    update_readback_verification:true,
    delete_verification:true,
    roster_batch_lock:true,
    p2h_transaction_lock:true,
    database_id:getDb_().getId()
  };
}


function apiPerformancePing(token, month) {
  const started = Date.now();
  const prime = apiFastPrime(token, month);

  return {
    ok:true,
    elapsed_ms:Date.now() - started,
    roster_count:(prime.roster || []).length,
    washing_count:(prime.washing || []).length,
    hauling_count:(prime.hauling || []).length,
    maintenance_count:(prime.maintenance || []).length,
    history_count:(prime.history || []).length
  };
}


/* ============================================================
 * 15E. DATABASE-SYNC OPERATIONAL SNAPSHOT
 * ============================================================
 */

function getConfigMap_() {
  const out = {};

  readObjects_(FJB_SHEETS.CONFIG)
    .forEach(function(r) {
      out[String(r.key)] = r.value;
    });

  return out;
}


function rosterStatusToShift_(status) {
  const s = String(status || '').toUpperCase();

  if (s === 'D') return 'DAY';
  if (s === 'N') return 'NIGHT';

  return '';
}


function getRosterRowsForDate_(dateKey) {
  return readObjects_(FJB_SHEETS.ROSTER)
    .filter(function(r) {
      return toIsoDate_(r.date) === dateKey;
    });
}


function getHaulingRowsForDate_(dateKey) {
  return readObjects_(FJB_SHEETS.HAULING)
    .filter(function(r) {
      return toIsoDate_(r.date) === dateKey;
    });
}


function getP2HRowsForDate_(dateKey) {
  return readObjects_(FJB_SHEETS.P2H)
    .filter(function(r) {
      return toIsoDate_(r.date) === dateKey;
    });
}


function buildOperationalUnits_(dateKey) {
  const units = readObjects_(FJB_SHEETS.UNITS)
    .filter(function(x) {
      return toBool_(x.active);
    });

  const personnel = readObjects_(FJB_SHEETS.PERSONNEL);
  const personnelMap = {};

  personnel.forEach(function(p) {
    personnelMap[String(p.nik)] = p;
  });

  const roster = getRosterRowsForDate_(dateKey);
  const rosterMap = {};

  roster.forEach(function(r) {
    rosterMap[String(r.nik)] = r;
  });

  const hauling = getHaulingRowsForDate_(dateKey);
  const haulAgg = {};

  hauling.forEach(function(h) {
    const key = String(h.hauler || '');

    if (!haulAgg[key]) {
      haulAgg[key] = {
        ritase:0,
        tonase:0
      };
    }

    haulAgg[key].ritase += 1;
    haulAgg[key].tonase += Number(h.net_ton || 0);
  });

  const p2h = getP2HRowsForDate_(dateKey);
  const p2hMap = {};

  p2h.forEach(function(x) {
    p2hMap[String(x.unit)] = x;
  });

  return units.map(function(u) {
    const nik = String(u.assigned_nik || '');
    const person = personnelMap[nik] || null;
    const rosterRow = rosterMap[nik] || null;
    const agg = haulAgg[String(u.unit_code)] || {
      ritase:0,
      tonase:0
    };

    return {
      unit_id:u.unit_id,
      unit_code:u.unit_code,
      unit_type:u.unit_type,
      owner:u.owner,
      status:u.status,
      operational_status:u.operational_status,
      hm_km:Number(u.hm_km || 0),
      availability_pct:Number(u.availability_pct || 0),
      achievement_pct:Number(u.achievement_pct || 0),

      tonase_today:Number(agg.tonase || 0),
      ritase_today:Number(agg.ritase || 0),

      shift:rosterRow
        ? rosterStatusToShift_(rosterRow.roster_status)
        : '',

      roster_status:rosterRow
        ? String(rosterRow.roster_status || '')
        : '',

      assigned_nik:nik,

      assigned_name:person
        ? person.name
        : String(u.assigned_name || ''),

      location:u.location,
      active:toBool_(u.active),
      updated_at:u.updated_at,

      p2h_today:!!p2hMap[String(u.unit_code)],

      p2h_finding_count:
        p2hMap[String(u.unit_code)]
          ? Number(
              p2hMap[String(u.unit_code)].finding_count || 0
            )
          : 0
    };
  });
}


function buildManpowerControl_(dateKey, selectedShift) {
  const people = readObjects_(FJB_SHEETS.PERSONNEL)
    .filter(function(p) {
      return String(p.status).toUpperCase() === 'ACTIVE';
    });

  const roster = getRosterRowsForDate_(dateKey);
  const rosterMap = {};

  roster.forEach(function(r) {
    rosterMap[String(r.nik)] = r;
  });

  const targetRoster =
    String(selectedShift || 'DAY').toUpperCase() === 'NIGHT'
      ? 'N'
      : 'D';

  const definitions = [
    {
      key:'DRIVER DT',
      match:function(p) {
        return String(p.category)
          .toUpperCase()
          .indexOf('DRIVER') >= 0;
      }
    },
    {
      key:'MECHANIC',
      match:function(p) {
        const c = String(p.category).toUpperCase();

        return c.indexOf('MECHANIC') >= 0 ||
          c.indexOf('MEKANIK') >= 0;
      }
    },
    {
      key:'GL / PENGAWAS',
      match:function(p) {
        return String(p.category)
          .toUpperCase()
          .indexOf('GL') >= 0;
      }
    },
    {
      key:'ADMIN OPS',
      match:function(p) {
        return String(p.category)
          .toUpperCase()
          .indexOf('ADMIN') >= 0;
      }
    }
  ];

  const groups = definitions.map(function(def) {
    const members = people.filter(def.match);

    const actual = members.filter(function(p) {
      const r = rosterMap[String(p.nik)];

      return r &&
        String(r.roster_status).toUpperCase() === targetRoster;
    }).length;

    const plan = members.length;
    const gap = actual - plan;

    return {
      position:def.key,
      plan:plan,
      actual:actual,
      gap:gap,
      status:gap < 0 ? 'SHORTAGE' : 'OK'
    };
  });

  const replacement = people
    .filter(function(p) {
      const c = String(p.category).toUpperCase();

      if (
        c.indexOf('DRIVER') < 0 &&
        c.indexOf('MECHANIC') < 0 &&
        c.indexOf('MEKANIK') < 0
      ) {
        return false;
      }

      const r = rosterMap[String(p.nik)];
      const status = r
        ? String(r.roster_status).toUpperCase()
        : '';

      return status !== targetRoster &&
        status !== 'SK' &&
        status !== 'CT';
    })
    .slice(0, 10)
    .map(function(p) {
      const r = rosterMap[String(p.nik)];

      return {
        nik:String(p.nik),
        name:p.name,
        category:p.category,
        position:p.position,
        assigned_unit:p.assigned_unit,
        roster_status:r
          ? String(r.roster_status || '')
          : ''
      };
    });

  return {
    groups:groups,
    replacement:replacement
  };
}


/**
 * Database source for Dashboard, Daily Control, Unit Monitoring and Manpower.
 */
function apiGetOperationalSnapshotJson(
  token,
  dateValue,
  shiftValue
) {
  requireSession_(token, ['ADMIN','GL']);

  const requestedDate =
    toIsoDate_(dateValue) ||
    Utilities.formatDate(
      new Date(),
      getTimezone_(),
      'yyyy-MM-dd'
    );

  const shift =
    String(shiftValue || 'DAY')
      .trim()
      .toUpperCase() === 'NIGHT'
        ? 'NIGHT'
        : 'DAY';

  /*
   * IMPORTANT PERFORMANCE RULE:
   * Each database sheet below is read only ONCE.
   */
  const configRows =
    readObjects_(FJB_SHEETS.CONFIG);

  const unitsAll =
    readObjects_(FJB_SHEETS.UNITS)
      .filter(function(x) {
        return toBool_(x.active);
      });

  const personnelAll =
    readObjects_(FJB_SHEETS.PERSONNEL);

  const rosterAll =
    readObjects_(FJB_SHEETS.ROSTER);

  // Transaction sheets are bounded to recent records for speed.
  const haulingRecent =
    readObjectsTail_(
      FJB_SHEETS.HAULING,
      5000
    );

  const p2hRecent =
    readObjectsTail_(
      FJB_SHEETS.P2H,
      2500
    );

  const maintenanceRecent =
    readObjectsTail_(
      FJB_SHEETS.MAINTENANCE,
      800
    );

  const config = {};
  configRows.forEach(function(r) {
    config[String(r.key)] = r.value;
  });

  /*
   * Determine operational date.
   * If current/requested date has no transaction yet,
   * automatically use latest available transaction date <= requested date.
   *
   * This avoids a dashboard that looks disconnected just because
   * today has not received Hauling/P2H input yet.
   */
  const activityDates = [];

  haulingRecent.forEach(function(x) {
    const d = toIsoDate_(x.date);
    if (d && d <= requestedDate) {
      activityDates.push(d);
    }
  });

  p2hRecent.forEach(function(x) {
    const d = toIsoDate_(x.date);
    if (d && d <= requestedDate) {
      activityDates.push(d);
    }
  });

  const requestedHasActivity =
    activityDates.indexOf(requestedDate) >= 0;

  let effectiveDate = requestedDate;

  if (!requestedHasActivity && activityDates.length) {
    effectiveDate = activityDates.sort().slice(-1)[0];
  }

  // Current effective-date slices.
  const rosterToday =
    rosterAll.filter(function(x) {
      return toIsoDate_(x.date) === effectiveDate;
    });

  const haulingToday =
    haulingRecent.filter(function(x) {
      return toIsoDate_(x.date) === effectiveDate;
    });

  const p2hToday =
    p2hRecent.filter(function(x) {
      return toIsoDate_(x.date) === effectiveDate;
    });

  // 7-day trend.
  const endDate =
    parseIsoDate_(effectiveDate);

  const startDate =
    new Date(endDate);

  startDate.setDate(
    startDate.getDate() - 6
  );

  const startKey =
    Utilities.formatDate(
      startDate,
      getTimezone_(),
      'yyyy-MM-dd'
    );

  const haulingTrend =
    haulingRecent.filter(function(x) {
      const d = toIsoDate_(x.date);

      return d &&
        d >= startKey &&
        d <= effectiveDate;
    });

  /*
   * Build maps only once.
   */
  const personnelMap = {};

  personnelAll.forEach(function(p) {
    personnelMap[String(p.nik)] = p;
  });

  const rosterMap = {};

  rosterToday.forEach(function(r) {
    rosterMap[String(r.nik)] = r;
  });

  const haulAgg = {};

  haulingToday.forEach(function(h) {
    const key = String(h.hauler || '');

    if (!haulAgg[key]) {
      haulAgg[key] = {
        ritase:0,
        tonase:0
      };
    }

    haulAgg[key].ritase += 1;
    haulAgg[key].tonase +=
      Number(h.net_ton || 0);
  });

  const p2hMap = {};

  p2hToday.forEach(function(x) {
    p2hMap[String(x.unit)] = x;
  });

  const units = unitsAll.map(function(u) {
    const nik =
      String(u.assigned_nik || '');

    const person =
      personnelMap[nik] || null;

    const rosterRow =
      rosterMap[nik] || null;

    const agg =
      haulAgg[String(u.unit_code)] || {
        ritase:0,
        tonase:0
      };

    return {
      unit_id:u.unit_id,
      unit_code:u.unit_code,
      unit_type:u.unit_type,
      owner:u.owner,
      status:u.status,
      operational_status:
        u.operational_status,
      hm_km:Number(u.hm_km || 0),
      availability_pct:
        Number(u.availability_pct || 0),
      achievement_pct:
        Number(u.achievement_pct || 0),

      tonase_today:
        Number(agg.tonase || 0),

      ritase_today:
        Number(agg.ritase || 0),

      shift:rosterRow
        ? rosterStatusToShift_(
            rosterRow.roster_status
          )
        : '',

      roster_status:rosterRow
        ? String(
            rosterRow.roster_status || ''
          )
        : '',

      assigned_nik:nik,

      assigned_name:person
        ? person.name
        : String(u.assigned_name || ''),

      location:u.location,
      active:toBool_(u.active),
      updated_at:u.updated_at,

      p2h_today:
        !!p2hMap[
          String(u.unit_code)
        ],

      p2h_finding_count:
        p2hMap[String(u.unit_code)]
          ? Number(
              p2hMap[
                String(u.unit_code)
              ].finding_count || 0
            )
          : 0
    };
  });

  /*
   * Manpower control using the SAME personnel+roster reads.
   */
  const targetRoster =
    shift === 'NIGHT'
      ? 'N'
      : 'D';

  const activePeople =
    personnelAll.filter(function(p) {
      return String(p.status)
        .toUpperCase() === 'ACTIVE';
    });

  const manpowerDefs = [
    {
      key:'DRIVER DT',
      match:function(p) {
        return String(p.category)
          .toUpperCase()
          .indexOf('DRIVER') >= 0;
      }
    },
    {
      key:'MECHANIC',
      match:function(p) {
        const c =
          String(p.category)
            .toUpperCase();

        return c.indexOf('MECHANIC') >= 0 ||
          c.indexOf('MEKANIK') >= 0;
      }
    },
    {
      key:'GL / PENGAWAS',
      match:function(p) {
        return String(p.category)
          .toUpperCase()
          .indexOf('GL') >= 0;
      }
    },
    {
      key:'ADMIN OPS',
      match:function(p) {
        return String(p.category)
          .toUpperCase()
          .indexOf('ADMIN') >= 0;
      }
    }
  ];

  const manpowerGroups =
    manpowerDefs.map(function(def) {
      const members =
        activePeople.filter(def.match);

      const actual =
        members.filter(function(p) {
          const r =
            rosterMap[String(p.nik)];

          return r &&
            String(r.roster_status)
              .toUpperCase() ===
              targetRoster;
        }).length;

      const plan = members.length;

      return {
        position:def.key,
        plan:plan,
        actual:actual,
        gap:actual - plan,
        status:
          actual < plan
            ? 'SHORTAGE'
            : 'OK'
      };
    });

  const replacement =
    activePeople
      .filter(function(p) {
        const c =
          String(p.category)
            .toUpperCase();

        if (
          c.indexOf('DRIVER') < 0 &&
          c.indexOf('MECHANIC') < 0 &&
          c.indexOf('MEKANIK') < 0
        ) {
          return false;
        }

        const r =
          rosterMap[String(p.nik)];

        const status =
          r
            ? String(
                r.roster_status || ''
              ).toUpperCase()
            : '';

        return status !== targetRoster &&
          status !== 'SK' &&
          status !== 'CT';
      })
      .slice(0, 10)
      .map(function(p) {
        const r =
          rosterMap[String(p.nik)];

        return {
          nik:String(p.nik),
          name:p.name,
          category:p.category,
          position:p.position,
          assigned_unit:p.assigned_unit,
          roster_status:r
            ? String(
                r.roster_status || ''
              )
            : ''
        };
      });

  return JSON.stringify({
    ok:true,

    requested_date:requestedDate,
    effective_date:effectiveDate,
    fallback_used:
      effectiveDate !== requestedDate,

    shift:shift,

    target_availability:
      Number(
        config.TARGET_AVAILABILITY || 85
      ),

    units:units.map(cleanObject_),

    hauling_today:
      haulingToday.map(cleanObject_),

    hauling_trend:
      haulingTrend.map(cleanObject_),

    p2h_today:
      p2hToday.map(cleanObject_),

    maintenance:
      maintenanceRecent.map(cleanObject_),

    manpower:{
      groups:manpowerGroups,
      replacement:replacement
    },

    source_counts:{
      units:unitsAll.length,
      personnel:personnelAll.length,
      roster:rosterAll.length,
      hauling_recent:
        haulingRecent.length,
      p2h_recent:
        p2hRecent.length,
      maintenance_recent:
        maintenanceRecent.length
    }
  });
}



/**
 * Jalankan langsung dari Apps Script dropdown.
 * Menampilkan database yang BENAR-BENAR sedang dipakai script.
 */
function diagnoseFJBDatabaseBinding() {
  const db = getDb_();

  const result = {
    version:FJB_VERSION,
    spreadsheet_id:db.getId(),
    spreadsheet_name:db.getName(),
    spreadsheet_timezone:
      db.getSpreadsheetTimeZone(),
    configured_property:
      PropertiesService
        .getScriptProperties()
        .getProperty(
          'FJB_SPREADSHEET_ID'
        ),
    sheets:{}
  };

  Object.keys(FJB_SCHEMA)
    .forEach(function(name) {
      const sh = db.getSheetByName(name);

      result.sheets[name] =
        sh
          ? Math.max(
              0,
              sh.getLastRow()-1
            )
          : 'MISSING';
    });

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


function apiDatabaseConnectionStatusJson(token) {
  const session = requireSession_(token);
  const db = getDb_();

  const result = {
    ok:true,

    database:{
      id:db.getId(),
      name:db.getName(),
      timezone:
        db.getSpreadsheetTimeZone()
    },

    role:session.role,
    nik:session.nik,
    assigned_unit:
      session.assigned_unit,

    counts:{}
  };

  Object.keys(FJB_SCHEMA)
    .forEach(function(name) {
      const sh = getSheet_(name);

      result.counts[name] =
        Math.max(
          0,
          sh.getLastRow() - 1
        );
    });

  return JSON.stringify(result);
}


/* ============================================================
 * 16. REPORT DATA API
 * ============================================================
 */

/**
 * Frontend dapat mengambil data report dalam satu call.
 * Khusus ADMIN / GL.
 */
function apiGetReportBundle(token, filters) {
  requireSession_(token, ['ADMIN','GL']);

  filters = filters || {};

  return {
    ok:true,
    hauling:filterCommon_(
      readObjectsFast_(FJB_SHEETS.HAULING),
      filters,
      {
        date:'date',
        shift:'shift',
        unit:'hauler',
        nik:'input_by_nik'
      }
    ).map(cleanObject_),

    units:filterCommon_(
      readObjectsFast_(FJB_SHEETS.UNITS),
      filters,
      {
        unit:'unit_code',
        nik:'assigned_nik'
      }
    ).map(cleanObject_),

    roster:filterCommon_(
      readObjectsFast_(FJB_SHEETS.ROSTER),
      filters,
      {
        date:'date',
        unit:'assigned_unit',
        nik:'nik'
      }
    ).map(cleanObject_),

    washing:filterCommon_(
      readObjectsFast_(FJB_SHEETS.WASHING),
      filters,
      {
        date:'plan_date',
        unit:'unit',
        nik:'pic_nik'
      }
    ).map(cleanObject_),

    p2h:filterCommon_(
      readObjectsFast_(FJB_SHEETS.P2H),
      filters,
      {
        date:'date',
        shift:'shift',
        unit:'unit',
        nik:'input_by_nik'
      }
    ).map(cleanObject_),

    maintenance:filterCommon_(
      readObjectsFast_(FJB_SHEETS.MAINTENANCE),
      filters,
      {
        date:'date',
        unit:'unit',
        nik:'mechanic_nik'
      }
    ).map(cleanObject_)
  };
}






/* ============================================================
 * 15D. ASSIGNMENT CONSISTENCY + SYSTEM SELF TEST
 * ============================================================
 */

function applyPersonnelAssignment_(nik, unitCode, actorSession) {
  nik = String(nik || '').trim();
  unitCode = String(unitCode || '').trim();

  const person = findOne_(FJB_SHEETS.PERSONNEL, 'nik', nik);
  if (!person) throw new Error('Personnel tidak ditemukan: ' + nik);

  let unit = null;
  if (unitCode) {
    unit = findOne_(FJB_SHEETS.UNITS, 'unit_code', unitCode);
    if (!unit) throw new Error('Unit tidak ditemukan: ' + unitCode);
  }

  const now = new Date();
  const oldUnit = String(person.assigned_unit || '');

  // Jika unit baru sebelumnya dimiliki orang lain, lepaskan master personnel lama.
  if (unit && unit.assigned_nik && String(unit.assigned_nik) !== nik) {
    const previousNik = String(unit.assigned_nik);
    const previousPerson = findOne_(FJB_SHEETS.PERSONNEL, 'nik', previousNik);
    if (previousPerson && String(previousPerson.assigned_unit || '') === unitCode) {
      updateRowObject_(FJB_SHEETS.PERSONNEL, previousPerson._row, {
        assigned_unit:'',
        updated_at:now
      });
      syncActiveSessionsForNik_(previousNik, { assigned_unit:'' });
    }
  }

  // Lepaskan semua unit lama dari personnel ini.
  readObjects_(FJB_SHEETS.UNITS)
    .filter(function(x) {
      return String(x.assigned_nik || '') === nik &&
        String(x.unit_code || '') !== unitCode;
    })
    .forEach(function(x) {
      updateRowObject_(FJB_SHEETS.UNITS, x._row, {
        assigned_nik:'',
        assigned_name:'',
        updated_at:now
      });
    });

  updateRowObject_(FJB_SHEETS.PERSONNEL, person._row, {
    assigned_unit:unitCode,
    updated_at:now
  });

  if (unit) {
    updateRowObject_(FJB_SHEETS.UNITS, unit._row, {
      assigned_nik:nik,
      assigned_name:person.name,
      updated_at:now
    });
  }

  syncActiveSessionsForNik_(nik, {
    name:person.name,
    position:person.position,
    assigned_unit:unitCode
  });

  if (actorSession) {
    appendHistory_(
      actorSession,
      'Roster',
      'UPDATE',
      person.name,
      'Assign batangan ' + (oldUnit || '-') + ' → ' + (unitCode || '-'),
      {
        nik:nik,
        name:person.name,
        before_unit:oldUnit,
        after_unit:unitCode
      }
    );
  }

  return {
    nik:nik,
    name:person.name,
    before_unit:oldUnit,
    assigned_unit:unitCode
  };
}


function runFJBSelfTest() {
  const startedAt = new Date();
  const result = performFJBSelfTest_();

  const lines = [
    'FJB SYSTEM SELF-TEST',
    'Version: ' + FJB_VERSION,
    'Database: ' + getDb_().getName(),
    'Spreadsheet ID: ' + getDb_().getId(),
    '',
    'PASS: ' + result.pass,
    'WARN: ' + result.warn,
    'FAIL: ' + result.fail,
    ''
  ];

  result.checks.forEach(function(x) {
    lines.push(
      '[' + x.status + '] ' +
      x.name +
      (x.detail ? ' — ' + x.detail : '')
    );
  });

  const finishedAt = new Date();

  lines.push('');
  lines.push(
    'Elapsed: ' +
    (finishedAt.getTime() - startedAt.getTime()) +
    ' ms'
  );

  const output = lines.join('\n');

  // Selalu log agar hasil dapat dilihat walaupun popup tersedia.
  console.log(output);

  // Popup hanya jika konteks mendukung Spreadsheet UI.
  safeUiAlert_(
    'FJB SYSTEM SELF-TEST',
    output
  );

  return result;
}


function apiHealthCheck(token) {
  requireSession_(token, ['ADMIN']);
  return performFJBSelfTest_();
}


function performFJBSelfTest_() {
  const checks = [];
  function add(status, name, detail) {
    checks.push({ status:status, name:name, detail:detail || '' });
  }

  const ss = getDb_();

  Object.keys(FJB_SCHEMA).forEach(function(sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) {
      add('FAIL', 'Sheet ' + sheetName, 'Tidak ditemukan');
      return;
    }

    const expected = FJB_SCHEMA[sheetName];
    const actual = sh.getRange(1,1,1,expected.length).getValues()[0].map(String);
    if (actual.join('|') === expected.join('|')) {
      add('PASS', 'Schema ' + sheetName, Math.max(0, sh.getLastRow()-1) + ' row');
    } else {
      add('FAIL', 'Schema ' + sheetName, 'Header tidak sesuai');
    }
  });

  const people = readObjects_(FJB_SHEETS.PERSONNEL);
  const peopleMap = {};
  people.forEach(function(x) { peopleMap[String(x.nik)] = x; });

  const units = readObjects_(FJB_SHEETS.UNITS);
  const unitMap = {};
  units.forEach(function(x) { unitMap[String(x.unit_code)] = x; });

  readObjects_(FJB_SHEETS.USERS).forEach(function(u) {
    if (!peopleMap[String(u.nik)]) {
      add('FAIL', 'User ' + u.nik, 'Tidak punya Personnel');
    }
  });

  people.forEach(function(p) {
    const assigned = String(p.assigned_unit || '');
    if (assigned && assigned.indexOf('DT-') === 0 && !unitMap[assigned]) {
      add('FAIL', 'Batangan ' + p.nik, assigned + ' tidak ada di Master Unit');
    }
  });

  readObjects_(FJB_SHEETS.ROSTER).forEach(function(r) {
    if (!peopleMap[String(r.nik)]) add('FAIL', 'Roster ' + r.roster_id, 'NIK tidak ada');
    if (['D','N','OFF','CT','SK'].indexOf(String(r.roster_status)) < 0) {
      add('FAIL', 'Roster ' + r.roster_id, 'Status invalid ' + r.roster_status);
    }
    const month = monthKeyFromValue_(r.date);
    if (month && normalizeMonthKey_(r.month) !== month) {
      add('WARN', 'Roster month ' + r.roster_id, String(r.month) + ' vs ' + month);
    }
  });

  readObjects_(FJB_SHEETS.WASHING).forEach(function(w) {
    if (!unitMap[String(w.unit)]) add('FAIL', 'Washing ' + w.washing_id, 'Unit tidak ada');
    if (['PLAN','DONE','RESCHEDULE'].indexOf(String(w.status)) < 0) {
      add('FAIL', 'Washing ' + w.washing_id, 'Status invalid');
    }
  });

  readObjects_(FJB_SHEETS.HAULING).forEach(function(h) {
    if (!unitMap[String(h.hauler)]) add('FAIL', 'Hauling ' + h.transaction_id, 'Hauler tidak ada');
    if (!unitMap[String(h.loader)]) add('FAIL', 'Hauling ' + h.transaction_id, 'Loader tidak ada');
  });

  const p2hHeaders = {};
  readObjects_(FJB_SHEETS.P2H).forEach(function(p) {
    p2hHeaders[String(p.p2h_id)] = p;
    if (!unitMap[String(p.unit)]) add('FAIL', 'P2H ' + p.p2h_id, 'Unit tidak ada');
  });
  readObjects_(FJB_SHEETS.P2H_DETAIL).forEach(function(d) {
    if (!p2hHeaders[String(d.p2h_id)]) add('FAIL', 'P2H Detail ' + d.detail_id, 'Header P2H tidak ada');
  });

  readObjects_(FJB_SHEETS.MAINTENANCE).forEach(function(m) {
    if (!unitMap[String(m.unit)]) add('FAIL', 'Maintenance ' + m.maintenance_id, 'Unit tidak ada');
    if (!peopleMap[String(m.mechanic_nik)]) add('WARN', 'Maintenance ' + m.maintenance_id, 'Mechanic NIK tidak ada');
  });

  const fail = checks.filter(function(x){ return x.status === 'FAIL'; }).length;
  const warn = checks.filter(function(x){ return x.status === 'WARN'; }).length;
  const pass = checks.filter(function(x){ return x.status === 'PASS'; }).length;

  return {
    ok:fail === 0,
    version:FJB_VERSION,
    pass:pass,
    warn:warn,
    fail:fail,
    checks:checks
  };
}


/* ============================================================
 * 16A. ASSIGN BATANGAN
 * ============================================================
 */

function apiAssignBatangan(token, nik, unitCode) {
  return withWriteLock_('apiAssignBatangan', function() {
      const session = requireSession_(token, ['ADMIN','GL']);
      const result = applyPersonnelAssignment_(nik, unitCode, session);
      return { ok:true, nik:result.nik, name:result.name, assigned_unit:result.assigned_unit };

  });
}


/* ============================================================
 * 16B. SYSTEM SETTINGS + CLIENT ACTION LOG
 * ============================================================
 */

function apiGetSystemSettings(token) {
  requireSession_(token, ['ADMIN']);

  const defaults = {
    APP_NAME:'FJB Operations Control System',
    SITE:'BIB',
    TIMEZONE:'Asia/Jakarta',
    DEFAULT_SHIFT:'DAY',
    TARGET_AVAILABILITY:'85',
    AUDIT_LOG:'TRUE',
    SOFT_DELETE:'TRUE',
    P2H_ALERT:'TRUE',
    LEGACY_SYNC:'FALSE',
    EMAIL_SUMMARY:'FALSE'
  };

  const rows = readObjectsFast_(FJB_SHEETS.CONFIG);
  const out = {};

  rows.forEach(function(r) {
    out[String(r.key)] = r.value;
  });

  Object.keys(defaults).forEach(function(k) {
    if (out[k] === undefined || out[k] === '') {
      out[k] = defaults[k];
    }
  });

  return {
    ok:true,
    settings:out
  };
}


function apiSaveSystemSettings(token, settings) {
  return withWriteLock_('apiSaveSystemSettings', function() {
      const session = requireSession_(token, ['ADMIN']);
      settings = settings || {};

      const allowed = [
        'APP_NAME',
        'SITE',
        'TIMEZONE',
        'DEFAULT_SHIFT',
        'TARGET_AVAILABILITY',
        'AUDIT_LOG',
        'SOFT_DELETE',
        'P2H_ALERT',
        'LEGACY_SYNC',
        'EMAIL_SUMMARY'
      ];

      const sheet = getSheet_(FJB_SHEETS.CONFIG);
      const rows = readObjects_(FJB_SHEETS.CONFIG);

      allowed.forEach(function(key) {
        if (settings[key] === undefined) return;

        const value = String(settings[key]);
        const existing = rows.find(function(r) {
          return String(r.key) === key;
        });

        if (existing) {
          updateRowObject_(
            FJB_SHEETS.CONFIG,
            existing._row,
            { value:value }
          );
        } else {
          appendObject_(
            FJB_SHEETS.CONFIG,
            {
              key:key,
              value:value,
              description:'System setting'
            }
          );
        }
      });

      if (settings.TIMEZONE) {
        const tz = String(settings.TIMEZONE);
        PropertiesService.getScriptProperties().setProperty('FJB_TIMEZONE', tz);
        try {
          getDb_().setSpreadsheetTimeZone(tz);
        } catch (err) {
          // Spreadsheet timezone update is best-effort.
        }
      }

      appendHistory_(
        session,
        'Settings',
        'UPDATE',
        'SYSTEM_CONFIG',
        'System settings diperbarui',
        settings
      );

      return { ok:true };

  });
}


function apiLogClientAction(
  token,
  module,
  activity,
  entity,
  summary,
  payload
) {
  return withWriteLock_('apiLogClientAction', function() {
      const session = requireSession_(token);

      appendHistory_(
        session,
        String(module || 'UI'),
        String(activity || 'CREATE'),
        String(entity || ''),
        String(summary || ''),
        payload || {}
      );

      return { ok:true };

  });
}



/* ============================================================
 * 16F. DATE SYNC DIAGNOSTIC
 * ============================================================
 */

/**
 * Read-only diagnostic.
 *
 * Compares the actual Sheet Date value with the serialized date
 * received by the Web App.
 */
function apiDateSyncDebugJson(token, nik, month) {
  const session = requireSession_(token);

  nik = String(nik || session.nik || '').trim();

  const monthKey =
    normalizeMonthKey_(month) ||
    Utilities.formatDate(
      new Date(),
      getTimezone_(),
      'yyyy-MM'
    );

  const roster = readObjects_(FJB_SHEETS.ROSTER)
    .filter(function(x) {
      return String(x.nik) === nik &&
        rowMonthMatches_(
          x,
          monthKey,
          'date'
        );
    })
    .slice(0, 5)
    .map(function(x) {
      return {
        nik:String(x.nik),
        raw_date_type:
          x.date instanceof Date
            ? 'Date'
            : typeof x.date,
        local_date:
          x.date instanceof Date
            ? Utilities.formatDate(
                x.date,
                getTimezone_(),
                'yyyy-MM-dd'
              )
            : toIsoDate_(x.date),
        serialized:
          cleanObject_(x).date,
        status:String(x.roster_status || '')
      };
    });

  const assignedUnit =
    session.role === 'OPERATOR'
      ? String(session.assigned_unit || '')
      : '';

  const washing = readObjects_(FJB_SHEETS.WASHING)
    .filter(function(x) {
      if (
        assignedUnit &&
        String(x.unit) !== assignedUnit
      ) {
        return false;
      }

      return rowMonthMatches_(
        x,
        monthKey,
        'plan_date'
      );
    })
    .slice(0, 5)
    .map(function(x) {
      const c = cleanObject_(x);

      return {
        unit:String(x.unit),
        local_plan_date:
          x.plan_date instanceof Date
            ? Utilities.formatDate(
                x.plan_date,
                getTimezone_(),
                'yyyy-MM-dd'
              )
            : toIsoDate_(x.plan_date),

        serialized_plan_date:
          c.plan_date,

        serialized_actual_date:
          c.actual_date,

        serialized_reschedule_date:
          c.reschedule_date,

        status:String(x.status || '')
      };
    });

  return JSON.stringify({
    ok:true,
    timezone:getTimezone_(),
    month:monthKey,
    nik:nik,
    assigned_unit:
      session.assigned_unit,
    roster:roster,
    washing:washing
  });
}


/**
 * Editor function, no web token needed.
 * Default: JOKO SUSILO / 3101027.
 */
function diagnoseFJBDateSerialization() {
  const nik = '3101027';
  const month = '2026-08';

  const roster = readObjects_(FJB_SHEETS.ROSTER)
    .filter(function(x) {
      return String(x.nik) === nik &&
        rowMonthMatches_(x, month, 'date');
    })
    .slice(0, 5)
    .map(function(x) {
      return {
        sheet_local:
          Utilities.formatDate(
            x.date,
            getTimezone_(),
            'yyyy-MM-dd'
          ),
        web_serialized:
          cleanObject_(x).date,
        status:x.roster_status
      };
    });

  const washing = readObjects_(FJB_SHEETS.WASHING)
    .filter(function(x) {
      return String(x.unit) === 'DT-008' &&
        rowMonthMatches_(
          x,
          month,
          'plan_date'
        );
    })
    .map(function(x) {
      return {
        unit:x.unit,
        sheet_plan:
          Utilities.formatDate(
            x.plan_date,
            getTimezone_(),
            'yyyy-MM-dd'
          ),
        web_plan:
          cleanObject_(x).plan_date,
        web_actual:
          cleanObject_(x).actual_date,
        status:x.status
      };
    });

  const result = {
    version:FJB_VERSION,
    timezone:getTimezone_(),
    roster:roster,
    washing:washing
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}





/* ============================================================
 * 16G. INITIAL APP CACHE — V2.4
 * ============================================================
 *
 * One clear login-data path:
 * Login -> Initial Cache -> Render.
 *
 * The initial cache deliberately contains:
 * - stable master data needed by UI;
 * - previous/current/next month Roster & Washing;
 * - recent operational transactions;
 * - recent History;
 * - Users/Settings for ADMIN.
 *
 * Older / future periods are loaded only when selected, using
 * existing module APIs. This keeps login fast without changing
 * the business flow.
 */


function shiftMonthKey_(monthKey, offset) {
  const m = String(monthKey || '')
    .match(/^(\d{4})-(\d{2})$/);

  if (!m) {
    throw new Error(
      'Month key tidak valid: ' + monthKey
    );
  }

  const d = new Date(
    Number(m[1]),
    Number(m[2]) - 1 + Number(offset || 0),
    1
  );

  return Utilities.formatDate(
    d,
    getTimezone_(),
    'yyyy-MM'
  );
}


function apiInitialAppCacheJson(token) {
  ensureOperationalExtensionV260_(false);

  const session = requireSession_(token);
  const role = String(session.role || '');
  const started = Date.now();

  const today = Utilities.formatDate(
    new Date(),
    getTimezone_(),
    'yyyy-MM-dd'
  );

  const currentMonth = today.substring(0, 7);

  const hotMonths = [
    shiftMonthKey_(currentMonth, -1),
    currentMonth,
    shiftMonthKey_(currentMonth, 1)
  ];

  const hotMonthMap = {};
  hotMonths.forEach(function(m) {
    hotMonthMap[m] = true;
  });

  const rangeFrom =
    hotMonths[0] + '-01';

  const rangeTo = today;

  /*
   * Small/stable masters.
   */
  const configRows =
    readObjectsFast_(FJB_SHEETS.CONFIG);

  const config = {};

  configRows.forEach(function(r) {
    config[String(r.key)] = r.value;
  });

  const personnelAll =
    readObjectsFast_(FJB_SHEETS.PERSONNEL);

  const personnelMap = {};

  personnelAll.forEach(function(p) {
    personnelMap[String(p.nik)] = p;
  });

  const self =
    personnelMap[String(session.nik)] ||
    null;

  const units =
    readObjectsFast_(FJB_SHEETS.UNITS)
      .filter(function(x) {
        return toBool_(x.active);
      })
      .map(cleanObject_);

  const products =
    readObjectsFast_(FJB_SHEETS.PRODUCT)
      .filter(function(x) {
        return toBool_(x.active);
      })
      .map(cleanObject_);

  const locations =
    readObjectsFast_(FJB_SHEETS.LOCATION)
      .filter(function(x) {
        return toBool_(x.active);
      })
      .map(cleanObject_);

  const options =
    readObjectsFast_(FJB_SHEETS.OPTIONS)
      .filter(function(x) {
        return toBool_(x.active);
      })
      .map(cleanObject_);

  const p2hMaster =
    readObjectsFast_(FJB_SHEETS.P2H_MASTER)
      .filter(function(x) {
        return toBool_(x.active);
      })
      .sort(function(a,b) {
        return Number(a.sort_order) -
          Number(b.sort_order);
      })
      .map(cleanObject_);

  const owners =
    readObjectsFast_(
      FJB_SHEETS.OWNER
    )
    .map(cleanObject_);

  const aliases =
    readObjectsFast_(
      FJB_SHEETS.UNIT_ALIAS
    )
    .map(cleanObject_);

  /*
   * Roster:
   * bounded tail + 3 hot months.
   * Avoid reading years of historical roster at every login.
   */
  let roster =
    readObjectsTailFast_(
      FJB_SHEETS.ROSTER,
      12000
    )
    .filter(function(x) {
      const d = toIsoDate_(x.date);

      return d &&
        hotMonthMap[
          d.substring(0,7)
        ];
    });

  if (
    role === 'OPERATOR' ||
    role === 'MECHANIC'
  ) {
    roster = roster.filter(function(x) {
      return String(x.nik) ===
        String(session.nik);
    });
  }

  roster.forEach(function(x) {
    const p =
      personnelMap[String(x.nik)];

    if (!p) return;

    x.name = p.name;
    x.category = p.category;
    x.position = p.position;
    x.assigned_unit = p.assigned_unit;
  });

  /*
   * Washing:
   * only previous/current/next month.
   */
  let washing = [];

  if (
    role === 'ADMIN' ||
    role === 'GL' ||
    role === 'OPERATOR'
  ) {
    washing =
      readObjectsTailFast_(
        FJB_SHEETS.WASHING,
        4000
      )
      .filter(function(x) {
        const d =
          toIsoDate_(x.plan_date);

        return d &&
          hotMonthMap[
            d.substring(0,7)
          ];
      });

    if (role === 'OPERATOR') {
      washing = washing.filter(function(x) {
        return String(x.unit) ===
          String(
            session.assigned_unit || ''
          );
      });
    }
  }

  /*
   * Recent operational data only.
   * Enough for Dashboard + current period pages.
   */
  let hauling = [];
  let p2h = [];
  let maintenance = [];
  let history = [];
  let users = [];

  if (
    role === 'ADMIN' ||
    role === 'GL' ||
    role === 'OPERATOR'
  ) {
    hauling =
      readObjectsTailFast_(
        FJB_SHEETS.HAULING,
        1400
      )
      .filter(function(x) {
        const d = toIsoDate_(x.date);

        return d &&
          d >= rangeFrom &&
          d <= rangeTo;
      });

    p2h =
      readObjectsTailFast_(
        FJB_SHEETS.P2H,
        1000
      )
      .filter(function(x) {
        const d = toIsoDate_(x.date);

        return d &&
          d >= rangeFrom &&
          d <= rangeTo;
      });

    if (role === 'OPERATOR') {
      hauling = hauling.filter(function(x) {
        return String(x.input_by_nik) ===
          String(session.nik);
      });

      p2h = p2h.filter(function(x) {
        return String(x.input_by_nik) ===
          String(session.nik);
      });
    }
  }

  if (
    role === 'ADMIN' ||
    role === 'GL' ||
    role === 'MECHANIC'
  ) {
    maintenance =
      readObjectsTailFast_(
        FJB_SHEETS.MAINTENANCE,
        600
      )
      .filter(function(x) {
        const d = toIsoDate_(x.date);

        return d &&
          d >= rangeFrom &&
          d <= rangeTo;
      });

    if (role === 'MECHANIC') {
      maintenance =
        maintenance.filter(function(x) {
          return String(
            x.mechanic_nik
          ) === String(session.nik);
        });
    }
  }

  if (
    role === 'ADMIN' ||
    role === 'GL'
  ) {
    history =
      readObjectsTailFast_(
        FJB_SHEETS.HISTORY,
        300
      )
      .filter(function(x) {
        const d =
          toIsoDate_(x.timestamp);

        return d &&
          d >= rangeFrom &&
          d <= rangeTo;
      });
  }

  if (role === 'ADMIN') {
    users =
      readObjectsFast_(FJB_SHEETS.USERS)
        .map(function(u) {
          const p =
            personnelMap[
              String(u.nik)
            ] || {};

          return {
            user_id:u.user_id,
            nik:String(u.nik),
            name:p.name || '',
            category:p.category || '',
            position:p.position || '',
            role:resolveRole_(u, p),
            role_override:
              String(
                u.role_override || ''
              ),
            status:u.status,
            must_change_password:
              toBool_(
                u.must_change_password
              ),
            last_login_at:
              serializeValue_(
                u.last_login_at,
                'last_login_at'
              )
          };
        });
  }

  return JSON.stringify({
    ok:true,

    elapsed_ms:
      Date.now() - started,

    cache_stamp:
      getFastViewStamp_(),

    server_time:
      new Date().toISOString(),

    coverage:{
      hot_months:hotMonths,
      range_from:rangeFrom,
      range_to:rangeTo
    },

    app:{
      name:
        config.APP_NAME ||
        'FJB Operations Control',
      company:
        config.COMPANY_NAME ||
        'PT. FORTUNA JAYA BERSAUDARA',
      version:FJB_VERSION
    },

    session:{
      nik:String(
        session.nik || ''
      ),
      name:String(
        session.name || ''
      ),
      role:role,
      position:String(
        session.position || ''
      ),
      assigned_unit:String(
        session.assigned_unit || ''
      ),
      menus:
        ROLE_MENUS[role] || []
    },

    personnel_self:
      self
        ? cleanObject_(self)
        : null,

    personnel:
      (
        role === 'ADMIN' ||
        role === 'GL'
      )
        ? personnelAll.map(cleanObject_)
        : (
            self
              ? [cleanObject_(self)]
              : []
          ),

    units:units,
    products:products,
    locations:locations,
    options:options,
    p2h_master:p2hMaster,

    roster:
      roster.map(cleanObject_),

    washing:
      washing.map(cleanObject_),

    hauling:
      hauling.map(cleanObject_),

    p2h:
      p2h.map(cleanObject_),

    maintenance:
      maintenance.map(cleanObject_),

    history:
      history.map(cleanObject_),

    users:users,

    masters:{
      owner:owners,
      unit_alias:aliases
    },

    settings:config
  });
}


/*
 * Deferred login audit.
 * Authentication no longer waits for this non-critical write.
 */
function apiRecordLoginAudit(token) {
  const session =
    requireSession_(token);

  try {
    const user = findOne_(
      FJB_SHEETS.USERS,
      'nik',
      String(session.nik)
    );

    withWriteLock_(
      'LOGIN_AUDIT',
      function() {
        const now = new Date();

        if (user && user._row) {
          updateRowObject_(
            FJB_SHEETS.USERS,
            user._row,
            {
              last_login_at:now,
              updated_at:now
            }
          );
        }

        appendHistory_(
          session,
          'Auth',
          'LOGIN',
          String(session.nik),
          String(session.role) +
            ' login',
          {
            position:
              session.position,
            assigned_unit:
              session.assigned_unit
          }
        );
      }
    );

    return {ok:true};

  } catch (err) {
    console.warn(
      'LOGIN_AUDIT_WARNING: ' +
      (
        err && err.message
          ? err.message
          : String(err)
      )
    );

    return {
      ok:true,
      warning:true
    };
  }
}




/* ============================================================
 * 17. GENERIC HELPERS
 * ============================================================
 */

function filterCommon_(rows, filters, mapping) {
  filters = filters || {};

  const from = filters.from
    ? toIsoDate_(filters.from)
    : '';

  const to = filters.to
    ? toIsoDate_(filters.to)
    : '';

  const shift = String(
    filters.shift || 'ALL'
  ).toUpperCase();

  const unit = String(
    filters.unit || 'ALL'
  );

  const nik = String(
    filters.nik || filters.driver_nik || 'ALL'
  );

  return rows.filter(function(x) {
    if (mapping.date) {
      const d = toIsoDate_(x[mapping.date]);

      if (from && d < from) return false;
      if (to && d > to) return false;
    }

    if (
      mapping.shift &&
      shift !== 'ALL' &&
      String(x[mapping.shift]).toUpperCase() !== shift
    ) {
      return false;
    }

    if (
      mapping.unit &&
      unit !== 'ALL' &&
      String(x[mapping.unit]) !== unit
    ) {
      return false;
    }

    if (
      mapping.nik &&
      nik !== 'ALL' &&
      String(x[mapping.nik]) !== nik
    ) {
      return false;
    }

    return true;
  });
}


function generateId_(prefix) {
  const now = new Date();

  const stamp = Utilities.formatDate(
    now,
    getTimezone_(),
    'yyyyMMdd-HHmmss'
  );

  return prefix + '-' + stamp + '-' +
    Utilities.getUuid().substring(0,8).toUpperCase();
}



function isValidIsoDate_(value) {
  const s = String(value || '').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (!m) return false;

  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);

  if (mo < 1 || mo > 12 || d < 1) {
    return false;
  }

  const maxDay = new Date(y, mo, 0).getDate();

  return d <= maxDay;
}


function parseIsoDate_(iso) {
  if (iso instanceof Date) {
    return new Date(
      iso.getFullYear(),
      iso.getMonth(),
      iso.getDate()
    );
  }

  const s = String(iso || '').trim();

  if (!isValidIsoDate_(s)) {
    throw new Error(
      'Tanggal tidak valid: ' + s
    );
  }

  const m = s.match(
    /^(\d{4})-(\d{2})-(\d{2})$/
  );

  return new Date(
    Number(m[1]),
    Number(m[2]) - 1,
    Number(m[3])
  );
}


function toIsoDate_(value) {
  if (!value) return '';

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      getTimezone_(),
      'yyyy-MM-dd'
    );
  }

  const s = String(value).trim();

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    return s.substring(0,10);
  }

  let m = s.match(
    /^(\d{2})\/(\d{2})\/(\d{4})/
  );

  if (m) {
    return m[3] + '-' + m[2] + '-' + m[1];
  }

  const d = new Date(s);

  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(
      d,
      getTimezone_(),
      'yyyy-MM-dd'
    );
  }

  return '';
}


function formatTime_(date) {
  return Utilities.formatDate(
    date || new Date(),
    getTimezone_(),
    'HH:mm:ss'
  );
}


function durationMinutes_(start, finish) {
  if (!start || !finish) return 0;

  function toMinutes(t) {
    const p = String(t).split(':');
    if (p.length < 2) return null;

    return Number(p[0]) * 60 + Number(p[1]);
  }

  const a = toMinutes(start);
  const b = toMinutes(finish);

  if (a === null || b === null) return 0;

  let diff = b - a;

  if (diff < 0) diff += 24 * 60;

  return diff;
}


function toBool_(value) {
  if (value === true) return true;
  if (value === false) return false;

  const s = String(value || '')
    .trim()
    .toUpperCase();

  return (
    s === 'TRUE' ||
    s === 'YES' ||
    s === 'Y' ||
    s === '1' ||
    s === 'ACTIVE'
  );
}


function serializeValue_(value, fieldName) {
  if (!(value instanceof Date)) {
    return value;
  }

  const field =
    String(fieldName || '')
      .trim()
      .toLowerCase();

  const tz = getTimezone_();

  /*
   * DATE-ONLY FIELDS
   *
   * NEVER use Date.toISOString() here.
   * Google Sheet date midnight in Indonesia becomes previous UTC date.
   *
   * Example:
   * Sheet: 16/08/2026 00:00 Asia/Jakarta
   * toISOString(): 2026-08-15T17:00:00.000Z  <-- WRONG DAY FOR UI
   *
   * Correct:
   * Utilities.formatDate(..., spreadsheetTimezone, 'yyyy-MM-dd')
   * => 2026-08-16
   */
  const dateOnlyFields = {
    date:true,
    join_date:true,
    plan_date:true,
    actual_date:true,
    reschedule_date:true
  };

  if (dateOnlyFields[field]) {
    return Utilities.formatDate(
      value,
      tz,
      'yyyy-MM-dd'
    );
  }

  /*
   * Month stored accidentally as Date.
   */
  if (field === 'month') {
    return Utilities.formatDate(
      value,
      tz,
      'yyyy-MM'
    );
  }

  /*
   * Time-only fields may be represented by Google Sheets
   * using the 1899/1900 base date.
   */
  const timeOnlyFields = {
    time:true,
    start_time:true,
    finish_time:true
  };

  const year =
    Number(
      Utilities.formatDate(
        value,
        tz,
        'yyyy'
      )
    );

  if (
    timeOnlyFields[field] ||
    year <= 1900
  ) {
    return Utilities.formatDate(
      value,
      tz,
      'HH:mm:ss'
    );
  }

  /*
   * Timestamp fields:
   * return local spreadsheet time WITHOUT UTC conversion.
   *
   * Frontend displays this as operational local timestamp.
   */
  return Utilities.formatDate(
    value,
    tz,
    "yyyy-MM-dd'T'HH:mm:ss"
  );
}


function cleanObject_(obj) {
  const out = {};

  Object.keys(obj || {})
    .forEach(function(k) {
      if (k === '_row') return;

      out[k] = serializeValue_(
        obj[k],
        k
      );
    });

  return out;
}


/* ============================================================
 * 18. WEB APP ENTRY POINT & REST API (GAS + VERCEL COMPATIBLE)
 * ============================================================
 *
 * Mendukung 2 mode operasi sekaligus:
 * 1. Mode Internal GAS: Diakses via Google Apps Script Web App.
 * 2. Mode External/Vercel: Diakses via HTTP POST/GET dari Vercel
 *    atau web/mobile browser pihak ketiga secara aman dan bebas CORS.
 */

function doGet(e) {
  // Jika dipanggil dari Vercel / fetch GET sebagai API
  if (e && e.parameter && (e.parameter.action || e.parameter.fn)) {
    return handleApiRequest_(e.parameter.action || e.parameter.fn, e.parameter.args);
  }

  // Health check endpoint untuk Vercel / tes koneksi
  if (e && e.parameter && e.parameter.ping === '1') {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        message: 'FJB Operations Control API is Live!',
        version: FJB_VERSION,
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Standar Apps Script Web App render template Index.html
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('FJB Operations Control')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    )
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1, viewport-fit=cover'
    );
}

function doPost(e) {
  var action = '';
  var args = [];

  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      action = body.action || body.fn || '';
      args = body.args || [];
    } catch (parseErr) {
      action = (e.parameter && (e.parameter.action || e.parameter.fn)) || '';
      args = (e.parameter && e.parameter.args) || [];
    }
  } else if (e && e.parameter) {
    action = e.parameter.action || e.parameter.fn || '';
    args = e.parameter.args || [];
  }

  return handleApiRequest_(action, args);
}

function handleApiRequest_(fnName, args) {
  try {
    if (typeof args === 'string') {
      try {
        args = JSON.parse(args);
      } catch (ex) {
        args = [args];
      }
    }
    if (!Array.isArray(args)) {
      args = (args !== undefined && args !== null) ? [args] : [];
    }

    if (!fnName) {
      return ContentService
        .createTextOutput(JSON.stringify({
          ok: false,
          error: 'ACTION_REQUIRED: Parameter action atau fn harus disertakan.'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Hanya fungsi API yang boleh dipanggil dari luar demi keamanan
    var isAllowed = typeof fnName === 'string' && (
      fnName.indexOf('api') === 0 ||
      fnName === 'setupFJBSystem' ||
      fnName === 'getSystemStatus'
    );

    var targetFn = this[fnName];
    if (!isAllowed || typeof targetFn !== 'function') {
      return ContentService
        .createTextOutput(JSON.stringify({
          ok: false,
          error: 'API_NOT_FOUND_OR_FORBIDDEN: ' + fnName
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var result = targetFn.apply(null, args);
    var output = (typeof result === 'string') ? result : JSON.stringify(result);

    return ContentService
      .createTextOutput(output)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: err.message || String(err)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

