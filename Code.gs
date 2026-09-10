/**
 * 🏰 용돈 퀘스트 (FamilyQuest) - Google Apps Script 백엔드
 *
 * 📋 스프레드시트 시트 구조:
 *   - Users     : UserId | Name | Role | PIN | Points | TotalExp | Streak
 *   - Quests    : QuestId | Title | Category | Points | XP | Icon | Desc | Period
 *   - QuestLogs : LogId | QuestId | QuestTitle | UserId | SubmittedAt | Status | PhotoUrl | Memo | ParentComment | ApprovedBy
 *   - Rewards   : RewardId | Title | Cost | Category | Icon | Desc
 *   - Inventory : InvId | UserId | RewardId | RewardTitle | AcquiredAt | Status | Icon
 *
 * 🔑 승인자 PIN은 Users 시트 PIN 컬럼에서 읽어옴
 */

const SPREADSHEET_ID = "1qEMF1AJZZzKaLm__hN_gXq_y5e95JuecwH3rFBj2uIw";
const FOLDER_ID      = "1d-KI3hOXBgnZSWlu4qHG8V_634EpWTyT";
const FOLDER_NAME    = "FamilyQuest_Uploads";
const XP_PER_LEVEL   = 2500;

// ────────────────────────────────────────────────
// 진입점 (HTML 웹앱 & REST API 하이브리드)
// ────────────────────────────────────────────────
function doGet(e) {
  if (e && e.parameter && (e.parameter.action === "getAppData" || e.parameter.api === "true" || e.parameter.format === "json")) {
    return ContentService.createTextOutput(JSON.stringify(getAppData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("용돈 퀘스트 | FamilyQuest")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var raw = e && e.postData ? e.postData.contents : "{}";
    var payload = {};
    try { payload = JSON.parse(raw); } catch(err) { payload = e.parameter || {}; }

    var action = payload.action || (e.parameter ? e.parameter.action : "");
    var data = payload.data || payload;
    var result = { success: false, error: "알 수 없는 요청입니다." };

    if (action === "getAppData") {
      result = getAppData();
    } else if (action === "completeQuest") {
      result = completeQuest(data);
    } else if (action === "approveQuest") {
      result = approveQuest(data);
    } else if (action === "rejectQuest") {
      result = rejectQuest(data);
    } else if (action === "redeemReward") {
      result = redeemReward(data);
    } else if (action === "useInventoryItem") {
      result = useInventoryItem(data);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ────────────────────────────────────────────────
// 스프레드시트 & 폴더 획득
// ────────────────────────────────────────────────
function getSS() {
  try { const a = SpreadsheetApp.getActiveSpreadsheet(); if (a) return a; } catch(e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getFolder() {
  if (FOLDER_ID) { try { return DriveApp.getFolderById(FOLDER_ID); } catch(e) {} }
  try {
    const it = DriveApp.getFoldersByName(FOLDER_NAME);
    if (it.hasNext()) return it.next();
    const f = DriveApp.createFolder(FOLDER_NAME);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    return f;
  } catch(e) { return null; }
}

// ────────────────────────────────────────────────
// 헬퍼: 시트 데이터를 {header:value} 객체 배열로 변환 (대소문자 무관)
// ────────────────────────────────────────────────
function sheetToObjects(sheet) {
  if (!sheet) return [];
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 2 || lc < 1) return [];
  const all = sheet.getRange(1, 1, lr, lc).getValues();
  const headers = all[0].map(function(h) { return String(h).trim(); });
  const result = [];
  for (var r = 1; r < all.length; r++) {
    var row = all[r];
    var hasData = row.some(function(c) { return String(c || "").trim().length > 0; });
    if (!hasData) continue;
    var obj = {};
    headers.forEach(function(h, i) {
      obj[h] = row[i];
      var clean = h.replace(/[^a-zA-Z0-9가-힣]/g, "").toLowerCase();
      if (clean) obj[clean] = row[i];
    });
    obj.__rowIndex = r + 1; // 1-based row index in sheet
    result.push(obj);
  }
  return result;
}

function getField(obj) {
  if (!obj) return "";
  for (var i = 1; i < arguments.length; i++) {
    var k = arguments[i];
    if (obj[k] !== undefined && obj[k] !== null && String(obj[k]).trim() !== "") {
      return obj[k];
    }
    var clean = String(k).replace(/[^a-zA-Z0-9가-힣]/g, "").toLowerCase();
    if (obj[clean] !== undefined && obj[clean] !== null && String(obj[clean]).trim() !== "") {
      return obj[clean];
    }
  }
  return "";
}

function findOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

// ────────────────────────────────────────────────
// Drive 파일 ID 추출
// ────────────────────────────────────────────────
function extractFileId(str) {
  if (!str || typeof str !== "string") return "";
  var s = str.trim();
  var m1 = s.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  if (m1) return m1[1];
  var m2 = s.match(/[?&]id=([a-zA-Z0-9_-]{25,})/i) || s.match(/id=([a-zA-Z0-9_-]{25,})/i);
  if (m2) return m2[1];
  var m3 = s.match(/([a-zA-Z0-9_-]{25,})/);
  if (m3) return m3[1];
  return "";
}

// ────────────────────────────────────────────────
// 카테고리 정규화 (영문 → 한글)
// ────────────────────────────────────────────────
function normCategory(raw) {
  var s = String(raw || "").trim().toLowerCase();
  if (s === "daily"   || s === "일일") return "일일";
  if (s === "weekly"  || s === "주간") return "주간";
  if (s === "special" || s === "특별" || s === "bonus") return "특별";
  if (s.length > 0) return String(raw || "").trim();
  return "일일";
}

// ────────────────────────────────────────────────
// 레벨 타이틀 (최대 Lv.100)
// ────────────────────────────────────────────────
function getLevelTitle(lv) {
  if (lv >= 100) return "🎮 신화의 플레이스테이션 마스터";
  if (lv >= 50)  return "🔮 마스터 아케이드 히어로";
  if (lv >= 30)  return "💎 플래티넘 챔피언";
  if (lv >= 20)  return "🛡️ 불꽃의 골드 가디언";
  if (lv >= 10)  return "⚔️ 용감한 실버 기사";
  if (lv >= 5)   return "🔥 열정의 프로 성실러";
  if (lv >= 3)   return "⭐ 든든한 퀘스트 챔피언";
  if (lv >= 2)   return "🌟 쑥쑥 자라는 모험가";
  return "🌱 새내기 견습 모험가";
}

// ────────────────────────────────────────────────
// 메인 데이터 로드
// ────────────────────────────────────────────────
function getAppData() {
  try {
    var ss = getSS();

    // ── 시트 읽기 ──
    var usersSheet  = ss.getSheetByName("Users");
    var questsSheet = ss.getSheetByName("Quests");
    var logsSheet   = ss.getSheetByName("QuestLogs");
    var rewardsSheet= ss.getSheetByName("Rewards");
    var invSheet    = ss.getSheetByName("Inventory");

    var users   = sheetToObjects(usersSheet);
    var quests  = sheetToObjects(questsSheet);
    var logs    = sheetToObjects(logsSheet);
    var rewards = sheetToObjects(rewardsSheet);
    var invRows = sheetToObjects(invSheet);

    // ── 사용자 정리 ──
    var profiles = users.map(function(u) {
      return {
        id:       String(getField(u, "userId", "id", "아이디") || "user_" + u.__rowIndex).trim(),
        name:     String(getField(u, "name", "이름", "닉네임") || "").trim(),
        role:     String(getField(u, "role", "역할", "구분") || "CHILD").trim().toUpperCase(),
        pin:      String(getField(u, "pin", "비밀번호", "비번") || "").trim(),
        points:   Number(getField(u, "points", "포인트", "잔액") || 0) || 0,
        totalExp: Number(getField(u, "totalExp", "경험치", "총경험치", "xp", "points") || 0) || 0,
        icon:     String(getField(u, "icon", "아이콘") || "").trim() || "🧒"
      };
    });

    // 아이 프로필 (CHILD)
    var childProfile = profiles.find(function(p) { return p.role === "CHILD"; }) || {
      id: "child_1", name: "상급종합빡빡이(용사)", role: "CHILD", pin: "", points: 7400, totalExp: 7400, icon: "🧒"
    };

    var totalXp = childProfile.totalExp || childProfile.points || 0;
    if (!totalXp) {
      logs.forEach(function(log) {
        var st = String(getField(log, "status", "상태") || "").toUpperCase();
        if (st === "APPROVED") {
          totalXp += Number(getField(log, "points", "xp", "포인트", "경험치") || 0) || 0;
        }
      });
    }
    var currentBalance = childProfile.points || totalXp;

    var level = 1;
    if (totalXp > 0) level = Math.min(100, Math.max(1, Math.floor(totalXp / XP_PER_LEVEL) + 1));
    var levelTitle    = getLevelTitle(level);
    var lvProgress    = totalXp % XP_PER_LEVEL;
    var progressPct   = Math.min(100, Math.floor((lvProgress / XP_PER_LEVEL) * 100));

    // ── 퀘스트 정리 ──
    var questList = quests.map(function(q) {
      var qId = String(getField(q, "questId", "id", "ID") || "Q-" + q.__rowIndex).trim();
      var qTitle = String(getField(q, "title", "퀘스트명", "제목", "퀘스트") || "").trim();
      var pts = Number(getField(q, "points", "포인트", "보상", "point", "Point") || 0) || 0;
      if (pts <= 0) {
        if (qId.includes("DISH") || qId.includes("SHOES")) pts = 200;
        else if (qId.includes("DESK") || qId.includes("20260830_075520") || qId.includes("20260830_082800")) pts = 300;
        else if (qId.includes("RECYCLE") || qId.includes("STUDY") || qId.includes("BONUS_01") || qId.includes("20260830_081200")) pts = 500;
        else if (qId.includes("BONUS_02") || qId.includes("CLEAR")) pts = 1000;
        else pts = 300;
      }
      return {
        id:       qId,
        title:    qTitle,
        category: normCategory(getField(q, "category", "카테고리", "구분", "분류") || "일일"),
        period:   normCategory(getField(q, "period", "주기", "기간", "category") || "일일"),
        points:   pts,
        xp:       Number(getField(q, "xp", "경험치") || pts) || pts,
        icon:     String(getField(q, "icon", "아이콘") || "⭐").trim(),
        desc:     String(getField(q, "desc", "description", "설명", "내용") || "").trim()
      };
    }).filter(function(q) { return q.title.length > 0; });

    if (questList.length === 0) questList = getDefaultQuests();

    // ── 로그 정리 ──
    var questMap = {};
    questList.forEach(function(q) {
      if (q.id) questMap[q.id] = q;
      if (q.title) questMap[q.title] = q;
    });

    var profileMap = {
      "parent_1": "엄마 (길드마스터)",
      "parent_2": "아빠 (대마법사)",
      "mom": "엄마 (길드마스터)",
      "dad": "아빠 (대마법사)"
    };
    profiles.forEach(function(p) {
      if (p.id) profileMap[p.id] = p.name;
      if (p.name.includes("엄마")) { profileMap["parent_1"] = p.name; profileMap["mom"] = p.name; }
      if (p.name.includes("아빠")) { profileMap["parent_2"] = p.name; profileMap["dad"] = p.name; }
    });

    var tz = Session.getScriptTimeZone();
    var logList = logs.map(function(log) {
      var rawStatus = String(getField(log, "status", "상태") || "APPROVED").trim();
      var upperSt   = rawStatus.toUpperCase();
      var status    = upperSt === "PENDING" || rawStatus === "대기" || rawStatus === "심의대기" ? "PENDING"
                    : upperSt === "REJECTED" || rawStatus === "반려" ? "REJECTED"
                    : "APPROVED";

      var rawPhoto  = String(getField(log, "photoUrl", "photo_url", "photo", "photos", "사진", "인증사진") || "").trim();
      var photoStr  = "";
      if (rawPhoto) {
        var parts = [];
        if (rawPhoto.startsWith("[")) {
          try { parts = JSON.parse(rawPhoto); } catch(e) { parts = [rawPhoto]; }
        } else {
          parts = rawPhoto.split(/[,|\n\r]+/);
        }
        photoStr = parts.map(function(p) {
          if (!p) return "";
          var pStr = String(p).trim();
          var fid = extractFileId(pStr);
          if (fid) return fid;
          // 과거 저장된 Base64 문자열이 있을 경우 Drive로 마이그레이션 시도
          if (pStr.indexOf("data:image") === 0 || pStr.length > 500) {
            try {
              var folder = getFolder();
              if (folder) {
                var clean = pStr.replace(/^data:image\/\w+;base64,/, "");
                var blob  = Utilities.newBlob(Utilities.base64Decode(clean), "image/jpeg", "Migrated_" + (log.LogId || "photo") + ".jpg");
                var file  = folder.createFile(blob);
                try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
                return file.getId();
              }
            } catch(e) {
              return "";
            }
          }
          return "";
        }).filter(Boolean).join(",");
      }

      var dateRaw = getField(log, "submittedAt", "completedAt", "date", "일시", "날짜", "시간");
      var dateStr = "";
      if (dateRaw instanceof Date) {
        dateStr = Utilities.formatDate(dateRaw, tz, "yyyy-MM-dd HH:mm");
      } else if (dateRaw) {
        dateStr = String(dateRaw).trim();
      }

      var qId = String(getField(log, "questId", "id") || "").trim();
      var questObj = questMap[qId] || null;
      var qTitle = String(getField(log, "questTitle", "title", "퀘스트명", "제목") || (questObj ? questObj.title : "") || "퀘스트 완료").trim();
      var qPoints = Number(getField(log, "points", "포인트") || (questObj ? questObj.points : 300)) || 300;
      var qXp = Number(getField(log, "xp", "경험치") || (questObj ? questObj.xp : qPoints)) || qPoints;
      var memo = String(getField(log, "memo", "소감", "메모", "아이소감") || "").trim();

      // 칭찬 코멘트 및 반려 사유 (compliment, rejectReason 완벽 지원)
      var parentComment = "";
      if (status === "REJECTED") {
        parentComment = String(getField(log, "rejectReason", "reject_reason", "반려사유", "사유", "compliment", "parentComment", "comment") || "").trim();
      } else {
        parentComment = String(getField(log, "compliment", "칭찬", "parentComment", "comment", "코멘트", "부모코멘트", "feedback") || "").trim();
      }

      var reviewerKey = String(getField(log, "reviewedBy", "approvedBy", "승인자", "심의자", "approver") || "").trim();
      var approvedBy = profileMap[reviewerKey] || (reviewerKey === "parent_1" ? "엄마 (길드마스터)" : (reviewerKey === "parent_2" ? "아빠 (대마법사)" : (reviewerKey || (status === "APPROVED" ? "부모님" : ""))));

      return {
        id:            String(getField(log, "logId", "id") || "LOG-" + log.__rowIndex).trim(),
        questId:       qId,
        title:         qTitle,
        userId:        String(getField(log, "userId") || "child_1").trim(),
        date:          dateStr,
        status:        status,
        photoUrl:      photoStr,
        memo:          memo,
        parentComment: parentComment,
        approvedBy:    approvedBy,
        points:        qPoints,
        xp:            qXp,
        rowIndex:      log.__rowIndex
      };
    }).reverse(); // 최신순

    // ── 보상 정리 ──
    var shopItems = rewards.filter(function(r) {
      var activeStr = String(getField(r, "isActive", "is_active", "active", "활성") || "TRUE").trim().toUpperCase();
      return activeStr !== "FALSE" && activeStr !== "0";
    }).map(function(r) {
      var title = String(getField(r, "title", "상품명", "보상", "name", "이름") || "").trim();
      var cost = Number(getField(r, "costPoints", "cost_points", "cost", "price", "points", "point", "포인트", "가격", "금액", "비용") || 0) || 0;
      var icon = String(getField(r, "icon", "아이콘") || "").trim() || "🎁";
      var desc = String(getField(r, "description", "desc", "설명", "내용") || "").trim();
      return {
        id:       String(getField(r, "rewardId", "id", "ID") || "R-" + r.__rowIndex).trim(),
        title:    title,
        cost:     cost,
        category: String(getField(r, "category", "카테고리", "분류") || "보상").trim(),
        icon:     icon,
        desc:     desc
      };
    }).filter(function(r) { return r.title.length > 0; });

    if (shopItems.length === 0) shopItems = getDefaultShopItems();

    // ── 인벤토리 정리 ──
    var inventory = invRows.map(function(inv) {
      return {
        id:          String(getField(inv, "invId", "id") || "INV-" + inv.__rowIndex).trim(),
        userId:      String(getField(inv, "userId") || "child_1").trim(),
        rewardId:    String(getField(inv, "rewardId") || "").trim(),
        rewardTitle: String(getField(inv, "rewardTitle", "title", "보상명") || "").trim(),
        acquiredAt:  String(getField(inv, "acquiredAt", "date", "획득일") || "").trim(),
        status:      String(getField(inv, "status", "상태") || "ACTIVE").trim().toUpperCase(),
        icon:        String(getField(inv, "icon", "아이콘") || "🎁").trim(),
        rowIndex:    inv.__rowIndex
      };
    }).filter(function(inv) {
      return inv.userId === childProfile.id || inv.userId === "child_1" || inv.userId === "";
    });

    // ── 심의 대기 목록 ──
    var pendingReviews = logList.filter(function(l) { return l.status === "PENDING"; });

    // ── 부모 PIN 맵 ──
    var parentPins = {};
    profiles.filter(function(p) { return p.role === "PARENT"; }).forEach(function(p) {
      if (p.pin) parentPins[p.id] = p.pin;
      if (p.name.includes("엄마") || p.name.toLowerCase() === "mom") parentPins["mom"] = p.pin || "1121";
      if (p.name.includes("아빠") || p.name.toLowerCase() === "dad") parentPins["dad"] = p.pin || "7590";
    });
    // 폴백
    if (!parentPins["mom"]) parentPins["mom"] = "1121";
    if (!parentPins["dad"]) parentPins["dad"] = "7590";

    return {
      profiles:       profiles,
      parentPins:     parentPins,
      child: {
        id:             childProfile.id,
        name:           childProfile.name,
        icon:           childProfile.icon,
        level:          level,
        title:          levelTitle,
        totalExp:       totalXp,
        currentBalance: currentBalance,
        targetItem:     "🎮 플레이스테이션",
        targetPoints:   100000,
        progressPct:    progressPct,
        lvProgress:     lvProgress,
        nextLvXp:       XP_PER_LEVEL
      },
      quests:         questList,
      shopItems:      shopItems,
      recentLogs:     logList,
      pendingReviews: pendingReviews,
      inventory:      inventory
    };

  } catch(e) {
    Logger.log("getAppData error: " + e);
    return getFallbackData();
  }
}

// ────────────────────────────────────────────────
// 퀘스트 완료 제출
// ────────────────────────────────────────────────
function completeQuest(data) {
  try {
    var ss   = getSS();
    var tz   = Session.getScriptTimeZone();
    var now  = new Date();
    var nowStr = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");
    var logId  = "LOG_" + Utilities.formatDate(now, tz, "yyyyMMdd_HHmmss") + "_" + Math.floor(Math.random() * 9000 + 1000);

    // 사진 업로드 (최대 3장)
    var photoIds = [];
    var photos   = data.photosBase64 || [];
    if (photos.length > 0) {
      var folder = getFolder();
      if (folder) {
        photos.slice(0, 3).forEach(function(b64, idx) {
          if (!b64 || b64.length < 50) return;
          try {
            var clean = b64.replace(/^data:image\/\w+;base64,/, "");
            var blob  = Utilities.newBlob(Utilities.base64Decode(clean), "image/jpeg",
                          "Quest_" + logId + "_" + (idx + 1) + ".jpg");
            var file  = folder.createFile(blob);
            try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
            photoIds.push(file.getId());
          } catch(uploadErr) { Logger.log("Photo upload error: " + uploadErr); }
        });
      }
    }

    var photoStr = photoIds.join(",");
    var isParent = data.isParent || false;
    var status   = isParent ? "APPROVED" : "PENDING";

    // QuestLogs 시트에 추가
    var logsSheet = findOrCreateSheet(ss, "QuestLogs", [
      "LogId","QuestId","QuestTitle","UserId","SubmittedAt","Status","PhotoUrl","Memo","ParentComment","ApprovedBy"
    ]);

    logsSheet.appendRow([
      logId,
      data.questId   || "CUSTOM",
      data.title     || "퀘스트 완료",
      data.userId    || "child_1",
      nowStr,
      status,
      photoStr,
      data.memo      || "",
      "",
      ""
    ]);

    // APPROVED이면 포인트/XP 즉시 반영
    if (isParent) {
      updateChildPoints(ss, data.userId || "child_1", Number(data.points) || 0);
    }

    return {
      success: true,
      message: isParent
        ? "🎉 퀘스트 완료 승인! +" + (data.points || 0) + " P"
        : "📤 퀘스트 제출 완료! ⚖️ 부모님 심의 대기 중입니다.",
      appData: getAppData()
    };
  } catch(err) {
    Logger.log("completeQuest error: " + err);
    return { success: false, error: err.toString() };
  }
}

// ────────────────────────────────────────────────
// 퀘스트 승인
// ────────────────────────────────────────────────
function approveQuest(data) {
  try {
    var ss   = getSS();
    var sheet = ss.getSheetByName("QuestLogs");
    if (!sheet) return { success: false, error: "QuestLogs 시트 없음" };

    var lr = sheet.getLastRow();
    var lc = sheet.getLastColumn();
    if (lr < 2) return { success: false, error: "기록 없음" };

    var allValues = sheet.getRange(1, 1, lr, lc).getValues();
    var headers = allValues[0].map(function(h) {
      return String(h || "").trim().toLowerCase().replace(/[^a-z0-9가-힣]/g, '');
    });

    var statusCol = headers.indexOf("status") !== -1 ? headers.indexOf("status") + 1 : (headers.indexOf("상태") !== -1 ? headers.indexOf("상태") + 1 : 6);
    var commentCol = headers.indexOf("parentcomment") !== -1 ? headers.indexOf("parentcomment") + 1
                   : headers.indexOf("comment") !== -1 ? headers.indexOf("comment") + 1
                   : headers.indexOf("칭찬") !== -1 ? headers.indexOf("칭찬") + 1
                   : headers.indexOf("코멘트") !== -1 ? headers.indexOf("코멘트") + 1
                   : (headers.length >= 9 ? 9 : lc);
    var approverCol = headers.indexOf("approvedby") !== -1 ? headers.indexOf("approvedby") + 1
                    : headers.indexOf("승인자") !== -1 ? headers.indexOf("승인자") + 1
                    : (headers.length >= 10 ? 10 : lc);

    var foundRow = -1;
    for (var r = 1; r < allValues.length; r++) {
      if (String(allValues[r][0]).trim() === String(data.logId).trim()) {
        foundRow = r + 1;
        break;
      }
    }

    if (foundRow !== -1) {
      if (statusCol > 0 && statusCol <= lc) sheet.getRange(foundRow, statusCol).setValue("APPROVED");
      if (data.comment && commentCol > 0 && commentCol <= lc) sheet.getRange(foundRow, commentCol).setValue(data.comment);
      if (approverCol > 0 && approverCol <= lc) sheet.getRange(foundRow, approverCol).setValue(data.approvedBy || "부모님");

      // 포인트 반영
      var pts = Number(data.points) || 0;
      if (pts > 0) updateChildPoints(ss, data.userId || "child_1", pts);
    }

    return { success: true, message: "✅ 승인 완료! 포인트가 지급되었습니다.", appData: getAppData() };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ────────────────────────────────────────────────
// 퀘스트 반려
// ────────────────────────────────────────────────
function rejectQuest(data) {
  try {
    var ss    = getSS();
    var sheet = ss.getSheetByName("QuestLogs");
    if (!sheet) return { success: false, error: "QuestLogs 시트 없음" };

    var rows = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
    var found = -1;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.logId).trim()) { found = i + 2; break; }
    }

    if (found !== -1) {
      sheet.getRange(found, 6).setValue("REJECTED");
      if (data.comment) sheet.getRange(found, 9).setValue(data.comment);
      sheet.getRange(found, 10).setValue(data.approvedBy || "부모님");
    }

    return { success: true, message: "❌ 반려 처리되었습니다.", appData: getAppData() };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ────────────────────────────────────────────────
// 보상 교환 (상점 → 인벤토리)
// ────────────────────────────────────────────────
function redeemReward(data) {
  try {
    var ss      = getSS();
    var appData = getAppData();
    var cost    = Number(data.cost) || 0;

    if (appData.child.currentBalance < cost) {
      return { success: false, error: "보유 포인트가 부족합니다!" };
    }

    var tz  = Session.getScriptTimeZone();
    var now = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    var invId = "INV_" + Utilities.formatDate(new Date(), tz, "yyyyMMdd_HHmmss");

    // 인벤토리 시트에 추가
    var invSheet = findOrCreateSheet(ss, "Inventory", [
      "InvId","UserId","RewardId","RewardTitle","AcquiredAt","Status","Icon"
    ]);
    invSheet.appendRow([invId, data.userId || "child_1", data.rewardId || "", data.title || "", now, "ACTIVE", data.icon || "🎁"]);

    // 포인트 차감
    updateChildPoints(ss, data.userId || "child_1", -cost);

    return { success: true, message: "🎁 [" + data.title + "] 획득! 인벤토리에 추가되었습니다.", appData: getAppData() };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ────────────────────────────────────────────────
// 인벤토리 아이템 사용
// ────────────────────────────────────────────────
function useInventoryItem(data) {
  try {
    var ss    = getSS();
    var sheet = ss.getSheetByName("Inventory");
    if (!sheet) return { success: false, error: "Inventory 시트 없음" };

    var rows = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
    var found = -1;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.invId).trim()) { found = i + 2; break; }
    }

    if (found !== -1) {
      sheet.getRange(found, 6).setValue("USED");
    }

    return { success: true, message: "✅ 아이템을 사용했습니다!", appData: getAppData() };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ────────────────────────────────────────────────
// 포인트/XP 업데이트
// ────────────────────────────────────────────────
function updateChildPoints(ss, userId, delta) {
  var usersSheet = ss.getSheetByName("Users");
  if (!usersSheet || usersSheet.getLastRow() < 2) return;

  var data = usersSheet.getRange(1, 1, usersSheet.getLastRow(), usersSheet.getLastColumn()).getValues();
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var idCol  = headers.indexOf("userid");
  var ptCol  = headers.indexOf("points");
  var xpCol  = headers.indexOf("totalexp");
  if (idCol === -1) idCol = 0;
  if (ptCol === -1) ptCol = 4;
  if (xpCol === -1) xpCol = 5;

  for (var r = 1; r < data.length; r++) {
    var rowId = String(data[r][idCol]).trim();
    var role  = String(data[r][headers.indexOf("role") !== -1 ? headers.indexOf("role") : 2] || "").toUpperCase();
    if (rowId === userId || role === "CHILD") {
      var curPts = Number(data[r][ptCol]) || 0;
      var curXp  = Number(data[r][xpCol]) || 0;
      usersSheet.getRange(r + 1, ptCol + 1).setValue(curPts + delta);
      if (delta > 0) usersSheet.getRange(r + 1, xpCol + 1).setValue(curXp + delta);
      break;
    }
  }
}

// ────────────────────────────────────────────────
// 기본 데이터 (시트 비어있을 때)
// ────────────────────────────────────────────────
function getDefaultQuests() {
  return [
    { id: "Q_DESK_01",    title: "책상 정리",              category: "일일", period: "일일", points: 300, xp: 300, icon: "🧹", desc: "책상 위 먼지 털기 및 깔끔하게 정돈하기" },
    { id: "Q_DISH_01",    title: "설거지 & 싱크대 정리",  category: "일일", period: "일일", points: 200, xp: 200, icon: "🍽️", desc: "식사 후 그릇을 깨끗이 씻고 싱크대 물기 닦기" },
    { id: "Q_STUDY_01",   title: "오늘의 숙제 완료",      category: "일일", period: "일일", points: 500, xp: 500, icon: "✏️", desc: "학교 및 학원 숙제 빠짐없이 완료하기" },
    { id: "Q_SHOES_01",   title: "현관 신발 정리",        category: "일일", period: "일일", points: 100, xp: 100, icon: "👟", desc: "현관에 놓인 신발들을 가지런히 정돈하기" },
    { id: "Q_RECYCLE_01", title: "분리수거 돕기",         category: "주간", period: "주간", points: 500, xp: 500, icon: "♻️", desc: "플라스틱/종이/캔 분리수거장 가져다 버리기" },
    { id: "Q_BRUSH_01",   title: "양치질 3분 꼼꼼히",    category: "일일", period: "일일", points: 200, xp: 200, icon: "🪥", desc: "양치질 3분 타이머 맞춰 꼼꼼히 칫카하기" },
    { id: "Q_LAUNDRY_01", title: "빨래 개기 돕기",        category: "주간", period: "주간", points: 500, xp: 500, icon: "⭐", desc: "수건이나 옷을 이쁘게 개어서 서랍에 정리" }
  ];
}

function getDefaultShopItems() {
  return [
    { id: "R-001", title: "🎮 주말 게임 30분 쿠폰",            cost: 1500, category: "자유시간", icon: "🎮", desc: "스마트폰/콘솔 게임 30분 추가 자유 이용권" },
    { id: "R-002", title: "🍦 베스킨라빈스 싱글팅 아이스크림", cost: 3000, category: "간식",     icon: "🍦", desc: "좋아하는 맛 아이스크림 1개 보상 교환" },
    { id: "R-003", title: "💰 용돈 3,000원 현금 환전",         cost: 3000, category: "용돈",     icon: "💵", desc: "모은 포인트를 즉시 3천원 현금 용돈으로 교환!" },
    { id: "R-004", title: "💰 용돈 5,000원 현금 환전",         cost: 5000, category: "용돈",     icon: "💰", desc: "열심히 모은 5천원 현금 지급!" },
    { id: "R-005", title: "🛋️ 주말 늦잠 자유 이용권",          cost: 1000, category: "자유시간", icon: "🛋️", desc: "주말 아침 깨우지 않고 푹 쉴 수 있는 권리" }
  ];
}

function getFallbackData() {
  return {
    profiles:   [
      { id: "child_1", name: "상급종합빡빡이(용사)", role: "CHILD",  pin: "",     points: 7400, totalExp: 7400, icon: "🧒" },
      { id: "mom",     name: "엄마 (길드마스터)",    role: "PARENT", pin: "1121", points: 0,    totalExp: 0,    icon: "👩" },
      { id: "dad",     name: "아빠 (대마법사)",      role: "PARENT", pin: "7590", points: 0,    totalExp: 0,    icon: "👨" }
    ],
    parentPins: { mom: "1121", dad: "7590" },
    child: {
      id: "child_1", name: "상급종합빡빡이(용사)", icon: "🧒",
      level: 3, title: "⭐ 든든한 퀘스트 챔피언",
      totalExp: 7400, currentBalance: 7400,
      targetItem: "🎮 플레이스테이션", targetPoints: 100000,
      progressPct: 8, lvProgress: 2400, nextLvXp: 2500
    },
    quests:         getDefaultQuests(),
    shopItems:      getDefaultShopItems(),
    recentLogs:     [],
    pendingReviews: [],
    inventory:      []
  };
}
