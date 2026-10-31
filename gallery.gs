/**
 * 뱅크시와 고정관념 깨기 - 반별 사진 갤러리 서버 (Google Apps Script)
 *
 * ▶ 설치 방법 (한 번만)
 *   1. script.google.com → [새 프로젝트] → 이 코드를 전부 붙여넣기 → 저장
 *   2. [배포] → [새 배포] → 유형: 웹 앱
 *        - 실행 사용자: 나
 *        - 액세스 권한: 모든 사용자
 *   3. 권한 승인 후 나오는 "웹 앱 URL"을 복사해서
 *      index.html 의 GALLERY_SCRIPT_URL = "" 따옴표 안에 붙여넣기
 *
 * 작품은 선생님 구글 드라이브의 "뱅크시 반별 갤러리" 폴더 안에 반별 폴더로 저장됩니다.
 * (드라이브 공유 설정을 바꾸지 않아도 학생들에게 보입니다.)
 */

// ===== 선생님이 바꿀 수 있는 설정 =====
const TEACHER_PASSWORD = "banksy2026";       // 작품 삭제용 교사 비밀번호 (꼭 바꾸세요)
const ROOT_FOLDER_NAME = "뱅크시 반별 갤러리";
const MAX_LIST = 80;                         // 한 반에서 보여줄 최대 작품 수
const CLASS_PASSWORDS = {                    // index.html 의 반 비밀번호와 같게
  "3-1": "1111", "3-2": "2222", "3-3": "3333", "3-4": "4444", "3-5": "5555",
  "3-6": "6666", "3-7": "7777", "3-8": "8888", "3-9": "9999", "3-10": "1010"
};
// ====================================

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === "list") return json_(listArtworks_(p.classId, p.password));
    if (p.action === "image") return json_(getImage_(p.classId, p.password, p.id));
    if (p.action === "ping") return json_({ status: "success", message: "gallery ok" });
    return json_({ status: "error", message: "알 수 없는 요청입니다." });
  } catch (err) {
    return json_({ status: "error", message: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (body.action === "upload") return json_(uploadArtwork_(body));
    if (body.action === "delete") return json_(deleteArtwork_(body));
    if (body.action === "like") return json_(likeArtwork_(body));
    return json_({ status: "error", message: "알 수 없는 요청입니다." });
  } catch (err) {
    return json_({ status: "error", message: String(err && err.message || err) });
  }
}

// ---------- 기능 ----------
function uploadArtwork_(b) {
  checkClass_(b.classId, b.password);
  const full = parseDataUrl_(b.image);
  const thumb = parseDataUrl_(b.thumb);
  if (full.bytes.length > 3 * 1024 * 1024) throw new Error("사진 용량이 너무 큽니다.");

  // 반 폴더 찾기/만들기만 잠깐 잠그고, 파일 저장은 여러 학생이 동시에 해도 되게 함
  // (한 반 30명이 한꺼번에 올려도 줄 서서 기다리다 실패하지 않도록)
  let folder;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    folder = classFolder_(b.classId);
  } finally {
    lock.releaseLock();
  }

  const ts = Date.now();
  const uid = ts + "_" + Math.floor(Math.random() * 1e6);
  const meta = {
    name: clean_(b.name, 20) || "익명",
    title: clean_(b.title, 40) || "제목 없음",
    ts: ts
  };
  const fullFile = folder.createFile(Utilities.newBlob(full.bytes, "image/jpeg", uid + "_full.jpg"));
  meta.fullId = fullFile.getId();
  const thumbFile = folder.createFile(Utilities.newBlob(thumb.bytes, "image/jpeg", uid + "_thumb.jpg"));
  thumbFile.setDescription(JSON.stringify(meta));
  fullFile.setDescription(JSON.stringify({ thumbId: thumbFile.getId(), ts: ts }));
  return { status: "success", id: thumbFile.getId() };
}

function listArtworks_(classId, password) {
  checkClass_(classId, password);
  const folder = classFolder_(classId);
  const files = folder.getFiles();
  const items = [];
  while (files.hasNext()) {
    const f = files.next();
    if (!/_thumb\.jpg$/.test(f.getName())) continue;
    let meta = {};
    try { meta = JSON.parse(f.getDescription() || "{}"); } catch (err) {}
    items.push({ file: f, meta: meta });
  }
  items.sort(function (a, b) { return (b.meta.ts || 0) - (a.meta.ts || 0); });
  const list = items.slice(0, MAX_LIST).map(function (it) {
    return {
      id: it.file.getId(),
      name: it.meta.name || "익명",
      title: it.meta.title || "제목 없음",
      ts: it.meta.ts || 0,
      likes: it.meta.likes || 0,
      thumb: "data:image/jpeg;base64," + Utilities.base64Encode(it.file.getBlob().getBytes())
    };
  });
  return { status: "success", items: list, total: items.length };
}

function getImage_(classId, password, thumbId) {
  checkClass_(classId, password);
  const thumbFile = fileInClass_(classId, thumbId);
  const meta = JSON.parse(thumbFile.getDescription() || "{}");
  const full = DriveApp.getFileById(meta.fullId);
  return {
    status: "success",
    image: "data:image/jpeg;base64," + Utilities.base64Encode(full.getBlob().getBytes())
  };
}

// 좋아요 (+1) / 좋아요 취소 (-1)
function likeArtwork_(b) {
  checkClass_(b.classId, b.password);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const thumbFile = fileInClass_(b.classId, b.id);
    let meta = {};
    try { meta = JSON.parse(thumbFile.getDescription() || "{}"); } catch (err) {}
    meta.likes = Math.max(0, (meta.likes || 0) + (Number(b.delta) < 0 ? -1 : 1));
    thumbFile.setDescription(JSON.stringify(meta));
    return { status: "success", likes: meta.likes };
  } finally {
    lock.releaseLock();
  }
}

function deleteArtwork_(b) {
  if (String(b.teacherPassword || "") !== TEACHER_PASSWORD) throw new Error("교사 비밀번호가 올바르지 않습니다.");
  const thumbFile = fileInClass_(b.classId, b.id);
  let meta = {};
  try { meta = JSON.parse(thumbFile.getDescription() || "{}"); } catch (err) {}
  if (meta.fullId) {
    try { DriveApp.getFileById(meta.fullId).setTrashed(true); } catch (err) {}
  }
  thumbFile.setTrashed(true);
  return { status: "success" };
}

// ---------- 도우미 ----------
function checkClass_(classId, password) {
  if (!CLASS_PASSWORDS.hasOwnProperty(classId)) throw new Error("반 정보가 올바르지 않습니다.");
  if (String(password || "") !== CLASS_PASSWORDS[classId]) throw new Error("반 비밀번호가 올바르지 않습니다.");
}

function rootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("GALLERY_ROOT_ID");
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (err) {}
  }
  const folder = DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty("GALLERY_ROOT_ID", folder.getId());
  return folder;
}

function classFolder_(classId) {
  const root = rootFolder_();
  const it = root.getFoldersByName(classId);
  return it.hasNext() ? it.next() : root.createFolder(classId);
}

// 다른 반의 사진 ID로 접근하지 못하게 반 폴더 안의 파일인지 확인
function fileInClass_(classId, fileId) {
  if (!CLASS_PASSWORDS.hasOwnProperty(classId)) throw new Error("반 정보가 올바르지 않습니다.");
  const folderId = classFolder_(classId).getId();
  const f = DriveApp.getFileById(String(fileId || ""));
  const parents = f.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === folderId) return f;
  }
  throw new Error("이 반의 작품이 아닙니다.");
}

function parseDataUrl_(dataUrl) {
  const m = /^data:image\/(jpeg|png|webp);base64,(.+)$/.exec(String(dataUrl || ""));
  if (!m) throw new Error("이미지 형식이 올바르지 않습니다.");
  return { bytes: Utilities.base64Decode(m[2]) };
}

function clean_(s, maxLen) {
  return String(s || "").replace(/[<>]/g, "").trim().slice(0, maxLen);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
