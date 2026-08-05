/**
 * STREAMING_CHUNK: Initializing Google Apps Script Backend API for SPENTIG Portal...
 * 
 * Pengelolaan Sheet pada Google Spreadsheet:
 * 1. Sheet "Users" -> Baris 1 Header: Username | Password | Nama | Role | Status
 *    Role yang didukung: Guru | Staff | Siswa | Orang Tua | Administrator | Umum
 * 2. Sheet "Links" -> Baris 1 Header: Title | Category | Url | Icon | Roles | Status
 *    Roles dapat berisi koma: Guru,Staff,Siswa,Orang Tua,Administrator,Umum
 */

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse(false, "Payload data kosong!");
    }

    var data = JSON.parse(e.postData.contents);
    var action = String(data.action || "login").toLowerCase();

    // Request khusus mengambil Tautan Dinamis berdasarkan Role (Misal untuk Tamu/Public)
    if (action === "get_links") {
      var role = String(data.role || "Umum").trim();
      var links = getLinksForRole(role);
      return createJsonResponse(true, "Berhasil mengambil data tautan.", null, links);
    }

    // Aksi Utama: Login Autentikasi Pengguna
    var username = String(data.username || "").trim();
    var password = String(data.password || "").trim();

    if (!username || !password) {
      return createJsonResponse(false, "Username dan Password wajib diisi!");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("Users");

    if (!userSheet) {
      return createJsonResponse(false, "Sheet bernama 'Users' tidak ditemukan di Spreadsheet!");
    }

    var userRows = userSheet.getDataRange().getValues();

    // Memeriksa akun dari baris ke-2 (index 1)
    for (var i = 1; i < userRows.length; i++) {
      var rowUser = String(userRows[i][0] || "").trim();
      var rowPass = String(userRows[i][1] || "").trim();
      var name = String(userRows[i][2] || "").trim();
      var role = String(userRows[i][3] || "").trim();
      var status = String(userRows[i][4] || "Aktif").trim();

      // Pencocokan Username Case-Insensitive & Password Eksak
      if (rowUser.toLowerCase() === username.toLowerCase() && rowPass === password) {
        if (status.toLowerCase() === "nonaktif") {
          return createJsonResponse(false, "Akun Anda sedang dinonaktifkan oleh Administrator.");
        }

        var userObj = {
          username: rowUser,
          name: name || rowUser,
          role: role || "Umum"
        };

        // Ambil daftar link dinamis khusus role user
        var userLinks = getLinksForRole(userObj.role);

        return createJsonResponse(true, "Login Berhasil", userObj, userLinks);
      }
    }

    return createJsonResponse(false, "NIP/NISN/Username atau Kata Sandi salah!");

  } catch (err) {
    return createJsonResponse(false, "Terjadi kesalahan server: " + err.toString());
  }
}

function doGet(e) {
  try {
    var role = (e && e.parameter && e.parameter.role) ? String(e.parameter.role).trim() : "Umum";
    var links = getLinksForRole(role);
    return createJsonResponse(true, "API Portal SPENTIG Aktif.", null, links);
  } catch (err) {
    return createJsonResponse(false, "Error: " + err.toString());
  }
}

/**
 * STREAMING_CHUNK: Filtering dynamic links from sheet Links based on user role...
 */
function getLinksForRole(userRole) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var linkSheet = ss.getSheetByName("Links");
  var linksList = [];

  if (!linkSheet) return linksList;

  var linkRows = linkSheet.getDataRange().getValues();
  if (linkRows.length <= 1) return linksList;

  var lowerUserRole = userRole.toLowerCase().trim();

  for (var j = 1; j < linkRows.length; j++) {
    var title = String(linkRows[j][0] || "").trim();
    var category = String(linkRows[j][1] || "").trim();
    var url = String(linkRows[j][2] || "").trim();
    var icon = String(linkRows[j][3] || "fa-solid fa-link").trim();
    var rawRoles = String(linkRows[j][4] || "Umum").trim();
    var status = String(linkRows[j][5] || "Aktif").trim();

    // Abaikan jika status nonaktif atau data tidak lengkap
    if (status.toLowerCase() === "nonaktif" || !title || !url) continue;

    var rolesArray = rawRoles.split(",").map(function(r) { return r.trim().toLowerCase(); });

    // Administrator dapat melihat seluruh kartu, ATAU cocok dengan role user, ATAU mengandung 'umum'
    var isAllowed = (lowerUserRole === "administrator") || 
                    rolesArray.indexOf(lowerUserRole) !== -1 || 
                    rolesArray.indexOf("umum") !== -1;

    if (isAllowed) {
      linksList.push({
        title: title,
        category: category,
        url: url,
        icon: icon,
        roles: rawRoles
      });
    }
  }

  return linksList;
}

/**
 * STREAMING_CHUNK: Formatting JSON response helper...
 */
function createJsonResponse(success, message, userObj, linksArray) {
  var responsePayload = {
    success: success,
    message: message
  };

  if (userObj) responsePayload.user = userObj;
  if (linksArray) responsePayload.links = linksArray;

  return ContentService
    .createTextOutput(JSON.stringify(responsePayload))
    .setMimeType(ContentService.MimeType.JSON);
}
