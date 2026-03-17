// 1. ฟังก์ชันสำหรับเปิดหน้าเว็บ (รองรับ Parameter ?page=)
function doGet(e) {
  var page = e.parameter.page;
  
  if (page === 'manage') {
    return HtmlService.createHtmlOutputFromFile('Manage')
      .setTitle('ระบบจัดการข้อมูล - BUEMFOOD')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      
  } else if (page === 'track') {
    return HtmlService.createHtmlOutputFromFile('Track')
      .setTitle('เช็คสถานะพัสดุ - BUEMFOOD')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      
  } else {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('เพิ่มออเดอร์ใหม่ - BUEMFOOD')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

// ฟังก์ชันดึงลิงก์แอป
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

// 2. ฟังก์ชันประมวลผล Match เลขพัสดุอัตโนมัติ
function processTrackingText(rawText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("Order_Data");
  const orderValues = orderSheet.getDataRange().getValues();
  const lines = rawText.split('\n');
  
  let results = { matched: 0, total: 0, unmatched: [], potential: [] };
  const clean = (n) => n.toString().toLowerCase().replace(/คุณ|นาย|นาง|นางสาว|เด็กชาย|เด็กหญิง/g, "").replace(/\s+/g, "").trim();

  lines.forEach(line => {
    if (!line.trim()) return;
    results.total++;
    let parts = line.trim().split(/\s+/);
    let track = parts.pop();
    let carrierName = parts.join(" ");
    let cleanCarrier = clean(carrierName);

    let found = false;
    let potentialForThisLine = [];

    for (let i = 1; i < orderValues.length; i++) {
      let cleanSheet = clean(orderValues[i][1]);
      if (cleanSheet === "") continue;

      if (cleanSheet === cleanCarrier) {
        orderSheet.getRange(i + 1, 8).setValue(track);
        orderSheet.getRange(i + 1, 1, 1, 9).setBackground("#d9ead3");
        results.matched++;
        found = true;
        break;
      }
      
      if (cleanSheet.includes(cleanCarrier) || cleanCarrier.includes(cleanSheet)) {
        potentialForThisLine.push({ 
          rowIndex: i + 1, sheetName: orderValues[i][1], carrierName: carrierName, track: track 
        });
      }
    }

    if (!found) {
      if (potentialForThisLine.length > 0) results.potential.push(potentialForThisLine[0]);
      else results.unmatched.push({ name: carrierName, track: track });
    }
  });
  return results;
}

// ฟังก์ชันยืนยันการ Match
function confirmSuggestedMatch(rowIndex, finalName, track) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order_Data");
  sheet.getRange(rowIndex, 2).setValue(finalName);
  sheet.getRange(rowIndex, 8).setValue(track);
  sheet.getRange(rowIndex, 1, 1, 9).setBackground("#d9ead3");
  return "success";
}

// 3. ฟังก์ชันบันทึกออเดอร์แบบชุด (Bulk)
function saveBulkOrders(rawOrderText) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Order_Data");
    if (!sheet) return { status: "error", message: "ไม่พบแท็บ Order_Data" };

    const blocks = rawOrderText.split(/(?=\d*คุณ)/g).filter(b => b.trim().length > 10);
    let lastRow = sheet.getLastRow();
    let lastId = (lastRow <= 1) ? 0 : Number(sheet.getRange(lastRow, 1).getValue()) || 0;
    let successCount = 0;

    blocks.forEach(block => {
      let lines = block.trim().split('\n').filter(l => l.trim() !== "");
      if (lines.length < 2) return;

      let rawName = lines[0].replace(/^\d+/, "").trim();
      let name = rawName;
      let mid = Math.floor(rawName.length / 2);
      if (rawName.substring(0, mid) === rawName.substring(mid)) { name = rawName.substring(0, mid); }

      let address = lines[1] ? lines[1].trim() : "";
      let rest = lines.slice(2).join(" ");
      let phoneMatch = rest.match(/(\d{2,3}-\d{3,4}-\d{4}|\d{10})/);
      let phone = phoneMatch ? phoneMatch[0] : "";
      let carriers = ["ไปรษณีย์", "Flash", "J&T", "Kerry", "Best"];
      let carrier = carriers.find(c => rest.includes(c)) || "Flash";
      
      let cod = "-", product = "อกไก่", flavor = "-";
      let afterCarrierParts = rest.split(carrier);
      if (afterCarrierParts.length > 1) {
        let afterCarrier = afterCarrierParts[1].trim();
        let codRegex = /^(\d+\.-|-)/;
        let codMatch = afterCarrier.match(codRegex);
        if (codMatch) {
          cod = codMatch[0];
          let remaining = afterCarrier.replace(cod, "").trim();
          let flavorKeywords = ["คละ", "รส", "หมาล่า", "สไปซี่", "ออริจินัล", "พริกไทยดำ"];
          let splitIdx = -1;
          for (let kw of flavorKeywords) {
            let idx = remaining.indexOf(kw);
            if (idx !== -1 && (splitIdx === -1 || idx < splitIdx)) splitIdx = idx;
          }
          if (splitIdx !== -1) {
            product = remaining.substring(0, splitIdx).trim();
            flavor = remaining.substring(splitIdx).trim();
          } else { product = remaining || "อกไก่"; }
        } else { product = afterCarrier || "อกไก่"; }
      }

      lastId++;
      sheet.appendRow([lastId, name, address, carrier, cod, product, flavor, "", phone]);
      successCount++;
    });
    return { status: "success", count: successCount };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// 4. ระบบจัดการข้อมูล (Admin Management)
function searchForAdmin(query) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order_Data");
    const data = sheet.getDataRange().getValues();
    const results = [];
    const cleanQuery = query ? query.toString().toLowerCase().trim() : "";
    for (let i = data.length - 1; i >= 1; i--) {
      const name = (data[i][1] || "").toString().toLowerCase();
      const phone = (data[i][8] || "").toString();
      if (cleanQuery === "" || name.includes(cleanQuery) || phone.includes(cleanQuery)) {
        results.push({
          rowIndex: i + 1, id: data[i][0] || "-", name: data[i][1] || "", address: data[i][2] || "",
          carrier: data[i][3] || "", cod: data[i][4] || "", product: data[i][5] || "",
          flavor: data[i][6] || "", track: data[i][7] || "", phone: data[i][8] || ""
        });
      }
      if (results.length >= 50) break;
    }
    return { status: "success", data: results };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function updateOrderRow(rowIndex, newData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order_Data");
    sheet.getRange(rowIndex, 2).setValue(newData.name);
    sheet.getRange(rowIndex, 3).setValue(newData.address);
    sheet.getRange(rowIndex, 4).setValue(newData.carrier);
    sheet.getRange(rowIndex, 5).setValue(newData.cod);
    sheet.getRange(rowIndex, 6).setValue(newData.product);
    sheet.getRange(rowIndex, 7).setValue(newData.flavor);
    sheet.getRange(rowIndex, 8).setValue(newData.track);
    sheet.getRange(rowIndex, 9).setValue(newData.phone);
    return "success";
  } catch (e) { return e.toString(); }
}

function deleteOrderRow(rowIndex) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order_Data");
    sheet.deleteRow(rowIndex);
    return "success";
  } catch (e) { return e.toString(); }
}

function searchByPhone(phoneNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName("Order_Data").getDataRange().getValues();
  const cleanSearch = phoneNumber.replace(/[^0-9]/g, "");
  return data.slice(1).filter(row => row[8].toString().replace(/[^0-9]/g, "").includes(cleanSearch) && cleanSearch.length >= 9)
    .map(row => ({ name: row[1], carrier: row[3], product: row[5], flavor: row[6], track: row[7] }));
}
