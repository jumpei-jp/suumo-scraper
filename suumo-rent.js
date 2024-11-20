function scrapeSuumo() {
    // 1. Suumoの賃貸情報ページのURL
    var url = 'https://suumo.jp/jj/chintai/ichiran/FR301FC001/?ar=030&bs=040&pc=30&smk=&po1=25&po2=99&shkr1=03&shkr2=03&shkr3=03&shkr4=03&sc=13109&ta=13&cb=13.0&ct=20.0&md=04&md=06&md=07&et=9999999&mb=40&mt=9999999&cn=9999999&tc=0400301&fw2=';

    // 2. URLからHTMLを取得
    var response = UrlFetchApp.fetch(url);
    var html = response.getContentText();

    // 物件名の取得
    var propertyNameRegex = /<div class="cassetteitem_content-title">(.*?)<\/div>/g;
    var propertyMatch = propertyNameRegex.exec(html);
    var propertyName = propertyMatch ? propertyMatch[1].trim() : "-";

    // 所在地の取得
    var locationRegex = /<li class="cassetteitem_detail-col1">(.*?)<\/li>/g;
    var locationMatch = locationRegex.exec(html);
    var location = locationMatch ? locationMatch[1].trim() : "-";

    // 最寄駅と駅距離の取得
    var stationRegex = /<div class="cassetteitem_detail-text">(.*?)<\/div>/g;
    var stations = [];
    var stationMatch;
    while (stationMatch = stationRegex.exec(html)) {
        stations.push(stationMatch[1].trim());
    }
    var stationInfo = stations.join(", ") || "-";

    // 築年数と階数の取得
    var buildingInfoRegex = /<li class="cassetteitem_detail-col3">.*?<div>(.*?)<\/div>.*?<div>(.*?)<\/div>/g;
    var buildingInfoMatch = buildingInfoRegex.exec(html);
    var age = buildingInfoMatch ? buildingInfoMatch[1].trim() : "-";
    var floors = buildingInfoMatch ? buildingInfoMatch[2].trim() : "-";

    // 部屋情報の取得
    var roomRegex = /<tr class="js-cassette_link">.*?<td.*?>.*?<span.*?class="cassetteitem_price cassetteitem_price--rent">.*?<span.*?>(.*?)<\/span>.*?<\/td>.*?<td.*?>.*?<span.*?>(.*?)<\/span>.*?<\/td>.*?<td.*?>.*?<span.*?>(.*?)<\/span>.*?<\/td>.*?<\/tr>/gs;
    var roomDetails = [];
    var roomMatch;
    while (roomMatch = roomRegex.exec(html)) {
        roomDetails.push({
            floor: roomMatch[1].trim(),
            rent: roomMatch[2].trim(),
            managementFee: roomMatch[3].trim()
        });
    }

    // Googleスプレッドシートに出力
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // ヘッダーの設定
    sheet.getRange('A1').setValue('物件名');
    sheet.getRange('B1').setValue('所在地');
    sheet.getRange('C1').setValue('最寄駅と駅距離');
    sheet.getRange('D1').setValue('築年数');
    sheet.getRange('E1').setValue('階数');
    sheet.getRange('F1').setValue('階');
    sheet.getRange('G1').setValue('賃料');
    sheet.getRange('H1').setValue('管理費');
    sheet.getRange('I1').setValue('敷金');
    sheet.getRange('J1').setValue('礼金');
    sheet.getRange('K1').setValue('間取り');
    sheet.getRange('L1').setValue('面積');

    // データの入力
    var startRow = 2;  // データは2行目から開始
    for (var i = 0; i < roomDetails.length; i++) {
      sheet.getRange(startRow + i, 1).setValue(propertyName);        // 物件名
      sheet.getRange(startRow + i, 2).setValue(location);           // 所在地
      sheet.getRange(startRow + i, 3).setValue(stationInfo);        // 最寄駅と駅距離
      sheet.getRange(startRow + i, 4).setValue(age);               // 築年数
      sheet.getRange(startRow + i, 5).setValue(floors);            // 階数
      sheet.getRange(startRow + i, 6).setValue(roomDetails[i].floor);  // 階
      sheet.getRange(startRow + i, 7).setValue(roomDetails[i].rent);   // 賃料
      sheet.getRange(startRow + i, 8).setValue(roomDetails[i].managementFee); // 管理費
      sheet.getRange(startRow + i, 9).setValue('-');                // 敷金（ここに処理を追加できます）
      sheet.getRange(startRow + i, 10).setValue('-');               // 礼金（ここに処理を追加できます）
      sheet.getRange(startRow + i, 11).setValue('-');               // 間取り（ここに処理を追加できます）
      sheet.getRange(startRow + i, 12).setValue('-');               // 面積（ここに処理を追加できます）
    }
}