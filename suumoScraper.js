function fetchAndSaveSuumoData() {
  const domain = "https://suumo.jp";
  const parameter = "/jj/chintai/ichiran/FR301FC001/?" +
                  "ar=030&" + // エリア
                  "bs=040&" + // 物件種別
                  "pc=50&" + // ページネーション
                  "smk=&" +
                  "po1=25&" + // 並び替え: 25=おすすめ順, 12= 賃料+管理費が安い順, ....
                  "po2=99&" +
                  "shkr1=03&" +
                  "shkr2=03&" +
                  "shkr3=03&" +
                  "shkr4=03&" +
                  "rsnflg=1&" +
                  "ra=013&" +
                  "cb=12.0&" + // 賃料 12万円以上
                  "ct=20.0&" + // 賃料 20万円
                  "md=06&" + // 間取り 2DK
                  "md=07&" + // 間取り 2LDK
                  "et=15&" + // 駅徒歩 15分以内
                  "mb=40&" + // 専有面積 40m2以上
                  "mt=9999999&" + // 専有面積 9999m2以下(上限なし)
                  "cn=9999999&" + // 築年数 9999年以下(上限なし)
                  "tc=0400301&"; // こだわり条件: 風呂トイレ別

  // ekコードとrnコードをセットにして配列に入れる
  const stationParameters = [
    ["品川", "rn=0005&ek=000517460&"], // 山手線 品川
    ["大井町", "rn=0125&ek=012505480&"], // 京浜東北線 大井町
    ["大森駅", "rn=0125&ek=012506360&"], // 京浜東北線 大森
    ["蒲田", "rn=0125&ek=012508940&"], // 京浜東北線 蒲田
    ["青物横丁", "rn=0095&ek=009500240&"], // 京急本線 青物横丁
    ["鮫洲", "rn=0095&ek=009516530&"], // 京急本線 鮫洲
    ["立会川", "rn=0095&ek=009523090&"], // 京急本線 立会川
    ["京急蒲田", "rn=0095&ek=009513410&"], // 京急本線 京急蒲田
  ]
  const stationName = 0;
  const stationCode = 1;

  // Add headers
  const headers = ["物件名", "住所", "最寄り駅", "築年数", "階数", "階", "賃料", "管理費", "敷金", "礼金", "間取り", "専有面積", "物件ID", "詳細URL"];

  try {
    // 各駅ごとにデータを取得する
    stationParameters.forEach((station, index) => {

      let sheetName = station[stationName];
      let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = spreadsheet.getSheetByName(sheetName);

      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
      } else {
        sheet.clear();
      }

      sheet.appendRow(headers);
      // 賃料、管理費の表示形式を通貨形式に設定し、端数を切り捨て
      const currencyColumns = sheet.getRange("G:H"); // 賃料の列はG列目
      currencyColumns.setNumberFormat('¥#,##0');

      let page = 1;
      let hasNextPage = true;

      while (hasNextPage) {
        const url = domain + parameter + station[stationCode] + ":fw2=" + "&page=" + page;
        Logger.log("Fetching data from URL: " + url);

        const response = UrlFetchApp.fetch(url);
        const htmlContent = response.getContentText();

        const properties = htmlContent.match(/<div class="cassetteitem">(.*?)<\/table>/gs);
        if (!properties) {
          Logger.log("No properties found in the HTML.");
          return;
        }

        properties.forEach(property => {
          // 物件名
          const propertyName = property.match(/<div class="cassetteitem_content-title">(.*?)<\/div>/)?.[1]?.trim();

          // 住所
          const address = property.match(/<li class="cassetteitem_detail-col1">(.*?)<\/li>/)?.[1]?.trim();

          // 最寄り駅
          const nearestStations = [...property.matchAll(/<div class="cassetteitem_detail-text"(?:\s+style="[^"]*")?>(.*?)<\/div>/g)]
            .map(match => match[1]?.trim())
            .filter(station => station); // Filter out empty strings
          const nearestStation = nearestStations.join("\n");

          // 築年数
          const age = property.match(/<div>(築.*?年|新築)<\/div>/)?.[1]?.trim();

          // 階数
          const floors = property.match(/<div>(\d+階建)<\/div>/)?.[1]?.trim();

          const details = property.match(/<tr class="js-cassette_link">(.*?)<\/tr>/gs);
          if (!details) return;

          details.forEach(detail => {

            // 階
            const floor = detail.match(/<td>(\d+階)<\/td>/)?.[1]?.trim();

            // 賃料
            // const rent = detail.match(/<span class="cassetteitem_other-emphasis ui-text--bold">(.*?)<\/span>/)?.[1]?.trim();
            const rentMatch = detail.match(/<span class="cassetteitem_other-emphasis ui-text--bold">(.*?)<\/span>/)?.[1]?.trim();
            let rent = rentMatch ? rentMatch.replace(/[^\d.]/g, '') : null;
            if (rent) {
              rent = parseFloat(rent) * 10000;
            }

            // 管理費
            const adminFeeMatch = detail.match(/<span class="cassetteitem_price cassetteitem_price--administration">(.*?)<\/span>/)?.[1]?.trim();
            let adminFee = adminFeeMatch ? adminFeeMatch.replace(/[^\d.]/g, '') : null;
            if (adminFee) {
              adminFee = parseFloat(adminFee);
            }
            // 敷金, 礼金
            const deposit = detail.match(/<span class="cassetteitem_price cassetteitem_price--deposit">(.*?)<\/span>/)?.[1]?.trim() || "-";
            const gratuity = detail.match(/<span class="cassetteitem_price cassetteitem_price--gratuity">(.*?)<\/span>/)?.[1]?.trim() || "-";

            // 間取り
            const layout = detail.match(/<span class="cassetteitem_madori">(.*?)<\/span>/)?.[1]?.trim();

            // 専有面積
            const area = detail.match(/<span class="cassetteitem_menseki">(.*?)<sup>.<\/sup><\/span>/)?.[1]?.trim();

            // 物件ID
            const propertyId = detail.match(/value="(\d+)"/)?.[1]?.trim();

            // 詳細URL
            const detailUrl = detail.match(/<a href="(\/chintai\/.*?)"/)?.[1]?.trim();

            // Append to sheet
            sheet.appendRow([
              propertyName, address, nearestStation, age, floors, floor, rent, adminFee, deposit, gratuity, layout, area, propertyId, "https://suumo.jp" + detailUrl
            ]);
          });
        });

        hasNextPage = htmlContent.includes(`page=${page + 1}`);
        page++;
      }

      Logger.log("データの取得が完了しました。");

    });
  } catch (e) {
    Logger.log("Error fetching data: " + e.message);
  }
}