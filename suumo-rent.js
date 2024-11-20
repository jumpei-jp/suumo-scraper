function fetchAndSaveSuumoData() {
  const url = "https://suumo.jp/jj/chintai/ichiran/FR301FC001/?ar=030&bs=040&pc=30&smk=&po1=25&po2=99&shkr1=03&shkr2=03&shkr3=03&shkr4=03&rsnflg=1&rn=0125&rn=0095&rn=0005&ek=012508940&ek=009513410&ek=012506360&ek=012505480&ek=009500240&ek=009523090&ek=009516530&ek=000517460&ra=013&cb=12.0&ct=20.0&md=06&md=07&et=15&mb=40&mt=9999999&cn=9999999&tc=0400301&fw2=";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  sheet.clear(); // Clear existing content in the sheet

  // Add headers
  const headers = ["物件名", "住所", "最寄り駅", "築年数", "階数", "階", "賃料", "管理費", "敷金", "礼金", "間取り", "専有面積", "物件ID", "詳細URL"];
  sheet.appendRow(headers);

  try {
    // Fetch HTML content from the URL
    const response = UrlFetchApp.fetch(url);
    const htmlContent = response.getContentText();

    // Parse HTML using regex to extract data
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
        const rent = detail.match(/<span class="cassetteitem_other-emphasis ui-text--bold">(.*?)<\/span>/)?.[1]?.trim();

        // 管理費
        const adminFee = detail.match(/<span class="cassetteitem_price cassetteitem_price--administration">(.*?)<\/span>/)?.[1]?.trim();

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

    Logger.log("Data fetched and saved to Google Sheet.");
  } catch (e) {
    Logger.log("Error fetching data: " + e.message);
  }
}