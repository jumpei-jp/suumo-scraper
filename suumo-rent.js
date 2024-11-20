/**
 * 指定されたSUUMOのURLから物件情報をスクレイピングし、スプレッドシートに出力します。
 */
function scrapeSuumoProperties() {
    // スプレッドシートとシートの取得
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = '物件情報'; // 出力先のシート名
    var sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      // シートが存在しない場合は新規作成
      sheet = spreadsheet.insertSheet(sheetName);
    } else {
      // シートが存在する場合は内容をクリア
      sheet.clearContents();
    }
    
    // ヘッダーの設定
    var headers = [
      '物件名',
      '所在地',
      '最寄り駅と駅距離',
      '築年数',
      '階数',
      '階',
      '賃料',
      '管理費',
      '敷金',
      '礼金',
      '間取り',
      '面積'
    ];
    sheet.appendRow(headers);
    
    // スクレイピング対象のURL
    var url = 'https://suumo.jp/jj/chintai/ichiran/FR301FC001/?ar=030&bs=040&pc=30&smk=&po1=25&po2=99&shkr1=03&shkr2=03&shkr3=03&shkr4=03&sc=13109&ta=13&cb=13.0&ct=20.0&md=04&md=06&md=07&et=9999999&mb=40&mt=9999999&cn=9999999&tc=0400301&fw2=';
    
    try {
      // HTMLコンテンツの取得
      var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      var htmlContent = response.getContentText('UTF-8');
      
      // プロパティの配列を初期化
      var properties = [];
      
      // 各物件を抽出するための正規表現
      var trRegex = /<tr class="js-cassette_link">([\s\S]*?)<\/tr>/g;
      var match;
      
      while ((match = trRegex.exec(htmlContent)) !== null) {
        var trContent = match[1];
        
        // 物件名の抽出
        var nameMatch = trContent.match(/<div class="cassetteitem_content-title">([^<]+)<\/div>/);
        var propertyName = nameMatch ? nameMatch[1].trim() : '';
        
        // 所在地の抽出
        var locationMatch = trContent.match(/<li class="cassetteitem_detail-col1">([^<]+)<\/li>/);
        var location = locationMatch ? locationMatch[1].trim() : '';
        
        // 最寄り駅と駅距離の抽出
        var stationMatches = trContent.match(/<div class="cassetteitem_detail-text">([^<]+)<\/div>/g);
        var nearestStations = '';
        if (stationMatches) {
          var stations = stationMatches.map(function(s) {
            var stationName = s.match(/<div class="cassetteitem_detail-text">([^<]+)<\/div>/)[1].trim();
            return stationName;
          });
          nearestStations = stations.join('; ');
        }
        
        // 築年数の抽出
        var ageMatch = trContent.match(/<div>(築[^<]+)<\/div>/);
        var age = ageMatch ? ageMatch[1].trim() : '';
        
        // 階数の抽出
        var floorMatch = trContent.match(/<div>(\d+)階建<\/div>/);
        var numberOfFloors = floorMatch ? floorMatch[1].trim() : '';
        
        // 階の抽出
        var floorMatchAlt = trContent.match(/<td>\s*<div>(\d+)階<\/div>\s*<\/td>/);
        var floor = floorMatchAlt ? floorMatchAlt[1] + '階' : '';
        
        // 賃料の抽出
        var rentMatch = trContent.match(/cassetteitem_price--rent">[^<]*<span[^>]*>([^<]+)<\/span>/);
        var rent = rentMatch ? rentMatch[1].trim() : '';
        
        // 管理費の抽出
        var adminFeeMatch = trContent.match(/cassetteitem_price cassetteitem_price--administration">([^<]+)<\/span>/);
        var adminFee = adminFeeMatch ? adminFeeMatch[1].trim() : '';
        
        // 敷金の抽出
        var depositMatch = trContent.match(/cassetteitem_price cassetteitem_price--deposit">([^<]+)<\/span>/);
        var deposit = depositMatch ? depositMatch[1].trim() : '';
        
        // 礼金の抽出
        var keyMoneyMatch = trContent.match(/cassetteitem_price cassetteitem_price--gratuity">([^<]+)<\/span>/);
        var keyMoney = keyMoneyMatch ? keyMoneyMatch[1].trim() : '';
        
        // 間取りの抽出
        var layoutMatch = trContent.match(/cassetteitem_madori">([^<]+)<\/span>/);
        var layout = layoutMatch ? layoutMatch[1].trim() : '';

        // 面積の抽出
        var areaMatch = trContent.match(/cassetteitem_menseki">([^<]+)<\/span>/);
        var area = areaMatch ? areaMatch[1].trim() : '';
        
        // プロパティ配列に追加
        properties.push([
          propertyName,
          location,
          nearestStations,
          age,
          numberOfFloors,
          floor,
          rent,
          adminFee,
          deposit,
          keyMoney,
          layout,
          area
        ]);
      }
      
      // スプレッドシートにデータを書き込む
      if (properties.length > 0) {
        sheet.getRange(2, 1, properties.length, headers.length).setValues(properties);
        SpreadsheetApp.flush();
        Logger.log('スクレイピング完了: ' + properties.length + '件の物件情報を取得しました。');
        SpreadsheetApp.getUi().alert('スクレイピング完了: ' + properties.length + '件の物件情報を取得しました。');
      } else {
        Logger.log('物件情報が見つかりませんでした。');
        SpreadsheetApp.getUi().alert('物件情報が見つかりませんでした。');
      }
      
    } catch (error) {
      Logger.log('エラーが発生しました: ' + error.message);
      SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
    }
  }