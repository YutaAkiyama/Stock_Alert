function myFunction() {
  // シート情報にアクセスできるようにします
  var sheet=SpreadsheetApp.getActiveSheet();
  var APP = sheet.getRange('B2').getValue();
  var DEV =  sheet.getRange('B3').getValue();
  var CERT = sheet.getRange('B4').getValue();
  var Token = sheet.getRange('B5').getValue();
  // 最後の行数を取得します
  var rowcount = sheet.getLastRow();
  // 開始行から最終行までまわします
  
  for(var i=8; i<rowcount+1; i++){
      url = sheet.getRange(i,1).getValue();
      keyword = sheet.getRange(i,2).getValue();
      ItemID = sheet.getRange(i,4).getValue();
      result = check_url(url,keyword);
      if(result){
      // 想定通りならH列に◯を
         sheet.getRange(i,3).setValue('○');
      } else {
      // 想定と違ったら×を
        sheet.getRange(i,3).setValue('☓');
        try {
          var endpoint = "https://api.ebay.com/ws/api.dll"
          var headers = {'X-EBAY-API-COMPATIBILITY-LEVEL': '1149',
                         'X-EBAY-API-DEV-NAME': DEV ,
                         'X-EBAY-API-APP-NAME': APP,
                         'X-EBAY-API-CERT-NAME': CERT,
                         'X-EBAY-API-CALL-NAME': 'EndItem',
                         'X-EBAY-API-SITEID': '0',
                         'Content-Type': 'text/xml'}
          
          var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><EndItemRequest xmlns=\"urn:ebay:apis:eBLBaseComponents\"><RequesterCredentials><eBayAuthToken>" + Token + "</eBayAuthToken></RequesterCredentials><EndingReason>NotAvailable</EndingReason><ItemID>" + ItemID + "</ItemID></EndItemRequest>"
          var options = {
            'method' : 'post',
            'headers' : headers,
            'payload' : xml
          };
          response = UrlFetchApp.fetch(endpoint,options)
          var responseBody = response.getContentText();
          if (responseBody.search("invalid") !== -1) {
              sheet.getRange(i,5).setValue("入力データが無効");
          } 
          else　if(responseBody.search("closed") !== -1){
            sheet.getRange(i,5).setValue("END済");
         } 
         else　if(responseBody.search("Item cannot be accessed") !== -1){
            sheet.getRange(i,5).setValue("ItemIDが無効");
         } 
         else{
            sheet.getRange(i,5).setValue("END成功！");
         } 
          
        } catch (e){
          sheet.getRange(i,5).setValue("");
        }
        
      }
  }
  }
 
 function check_url(url,keyword){
     html = get_html(url);
  
     if (html.match(keyword)) {
     
        return true;
      }
   
     else{
      return false;
       
     }
    }  
 
  
 function get_html(url) {
    try{
       Utilities.sleep(1000);
       var response = UrlFetchApp.fetch(url).getContentText();
       return response
     }
    catch(e){
       return '';
     }
   }
    
 