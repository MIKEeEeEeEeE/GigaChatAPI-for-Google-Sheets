function generateUUID() {
  // Создаём массив байтов
  var bytes = [];
  for (var i = 0; i < 16; i++) {
    bytes.push(Math.floor(Math.random() * 256));
  }

  // Устанавливаем версию UUID (4) и варианты
  bytes[6] = (bytes[6] & 0x0f) | 0x40; // версия 4
  bytes[8] = (bytes[8] & 0x3f) | 0x80; // вариант

  // Преобразуем байты в строку формата UUID
  return bytes.map(function(byte, index) {
    var hex = byte.toString(16).padStart(2, '0');
    if (index === 4 || index === 6 || index === 8 || index === 10) {
      return '-' + hex;
    }
    return hex;
  }).join('');
}

function getGigaChatToken() {
  var scope, secret;
  if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    secret = PropertiesService.getUserProperties().getProperty("px22MjrEcx3qjpcuMsEkNgTvxkpCzwMw"); 
    scope = PropertiesService.getUserProperties().getProperty("FQ6LEHgYragvzZMuvr7m4AT5wvX5T78r");
  }
  else if (PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    secret = PropertiesService.getDocumentProperties().getProperty("px22MjrEcx3qjpcuMsEkNgTvxkpCzwMw"); 
    scope = PropertiesService.getDocumentProperties().getProperty("FQ6LEHgYragvzZMuvr7m4AT5wvX5T78r");
  }
  else {
    throw "getGigaChatToken exception";
  }    
  const response = UrlFetchApp.fetch(
    "https://ngw.devices.sberbank.ru:9443/api/v2/oauth", {
      "method": "POST",
      "validateHttpsCertificates": false,
      "muteHttpExceptions": true,
      "headers": {
        "RqUID": generateUUID(),
        "Authorization": `Basic ${secret}`
      },
      "payload": {
        "scope": scope
      }
    }
  );
  if (response.getResponseCode() == 200) {
    const data = JSON.parse(response);
    const token = data.access_token;
    if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
      secret = PropertiesService.getUserProperties().setProperty("bx4kQGdfTxr787HsEgDgch979zwXJMwv", token);
    }
    else if (PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
      secret = PropertiesService.getDocumentProperties().setProperty("bx4kQGdfTxr787HsEgDgch979zwXJMwv", token);
    }
    return "Success";
  }
  else {
    throw "Ошибка получения токена";
  }
};

function getGigaChatModels() {
  var token;
  if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    token = PropertiesService.getUserProperties().getProperty("bx4kQGdfTxr787HsEgDgch979zwXJMwv");
  }
  else if (PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) { 
    token = PropertiesService.getDocumentProperties().getProperty("bx4kQGdfTxr787HsEgDgch979zwXJMwv");
  }
  else {
    throw "getGigaChatModels exception";
  }
  const response = UrlFetchApp.fetch(
    "https://gigachat.devices.sberbank.ru/api/v1/models", {
      "method": "GET",
      "validateHttpsCertificates": false,
      "muteHttpExceptions": true,
      "headers": {
        "Authorization": `Bearer ${token}`
      }
    }
  ); 
  if (response.getResponseCode() == 200) {
    const data = JSON.parse(response);
    const models = data.data;
    return models.map(item => item.id);
  }
  else {
    throw response;
  }
};

function askGigaChat(template, ...args) {
  var token, model;
  if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    token = PropertiesService.getUserProperties().getProperty("bx4kQGdfTxr787HsEgDgch979zwXJMwv");
    model = PropertiesService.getUserProperties().getProperty("BLduEZTnWw3RuwPvn6ju5zJ5J7PrzfDv");
  }
  else if (PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) { 
    token = PropertiesService.getDocumentProperties().getProperty("bx4kQGdfTxr787HsEgDgch979zwXJMwv");
    model = PropertiesService.getDocumentProperties().getProperty("BLduEZTnWw3RuwPvn6ju5zJ5J7PrzfDv");
  }
  else {
    throw "askGigaChat exception";
  }
  try {
    for (index = 0; index < args.length; index++) {
      template = template.replace(`{{${index}}}`, args[index]);
    }
  }
  catch(e) {
    throw e;
  }
  const response = UrlFetchApp.fetch(
    "https://gigachat.devices.sberbank.ru/api/v1/chat/completions", {
      "method": "POST",
      "validateHttpsCertificates": false,
      "muteHttpExceptions": true,
      "headers": {
        "Authorization": `Bearer ${token}`
      },
      "payload": JSON.stringify({
        "model": model,
        "messages": [
          {
            "role": "user",
            "content": template
          }
        ]
      })
    }
  );
  if (response.getResponseCode() == 200) {
    const data = JSON.parse(response);
    const content = data.choices[0].message.content;
    return content;
  }
  else {
    throw response;
  }
};

function translateGigaChat(cell, into="русский") {
  var token, model;
  if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    token = PropertiesService.getUserProperties().getProperty("bx4kQGdfTxr787HsEgDgch979zwXJMwv");
    model = PropertiesService.getUserProperties().getProperty("BLduEZTnWw3RuwPvn6ju5zJ5J7PrzfDv");
  }
  else if (PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    token = PropertiesService.getDocumentProperties().getProperty("bx4kQGdfTxr787HsEgDgch979zwXJMwv");
    model = PropertiesService.getDocumentProperties().getProperty("BLduEZTnWw3RuwPvn6ju5zJ5J7PrzfDv");
  }
  else {
    throw "translateGigaChat exception";
  }
  const response = UrlFetchApp.fetch(
    "https://gigachat.devices.sberbank.ru/api/v1/chat/completions", {
      "method": "POST",
      "validateHttpsCertificates": false,
      "muteHttpExceptions": true,
      "headers": {
        "Authorization": `Bearer ${token}`
      },
      "payload": JSON.stringify({
        "model": model,
        "messages": [
          {
            "role": "system",
            "content": `Переведи текст на ${into} язык`
          },
          {
            "role": "user",
            "content": cell
          }
        ]
      })
    }
  );
  if (response.getResponseCode() == 200) {
    const data = JSON.parse(response);
    const content = data.choices[0].message.content;
    return content;
  }
  else {
    throw response;
  }
};

function makeGigaChatTrigger() {
  ScriptApp.newTrigger("getGigaChatToken")
    .timeBased()
    .everyMinutes(30)
    .create();
};

function authGigaChat() {
  var ui = SpreadsheetApp.getUi();
  var secret = ui.prompt("Авторизационные данные:").getResponseText();
  if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getUserProperties().setProperty("px22MjrEcx3qjpcuMsEkNgTvxkpCzwMw", secret);
  } 
  else if (PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getDocumentProperties().setProperty("px22MjrEcx3qjpcuMsEkNgTvxkpCzwMw", secret);
  }
  else {
    throw "authGigaChat exception";
  }
  if (getGigaChatToken() == "Success") {
    ui.alert("Токен получен!");
    makeGigaChatTrigger();
  };
};

function setUserProps() {
  PropertiesService.getUserProperties().setProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu", 1);
};

function setDocumentProps() {
  PropertiesService.getDocumentProperties().setProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu", 1);
};

function setPersScope() {
  if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getUserProperties().setProperty("FQ6LEHgYragvzZMuvr7m4AT5wvX5T78r", "GIGACHAT_API_PERS");
  } 
  else if (PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getDocumentProperties().setProperty("FQ6LEHgYragvzZMuvr7m4AT5wvX5T78r", "GIGACHAT_API_PERS");
  }
  else {
    throw "setPersScope exception";
  }
};

function setB2BScope() {
  if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getUserProperties().setProperty("FQ6LEHgYragvzZMuvr7m4AT5wvX5T78r", "GIGACHAT_API_B2B");
  }
  else if (PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getDocumentProperties().setProperty("FQ6LEHgYragvzZMuvr7m4AT5wvX5T78r", "GIGACHAT_API_B2B");
  }
  else {
    throw "setB2BScope exception";
  }
};

function setCorpScope() {
  if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getUserProperties().setProperty("FQ6LEHgYragvzZMuvr7m4AT5wvX5T78r", "GIGACHAT_API_CORP");
  }
  else if (PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getUserProperties().setProperty("FQ6LEHgYragvzZMuvr7m4AT5wvX5T78r", "GIGACHAT_API_CORP");
  }
  else {
    throw "setCorpScope exception";
  }
};

function setGigaChat() {
  if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getUserProperties().setProperty("BLduEZTnWw3RuwPvn6ju5zJ5J7PrzfDv", "GigaChat");
  }
  else if(PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getDocumentProperties().setProperty("BLduEZTnWw3RuwPvn6ju5zJ5J7PrzfDv", "GigaChat");
  }
  else {
    throw "setGigaChat exception";
  }
};

function setGigaChatPlus() {
  if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getUserProperties().setProperty("BLduEZTnWw3RuwPvn6ju5zJ5J7PrzfDv", "GigaChat-Plus");
  }
  else if(PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getDocumentProperties().setProperty("BLduEZTnWw3RuwPvn6ju5zJ5J7PrzfDv", "GigaChat-Plus");
  }
  else {
    throw "setGigaChatPlus exception";
  }
};

function setGigaChatPro() {
  if (PropertiesService.getUserProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getUserProperties().setProperty("BLduEZTnWw3RuwPvn6ju5zJ5J7PrzfDv", "GigaChat-Pro");
  }
  else if(PropertiesService.getDocumentProperties().getProperty("MteUVTbxPEfHfCph4BH7Lkzhu7ksQYCu")) {
    PropertiesService.getDocumentProperties().setProperty("BLduEZTnWw3RuwPvn6ju5zJ5J7PrzfDv", "GigaChat-Pro");
  }
  else {
    throw "setGigaChatPro exception";
  }
};
 
function onOpen(e){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("GigaChatAPI")
    .addSubMenu(ui.createMenu("Уровень доступа")
                  .addItem("Пользователь", "setUserProps")
                  .addItem("Документ", "setDocumentProps"))
    .addSeparator()
    .addSubMenu(ui.createMenu("Версия API")
                  .addItem("GIGACHAT_API_PERS", "setPersScope")
                  .addItem("GIGACHAT_API_B2B", "setB2BScope")
                  .addItem("GIGACHAT_API_CORP", "setCorpScope"))
    .addSeparator()
    .addSubMenu(ui.createMenu("Модель GigaChat")
                  .addItem("GigaChat", "setGigaChat")
                  .addItem("GigaChat-Plus", "setGigaChatPlus")
                  .addItem("GigaChat-Pro", "setGigaChatPro"))
    .addSeparator()
    .addItem("Авторизация", "authGigaChat")
    .addToUi();
};

function onInstall(e) {
  onOpen(e);
}
