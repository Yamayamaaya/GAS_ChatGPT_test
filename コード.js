// API keyとmodelを設定
var apiKey = "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";
var model = "gpt-3.5-turbo";
var url = "https://api.openai.com/v1/chat/completions";

// メニューを追加
function addMenu() {
  DocumentApp.getUi().createMenu("ChatGPT")
  .addItem("記事を生成する", "generateArticle")
  .addToUi();
}

function generateArticle() {
  // 選択テキストの取得　
  var doc = DocumentApp.getActiveDocument();
  var selectedText = doc.getSelection().getRangeElements()[0].getElement().asText().getText()

  // もしテキストが選択されていたら
  if(selectedText){
    var content = selectedText + "についての記事を300文字以内で書いてください";

    // ChatGPTに投げる質問/必要な情報を設定
    var temperature= 0; // 生成する応答の多様性
    var maxTokens = 2048; // 生成する応答の最大トークン数（文字数）

    const requestBody = {
      "model": model,
      "messages": [{'role': 'user', 'content': content}],
      "temperature": temperature,
      "max_tokens": maxTokens,
    };
    const requestOptions = {
      "method": "POST",
      "headers": {
        "Content-Type": "application/json",
        "Authorization": "Bearer "+apiKey
      },
      "payload": JSON.stringify(requestBody)
    }

    // APIを叩いてChatGPTから回答を取得
    var response = UrlFetchApp.fetch(url, requestOptions);
    var responseText = response.getContentText(); 
    var json = JSON.parse(responseText);
    var result_text = json.choices[0].message.content.trim();

    //ドキュメントに出力
    var addparagraph = doc.getBody().appendParagraph(result_text);
  }
}