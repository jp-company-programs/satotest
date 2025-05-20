
//テスト
(() => {
    'use strick';

    let syurui = "ドロップダウン_4";
    let sizeFree = "文字列__1行__11";

    kintone.events.on([
        'app.record.create.change.${syurui}',
        'app.record.edit.change.${syurui}'

    ], (event) => {
        let record = event.record;
        let selectType = record[syurui].value;

        if (selectType === "Ring") {
            record[sizeFree].required = true;  // 種類が「Ring」の場合、サイズフリーは必須
          } else {
            record[sizeFree].required = false; // 「Ring」以外はサイズフリーの必須を解除
          }
      
          return event;
    });

    kintone.events.on([
        'app.record.create.submit',
        'app.record.edit.submit'    

    ], (event) => {
        let record = event.record;
        let selectType = record[syurui].value;
        let sizeValue = record[sizeFree].value;
        let brandMain = record.ドロップダウン_16;
        let brandFree = record.文字列__1行__1;

        if (selectType === "Ring" && !sizeValue) {
            event.error ='種類で「Ring」を選択した場合、サイズフリーは必須です。';
            return event;  // エラーがある場合、保存をキャンセルする
        }
       if (!brandMain.value && !brandFree.value) {
            event.error = '「メインブランド」か「ブランド」のいずれかの入力欄に入力してください。';
            return event;  // エラーがある場合、保存をキャンセルする
        }

        return event;  // エラーがなければそのまま登録・更新
    });

})();