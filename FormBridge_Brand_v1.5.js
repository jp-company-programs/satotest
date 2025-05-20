(function() {
  'use strict';

  // フォーム表示時のログ（必須ではありません）
  formBridge.events.on('form.show', function (context) {
    console.log('フォーム表示完了');
  });

  // 送信時のバリデーションをひとつのハンドラにまとめる
  formBridge.events.on('form.submit', function (context) {
    const record = context.getRecord();

    // 「メインブランド」か「ブランド」のいずれか必須
    const brandMainValue = record['ドロップダウン_16'].value;
    const brandFreeValue = record['文字列__1行__1'].value;
    console.log(brandMainValue)
    console.log(brandFreeValue)
    if (!brandMainValue && !brandFreeValue) {
      alert('「メインブランド」か「ブランド」のどちらかは記載してください。');
      context.preventDefault();
      return;
    }

    // 「種類」が Ring のときは「サイズフリー」必須
    const syurui = record['ドロップダウン_4']?.value;
    const sizeFree = record['文字列__1行__11']?.value;
    if (syurui === 'Ring' && !sizeFree) {
      alert('「Ring」の場合、「サイズフリー」は必須です。');
      context.preventDefault();
      return;
    }

    // ここまで来たらエラーなし → true を返せば送信を続行
    return true;
  });

})();
