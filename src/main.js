/**
 * Gmailの領収書を抽出し、Googleドライブの月別フォルダへ自動保存する
 */
function autoArchiveReceipts() {
  const props = PropertiesService.getScriptProperties();
  const CONFIG = {
    TARGET_LABEL: props.getProperty('TARGET_LABEL') || '領収書振り分け',
    DONE_LABEL: props.getProperty('DONE_LABEL') || '領収書',
    PARENT_FOLDER_ID: props.getProperty('PARENT_FOLDER_ID'),
    TIME_ZONE: 'Asia/Tokyo'
  };

  if (!CONFIG.PARENT_FOLDER_ID) {
    console.error('スクリプトプロパティに PARENT_FOLDER_ID が設定されていません。');
    return;
  }

  const srcLabel = GmailApp.getUserLabelByName(CONFIG.TARGET_LABEL);
  if (!srcLabel) {
    console.error(`ラベル「${CONFIG.TARGET_LABEL}」が見つかりません。`);
    return;
  }

  const destLabel = GmailApp.getUserLabelByName(CONFIG.DONE_LABEL) || GmailApp.createLabel(CONFIG.DONE_LABEL);
  console.log(`destLabel: ${destLabel.getName()}`);

  const threads = srcLabel.getThreads();
  console.log(`対象スレッド数: ${threads.length}`);

  const parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);

  threads.forEach(thread => {
    const subject = thread.getFirstMessageSubject();
    try {
      const messages = thread.getMessages();
      console.log(`処理中: "${subject}" (メッセージ数: ${messages.length})`);

      messages.forEach(message => {
        const date = message.getDate();
        const yyyymmdd = Utilities.formatDate(date, CONFIG.TIME_ZONE, 'yyyyMMdd');
        const yearStr = Utilities.formatDate(date, CONFIG.TIME_ZONE, 'yyyy');
        const monthStr = Utilities.formatDate(date, CONFIG.TIME_ZONE, 'MM');

        // 差出人名から企業名を抽出（ファイル名に使えない文字を除去）
        const companyName = message.getFrom()
          .replace(/<.*>/, '')
          .replace(/"/g, '')
          .trim()
          .replace(/[\/\\:*?"<>|,]/g, '_')
          .replace(/\./g, '_')
          .replace(/ /g, '_')
          .replace(/_+/g, '_');

        console.log(`  メッセージ: ${yyyymmdd} / ${companyName}`);

        const yearFolderName = `${yearStr}_${toReiwa(yearStr)}`;
        const yearFolder = getOrCreateFolder(parentFolder, yearFolderName);
        const receiptFolder = getOrCreateFolder(yearFolder, '01_領収書');
        const monthFolder = getOrCreateFolder(receiptFolder, monthStr);

        const attachments = message.getAttachments();
        const pdfs = attachments.filter(a => a.getContentType() === 'application/pdf');
        console.log(`  添付PDF数: ${pdfs.length}`);

        let fileSaved = false;

        // 1. 添付PDFの処理（Receipt優先 → Invoice → 最初のPDF）
        if (pdfs.length > 0) {
          const fileToSave = pdfs.find(p => p.getName().toLowerCase().includes('receipt'))
            || pdfs.find(p => p.getName().toLowerCase().includes('invoice'))
            || pdfs[0];

          const newFileName = `${yyyymmdd}_${companyName}.pdf`;
          monthFolder.createFile(fileToSave).setName(newFileName);
          console.log(`  保存: ${newFileName}`);
          fileSaved = true;
        }

        // 2. 添付がない場合はHTML本文をPDF化（画像をbase64埋め込み）
        if (!fileSaved) {
          const htmlBody = message.getBody();
          const msgSubject = message.getSubject();
          const html = `<html><head><meta charset="utf-8"></head><body>
            <h2>${msgSubject}</h2>
            <p><strong>From:</strong> ${message.getFrom()}</p>
            <p><strong>Date:</strong> ${date}</p>
            <hr>
            ${htmlBody}
          </body></html>`;

          const embeddedHtml = embedImages(html);
          const pdfBlob = Utilities.newBlob(embeddedHtml, 'text/html', `${yyyymmdd}.html`)
            .getAs('application/pdf');
          const newFileName = `${yyyymmdd}_${companyName}_本文.pdf`;
          monthFolder.createFile(pdfBlob).setName(newFileName);
          console.log(`  保存（本文PDF化）: ${newFileName}`);
        }
      });

      // ラベルの付替え（振り分け → 領収書）
      thread.addLabel(destLabel);
      console.log(`  ラベル追加: ${destLabel.getName()}`);
      thread.removeLabel(srcLabel);
      console.log(`  ラベル削除: ${srcLabel.getName()}`);
      console.log(`完了: "${subject}"`);
    } catch (e) {
      console.error(`エラー: "${subject}" - ${e.message}`);
      console.error(e.stack);
    }
  });

  console.log('全処理完了');
}

/**
 * 西暦から令和表記を返す（例: "2026" → "令和8年"）
 */
function toReiwa(yearStr) {
  const reiwaYear = parseInt(yearStr, 10) - 2018;
  return `令和${reiwaYear}年`;
}

/**
 * HTML内の外部画像をbase64データURIに変換して埋め込む
 */
function embedImages(html) {
  return html.replace(/<img[^>]+src=["']?(https?:\/\/[^"'\s>]+)["']?/gi, function(match, url) {
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        const contentType = response.getHeaders()['Content-Type'] || 'image/png';
        const base64 = Utilities.base64Encode(response.getContent());
        return match.replace(url, `data:${contentType};base64,${base64}`);
      }
    } catch (e) {
      console.log(`  画像取得スキップ: ${url}`);
    }
    return match;
  });
}

/**
 * フォルダの存在を確認し、なければ作成して返す
 */
function getOrCreateFolder(parent, folderName) {
  const folders = parent.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parent.createFolder(folderName);
}
