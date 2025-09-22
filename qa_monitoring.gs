// 質疑応答連絡書監視システム
// コンテナバインドGASスクリプト

// ============================================
// 初期設定とプロパティ管理
// ============================================

/**
 * スクリプトプロパティの初期設定
 * 手動で実行してプロパティを設定してください
 */
function setupScriptProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();

  // 以下の値を実際の値に置き換えてください
  scriptProperties.setProperties({
    // Discordウェブフック（カンマ区切りで複数設定可能）
    'DISCORD_WEBHOOKS': 'https://discord.com/api/webhooks/xxx1,https://discord.com/api/webhooks/xxx2',

    // 通知先メールアドレス（カンマ区切りで複数設定可能）
    'EMAIL_ADDRESSES': 'email1@example.com,email2@example.com',

    // バックアップ先のスプレッドシートID
    'BACKUP_SHEET_ID': 'your-backup-spreadsheet-id-here',

    // 処理済み質疑番号を記録（初期値は空）
    'PROCESSED_QUESTIONS': '[]',

    // 処理済み回答番号を記録（初期値は空）
    'PROCESSED_ANSWERS': '[]'
  });

  console.log('スクリプトプロパティを設定しました');
}

// ============================================
// メイン監視関数
// ============================================

/**
 * メイン処理関数（時間ベーストリガーで実行）
 */
function monitorQASheet() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();

    // ヘッダー行をスキップ
    if (data.length <= 1) return;

    // 各監視処理を実行
    checkNewQuestions(data);
    checkNewAnswers(data);
    backupAndCheckChanges(sheet, data);

  } catch (error) {
    console.error('メイン監視処理でエラーが発生しました:', error);
    sendErrorNotification('メイン監視処理エラー', error.toString());
  }
}

// ============================================
// 新規質疑監視
// ============================================

/**
 * 新規質疑をチェックして通知
 */
function checkNewQuestions(data) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const processedQuestions = JSON.parse(scriptProperties.getProperty('PROCESSED_QUESTIONS') || '[]');
  const newQuestions = [];

  // データをチェック（行番号1から開始、0はヘッダー）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const questionNumber = String(row[0]); // A列: 番号
    const sender = String(row[1]); // B列: 発信者
    const recipient = String(row[7]); // H列: 宛先
    const content = String(row[6]); // G列: 質疑内容

    // A列からH列まで全て埋まっているかチェック
    if (questionNumber && sender && row[2] && row[3] && row[4] && row[5] && content && recipient) {
      // まだ処理していない質疑の場合
      if (!processedQuestions.includes(questionNumber)) {
        newQuestions.push({
          number: questionNumber,
          sender: sender,
          recipient: recipient,
          content: content
        });
        processedQuestions.push(questionNumber);
      }
    }
  }

  // 新規質疑があれば通知
  if (newQuestions.length > 0) {
    sendNewQuestionsNotification(newQuestions);
    // 処理済みリストを更新
    scriptProperties.setProperty('PROCESSED_QUESTIONS', JSON.stringify(processedQuestions));
  }
}

// ============================================
// 新規回答監視
// ============================================

/**
 * 新規回答をチェックして通知
 */
function checkNewAnswers(data) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const processedAnswers = JSON.parse(scriptProperties.getProperty('PROCESSED_ANSWERS') || '[]');
  const newAnswers = [];

  // データをチェック
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const questionNumber = String(row[0]); // A列: 番号
    const originalSender = String(row[1]); // B列: 元の発信者
    const answerer = String(row[7]); // H列: 回答者（元の宛先）
    const questionContent = String(row[6]); // G列: 質疑内容
    const answerDate = row[8]; // I列: 回答日
    const answerContent = String(row[9]); // J列: 回答内容

    // I列とJ列が埋まっているかチェック
    if (answerDate && answerContent && questionNumber) {
      const answerKey = `${questionNumber}_answer`;

      // まだ処理していない回答の場合
      if (!processedAnswers.includes(answerKey)) {
        newAnswers.push({
          number: questionNumber,
          answerer: answerer,
          questioner: originalSender,
          questionContent: questionContent,
          answerContent: answerContent
        });
        processedAnswers.push(answerKey);
      }
    }
  }

  // 新規回答があれば通知
  if (newAnswers.length > 0) {
    sendNewAnswersNotification(newAnswers);
    // 処理済みリストを更新
    scriptProperties.setProperty('PROCESSED_ANSWERS', JSON.stringify(processedAnswers));
  }
}

// ============================================
// バックアップと変更監視
// ============================================

/**
 * データのバックアップと変更監視
 */
function backupAndCheckChanges(sheet, currentData) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const backupSheetId = scriptProperties.getProperty('BACKUP_SHEET_ID');

  if (!backupSheetId) {
    console.warn('バックアップシートIDが設定されていません');
    return;
  }

  try {
    // バックアップシートを取得
    const backupSpreadsheet = SpreadsheetApp.openById(backupSheetId);
    const backupSheet = backupSpreadsheet.getSheetByName('Backup') ||
                       backupSpreadsheet.insertSheet('Backup');

    // 前回のデータを取得
    const lastData = backupSheet.getDataRange().getValues();

    // 変更を検知（B列からJ列のみ比較）
    const changes = detectChanges(lastData, currentData);

    if (changes.length > 0) {
      sendChangeNotification(changes);
    }

    // 現在のデータをバックアップ
    backupSheet.clear();
    if (currentData.length > 0) {
      backupSheet.getRange(1, 1, currentData.length, currentData[0].length).setValues(currentData);
    }

    // タイムスタンプを記録
    const timestampSheet = backupSpreadsheet.getSheetByName('Timestamp') ||
                          backupSpreadsheet.insertSheet('Timestamp');
    timestampSheet.getRange('A1').setValue('最終バックアップ日時');
    timestampSheet.getRange('B1').setValue(new Date());

  } catch (error) {
    console.error('バックアップ処理でエラーが発生しました:', error);
    sendErrorNotification('バックアップ処理エラー', error.toString());
  }
}

/**
 * データの変更を検知
 */
function detectChanges(oldData, newData) {
  const changes = [];

  // 両方のデータが存在しない場合は変更なし
  if (!oldData || oldData.length <= 1 || !newData || newData.length <= 1) {
    return changes;
  }

  // 削除されたエントリをチェック
  for (let i = 1; i < oldData.length; i++) {
    const oldRow = oldData[i];
    const oldNumber = String(oldRow[0]);

    if (oldNumber) {
      // 新データで同じ番号を探す
      const found = newData.some((newRow, index) => {
        return index > 0 && String(newRow[0]) === oldNumber;
      });

      if (!found) {
        changes.push({
          type: '削除',
          number: oldNumber,
          content: `番号 ${oldNumber} のエントリが削除されました`
        });
      }
    }
  }

  // 変更されたエントリをチェック（B列からJ列）
  for (let i = 1; i < newData.length; i++) {
    const newRow = newData[i];
    const number = String(newRow[0]);

    if (number) {
      // 古いデータで同じ番号を探す
      for (let j = 1; j < oldData.length; j++) {
        const oldRow = oldData[j];
        if (String(oldRow[0]) === number) {
          // B列からJ列を比較
          for (let col = 1; col <= 9; col++) {
            const oldValue = String(oldRow[col] || '');
            const newValue = String(newRow[col] || '');

            if (oldValue !== newValue && oldValue !== '') {
              const columnName = getColumnName(col);
              changes.push({
                type: '変更',
                number: number,
                content: `番号 ${number} の${columnName}が変更されました: "${oldValue}" → "${newValue}"`
              });
            }
          }
          break;
        }
      }
    }
  }

  return changes;
}

/**
 * 列番号から列名を取得
 */
function getColumnName(colIndex) {
  const columnNames = {
    1: '発信者(B列)',
    2: '発信者部署(C列)',
    3: 'フェーズ(D列)',
    4: 'カテゴリ(E列)',
    5: '参照資料(F列)',
    6: '質疑内容(G列)',
    7: '宛先(H列)',
    8: '回答日(I列)',
    9: '回答(J列)'
  };
  return columnNames[colIndex] || `列${colIndex + 1}`;
}

// ============================================
// 通知機能
// ============================================

/**
 * 新規質疑の通知を送信
 */
function sendNewQuestionsNotification(questions) {
  const subject = '新規質疑通知';

  // Discord用フォーマット
  const discordMessage = formatQuestionsForDiscord(questions);

  // Email用データ
  const emailData = {
    type: 'question',
    items: questions
  };

  sendOptimizedNotification(subject, discordMessage, emailData);
}

/**
 * 新規回答の通知を送信
 */
function sendNewAnswersNotification(answers) {
  const subject = '新規回答通知';

  // Discord用フォーマット
  const discordMessage = formatAnswersForDiscord(answers);

  // Email用データ
  const emailData = {
    type: 'answer',
    items: answers
  };

  sendOptimizedNotification(subject, discordMessage, emailData);
}

/**
 * 変更通知を送信
 */
function sendChangeNotification(changes) {
  const subject = 'データ変更通知';

  // Discord用フォーマット
  const discordMessage = formatChangesForDiscord(changes);

  // Email用データ
  const emailData = {
    type: 'change',
    items: changes
  };

  sendOptimizedNotification(subject, discordMessage, emailData);
}

/**
 * エラー通知を送信
 */
function sendErrorNotification(title, errorMessage) {
  const discordMessage = `⚠️ **エラー発生**\n\`\`\`\n${errorMessage}\n\`\`\``;

  const emailData = {
    type: 'error',
    errorMessage: errorMessage
  };

  sendOptimizedNotification(title, discordMessage, emailData);
}

/**
 * Discord用: 新規質疑フォーマット
 */
function formatQuestionsForDiscord(questions) {
  let message = '📝 **新規質疑が登録されました**\n\n';

  questions.forEach((q, index) => {
    message += `**【質疑 #${q.number}】**\n`;
    message += `👤 **発信者:** ${q.sender}\n`;
    message += `📮 **宛先:** ${q.recipient}\n`;
    message += `\`\`\`\n${q.content}\n\`\`\`\n`;

    if (index < questions.length - 1) {
      message += '━━━━━━━━━━━━━━━\n';
    }
  });

  message += `\n📊 **件数:** ${questions.length}件`;

  return message;
}

/**
 * Discord用: 新規回答フォーマット
 */
function formatAnswersForDiscord(answers) {
  let message = '✅ **新規回答が登録されました**\n\n';

  answers.forEach((a, index) => {
    message += `**【回答 #${a.number}】**\n`;
    message += `👤 **回答者:** ${a.answerer}\n`;
    message += `📤 **質問者:** ${a.questioner}\n\n`;
    message += `**質疑内容:**\n> ${a.questionContent.replace(/\n/g, '\n> ')}\n\n`;
    message += `**回答内容:**\n\`\`\`\n${a.answerContent}\n\`\`\`\n`;

    if (index < answers.length - 1) {
      message += '━━━━━━━━━━━━━━━\n';
    }
  });

  message += `\n📊 **件数:** ${answers.length}件`;

  return message;
}

/**
 * Discord用: 変更通知フォーマット
 */
function formatChangesForDiscord(changes) {
  let message = '🔄 **データ変更が検出されました**\n\n';

  const deletions = changes.filter(c => c.type === '削除');
  const modifications = changes.filter(c => c.type === '変更');

  if (deletions.length > 0) {
    message += '**🗑️ 削除:**\n';
    deletions.forEach(d => {
      message += `• ${d.content}\n`;
    });
    message += '\n';
  }

  if (modifications.length > 0) {
    message += '**✏️ 変更:**\n';
    modifications.forEach(m => {
      message += `• ${m.content}\n`;
    });
  }

  message += `\n📊 **変更数:** ${changes.length}件`;

  return message;
}

/**
 * Email用: HTMLテンプレート生成
 */
function generateEmailHTML(subject, emailData) {
  let htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body {
      font-family: 'Noto Sans JP', 'Hiragino Sans', Arial, sans-serif;
      line-height: 1.6;
      color: #333;
      max-width: 800px;
      margin: 0 auto;
      padding: 20px;
    }
    .header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 30px;
      border-radius: 10px 10px 0 0;
      margin: -20px -20px 20px -20px;
    }
    h1 {
      margin: 0;
      font-size: 24px;
    }
    .content {
      background: white;
      padding: 30px;
      border-radius: 0 0 10px 10px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    .item {
      background: #f8f9fa;
      border-left: 4px solid #667eea;
      padding: 15px;
      margin: 20px 0;
      border-radius: 5px;
    }
    .item-header {
      font-weight: bold;
      color: #495057;
      margin-bottom: 10px;
      font-size: 16px;
    }
    .label {
      display: inline-block;
      background: #e9ecef;
      padding: 2px 8px;
      border-radius: 3px;
      font-size: 12px;
      color: #495057;
      margin-right: 5px;
    }
    .question-content, .answer-content {
      background: white;
      border: 1px solid #dee2e6;
      padding: 15px;
      margin: 10px 0;
      border-radius: 5px;
      white-space: pre-wrap;
      font-family: 'Courier New', monospace;
    }
    .original-question {
      background: #fff3cd;
      border-left: 3px solid #ffc107;
      padding: 10px;
      margin: 10px 0;
      border-radius: 3px;
      font-style: italic;
    }
    .stats {
      background: #e7f3ff;
      border: 1px solid #b3d7ff;
      padding: 10px 15px;
      border-radius: 5px;
      margin-top: 20px;
      text-align: center;
      font-weight: bold;
    }
    .footer {
      margin-top: 30px;
      padding-top: 20px;
      border-top: 1px solid #dee2e6;
      color: #6c757d;
      font-size: 12px;
      text-align: center;
    }
    .error-box {
      background: #f8d7da;
      border: 1px solid #f5c6cb;
      color: #721c24;
      padding: 15px;
      border-radius: 5px;
      margin: 20px 0;
    }
    .change-item {
      padding: 8px;
      margin: 5px 0;
      background: white;
      border-radius: 3px;
    }
    .deletion {
      background: #f8d7da;
      border-left: 3px solid #dc3545;
    }
    .modification {
      background: #d1ecf1;
      border-left: 3px solid #17a2b8;
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>📋 ${subject}</h1>
    <p style="margin: 5px 0 0 0; opacity: 0.9;">質疑応答連絡書 自動監視システム</p>
  </div>

  <div class="content">
`;

  // タイプ別にコンテンツを生成
  if (emailData.type === 'question') {
    htmlBody += generateQuestionEmailContent(emailData.items);
  } else if (emailData.type === 'answer') {
    htmlBody += generateAnswerEmailContent(emailData.items);
  } else if (emailData.type === 'change') {
    htmlBody += generateChangeEmailContent(emailData.items);
  } else if (emailData.type === 'error') {
    htmlBody += `
    <div class="error-box">
      <strong>⚠️ エラーが発生しました</strong>
      <pre style="margin-top: 10px;">${emailData.errorMessage}</pre>
    </div>`;
  }

  htmlBody += `
    </div>
    <div class="footer">
      <p>このメールは質疑応答連絡書監視システムから自動送信されています。</p>
      <p>送信日時: ${new Date().toLocaleString('ja-JP')}</p>
    </div>
  </div>
</body>
</html>`;

  return htmlBody;
}

/**
 * Email用: 質疑コンテンツ生成
 */
function generateQuestionEmailContent(questions) {
  let content = '<h2>新規質疑一覧</h2>';

  questions.forEach(q => {
    content += `
    <div class="item">
      <div class="item-header">
        <span class="label">質疑番号</span> ${q.number}
      </div>
      <p><strong>発信者:</strong> ${q.sender}</p>
      <p><strong>宛先:</strong> ${q.recipient}</p>
      <div class="question-content">${q.content}</div>
    </div>`;
  });

  content += `<div class="stats">📊 登録件数: ${questions.length}件</div>`;

  return content;
}

/**
 * Email用: 回答コンテンツ生成
 */
function generateAnswerEmailContent(answers) {
  let content = '<h2>新規回答一覧</h2>';

  answers.forEach(a => {
    content += `
    <div class="item">
      <div class="item-header">
        <span class="label">回答番号</span> ${a.number}
      </div>
      <p><strong>回答者:</strong> ${a.answerer}</p>
      <p><strong>質問者:</strong> ${a.questioner}</p>

      <h4>元の質疑内容:</h4>
      <div class="original-question">${a.questionContent}</div>

      <h4>回答内容:</h4>
      <div class="answer-content">${a.answerContent}</div>
    </div>`;
  });

  content += `<div class="stats">📊 回答件数: ${answers.length}件</div>`;

  return content;
}

/**
 * Email用: 変更コンテンツ生成
 */
function generateChangeEmailContent(changes) {
  let content = '<h2>変更内容一覧</h2>';

  const deletions = changes.filter(c => c.type === '削除');
  const modifications = changes.filter(c => c.type === '変更');

  if (deletions.length > 0) {
    content += '<h3>🗑️ 削除されたエントリ</h3>';
    deletions.forEach(d => {
      content += `<div class="change-item deletion">${d.content}</div>`;
    });
  }

  if (modifications.length > 0) {
    content += '<h3>✏️ 変更されたエントリ</h3>';
    modifications.forEach(m => {
      content += `<div class="change-item modification">${m.content}</div>`;
    });
  }

  content += `<div class="stats">📊 総変更数: ${changes.length}件</div>`;

  return content;
}

/**
 * 最適化された通知送信
 */
function sendOptimizedNotification(subject, discordMessage, emailData) {
  const scriptProperties = PropertiesService.getScriptProperties();

  // Discord通知
  const discordWebhooks = scriptProperties.getProperty('DISCORD_WEBHOOKS');
  if (discordWebhooks) {
    sendDiscordNotifications(discordWebhooks.split(','), discordMessage);
  }

  // Email通知
  const emailAddresses = scriptProperties.getProperty('EMAIL_ADDRESSES');
  if (emailAddresses) {
    const htmlBody = generateEmailHTML(subject, emailData);
    sendEmailNotifications(emailAddresses.split(','), subject, htmlBody);
  }
}

/**
 * Discordに通知を送信（Embed対応）
 */
function sendDiscordNotifications(webhooks, message) {
  webhooks.forEach(webhook => {
    if (webhook.trim()) {
      try {
        // 文字数制限チェック
        if (message.length > 1900) {
          // 長いメッセージは分割
          const chunks = splitMessage(message, 1900);
          chunks.forEach(chunk => {
            const payload = {
              content: chunk,
              username: '質疑応答監視システム',
              avatar_url: 'https://cdn.discordapp.com/embed/avatars/0.png'
            };

            const options = {
              method: 'post',
              contentType: 'application/json',
              payload: JSON.stringify(payload),
              muteHttpExceptions: true
            };

            UrlFetchApp.fetch(webhook.trim(), options);
          });
        } else {
          // 通常送信
          const payload = {
            content: message,
            username: '質疑応答監視システム',
            avatar_url: 'https://cdn.discordapp.com/embed/avatars/0.png'
          };

          const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
          };

          const response = UrlFetchApp.fetch(webhook.trim(), options);

          if (response.getResponseCode() !== 204) {
            console.error(`Discord通知失敗: ${response.getContentText()}`);
          }
        }
      } catch (error) {
        console.error(`Discord通知エラー: ${error}`);
      }
    }
  });
}

/**
 * メール通知を送信（HTML最適化版）
 */
function sendEmailNotifications(emails, subject, htmlBody) {
  emails.forEach(email => {
    if (email.trim()) {
      try {
        // プレーンテキスト版も生成（HTMLが表示できない環境用）
        const plainBody = htmlBody.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();

        MailApp.sendEmail({
          to: email.trim(),
          subject: `[質疑応答監視] ${subject}`,
          body: plainBody,
          htmlBody: htmlBody
        });

      } catch (error) {
        console.error(`メール送信エラー (${email}): ${error}`);
      }
    }
  });
}

/**
 * 長いメッセージを分割
 */
function splitMessage(message, maxLength) {
  const chunks = [];
  let currentChunk = '';
  const lines = message.split('\n');

  for (const line of lines) {
    if (currentChunk.length + line.length + 1 > maxLength) {
      if (currentChunk) {
        chunks.push(currentChunk);
        currentChunk = '';
      }
    }
    currentChunk += (currentChunk ? '\n' : '') + line;
  }

  if (currentChunk) {
    chunks.push(currentChunk);
  }

  return chunks;
}

// ============================================
// トリガー設定
// ============================================

/**
 * 時間ベーストリガーを設定（1時間ごと）
 * この関数を一度だけ実行してトリガーを設定してください
 */
function setupTimeTrigger() {
  // 既存のトリガーをクリア
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'monitorQASheet') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新しいトリガーを設定（1時間ごと）
  ScriptApp.newTrigger('monitorQASheet')
    .timeBased()
    .everyHours(1)
    .create();

  console.log('時間ベーストリガーを設定しました（1時間ごと）');
}

/**
 * テスト実行用関数
 * 動作確認のために手動で実行できます
 */
function testRun() {
  console.log('テスト実行を開始します...');
  monitorQASheet();
  console.log('テスト実行が完了しました');
}

/**
 * 処理済みリストをリセット
 * 新規テストの際に使用
 */
function resetProcessedLists() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('PROCESSED_QUESTIONS', '[]');
  scriptProperties.setProperty('PROCESSED_ANSWERS', '[]');
  console.log('処理済みリストをリセットしました');
}