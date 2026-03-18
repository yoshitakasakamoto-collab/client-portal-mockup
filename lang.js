/* ==========================================
   Language Switching - EN / JP
   ========================================== */

const translations = {
  // Sidebar
  'Project management': 'プロジェクト管理',
  'User management': 'ユーザー管理',

  // Header
  'Project Management': 'プロジェクト管理',
  'Notifications': '通知',
  'Mark all read': 'すべて既読にする',
  'New expert proposed': '新しいエキスパートが提案されました',
  '2 new experts proposed for A109919 BCG US GFPs project': 'A109919 BCG US GFPsプロジェクトに新たに2名のエキスパートが提案されました',
  '10 minutes ago': '10分前',
  'Interview scheduled': 'インタビュー予定確定',
  'Pretty Dewati (Cathy) - Interview confirmed for Mar 17, 9:00 AM': 'Pretty Dewati (Cathy) - 3月17日 9:00 AMのインタビューが確定しました',
  '1 hour ago': '1時間前',
  'Transcript ready': '書き起こし完了',
  'Recording & transcript available for Luci Dao interview': 'Luci Daoのインタビューの録音と書き起こしが利用可能です',
  '3 hours ago': '3時間前',
  'Expert approved': 'エキスパート承認済み',
  'Enrico Tranchina approved Pretty Dewati (Cathy) for interview': 'Enrico TranchinaがPretty Dewati (Cathy)のインタビューを承認しました',
  'Yesterday': '昨日',

  // Project list page
  'Ongoing Projects': '進行中のプロジェクト',
  'Past Projects': '過去のプロジェクト',
  'Date range:': '期間：',
  'Search': '検索',
  'Project': 'プロジェクト',
  'Status': 'ステータス',
  'Start date': '開始日',
  'CDD stats': 'CDD統計',
  'IV stats': 'IV統計',
  'Total price': '合計金額',
  'Arches team': 'Archesチーム',
  'Geography': '地域',
  'Project/Expert tags': 'プロジェクト/エキスパートタグ',
  'On going': '進行中',
  'Closed': 'クローズ',
  'Proposed': '提案済み',
  'Approved': '承認済み',
  'Waiting': '待機中',
  'Accept': '受諾',
  'Finished': '完了',

  // Project detail tabs
  'General': '概要',
  'Candidates': '候補者',
  'Interview': 'インタビュー',
  'Billing': '請求',

  // General tab
  'Project Overview': 'プロジェクト概要',
  'Start date:': '開始日：',
  'Inquiry date:': '問い合わせ日：',
  'Your team:': 'クライアントチーム：',
  'Arches team:': 'Archesチーム：',
  'Geography:': '地域：',
  'Project/Expert tags:': 'プロジェクト/エキスパートタグ：',
  'Billing code:': '請求コード：',
  'Project Dashboard': 'プロジェクトダッシュボード',
  'Experts Proposed': '提案エキスパート数',
  'Calls Conducted': '実施済み通話数',
  'Total Spent': '合計支出',
  'Pending Interviews': '予定インタビュー',
  'by segment': 'セグメント別',
  'avg duration': '平均時間',
  'interviews booked': 'インタビュー予約済み',
  'Team Activity': 'チーム活動',
  'Briefing': 'ブリーフィング',
  'approved expert': 'エキスパート承認',
  'scheduled interview with': 'インタビュー予約：',
  'declined expert': 'エキスパート辞退',
  'conducted interview with': 'インタビュー実施：',

  // Candidates tab
  'Filter by:': 'フィルター：',
  'All Statuses': '全ステータス',
  'Sort by:': '並び替え：',
  'Date Proposed': '提案日',
  'Cost': 'コスト',
  'Name': '名前',
  'Segment:': 'セグメント：',
  'All Segments': '全セグメント',
  'Bulk Actions': '一括操作',
  'Export List': 'リスト出力',
  'Compare': '比較',
  'Expert info': 'エキスパート情報',
  'Activities': '活動',
  'Availability': '空き状況',
  'Experience': '経歴',
  'Working history': '職歴',
  'Decline': '辞退',
  'Approve': '承認',
  'Book Interview': 'インタビュー予約',
  'Hourly Rate:': '時間単価：',
  'Location:': '所在地：',
  'Languages:': '言語：',
  'Years of Exp:': '経験年数：',
  'Next available slots:': '次の空き枠：',
  'Copy availability as text': '空き状況をテキストでコピー',
  'Key Skills': '主要スキル',
  'Decline Reason': '辞退理由',
  'Why are you declining this expert? (optional)': 'このエキスパートを辞退する理由は？（任意）',
  'Not relevant experience': '関連する経験がない',
  'Too expensive': '費用が高い',
  'Already have enough experts': '十分なエキスパートが確保済み',
  'Other': 'その他',
  'Please specify...': '詳細を記入してください...',
  'Cancel': 'キャンセル',
  'Submit & Decline': '送信して辞退',
  'experts': '名のエキスパート',
  'calls done': '通話完了',
  'updated': '更新',

  // Interview tab
  'Interview Management': 'インタビュー管理',
  'Filter:': 'フィルター：',
  'All': 'すべて',
  'Booked': '予約済み',
  'To Be Booked': '予約待ち',
  'Conducted': '実施済み',
  'Canceled': 'キャンセル済み',
  'Expert': 'エキスパート',
  'Segment': 'セグメント',
  'Date & Time': '日時',
  'Duration': '所要時間',
  'Cost': 'コスト',
  'Recording': '録音',
  'Actions': 'アクション',
  'View Recording': '録音を見る',
  'AI Summary': 'AI要約',
  'Transcript': '書き起こし',
  'Confirm Duration': '時間を確認',
  'Contest Duration': '時間に異議',
  'Cancel Interview': 'インタビューをキャンセル',
  'Book Now': '今すぐ予約',
  'Rate Expert': 'エキスパートを評価',
  'Leave Feedback': 'フィードバック',
  'min': '分',
  'Cancellation Reason': 'キャンセル理由',
  'Please select a reason for cancellation:': 'キャンセル理由を選択してください：',
  'Schedule conflict': 'スケジュールの競合',
  'No longer needed': '不要になった',
  'Expert unavailable': 'エキスパートが不在',
  'Additional comments...': '追加コメント...',
  'Submit Cancellation': 'キャンセルを送信',
  'Contest Call Duration': '通話時間に異議を申し立てる',
  'Recorded duration:': '記録された時間：',
  'Your estimated duration:': 'ご認識の時間：',
  'Please describe the discrepancy:': '相違の詳細を記入してください：',
  'Submit Contest': '異議を送信',
  'Expert Feedback': 'エキスパートフィードバック',
  'How would you rate this expert?': 'このエキスパートの評価は？',
  'Your feedback:': 'フィードバック：',
  'Share your experience with this expert...': 'このエキスパートとのやり取りについてご記入ください...',
  'Submit Feedback': 'フィードバックを送信',
  'Recording & Transcript': '録音と書き起こし',
  'AI-Generated Summary': 'AI生成要約',
  'Full Transcript': '全文書き起こし',
  'Download Transcript': '書き起こしをダウンロード',
  'Download Recording': '録音をダウンロード',
  'Close': '閉じる',

  // Billing tab
  'Billing & Invoices': '請求と請求書',
  'Billing Summary': '請求概要',
  'Total Billed': '請求総額',
  'Invoices Issued': '発行済み請求書',
  'Pending Payment': '未払い',
  'Paid': '支払い済み',
  'Invoice Number': '請求書番号',
  'Issue Date': '発行日',
  'Due Date': '支払期限',
  'Amount': '金額',
  'Download': 'ダウンロード',
  'Download PDF': 'PDFダウンロード',
  'Download All': 'すべてダウンロード',
  'Overdue': '期限超過',
};

let currentLang = localStorage.getItem('portalLang') || 'en';

function initLang() {
  // Set up the language toggle button
  const langBtn = document.querySelector('.lang-toggle');
  if (langBtn) {
    updateLangButton(langBtn);
    langBtn.addEventListener('click', function(e) {
      e.stopPropagation();
      currentLang = currentLang === 'en' ? 'jp' : 'en';
      localStorage.setItem('portalLang', currentLang);
      applyLang();
      updateLangButton(langBtn);
    });
  }

  applyLang();
}

function updateLangButton(btn) {
  if (currentLang === 'en') {
    btn.innerHTML = '🇬🇧 EN';
  } else {
    btn.innerHTML = '🇯🇵 JP';
  }
}

function applyLang() {
  const elements = document.querySelectorAll('[data-lang]');
  elements.forEach(el => {
    const key = el.getAttribute('data-lang');
    if (currentLang === 'jp' && translations[key]) {
      el.textContent = translations[key];
    } else if (currentLang === 'en') {
      el.textContent = key;
    }
  });
}

document.addEventListener('DOMContentLoaded', initLang);
