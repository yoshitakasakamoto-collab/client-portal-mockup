const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat,
        HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageNumber, PageBreak, TableOfContents } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };
const headerShading = { fill: "2C3E50", type: ShadingType.CLEAR };
const altShading = { fill: "F8F9FA", type: ShadingType.CLEAR };

function hCell(text, width) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: headerShading, margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 20 })] })]
  });
}
function cell(text, width, shading) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: shading || undefined, margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 20 })] })]
  });
}
function heading1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 32 })] });
}
function heading2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 26 })] });
}
function heading3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 22 })] });
}
function para(text, opts = {}) {
  return new Paragraph({ spacing: { after: 120 }, ...opts,
    children: [new TextRun({ text, font: "Arial", size: 20, ...opts.run })] });
}
function boldPara(label, text) {
  return new Paragraph({ spacing: { after: 100 },
    children: [
      new TextRun({ text: label, bold: true, font: "Arial", size: 20 }),
      new TextRun({ text, font: "Arial", size: 20 })
    ]
  });
}
function bulletItem(text, ref) {
  return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20 })] });
}
function numberItem(text, ref) {
  return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20 })] });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "2C3E50" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "E67E22" },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "34495E" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers2", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers3", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers4", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers5", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets2", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets3", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets4", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [
    // ===== 表紙 =====
    {
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      children: [
        new Paragraph({ spacing: { before: 3000 } }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
          children: [new TextRun({ text: "クライアントポータル", font: "Arial", size: 56, bold: true, color: "2C3E50" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "開発仕様書", font: "Arial", size: 40, color: "E67E22" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "E67E22", space: 1 } },
          children: [new TextRun({ text: " ", font: "Arial", size: 20 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "Arches Consulting", font: "Arial", size: 28, color: "7F8C8D" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "Version 1.0 | 2026年3月", font: "Arial", size: 22, color: "7F8C8D" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "社外秘", font: "Arial", size: 20, bold: true, color: "E74C3C" })] }),
      ]
    },

    // ===== 本文 =====
    {
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      headers: {
        default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "クライアントポータル - 開発仕様書", font: "Arial", size: 16, color: "999999" })] })] })
      },
      footers: {
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Page ", font: "Arial", size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16 })] })] })
      },
      children: [
        new TableOfContents("目次", { hyperlink: true, headingStyleRange: "1-3" }),
        new Paragraph({ children: [new PageBreak()] }),

        // ===== 1. システム概要 =====
        heading1("1. システム概要"),
        para("クライアントポータルは、AIS（Arches Intelligence System）の一モジュールであり、クライアントにエキスパート管理、インタビュー日程調整、請求管理のセルフサービス機能を提供します。本書はポータル改修の機能要件および技術要件を定義します。"),

        heading2("1.1 アーキテクチャ"),
        boldPara("マスターシステム: ", "AIS（Arches Intelligence System）- 社内のエキスパートDB・案件管理システム"),
        boldPara("クライアントポータル: ", "AISのクライアント向けモジュール。全データはAISが正本。"),
        boldPara("データフロー: ", "AIS（正本）-> ポータル（参照 + 限定的な書き込み）。クライアントのアクション（予約、コメント、辞退）はAISに書き戻され通知を発火。"),

        heading2("1.2 通知アーキテクチャ"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 3680, 3680],
          rows: [
            new TableRow({ children: [hCell("チャネル", 2000), hCell("クライアント側", 3680), hCell("Arches側", 3680)] }),
            new TableRow({ children: [cell("ポータル内", 2000), cell("通知ベルアイコン（バッジ数 + ドロップダウン一覧）", 3680), cell("N/A - Archesは社内AISダッシュボードを使用", 3680)] }),
            new TableRow({ children: [cell("メール", 2000, altShading), cell("主要イベント時にメール送信（新エキスパート提案、IV確定、書起こし完了）", 3680, altShading), cell("N/A", 3680, altShading)] }),
            new TableRow({ children: [cell("Slack", 2000), cell("N/A", 3680), cell("案件ごとに専用Slackチャンネルを作成。Slackメールアドレスを案件に登録し、クライアントのアクションはすべてこのアドレスに通知→Slackチャンネルに投下", 3680)] }),
          ]
        }),

        heading2("1.3 ユーザー権限（初期リリース）"),
        para("クライアントユーザーは全員同一権限。ロールベースのアクセス制御（Admin/Viewer/Booker）はv1対象外だが、将来の拡張性を考慮した設計とすること。"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 2. データモデル =====
        heading1("2. データモデル"),

        heading2("2.1 エンティティ関係"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 4360, 3000],
          rows: [
            new TableRow({ children: [hCell("エンティティ", 2000), hCell("説明", 4360), hCell("データソース", 3000)] }),
            new TableRow({ children: [cell("Project", 2000), cell("クライアント案件（セグメント、チームメンバー、ブリーフィングを含む）", 4360), cell("AIS（正本）", 3000)] }),
            new TableRow({ children: [cell("Expert", 2000, altShading), cell("専門家（プロフィール、経歴、対応可能日程）", 4360, altShading), cell("AIS（正本）", 3000, altShading)] }),
            new TableRow({ children: [cell("Interview", 2000), cell("クライアントとエキスパート間の予定/完了済み通話", 4360), cell("AIS + ポータル（書き戻し）", 3000)] }),
            new TableRow({ children: [cell("Billing Item", 2000, altShading), cell("費用明細（インタビュー、通訳、フォローアップ）", 4360, altShading), cell("AIS（正本、既存ロジック）", 3000, altShading)] }),
            new TableRow({ children: [cell("Comment", 2000), cell("エキスパートへのスレッド形式コメント", 4360), cell("ポータル（書込）-> AIS", 3000)] }),
            new TableRow({ children: [cell("Notification", 2000, altShading), cell("アプリ内+メール通知", 4360, altShading), cell("AISが生成", 3000, altShading)] }),
          ]
        }),

        heading2("2.2 エキスパートステータスフロー"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 4680, 2680],
          rows: [
            new TableRow({ children: [hCell("ステータス", 2000), hCell("説明", 4680), hCell("ポータル表示", 2680)] }),
            new TableRow({ children: [cell("Prospect", 2000), cell("Archesがリーチ中の候補者。まだ精査されていない。", 4680), cell("非表示（社内のみ）", 2680)] }),
            new TableRow({ children: [cell("Proposed", 2000, altShading), cell("Archesが精査しクライアントに提案した状態。", 4680, altShading), cell("表示（最初に見える状態）", 2680, altShading)] }),
            new TableRow({ children: [cell("Approved", 2000), cell("クライアントがインタビュー承認。", 4680), cell("表示", 2680)] }),
            new TableRow({ children: [cell("Declined", 2000, altShading), cell("クライアントが辞退（理由付き）。", 4680, altShading), cell("表示（グレーアウト）", 2680, altShading)] }),
            new TableRow({ children: [cell("Interview", 2000), cell("インタビュー予約済みまたは進行中。", 4680), cell("表示（IV画面に表示）", 2680)] }),
            new TableRow({ children: [cell("Billing", 2000, altShading), cell("インタビュー完了、請求発生。", 4680, altShading), cell("表示（Billing画面）", 2680, altShading)] }),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 3. 画面仕様 =====
        heading1("3. 画面仕様"),

        heading2("3.1 プロジェクトダッシュボード"),
        para("クライアントがアクセス可能な全プロジェクトを一覧表示。「進行中」「過去」タブで切替。"),
        heading3("3.1.1 必要なAPI"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [1500, 3680, 4180],
          rows: [
            new TableRow({ children: [hCell("メソッド", 1500), hCell("エンドポイント", 3680), hCell("説明", 4180)] }),
            new TableRow({ children: [cell("GET", 1500), cell("/api/projects", 3680), cell("当該クライアント組織の全プロジェクト一覧", 4180)] }),
            new TableRow({ children: [cell("GET", 1500, altShading), cell("/api/notifications", 3680, altShading), cell("現在のユーザーの未読通知", 4180, altShading)] }),
          ]
        }),

        heading2("3.2 プロジェクト概要ページ"),
        para("プロジェクトのサマリー、セグメント内訳、チーム活動、ブリーフィングを表示するダッシュボード。"),

        heading2("3.3 候補者ページ（エキスパート管理）"),
        para("左右分割ビュー：左パネルにセグメント別エキスパートリスト、右パネルに選択したエキスパートの詳細。"),

        heading3("3.3.1 左パネル - エキスパートリスト"),
        bulletItem("セグメント別グループ化（折りたたみ可能）", "bullets"),
        bulletItem("各エキスパート：ID、名前、会社、役職、コスト、ステータスバッジ表示", "bullets"),
        bulletItem("チェックボックスによる一括選択", "bullets"),
        bulletItem("Export Listボタン（リスト上部）- 全エキスパートをExcel出力", "bullets"),
        bulletItem("バルクアクションバー（チェック時表示）- Export Selected", "bullets"),

        heading3("3.3.2 右パネル - エキスパート詳細（タブなし・縦スクロール）"),
        para("タブ廃止。以下の順番で全セクションを1ページに縦並び表示："),
        numberItem("Availability - 対応可能日程（クリップボードコピー機能付き）", "numbers"),
        numberItem("Working History - 職歴テーブル（会社、役職、期間）", "numbers"),
        numberItem("Experience - 経歴の段落テキスト", "numbers"),
        numberItem("Screening Answers - スクリーニングQ&A", "numbers"),
        numberItem("Comments & Activity - コメント（スレッド形式）+ 折りたたみ可能なアクティビティログ", "numbers"),

        heading3("3.3.3 アクションボタン"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2800, 3280, 3280],
          rows: [
            new TableRow({ children: [hCell("ボタン", 2800), hCell("表示条件", 3280), hCell("動作", 3280)] }),
            new TableRow({ children: [cell("Book for Interview", 2800), cell("空き日程あり", 3280), cell("予約モーダルを開く", 3280)] }),
            new TableRow({ children: [cell("Request Availability", 2800, altShading), cell("空き日程なし", 3280, altShading), cell("Arches Slackに通知を送信", 3280, altShading)] }),
            new TableRow({ children: [cell("Not Interested", 2800), cell("未辞退のエキスパート", 3280), cell("辞退モーダル（理由選択）を開く", 3280)] }),
          ]
        }),

        heading3("3.3.4 コメントシステム"),
        bulletItem("クライアントがエキスパートにコメントを投稿可能", "bullets2"),
        bulletItem("Archesスタッフが返信可能（AISから投稿、ポータルに表示）", "bullets2"),
        bulletItem("新規コメント時：案件のSlackメールアドレスに通知→Slackチャンネルに投下", "bullets2"),

        new Paragraph({ children: [new PageBreak()] }),

        // --- 3.4 予約モジュール ---
        heading2("3.4 予約モジュール（モーダル）"),
        para("最も複雑なUIコンポーネント。タイムゾーン変換、カレンダー表示、時間バリデーション、参加者管理、ミーティング方式選択を処理。"),

        heading3("3.4.1 タイムゾーン処理"),
        bulletItem("ページ読込時にブラウザのIntl APIでクライアントのタイムゾーンを自動検出（DST対応）", "bullets"),
        bulletItem("一般的なタイムゾーンラベルのドロップダウン（例：UTC-05:00 Eastern Time）", "bullets"),
        bulletItem("エキスパートの空き日程はAIS内でエキスパートのローカルTZで保存", "bullets"),
        bulletItem("カレンダー表示時にクライアント選択TZに変換して表示", "bullets"),

        heading3("3.4.2 カレンダータイムライン"),
        bulletItem("2週間のローリングウィンドウでエキスパートの空き日程を表示", "bullets2"),
        bulletItem("15分単位のクリック可能なグリッド", "bullets2"),
        bulletItem("色分け：緑=空き、オレンジ=選択中、赤=予約済み", "bullets2"),

        heading3("3.4.3 インタビュー時間"),
        boldPara("形式: ", "ドロップダウン（自由入力不可）"),
        boldPara("選択肢: ", "30, 45, 60, 75, 90, 105, 120分（15分刻み）"),

        heading3("3.4.4 空き時間バリデーション（重要）"),
        para("開始時間とインタビュー時間を選択した際、インタビュー全体がエキスパートの空き時間内に収まることを必ず検証すること：", { run: { bold: true, color: "E74C3C" } }),
        numberItem("選択した開始時間+インタビュー時間をエキスパートのTZに変換", "numbers2"),
        numberItem("開始から終了までの全15分ブロックが空き日程内にあることを確認", "numbers2"),
        numberItem("空き時間外のブロックがある場合、警告を表示し確定をブロック", "numbers2"),
        numberItem("エラーメッセージ：「選択した時間がエキスパートの対応可能時間を超えています。短い時間を選択するか、別の時間枠を選んでください。」", "numbers2"),

        heading3("3.4.5 参加者（Attendees）"),
        bulletItem("デフォルト：ログインユーザーのメアドが入力済み（読み取り専用）", "bullets3"),
        bulletItem("「+」ボタンで追加のメール入力フィールドを追加", "bullets3"),
        bulletItem("「x」ボタンで追加した参加者を削除（デフォルトは削除不可）", "bullets3"),

        heading3("3.4.6 ミーティング方式"),
        boldPara("形式: ", "ドロップダウン"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2800, 3280, 3280],
          rows: [
            new TableRow({ children: [hCell("選択肢", 2800), hCell("動作", 3280), hCell("連携", 3280)] }),
            new TableRow({ children: [cell("Arches Zoom Link", 2800), cell("Zoomミーティングを自動生成", 3280), cell("Zoom API: ミーティング作成、カレンダーにリンク埋込", 3280)] }),
            new TableRow({ children: [cell("Arches Zoom (Call-in)", 2800, altShading), cell("ダイヤルイン番号付きZoomを生成", 3280, altShading), cell("Zoom API: テレフォニー有効でミーティング作成", 3280, altShading)] }),
            new TableRow({ children: [cell("Client-provided Link", 2800), cell("クライアントがURL入力", 3280), cell("入力URLをそのままカレンダー招待に反映", 3280)] }),
          ]
        }),

        heading3("3.4.7 カレンダー招待（個人情報保護）"),
        para("重要：プライバシー保護のため、カレンダー招待はクライアント側とエキスパート側で別々に作成すること：", { run: { bold: true, color: "E74C3C" } }),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 3680, 3680],
          rows: [
            new TableRow({ children: [hCell("招待", 2000), hCell("送信先", 3680), hCell("記載内容", 3680)] }),
            new TableRow({ children: [cell("クライアント側", 2000), cell("クライアントの参加者メールのみ", 3680), cell("エキスパート名（個人メール/電話番号なし）、会議リンク、日時", 3680)] }),
            new TableRow({ children: [cell("エキスパート側", 2000, altShading), cell("エキスパートのメールのみ", 3680, altShading), cell("クライアント企業名（個人名/メールなし）、会議リンク、日時", 3680, altShading)] }),
          ]
        }),

        heading3("3.4.8 予約後のフロー"),
        numberItem("AISに予約リクエストを「Pending」ステータスで作成", "numbers3"),
        numberItem("案件のSlackチャンネルに通知", "numbers3"),
        numberItem("Archesがエキスパートに日程確認の連絡", "numbers3"),
        numberItem("エキスパート承諾時：Zoom APIでミーティング自動作成", "numbers3"),
        numberItem("クライアント側・エキスパート側それぞれにカレンダー招待を送信", "numbers3"),
        numberItem("AISでエキスパートのステータスを「Interview」に更新", "numbers3"),
        numberItem("クライアントにメール+ポータル内通知：「インタビュー確定」", "numbers3"),

        new Paragraph({ children: [new PageBreak()] }),

        // --- 3.5 インタビュー ---
        heading2("3.5 インタビューページ"),
        para("全インタビューをセグメント別に表示。ステータス管理、録音アクセス、事後アクション。"),

        heading3("3.5.1 ステータスカード"),
        bulletItem("Booked: 予定されているインタビュー数", "bullets"),
        bulletItem("Conducted: 完了済みインタビュー数", "bullets"),
        bulletItem("Canceled: キャンセル済みインタビュー数", "bullets"),

        heading3("3.5.2 インタビュー後の自動パイプライン"),
        numberItem("Zoomでインタビュー完了", "numbers4"),
        numberItem("Zoom APIで録音を自動取得", "numbers4"),
        numberItem("音声を文字起こしサービスに送信（Whisper, Deepgram等）", "numbers4"),
        numberItem("文字起こしをLLMでAI要約を自動生成", "numbers4"),
        numberItem("録音・文字起こし・要約をインタビューレコードに紐付けて保存", "numbers4"),
        numberItem("クライアントに通知：「[エキスパート名]の書き起こしが完了しました」", "numbers4"),

        heading3("3.5.3 クライアントのアクション"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2500, 3430, 3430],
          rows: [
            new TableRow({ children: [hCell("アクション", 2500), hCell("利用可能条件", 3430), hCell("詳細", 3430)] }),
            new TableRow({ children: [cell("録音再生", 2500), cell("文字起こし完了後", 3430), cell("録音のストリーミングまたはダウンロード", 3430)] }),
            new TableRow({ children: [cell("文字起こし表示", 2500, altShading), cell("文字起こし完了後", 3430, altShading), cell("全文テキスト + AI要約の表示", 3430, altShading)] }),
            new TableRow({ children: [cell("エキスパート評価", 2500), cell("インタビュー完了後", 3430), cell("星評価（1-5）+ テキストフィードバック。AISに蓄積。", 3430)] }),
            new TableRow({ children: [cell("キャンセル", 2500, altShading), cell("Bookedステータス時", 3430, altShading), cell("理由選択 + コメント。Arches Slackに通知。", 3430, altShading)] }),
            new TableRow({ children: [cell("時間異議", 2500), cell("完了後", 3430), cell("録音時間と実際の時間に差異がある場合の申告", 3430)] }),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // --- 3.6 Billing ---
        heading2("3.6 請求ページ"),
        para("プロジェクトの費用、通話内訳、請求書管理を表示。"),

        heading3("3.6.1 請求サマリー"),
        bulletItem("Total Billed: 全完了通話の合計費用", "bullets"),
        bulletItem("Discount Applied: ボリュームディスカウント等", "bullets"),
        bulletItem("Billing Code: クライアントの経理コード", "bullets"),

        heading3("3.6.2 費用計算"),
        para("課金ロジックはAISに既存。ポータルはAISから計算済みの値を表示。費用タイプ："),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2500, 3430, 3430],
          rows: [
            new TableRow({ children: [hCell("タイプ", 2500), hCell("説明", 3430), hCell("単価基準", 3430)] }),
            new TableRow({ children: [cell("Interview", 2500), cell("標準エキスパート通話", 3430), cell("1回あたりの固定料金（エキスパートにより異なる、30/60分で変動）", 3430)] }),
            new TableRow({ children: [cell("Follow-up Q&A", 2500, altShading), cell("インタビュー後の書面フォローアップ", 3430, altShading), cell("セッション固定料金", 3430, altShading)] }),
            new TableRow({ children: [cell("Interpretation", 2500), cell("通訳サービス", 3430), cell("セッション固定料金", 3430)] }),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 4. 外部連携 =====
        heading1("4. 外部連携"),

        heading2("4.1 Zoom API"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2500, 6860],
          rows: [
            new TableRow({ children: [hCell("機能", 2500), hCell("詳細", 6860)] }),
            new TableRow({ children: [cell("ミーティング作成", 2500), cell("エキスパート承諾時にZoomリンクを自動生成。Call-inの場合はダイヤルイン番号を含む。", 6860)] }),
            new TableRow({ children: [cell("録音取得", 2500, altShading), cell("ミーティング終了後、クラウド録音をZoom Webhook/APIで自動ダウンロード。", 6860, altShading)] }),
            new TableRow({ children: [cell("認証", 2500), cell("Server-to-Server OAuthアプリ（Archesのアカウント）。クライアント認証不要。", 6860)] }),
          ]
        }),

        heading2("4.2 文字起こしパイプライン"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2500, 6860],
          rows: [
            new TableRow({ children: [hCell("ステップ", 2500), hCell("詳細", 6860)] }),
            new TableRow({ children: [cell("音声入力", 2500), cell("Zoom録音ファイル（mp4/m4a）", 6860)] }),
            new TableRow({ children: [cell("文字起こし", 2500, altShading), cell("音声認識API（Whisper, Deepgram等）。出力：タイムスタンプ付きテキスト。", 6860, altShading)] }),
            new TableRow({ children: [cell("AI要約", 2500), cell("LLMで構造化要約を生成（要点、テーマ、アクションアイテム）", 6860)] }),
            new TableRow({ children: [cell("保存", 2500, altShading), cell("録音（オブジェクトストレージ）、文字起こし+要約（データベース）", 6860, altShading)] }),
          ]
        }),

        heading2("4.3 Slack通知"),
        bulletItem("案件ごとに専用Slackメールアドレスを登録", "bullets"),
        bulletItem("クライアントのアクション（コメント、予約、辞退、キャンセル）→Slackメールに通知", "bullets"),
        bulletItem("メール内容がSlackチャンネルに自動投稿", "bullets"),

        // ===== 5. 国際化 =====
        heading1("5. 国際化（i18n）"),
        bulletItem("英語・日本語の切替をヘッダーのトグルで実現", "bullets2"),
        bulletItem("全UIテキストに翻訳キーを使用（モックアップのdata-lang属性）", "bullets2"),
        bulletItem("言語設定はユーザープロフィールに保存（本番環境）", "bullets2"),
        bulletItem("200以上の翻訳キーがlang.jsに定義済み（モックアップソースに同梱）", "bullets2"),

        // ===== 6. セキュリティ =====
        heading1("6. セキュリティ要件"),
        bulletItem("認証：既存AIS認証システム（SSO/OAuth）", "bullets3"),
        bulletItem("認可：全APIコールをクライアント組織にスコープ（組織間データアクセス不可）", "bullets3"),
        bulletItem("カレンダープライバシー：クライアントとエキスパートで別々の招待（3.4.7参照）", "bullets3"),
        bulletItem("エキスパートPII：個人メール/電話番号をポータルに露出させない", "bullets3"),
        bulletItem("通信暗号化：全API通信にHTTPS/TLS", "bullets3"),

        // ===== 7. モックアップ参照 =====
        heading1("7. モックアップ参照"),
        para("インタラクティブHTMLモックアップは以下で確認可能："),
        boldPara("リポジトリ: ", "https://github.com/yoshitakasakamoto-collab/client-portal-mockup"),
        boldPara("パスワード: ", "Client216"),
        para("注：モックアップは静的JavaScriptデータを使用。本番では本書記載のAIS APIエンドポイントからデータを取得すること。"),
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("docs/Client_Portal_Dev_Spec_JP.docx", buffer);
  console.log("JP spec generated: docs/Client_Portal_Dev_Spec_JP.docx");
});
