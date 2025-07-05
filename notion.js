/**
 * Notion 公開／非公開ページ URL を Markdown に変換して返す
 * @param {string} pageUrl  例: https://www.notion.so/Your-Page-Title-0123456789abcdef0123456789abcdef
 * @return {string}         Markdown
 */
function fetchNotionMarkdown(pageUrl) {
  const token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('NOTION_TOKEN が未設定');

  // 1) URL からページ ID を抽出（32桁 hex）
  const idMatch = pageUrl.match(/[0-9a-f]{32}/i);
  if (!idMatch) throw new Error('ページ ID を URL から取得できませんでした');
  const pageId = idMatch[0].replace(/(.{8})(.{4})(.{4})(.{4})(.{12})/, '$1-$2-$3-$4-$5');

  // 2) 再帰的にブロックを取得
  const blocks = listAllBlocks(pageId, token);

  // 3) ブロック → Markdown へ変換
  Logger.log(blocksToMarkdown(blocks).trim());
  return blocksToMarkdown(blocks).trim();
}

/* -------- 内部関数 -------- */

/**
 * Notion blocks/children を page_size=100 で再帰取得
 * @return {Object[]} フル展開されたブロック配列
 */
function listAllBlocks(blockId, token) {
  const api = 'https://api.notion.com/v1/blocks/';
  const headers = {
    Authorization: `Bearer ${token}`,
    'Notion-Version': '2022-06-28'
  };
  let blocks = [], next;

  do {
    const url = `${api}${blockId}/children?page_size=100${next ? '&start_cursor=' + next : ''}`;
    const res = UrlFetchApp.fetch(url, {headers, muteHttpExceptions: true});
    if (res.getResponseCode() !== 200) throw new Error('Notion API error: ' + res.getContentText());

    const json = JSON.parse(res.getContentText());
    blocks = blocks.concat(json.results);
    next = json.has_more ? json.next_cursor : null;
  } while (next);

  // 子ページなど入れ子を再帰的に取る
  return blocks.map(b => {
    if (b.has_children) b.children = listAllBlocks(b.id, token);
    return b;
  });
}

function blocksToMarkdown(blocks, depth = 0) {
  const md      = [];
  const indent  = '  '.repeat(depth);

  blocks.forEach(b => {
    const txt = getPlainText(b);

    switch (b.type) {
      /* ----- テキスト系はそのまま ----- */
      case 'paragraph':              if (txt.trim()) md.push(indent + txt); md.push(''); break;
      case 'heading_1':              md.push(indent + '# '  + txt, '');              break;
      case 'heading_2':              md.push(indent + '## ' + txt, '');              break;
      case 'heading_3':              md.push(indent + '### '+ txt, '');              break;
      case 'bulleted_list_item':     md.push(indent + '- '  + txt);                  break;
      case 'numbered_list_item':     md.push(indent + '1. ' + txt);                  break;
      case 'to_do':                  md.push(indent + '- [ ] ' + txt);               break;
      case 'quote':                  md.push(indent + '> ' + txt, '');               break;

      /* ----- コードブロック ----- */
      case 'code':
        const lang = b.code?.language || '';
        md.push(indent + '```' + lang, txt, indent + '```', '');
        break;

      /* ----- 区切り線 ----- */
      case 'divider':
        md.push(indent + '---', '');
        break;

      /* ----- 画像 ----- */
      case 'image':
        const url = b.image?.file?.url || b.image?.external?.url || '';
        md.push(indent + `![](${url})`, '');
        break;

      /* ======== テーブル (NEW) ======== */
      case 'table':
        md.push(...tableBlockToMarkdown(b, indent), '');
        break;

      /* ----- その他 ----- */
      default:
        if (txt) md.push(indent + txt);
    }

    /* 子ブロックを再帰処理 */
    if (Array.isArray(b.children) && b.children.length) {
      md.push(blocksToMarkdown(b.children, depth + 1));
      if (!['bulleted_list_item','numbered_list_item'].includes(b.type)) md.push('');
    }
  });

  return md.join('\n');
}

/* ------------ テーブル変換 ------------ */
function tableBlockToMarkdown(tableBlock, indent) {
  const hasHeaderRow = tableBlock.table?.has_column_header;
  const rows         = (tableBlock.children || []).filter(r => r.type === 'table_row');
  if (!rows.length) return [];

  const lines = [];
  rows.forEach((row, idx) => {
    const cells = (row.table_row?.cells || []).map(cell =>
      cell.map(r => r.plain_text).join('')   // マルチリッチテキスト → 1セル文字列
    );

    // Markdown 行: | cell1 | cell2 |
    lines.push(indent + '| ' + cells.join(' | ') + ' |');

    // ヘッダ行の直後に区切り線
    if (idx === 0 && hasHeaderRow) {
      lines.push(indent + '| ' + cells.map(() => '---').join(' | ') + ' |');
    }
  });
  return lines;
}

/* ------------ テキスト抽出 ------------ */
function getPlainText(block) {
  const obj   = block[block.type] || {};
  const rich  = obj.rich_text || obj.text || [];
  return rich.map(t => t.plain_text).join('');
}

/**
 * Markdown を Notion 既存ページ（block_id=page_id）へ書き出す
 * 必要スコープ：script.external_request
 * プロパティ：NOTION_TOKEN=secret_xxx
 *
 * @param {string} pageUrl   例: https://www.notion.so/Title-0123456789abcdef0123456789abcdef
 * @param {string} markdown  書き込みたい Markdown
 */
function writeMarkdownToNotion(pageUrl, markdown, outputType="output") {
  const token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('NOTION_TOKEN が未設定');

  const parentId = extractId(pageUrl);       // 親ページの ID
  const blocks   = markdownToBlocks(markdown);

  // ── 子ページを作成（最初の 100 ブロックだけ同梱）────────────
  const title   = outputType + ":" + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'); // ← タイトル
  const first   = blocks.slice(0, 100);      // Notion 制限：children ≤ 100

  const res = UrlFetchApp.fetch('https://api.notion.com/v1/pages', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      parent: { page_id: parentId },
      properties: {
        title: {                         // 既定のタイトルプロパティ
          title: [{ text: { content: title } }]
        }
      },
      children: first                    // 先頭 100 ブロックだけ同梱
    }),
    headers: {
      Authorization: `Bearer ${token}`,
      'Notion-Version': '2022-06-28'
    }
  });

  if (res.getResponseCode() !== 200) {
    throw new Error(`page 作成失敗 (${res.getResponseCode()}) → ${res.getContentText()}`);
  }
  const newPageId = JSON.parse(res.getContentText()).id.replace(/-/g, '');

  // ── 100 ブロック超過分は append で追追加 ────────────────────
  for (let i = 100; i < blocks.length; i += 100) {
    appendBlocks(newPageId, blocks.slice(i, i + 100), token);
  }
  return newPageId;
}

/**
 * blocks/{id}/children へ追追加するだけの軽量版
 * @param {string} parentId   追加先ブロック（今回は子ページ）の ID
 * @param {Array}  children   Notion ブロック配列 (≤100)
 * @param {string} token      Integration トークン
 */
function appendBlocks(parentId, children, token) {
  const url = `https://api.notion.com/v1/blocks/${parentId}/children`;

  const res = UrlFetchApp.fetch(url, {
    method: 'patch',
    contentType: 'application/json',
    payload: JSON.stringify({ children }),
    headers: {
      Authorization: `Bearer ${token}`,
      'Notion-Version': '2022-06-28'
    }
  });

  if (res.getResponseCode() !== 200) {
    throw new Error(`append 失敗 (${res.getResponseCode()}) → ${res.getContentText()}`);
  }
}

/* ---------- Markdown → Block オブジェクト ---------- */
function markdownToBlocks(md) {
  const lines  = md.replace(/\r\n/g, '\n').split('\n');
  const blocks = [];
  let bufCode  = null;

  lines.forEach(raw => {
    const line = raw.trimEnd();

    // ── コードブロック判定
    if (line.startsWith('```')) {
      if (bufCode) {            // 閉じタグ
        blocks.push(codeBlock(bufCode.code.join('\n'), bufCode.lang));
        bufCode = null;
      } else {                  // 開始タグ
        bufCode = {lang: line.slice(3).trim(), code: []};
      }
      return;
    }
    if (bufCode) { bufCode.code.push(line); return; }

    // ── テーブル行 (簡易) ──────────────
    if (/^\|.+\|$/.test(line)) {
      const cells = line.split('|').slice(1, -1).map(c => c.trim());
      // 区切り行ならスキップ
      if (cells.every(c => /^:?--+?:?$/.test(c))) return;

      const row  = tableRow(cells);
      const last = blocks[blocks.length - 1];

      const needNewTable =
        !(last && last.type === 'table' &&
          last.table.table_width === cells.length);

      if (needNewTable) {
        // ─ 新しいテーブルを開始
        blocks.push({
          object: 'block',
          type:   'table',
          table: {
            table_width:        cells.length,
            has_column_header:  true,
            has_row_header:     false,
            children:           [row]
          }
        });
      } else {
        // ─ 既存テーブルに行を追加
        last.table.children.push(row);
      }
      return;
    }

    // ── 見出し ────────────────────────
    if (line.startsWith('### ')) { blocks.push(heading(3, line.slice(4))); return; }
    if (line.startsWith('## '))  { blocks.push(heading(2, line.slice(3))); return; }
    if (line.startsWith('# '))   { blocks.push(heading(1, line.slice(2))); return; }

    // ── リスト ────────────────────────
    if (/^- /.test(line)) { blocks.push(listItem('bulleted_list_item', line.slice(2))); return; }
    if (/^\d+\.\s/.test(line)) { blocks.push(listItem('numbered_list_item', line.replace(/^\d+\.\s/, ''))); return; }

    // ── 引用 ──────────────────────────
    if (line.startsWith('> ')) { blocks.push(quote(line.slice(2))); return; }

    // ── 区切り線 ──────────────────────
    if (line === '---') { blocks.push({object:'block',type:'divider',divider:{}}); return; }

    // ── 段落 ──────────────────────────
    if (line) blocks.push(paragraph(line));
  });

  return blocks;
}

/* ---------- Block ビルダ ---------- */
const rich = t => [{type:'text',text:{content:t}}];

const paragraph = txt => ({object:'block',type:'paragraph',paragraph:{rich_text:rich(txt)}});
const heading   = (lvl, txt) => ({object:'block',type:`heading_${lvl}`,[`heading_${lvl}`]:{rich_text:rich(txt)}});
const listItem  = (t, txt) => ({object:'block',type:t,[t]:{rich_text:rich(txt)}});
const quote     = txt => ({object:'block',type:'quote',quote:{rich_text:rich(txt)}});
const codeBlock = (code, lang='') => ({object:'block',type:'code',code:{rich_text:rich(code),language:lang||'plain text'}});
const tableRow  = cells => ({
  object:'block',
  type:'table_row',
  table_row:{cells:cells.map(c => [{type:'text',text:{content:c}}])}
});

/* ---------- ユーティリティ ---------- */
function extractId(url) {
  const m = url.match(/[0-9a-f]{32}/i);
  if (!m) throw new Error('ページ ID を URL から取得できません');
  return m[0].replace(/(.{8})(.{4})(.{4})(.{4})(.{12})/, '$1-$2-$3-$4-$5');
}

DUMMY_MD =`
# サマリ
---

1. 主要な進捗・懸念点のまとめ  
   - 4 月の総消化金額は 3,170,543 円で、月次予算 5,750,000 円に対して 55.2％の進捗。営業日消化率（19/22＝86.4％）を考慮すると着地見込みは約 3,670 万円となり、依然として約 2,080 万円の未消化が見込まれる。  
   - CV は 159 件（着地見込み 184 件）で前月比 +35 件の増加見込み。特に Google P-MAX と Meta lcaWP が牽引している一方、Microsoft 指名、Meta cdp など一部キャンペーンは失速。  
   - CPA は全体で 19,941 円（Salesforce ベース 22,176 円）。Google 一般+注力kw や Microsoft 指名で CPA が急騰しており、早急な入札・キーワード精査が必要。  

2. 全体的な所感や方針転換など  
   - 予算未消化リスクが高いキャンペーン（Google 指名、Yahoo 指名・一般、Microsoft 指名）は、入札引き上げと広告バリエーション追加でボリューム確保を検討。  
   - 成果効率が高い Google P-MAX、Meta lcaWP には追加予算シフトを推奨。Meta サプライチェーンWP は微増の CV に対して CPA が悪化しており、クリエイティブ AB テストでの改善を優先。  
   - 訪問・案件はタイムラグ要素が大きく、現時点評価は限定的。来月以降に統計的有意性を持たせるため、Salesforce 連携データの遅延解消を進める。  

# 数値進捗
---

### ■予算の消化状況  
2025年4月27日時点での、配信月＝2025-04-01 の予算進捗状況は、次の表です。

| 媒体 | カテゴリー | 4月目標予算_増額 | 4月目標予算 | 4月消化金額 | 予算進捗率 | CPA | CPA(SF) | CV | CV(SF) | forecast_CV | forecast_CV(SF) |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Google | 一般+注力kw(lca、cdp) |  | 750,000 | 488,256 | 65.1% | 90,418 | 81,376 | 5 | 6 | 6 | 7 |
| Google | 指名 |  | 250,000 | 81,388 | 32.6% | 11,627 | 9,043 | 7 | 9 | 8 | 10 |
| Google | P-MAX |  | 1,000,000 | 830,465 | 83.0% | 30,089 | 31,941 | 28 | 26 | 32 | 30 |
| Yahoo | 指名 |  | 250,000 | 57,781 | 23.1% | 11,556 | 14,445 | 5 | 4 | 6 | 5 |
| Yahoo | 一般 |  | 250,000 | 147,107 | 58.8% | 24,518 | 29,421 | 6 | 5 | 7 | 6 |
| Yahoo | YDA |  | 250,000 | 0 | 0.0% | 0 | 0 | 0 | 0 | 0 | 0 |
| Microsoft | 一般+注力kw(lca、cdp) |  | 500,000 | 239,027 | 47.8% | 34,147 | 47,805 | 7 | 5 | 8 | 6 |
| Microsoft | 指名 |  | 250,000 | 54,368 | 21.7% | 4,943 | 5,437 | 11 | 10 | 13 | 12 |
| Meta | サプライチェーンWP |  | 1,000,000 | 775,325 | 77.5% | 14,910 | 18,031 | 52 | 43 | 60 | 50 |
| Meta | lcaWP |  | 500,000 | 243,737 | 48.7% | 18,749 | 20,311 | 13 | 12 | 15 | 14 |
| Meta | 人権DD |  | 250,000 | 0 | 0.0% | 0 | 0 | 0 | 0 | 0 | 0 |
| Meta | cdp WP |  | 500,000 | 253,089 | 50.6% | 10,124 | 11,004 | 25 | 23 | 29 | 27 |
| Meta | e-learing |  | 250,000 | 0 | 0.0% | 0 | 0 | 0 | 0 | 0 | 0 |
| Meta | セミナー |  | 250,000 | 0 | 0.0% | 0 | 0 | 0 | 0 | 0 | 0 |
| 全体 |  |  | 5,750,000 | 3,170,543 | 55.2% | 19,941 | 22,176 | 159 | 143 | 184 | 165 |

#### コメント
- 4 月全体の消化率は 55.2％で営業日進捗（86.4％）を大幅に下回り、残予算 2.58 百万円の消化が急務。  
- Google P-MAX（83.0％）、Meta サプライチェーンWP（77.5％）が高進捗。一方で Yahoo・Microsoft の指名系は 25％前後と低水準。  
- CV 観点では Google P-MAX と Meta lcaWP が伸長。CPA は Google 一般+注力kw と Microsoft 指名が急騰し、予算再配分の検討余地が大きい。  
`