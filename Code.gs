/**
 * @OnlyCurrentDoc
 */

/* =========================================================
 *  設定（フィールド名・選択肢・既定値）
 *  キー名がそのままフォームのフィールド名になる
 * ========================================================= */

const FIELDS = {
  mode: {
    default: 'move',
    options: [
      { key: 'move',  text: '○字ずらす（今の位置から）' },
      { key: 'set',   text: '左端から○字にする' },
      { key: 'reset', text: 'インデントを解除' },
      { key: 'none',  text: 'インデントは変えない（整形のみ）' },
    ],
  },
  move_target: {
    default: 'all',
    options: [
      { key: 'all',       text: '全体' },
      { key: 'first',     text: '1行目のみ' },
      { key: 'following', text: '2行目以降のみ' },
      { key: 'right',     text: '右端（右インデント）' },
    ],
  },
  move_amount: { default: '1', required: true },
  set_target: {
    default: 'all',
    options: [
      { key: 'all',       text: '全体' },
      { key: 'first',     text: '1行目のみ' },
      { key: 'following', text: '2行目以降のみ' },
      { key: 'right',     text: '右端（右インデント）' },
    ],
  },
  position: { default: '0', required: true },
  shape: {
    default: 'flush',
    options: [
      { key: 'flush',   text: '揃える' },
      { key: 'indent',  text: '字下げ（1行目だけ右へ）' },
      { key: 'hanging', text: 'ぶら下げ（2行目以降を右へ）' },
    ],
  },
  shape_amount: { default: '1', required: true },
  shape_base: {
    default: 'following',
    options: [
      { key: 'following', text: '2行目以降を指定位置に置く' },
      { key: 'first',     text: '1行目を指定位置に置く' },
    ],
  },
  space_chars: {
    default: 'none',
    options: [
      { key: 'none',    text: '削除しない' },
      { key: 'leading', text: '段落先頭の空白を削除' },
      { key: 'trim',    text: '段落の先頭と末尾の空白を削除' },
      { key: 'all',     text: '段落内のすべての空白を削除' },
      { key: 'all_but_label', text: '番号の後ろを除き、すべての空白を削除' },
    ],
  },
  para_spacing: {
    default: 'keep',
    options: [
      { key: 'keep',   text: '変更しない' },
      { key: 'zero',   text: '前後とも 0 にする（詰める）' },
      { key: 'custom', text: '値を指定する' },
    ],
  },
  spacing_before: { default: '' },
  spacing_after:  { default: '' },
  preview:        { default: '' }, // 'on' のときプレビューを表示
};

const LIMITS = {
  MAX_CHARS: 100,      // インデント量の上限（文字）
  MAX_SPACING_PT: 500, // 段落間隔の上限（pt）
};

const FALLBACK_FONT_SIZE = 11; // 文字サイズが取得できない場合（Googleドキュメント標準テキストの既定）
const SETTINGS_KEY = 'indentAddon.settings.v2';

const MSG = {
  NO_TARGET: '操作対象の段落が見つかりません。段落を選択するか、カーソルを段落内に置いてください。',
  DOC_UNAVAILABLE: 'このファイルは操作できません。Wordファイル(.docx)の場合は［ファイル］→［Googleドキュメントとして保存］で変換してからお使いください。',
  TIMEOUT: '処理がタイムアウトしました。選択範囲を小さくして再度お試しください。',
  UNEXPECTED: '予期せぬエラーが発生しました。',
  NOTHING_TO_DO: '実行する内容がありません。「操作」か「整形」のいずれかを設定してください。',
};

/* =========================================================
 *  トリガー／アクション（グローバル関数）
 * ========================================================= */

/** アドオンを開いたとき */
function onDocsHomepage(e) {
  return buildCard_(resolveSettings_(null));
}

/**
 * ドロップダウンが変わったとき：
 * 設定を保存し、関係する入力欄だけを表示し直す（＝無関係な項目を出さない）
 */
function onSettingsChange(e) {
  const s = resolveSettings_(e);
  saveSettings_(s);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildCard_(s)))
    .build();
}

/** 実行ボタン */
function runAction(e) {
  try {
    const s = resolveSettings_(e);
    saveSettings_(s);

    const direction = (e && e.parameters && e.parameters.direction) || null;
    const plan = buildPlan_(s, direction);

    const doc = getDocument_();
    const paragraphs = getTargetParagraphs_(doc);
    if (paragraphs.length === 0) return notify_(MSG.NO_TARGET);

    const ctx = createContext_(doc);
    paragraphs.forEach(p => applyToParagraph_(p, plan, ctx));

    return notify_(summarize_(plan, paragraphs.length, ctx));
  } catch (err) {
    return handleException_(err);
  }
}

/** 選択中（先頭）の段落の現在値を表示 */
function showCurrentValues(e) {
  try {
    const doc = getDocument_();
    const paragraphs = getTargetParagraphs_(doc);
    if (paragraphs.length === 0) return notify_(MSG.NO_TARGET);

    const p = paragraphs[0];
    const em = getCharWidth_(p, createContext_(doc));
    const start = num_(p.getIndentStart());
    const first = num_(p.getIndentFirstLine());
    const end   = num_(p.getIndentEnd());
    const ch = v => formatNumber_(v / em) + '字';
    const pt = v => (v === null || v === undefined) ? '既定' : formatNumber_(v) + 'pt';

    let shape = 'そろえ';
    if (first - start > 0.01) shape = '字下げ' + ch(first - start);
    else if (start - first > 0.01) shape = 'ぶら下げ' + ch(start - first);

    const msg =
      (paragraphs.length > 1 ? '【先頭の段落】' : '') +
      '1行目 ' + ch(first) + '・2行目以降 ' + ch(start) + '（' + shape + '）／右 ' + ch(end) +
      '／段落前 ' + pt(p.getSpacingBefore()) + '・段落後 ' + pt(p.getSpacingAfter()) +
      '（1字=' + formatNumber_(em) + 'pt）';
    // プレビュー表示中なら、図も選択中の段落で描き直す
    const s = resolveSettings_(e);
    const res = CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(msg));
    if (s.preview === 'on') res.setNavigation(CardService.newNavigation().updateCard(buildCard_(s)));
    return res.build();
  } catch (err) {
    return handleException_(err);
  }
}

/**
 * プレビュー用：選択中（先頭）の段落のインデントを文字数で返す。
 * 読めないとき（カーソルなし・権限なし等）は null。
 */
function readCurrentState_() {
  try {
    const doc = DocumentApp.getActiveDocument();
    if (!doc) return null;
    const paragraphs = getTargetParagraphs_(doc);
    if (paragraphs.length === 0) return null;
    const p = paragraphs[0];
    const em = getCharWidth_(p, createContext_(doc));
    const ch = v => Math.round(num_(v) / em * 10) / 10;
    return { first: ch(p.getIndentFirstLine()), start: ch(p.getIndentStart()), end: ch(p.getIndentEnd()) };
  } catch (err) {
    return null;
  }
}

/* =========================================================
 *  設定の読み込み・保存
 * ========================================================= */

function loadSavedSettings_() {
  try {
    const raw = PropertiesService.getUserProperties().getProperty(SETTINGS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (err) {
    return {};
  }
}

function saveSettings_(s) {
  try {
    PropertiesService.getUserProperties().setProperty(SETTINGS_KEY, JSON.stringify(s));
  } catch (err) {
    console.warn(err);
  }
}

/**
 * 既定値 ← 前回保存した値 ← 今回のフォーム入力 の順で重ねて設定を確定する。
 * 非表示の入力欄はフォームに含まれないため、保存値で補う。
 * 一方「表示されていたのに値が無い」入力欄は、ユーザーが空欄にしたものとして扱う。
 */
function resolveSettings_(e) {
  const s = {};
  const saved = loadSavedSettings_();
  Object.keys(FIELDS).forEach(key => {
    s[key] = (key in saved) ? String(saved[key]) : FIELDS[key].default;
    // 必須の数値欄が空のまま保存されていたら既定値に戻す（開いた直後から空欄にならないように）
    if (FIELDS[key].required && s[key].trim() === '') s[key] = FIELDS[key].default;
  });

  const input = (e && e.formInput) || {};
  const renderedParam = e && e.parameters && e.parameters.rendered;
  const rendered = renderedParam ? String(renderedParam).split(',') : [];

  Object.keys(FIELDS).forEach(key => {
    if (Object.prototype.hasOwnProperty.call(input, key)) {
      s[key] = String(input[key]);
    } else if (rendered.indexOf(key) !== -1) {
      s[key] = '';
    }
    if (key === 'space_chars' && s[key] === 'all_but_second') s[key] = 'all_but_label'; // 旧設定の引き継ぎ
    // 選択肢にない値（古い保存値など）は既定値に戻す
    const options = FIELDS[key].options;
    if (options && !options.some(o => o.key === s[key])) {
      s[key] = FIELDS[key].default;
    }
  });
  return s;
}

/** 現在の設定で表示される「空にできる」入力欄（テキスト入力・スイッチ）の一覧 */
function renderedTextFields_(s) {
  const fields = [];
  if (s.mode === 'move') fields.push('move_amount');
  if (s.mode === 'set') {
    fields.push('position');
    if (s.set_target === 'all' && s.shape !== 'flush') fields.push('shape_amount');
  }
  if (s.para_spacing === 'custom') fields.push('spacing_before', 'spacing_after');
  fields.push('preview'); // スイッチはオフのときフォームに値が入らないため常に記録する
  return fields;
}

/* =========================================================
 *  UI
 * ========================================================= */

function buildCard_(s) {
  const rendered = renderedTextFields_(s).join(',') || '-';
  const onChange = () => CardService.newAction()
    .setFunctionName('onSettingsChange')
    .setParameters({ rendered: rendered });
  const run = params => CardService.newAction()
    .setFunctionName('runAction')
    .setParameters(Object.assign({ rendered: rendered }, params || {}));
  // プレビュー表示中だけ、数値の入力や図に関わる項目の変更でも再描画する
  const live = s.preview === 'on';
  const liveChange = () => (live ? onChange() : null);
  const card = CardService.newCardBuilder();

  // ---- インデント ----
  const indent = CardService.newCardSection().setHeader('インデント');
  // 「左端から○字」で対象が右端のときは、選択肢の名前も「右端から○字」にする
  const modeLabels = (s.mode === 'set' && s.set_target === 'right') ? { set: '右端から○字にする' } : null;
  indent.addWidget(dropdown_('mode', '操作', s.mode, onChange(), modeLabels));

  switch (s.mode) {
    case 'move':
      indent.addWidget(dropdown_('move_target', '対象', s.move_target, onChange()));
      indent.addWidget(textInput_('move_amount', '文字数', s.move_amount,
        '0.5文字単位も可。1文字＝その段落の文字サイズ', liveChange()));
      break;

    case 'set':
      indent.addWidget(dropdown_('set_target', '対象', s.set_target, onChange()));
      indent.addWidget(textInput_('position', positionTitle_(s), s.position, positionHint_(s), liveChange()));
      if (s.set_target === 'all') {
        indent.addWidget(dropdown_('shape', '1行目の形', s.shape, onChange()));
        if (s.shape !== 'flush') {
          const label = s.shape === 'indent' ? '字下げ量' : 'ぶら下げ量';
          indent.addWidget(textInput_('shape_amount', label + '（文字数）',
            s.shape_amount || FIELDS.shape_amount.default, // 空欄のときは既定の 1 を表示
            '1行目と2行目以降の差', liveChange()));
          indent.addWidget(dropdown_('shape_base', '位置の基準', s.shape_base, liveChange()));
        }
      }
      break;

    case 'reset':
      indent.addWidget(CardService.newTextParagraph()
        .setText('選択した段落の左・1行目・右のインデントをすべて 0 にします。'));
      break;

    case 'none':
      indent.addWidget(CardService.newTextParagraph()
        .setText('インデントは変更せず、下の「整形」だけを実行します。'));
      break;
  }
  card.addSection(indent);

  // ---- 整形（インデント操作と同時に実行） ----
  const format = CardService.newCardSection().setHeader('同時に行う整形');
  format.addWidget(dropdown_('space_chars', '文字の空白', s.space_chars, liveChange()));
  format.addWidget(dropdown_('para_spacing', '段落の前後の間隔', s.para_spacing, onChange()));
  if (s.para_spacing === 'custom') {
    format.addWidget(textInput_('spacing_before', '段落前（pt）', s.spacing_before, '空欄なら変更しません', liveChange()));
    format.addWidget(textInput_('spacing_after',  '段落後（pt）', s.spacing_after,  '空欄なら変更しません', liveChange()));
  }
  card.addSection(format);

  // ---- プレビュー（スイッチで表示／非表示） ----
  card.addSection(buildPreviewSection_(s, onChange()));

  // ---- 実行ボタン ----
  const actions = CardService.newCardSection();
  if (s.mode === 'move') {
    const isRight = s.move_target === 'right';
    actions.addWidget(CardService.newButtonSet()
      .addButton(filledButton_(isRight ? '◀ 右端を左へ' : '◀ 左へ', run({ direction: 'left' })))
      .addButton(filledButton_(isRight ? '右端を右へ ▶' : '右へ ▶', run({ direction: 'right' }))));
  } else {
    const label = { set: '設定する', reset: 'インデントを解除', none: '整形を実行' }[s.mode];
    actions.addWidget(filledButton_(label, run()));
  }
  actions.addWidget(CardService.newTextButton()
    .setText('選択中の段落の現在値を確認')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showCurrentValues')
      .setParameters({ rendered: rendered })));
  card.addSection(actions);

  // ---- ヘルプ（折りたたみ） ----
  card.addSection(buildHelpSection_());

  return card.build();
}

function buildHelpSection_() {
  const usage =
    '<b>使い方</b>\n' +
    '① 段落を選択（カーソルを置くだけなら、その段落が対象）\n' +
    '② 「操作」と数値を設定\n' +
    '③ ボタンを押す\n' +
    '・1文字＝段落の先頭文字の文字サイズです。\n' +
    '・設定内容は次回開いたときも引き継がれます。';
  const notes =
    '<b>注意</b>\n' +
    '・Wordファイル(.docx)は［ファイル］→［Googleドキュメントとして保存］で変換してからお使いください。\n' +
    '・メイリオ等のプロポーショナルフォントでは、文字数ぴったりにならない場合があります。\n' +
    '・提案モードには対応していません（スクリプトの編集は常に直接編集になります）。提案は承認／拒否してからお使いください。';
  return CardService.newCardSection()
    .setHeader('ヘルプ・注意事項')
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(0)
    .addWidget(CardService.newTextParagraph().setText(usage))
    .addWidget(CardService.newTextParagraph().setText(notes));
}

function positionTitle_(s) {
  switch (s.set_target) {
    case 'first':     return '1行目の文字数（左端から）';
    case 'following': return '2行目以降の文字数（左端から）';
    case 'right':     return '文字数（右端から）';
    default:          return '文字数（左端から）';
  }
}

function positionHint_(s) {
  if (s.set_target === 'right') return '右インデントは右端から数えます';
  if (s.set_target === 'all' && s.shape !== 'flush') return '「位置の基準」で選んだ行をこの位置に置きます';
  return '0.5文字単位も可';
}

function dropdown_(name, title, current, onChangeAction, labelOverrides) {
  const w = CardService.newSelectionInput()
    .setFieldName(name)
    .setTitle(title)
    .setType(CardService.SelectionInputType.DROPDOWN);
  FIELDS[name].options.forEach(o => {
    const text = (labelOverrides && labelOverrides[o.key]) || o.text;
    w.addItem(text, o.key, o.key === current);
  });
  if (onChangeAction) w.setOnChangeAction(onChangeAction);
  return w;
}

function textInput_(name, title, value, hint, onChangeAction) {
  const w = CardService.newTextInput().setFieldName(name).setTitle(title);
  if (value !== undefined && value !== null && value !== '') w.setValue(String(value));
  if (hint) w.setHint(hint);
  if (onChangeAction) w.setOnChangeAction(onChangeAction);
  return w;
}

function filledButton_(text, action) {
  return CardService.newTextButton()
    .setText(text)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(action);
}

/* =========================================================
 *  プレビュー（ポンチ絵）
 *  文書には触れず、設定値だけから図を作る
 * ========================================================= */

const PREVIEW = {
  WIDTH: 16,                                // 図の1行の幅（文字）
  EXAMPLE: { first: 2, start: 2, end: 0 },  // 「今の段落」の例
  SAMPLE_TEXT: '　①　ああ　いい ',          // 空白処理の例文
  COLOR_TEXT: '#8d6e63',
  COLOR_SPACE: '#bdbdbd',
  COLOR_NOTE: '#5f6368',
  COLOR_ERROR: '#d93025',
};

function buildPreviewSection_(s, onChangeAction) {
  const live = s.preview === 'on';
  const section = CardService.newCardSection().setHeader('プレビュー');
  section.addWidget(CardService.newDecoratedText()
    .setText('動きのイメージを表示')
    .setBottomLabel('設定を変えるたびに図を更新します（そのぶん少し待ちが出ます）')
    .setWrapText(true)
    .setSwitchControl(CardService.newSwitch()
      .setFieldName('preview')
      .setValue('on')
      .setSelected(live)
      .setOnChangeAction(onChangeAction)));
  if (!live) return section;

  let html;
  try {
    html = previewHtml_(s, readCurrentState_());
  } catch (err) {
    html = font_(PREVIEW.COLOR_ERROR, (err && err.userFacing) ? err.message : '図を作れませんでした。');
  }
  section.addWidget(CardService.newTextParagraph().setText(html));
  return section;
}

function previewHtml_(s, current) {
  const plan = buildPlan_(s, 'right'); // 入力チェックも兼ねる（エラーは図の代わりに表示）
  // 右端を動かす例では、右側にも少し余白がある段落を例にする
  // 選択中の段落が読めればそれを「今」とし、読めなければ例の段落で描く
  const ex = current ||
    Object.assign({}, PREVIEW.EXAMPLE, plan.target === 'right' ? { end: 2 } : {});
  const nowLabel = current ? '選択中の段落（今）' : '例：今の段落';
  const blocks = [];
  let fractional = false;
  const fig = (title, st) => {
    if ([st.first, st.start, st.end].some(v => v % 1 !== 0)) fractional = true;
    return '<b>' + title + '</b>\n' + drawParagraph_(st) + '\n' + font_(PREVIEW.COLOR_NOTE, describe_(st));
  };

  switch (plan.mode) {
    case 'move': {
      const isRight = plan.target === 'right';
      blocks.push(fig(nowLabel, ex));
      blocks.push(fig(isRight ? '◀ 右端を左へ' : '◀ 左へ', computeMove_(ex, plan.target, -plan.chars)));
      blocks.push(fig(isRight ? '右端を右へ ▶' : '右へ ▶', computeMove_(ex, plan.target, plan.chars)));
      break;
    }
    case 'set': {
      const after = { first: ex.first, start: ex.start, end: ex.end };
      if (plan.target === 'all')       { after.first = plan.firstChars; after.start = plan.startChars; }
      if (plan.target === 'first')     after.first = plan.chars;
      if (plan.target === 'following') after.start = plan.chars;
      if (plan.target === 'right')     after.end = plan.chars;
      if (plan.target !== 'all') blocks.push(fig(nowLabel, ex));
      blocks.push(fig('設定後', after));
      break;
    }
    case 'reset':
      blocks.push(fig('解除後', { first: 0, start: 0, end: 0 }));
      break;
    default:
      blocks.push(font_(PREVIEW.COLOR_NOTE, 'インデントは変わりません。'));
  }
  if (fractional) blocks.push(font_(PREVIEW.COLOR_NOTE, '※図では0.5字を四捨五入しています'));

  if (plan.spaceChars !== 'none') {
    let str = PREVIEW.SAMPLE_TEXT;
    const fake = {
      getText: () => str,
      deleteText: (a, b) => { str = str.slice(0, a) + str.slice(b + 1); },
    };
    handleSpaceChars_({ editAsText: () => fake }, plan.spaceChars);
    blocks.push('<b>文字の空白</b>\n' + showSpaces_(PREVIEW.SAMPLE_TEXT) + '\n→ ' + showSpaces_(str));
  }

  if (plan.spacing) {
    const v = x => (x === null ? '変更なし' : formatNumber_(x) + 'pt');
    blocks.push('<b>段落の前後の間隔</b>\n段落前 ' + v(plan.spacing.before) + '・段落後 ' + v(plan.spacing.after));
  }
  return blocks.join('\n\n');
}

/** 段落の図：1行目と2行目以降の2行。・＝インデント、■＝本文 */
function drawParagraph_(st) {
  const W = PREVIEW.WIDTH;
  const right = Math.min(Math.round(st.end), W - 1);
  const line = left => {
    const l = Math.min(Math.round(left), W - right - 1);
    return '｜' +
      font_(PREVIEW.COLOR_SPACE, '・'.repeat(Math.max(0, l))) +
      font_(PREVIEW.COLOR_TEXT, '■'.repeat(W - Math.max(0, l) - right)) +
      font_(PREVIEW.COLOR_SPACE, '・'.repeat(right)) + '｜';
  };
  return line(st.first) + '\n' + line(st.start);
}

function describe_(st) {
  return '1行目 ' + formatNumber_(st.first) + '字・2行目以降 ' + formatNumber_(st.start) + '字' +
    (st.end ? '・右 ' + formatNumber_(st.end) + '字' : '');
}

/** 空白を見える記号にする（全角□・半角␣） */
function showSpaces_(str) {
  return '「' + str.split('').map(ch => {
    if (ch === '\u3000') return font_(PREVIEW.COLOR_SPACE, '□');
    if (isSpaceChar_(ch)) return font_(PREVIEW.COLOR_SPACE, '␣');
    return ch;
  }).join('') + '」';
}

function font_(color, text) {
  return text ? '<font color="' + color + '">' + text + '</font>' : '';
}

/* =========================================================
 *  入力値の検証 → 実行プラン
 * ========================================================= */

function buildPlan_(s, direction) {
  const plan = { mode: s.mode, spaceChars: s.space_chars, spacing: null };

  if (s.mode === 'move') {
    if (direction !== 'left' && direction !== 'right') throw userError_(MSG.UNEXPECTED);
    plan.target = s.move_target;
    plan.chars = parseChars_(s.move_amount, '文字数', false);
    plan.sign = direction === 'right' ? 1 : -1;

  } else if (s.mode === 'set') {
    plan.target = s.set_target;
    const n = parseChars_(s.position, '文字数', true);
    plan.chars = n;

    if (plan.target === 'all') {
      let first = n;
      let start = n;
      if (s.shape !== 'flush') {
        const label = s.shape === 'indent' ? '字下げ量' : 'ぶら下げ量';
        const m = parseChars_(s.shape_amount, label, false);
        const diff = s.shape === 'indent' ? m : -m; // 1行目 − 2行目以降
        if (s.shape_base === 'first') { first = n; start = n - diff; }
        else                          { start = n; first = n + diff; }
        if (first < 0 || start < 0) {
          const line = first < 0 ? '1行目' : '2行目以降';
          throw userError_(label + '（' + formatNumber_(m) + '字）が文字数（' + formatNumber_(n) +
            '字）より大きいため、' + line + 'が左端より左に出てしまいます。文字数を' +
            formatNumber_(m) + '字以上にするか、「位置の基準」を変えてください。');
        }
      }
      plan.firstChars = first;
      plan.startChars = start;
    }
  }

  if (s.para_spacing === 'zero') {
    plan.spacing = { before: 0, after: 0 };
  } else if (s.para_spacing === 'custom') {
    const before = parseOptionalPt_(s.spacing_before, '段落前');
    const after  = parseOptionalPt_(s.spacing_after,  '段落後');
    if (before === null && after === null) {
      throw userError_('段落前・段落後の少なくとも一方に数値を入力してください（単位: pt）。');
    }
    plan.spacing = { before: before, after: after };
  }

  if (s.mode === 'none' && s.space_chars === 'none' && !plan.spacing) {
    throw userError_(MSG.NOTHING_TO_DO);
  }
  return plan;
}

function parseChars_(raw, label, allowZero) {
  const str = toHalfWidth_(String(raw === undefined || raw === null ? '' : raw)).trim();
  if (str === '') throw userError_(label + 'を入力してください。');
  const v = parseNumber_(str);
  if (isNaN(v)) throw userError_(label + 'には数値を入力してください（例: 1、1.5）。');
  if (!allowZero && v === 0) throw userError_(label + 'には 0 より大きい数値を入力してください。');
  if (v > LIMITS.MAX_CHARS) throw userError_(label + 'は ' + LIMITS.MAX_CHARS + ' 字以下にしてください。');
  return v;
}

function parseOptionalPt_(raw, label) {
  const str = toHalfWidth_(String(raw === undefined || raw === null ? '' : raw)).trim();
  if (str === '') return null;
  const v = parseNumber_(str);
  if (isNaN(v)) throw userError_(label + 'には数値（pt）を入力してください。');
  if (v > LIMITS.MAX_SPACING_PT) throw userError_(label + 'は ' + LIMITS.MAX_SPACING_PT + 'pt 以下にしてください。');
  return v;
}

/** 0以上の数値のみ受け付ける（負数・"1abc" などは NaN） */
function parseNumber_(str) {
  return /^(\d+(\.\d*)?|\.\d+)$/.test(str) ? parseFloat(str) : NaN;
}

/** 全角の数字と小数点を半角に変換する */
function toHalfWidth_(str) {
  return str.replace(/[０-９．]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
}

/* =========================================================
 *  段落への適用
 * ========================================================= */

function createContext_(doc) {
  let body = null;
  try { body = doc.getBody(); } catch (err) { /* noop */ }
  return { body: body, headingSizes: {}, clamped: 0 };
}

function applyToParagraph_(p, plan, ctx) {
  // 空白の削除はインデント計算（先頭文字のサイズ取得）より先に行う
  handleSpaceChars_(p, plan.spaceChars);

  if (plan.spacing) {
    if (plan.spacing.before !== null) p.setSpacingBefore(plan.spacing.before);
    if (plan.spacing.after  !== null) p.setSpacingAfter(plan.spacing.after);
  }

  switch (plan.mode) {
    case 'move':
      moveIndent_(p, plan, getCharWidth_(p, ctx), ctx);
      break;
    case 'set':
      setIndent_(p, plan, getCharWidth_(p, ctx));
      break;
    case 'reset':
      p.setIndentStart(0);
      p.setIndentFirstLine(0);
      p.setIndentEnd(0);
      break;
  }
}

/*
 * ★ setAttributes() は段落配下の文字書式まで書き換えることがあるため、
 *   インデントは専用セッター（setIndentStart など）だけで変更する。
 */

/** 相対移動 */
function moveIndent_(p, plan, em, ctx) {
  const cur = {
    first: num_(p.getIndentFirstLine()),
    start: num_(p.getIndentStart()),
    end:   num_(p.getIndentEnd()),
  };
  const r = computeMove_(cur, plan.target, plan.chars * em * plan.sign);
  if (r.clamped) ctx.clamped++;
  if (plan.target === 'all' || plan.target === 'first') p.setIndentFirstLine(round_(r.first));
  if (plan.target === 'all' || plan.target === 'following') p.setIndentStart(round_(r.start));
  if (plan.target === 'right') p.setIndentEnd(round_(r.end));
}

/**
 * 移動後のインデントを計算する（実行とプレビューで共用）
 * delta は右向きが正。単位は呼び出し側に合わせる（pt でも文字数でも可）
 */
function computeMove_(cur, target, delta) {
  const r = { first: cur.first, start: cur.start, end: cur.end, clamped: false };
  switch (target) {
    case 'all': {
      // 左端に当たったら、そこで止める（1行目と2行目以降の差＝形は崩さない）
      const d = Math.max(delta, -Math.min(cur.start, cur.first));
      r.clamped = d !== delta;
      r.first += d;
      r.start += d;
      break;
    }
    case 'first':
      r.clamped = cur.first + delta < 0;
      r.first = Math.max(0, cur.first + delta);
      break;
    case 'following':
      r.clamped = cur.start + delta < 0;
      r.start = Math.max(0, cur.start + delta);
      break;
    case 'right':
      // 右端を「右へ」＝右インデントを減らす
      r.clamped = cur.end - delta < 0;
      r.end = Math.max(0, cur.end - delta);
      break;
  }
  return r;
}

/** 絶対値で設定 */
function setIndent_(p, plan, em) {
  switch (plan.target) {
    case 'all':
      p.setIndentStart(round_(plan.startChars * em));
      p.setIndentFirstLine(round_(plan.firstChars * em));
      break;
    case 'first':
      p.setIndentFirstLine(round_(plan.chars * em));
      break;
    case 'following':
      p.setIndentStart(round_(plan.chars * em));
      break;
    case 'right':
      p.setIndentEnd(round_(plan.chars * em));
      break;
  }
}

/**
 * 1文字分の幅（pt）＝段落の最初の「空白以外の文字」のサイズ。
 * 文字に直接サイズが付いていない場合は、段落スタイル（見出し／標準テキスト）のサイズを使う。
 */
function getCharWidth_(p, ctx) {
  try {
    const text = p.editAsText();
    const str = text.getText();
    if (str.length > 0) {
      let i = 0;
      while (i < str.length - 1 && isSpaceChar_(str.charAt(i))) i++;
      const size = text.getFontSize(i);
      if (size) return size;
    }
  } catch (err) { /* noop */ }

  try {
    const size = p.getAttributes()[DocumentApp.Attribute.FONT_SIZE];
    if (size) return size;
  } catch (err) { /* noop */ }

  try {
    const heading = p.getHeading();
    const key = String(heading);
    if (!(key in ctx.headingSizes)) {
      const attrs = ctx.body ? ctx.body.getHeadingAttributes(heading) : null;
      ctx.headingSizes[key] = (attrs && attrs[DocumentApp.Attribute.FONT_SIZE]) || null;
    }
    if (ctx.headingSizes[key]) return ctx.headingSizes[key];
  } catch (err) { /* noop */ }

  return FALLBACK_FONT_SIZE;
}

function num_(v) { return Number(v) || 0; }
function round_(v) { return Math.round(v * 100) / 100; }
function formatNumber_(v) { return String(Math.round(v * 10) / 10); }

/* =========================================================
 *  文字の空白の削除
 *  ★ setText() は段落を丸ごと置き換え、文字書式・リンク・画像・提案が失われるため、
 *    deleteText() で必要な範囲だけを削除する。
 * ========================================================= */

/** 空白とみなす文字（改行は含めない＝段落内のソフト改行を壊さない） */
const SPACE_CHARS_ = ' \t        ' +
                     '      　';

/** 番号（①、(1)、（１３４）など）とみなす最大の文字数 */
const LABEL_MAX_CHARS = 6;

function isSpaceChar_(ch) {
  return SPACE_CHARS_.indexOf(ch) !== -1;
}

function handleSpaceChars_(p, mode) {
  if (mode === 'none' || !p.editAsText) return;
  const text = p.editAsText();
  switch (mode) {
    case 'leading':
      removeLeadingSpaces_(text);
      break;
    case 'trim':
      removeTrailingSpaces_(text);
      removeLeadingSpaces_(text);
      break;
    case 'all':
      removeAllSpaces_(text, -1);
      break;
    case 'all_but_label': {
      // 「①　本文」「(1) 本文」「（１３４）　本文」用：
      // 先頭の空白を消したうえで、最初の空白が先頭から LABEL_MAX_CHARS 字以内にあれば
      // それを「番号と本文の間の空白」とみなして1文字だけ残す
      removeLeadingSpaces_(text);
      const str = text.getText();
      let first = -1;
      for (let i = 0; i < str.length; i++) {
        if (isSpaceChar_(str.charAt(i))) { first = i; break; }
      }
      removeAllSpaces_(text, (first !== -1 && first <= LABEL_MAX_CHARS) ? first : -1);
      break;
    }
  }
}

function removeLeadingSpaces_(text) {
  const s = text.getText();
  let n = 0;
  while (n < s.length && isSpaceChar_(s.charAt(n))) n++;
  if (n > 0) text.deleteText(0, n - 1);
}

function removeTrailingSpaces_(text) {
  const s = text.getText();
  let i = s.length;
  while (i > 0 && isSpaceChar_(s.charAt(i - 1))) i--;
  if (i < s.length) text.deleteText(i, s.length - 1);
}

/**
 * 後ろから連続する空白のまとまりごとに削除（残る文字の書式を保持）
 * keepIndex の位置の文字は空白でも残す（-1 なら残さない）
 */
function removeAllSpaces_(text, keepIndex) {
  const s = text.getText();
  let end = -1;
  for (let i = s.length - 1; i >= 0; i--) {
    if (i !== keepIndex && isSpaceChar_(s.charAt(i))) {
      if (end === -1) end = i;
    } else if (end !== -1) {
      text.deleteText(i + 1, end);
      end = -1;
    }
  }
  if (end !== -1) text.deleteText(0, end);
}

/* =========================================================
 *  対象段落の取得
 * ========================================================= */

function getDocument_() {
  let doc = null;
  try {
    doc = DocumentApp.getActiveDocument();
  } catch (err) {
    console.warn(err);
  }
  if (!doc) throw userError_(MSG.DOC_UNAVAILABLE);
  return doc;
}

function getTargetParagraphs_(doc) {
  const T = DocumentApp.ElementType;
  const result = [];
  const seen = {};

  const isParagraph = el => {
    const t = el.getType();
    return t === T.PARAGRAPH || t === T.LIST_ITEM;
  };
  const isBoundary = t =>
    t === T.BODY_SECTION || t === T.TABLE_CELL || t === T.HEADER_SECTION ||
    t === T.FOOTER_SECTION || t === T.FOOTNOTE_SECTION;

  // ★ Apps Script は同じ段落でも取得経路ごとに別オブジェクトを返すため、
  //   Set では重複を除けない（＝同じ段落が2回移動される）。文書内の位置で判定する。
  const add = el => {
    const key = elementKey_(el);
    if (!seen[key]) {
      seen[key] = true;
      result.push(el);
    }
  };

  const findParagraph = el => {
    let cur = el;
    while (cur) {
      if (isParagraph(cur)) return cur;
      if (isBoundary(cur.getType())) return null;
      cur = cur.getParent();
    }
    return null;
  };

  const collect = el => {
    if (isParagraph(el)) { add(el); return; }
    const t = el.getType();
    if (t === T.TABLE || t === T.TABLE_ROW || t === T.TABLE_CELL) {
      for (let i = 0; i < el.getNumChildren(); i++) collect(el.getChild(i));
      return;
    }
    const p = findParagraph(el);
    if (p) add(p);
  };

  const selection = doc.getSelection();
  if (selection) {
    selection.getRangeElements().forEach(re => {
      const el = re.getElement();
      if (re.isPartial()) {
        const p = findParagraph(el);
        if (p) add(p);
      } else {
        collect(el);
      }
    });
  } else {
    const cursor = doc.getCursor();
    if (cursor) {
      const p = findParagraph(cursor.getElement());
      if (p) add(p);
    }
  }
  return result;
}

/** 文書内の位置（親からの添字の列）を文字列化したもの */
function elementKey_(el) {
  const path = [];
  let cur = el;
  try {
    while (cur) {
      const parent = cur.getParent();
      if (!parent) break;
      path.push(parent.getChildIndex(cur));
      cur = parent;
    }
  } catch (err) { /* noop */ }
  path.push(cur ? String(cur.getType()) : '?');
  return path.reverse().join('/');
}

/* =========================================================
 *  通知・エラー
 * ========================================================= */

function summarize_(plan, count, ctx) {
  let msg;
  switch (plan.mode) {
    case 'move': {
      const dir = plan.sign > 0 ? '右' : '左';
      const what = plan.target === 'right' ? '右端を' : '';
      msg = count + '段落の' + what + dir + 'へ' + formatNumber_(plan.chars) + '字移動しました。';
      break;
    }
    case 'set':
      msg = count + '段落に設定しました' + (plan.target === 'all'
        ? '（1行目 ' + formatNumber_(plan.firstChars) + '字・2行目以降 ' + formatNumber_(plan.startChars) + '字）。'
        : '。');
      break;
    case 'reset':
      msg = count + '段落のインデントを解除しました。';
      break;
    default:
      msg = count + '段落を整形しました。';
  }
  if (ctx.clamped > 0) msg += ' ※' + ctx.clamped + '段落は端に達したため途中で止めました。';
  return msg;
}

function notify_(text) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(text))
    .build();
}

function userError_(message) {
  const err = new Error(message);
  err.userFacing = true;
  return err;
}

function handleException_(err) {
  if (err && err.userFacing) return notify_(err.message);
  console.error(err);
  const message = String((err && err.message) || err);
  if (/Exceeded maximum execution time|最大実行時間/.test(message)) return notify_(MSG.TIMEOUT);
  return notify_(MSG.UNEXPECTED + '（' + message.slice(0, 80) + '）');
}
