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
      { key: 'move',  text: '↔️ ずらす' },
      { key: 'set',   text: '📏 位置指定' },
      { key: 'reset', text: '↩️ 解除' },
      { key: 'none',  text: '✨ 整形のみ' },
    ],
  },
  move_target: {
    default: 'all',
    options: [
      { key: 'all',       text: '▤ 全体' },
      { key: 'first',     text: '1️⃣ 1行目' },
      { key: 'following', text: '2️⃣ 2行目〜' },
      { key: 'right',     text: '➡️ 右端' },
    ],
  },
  // 文字数はドロップダウンで選ぶ（最大5字）
  move_amount: { default: '1', dropdown: true, options: [{ key: '0.5', text: '0.5字' }, { key: '1', text: '1字' }, { key: '2', text: '2字' }, { key: '3', text: '3字' }, { key: '4', text: '4字' }, { key: '5', text: '5字' }] },
  set_target: {
    default: 'all',
    options: [
      { key: 'all',       text: '▤ 全体' },
      { key: 'first',     text: '1️⃣ 1行目' },
      { key: 'following', text: '2️⃣ 2行目〜' },
      { key: 'right',     text: '➡️ 右端' },
    ],
  },
  position: { default: '0', dropdown: true, options: [{ key: '0', text: '0字' }, { key: '0.5', text: '0.5字' }, { key: '1', text: '1字' }, { key: '2', text: '2字' }, { key: '3', text: '3字' }, { key: '4', text: '4字' }, { key: '5', text: '5字' }] },
  shape: {
    default: 'flush',
    options: [
      { key: 'flush',   text: '≡ 揃える' },
      { key: 'indent',  text: '⇥ 字下げ' },
      { key: 'hanging', text: '⇤ ぶら下げ' },
    ],
  },
  shape_base: {
    default: 'following',
    options: [
      { key: 'following', text: '2️⃣ 2行目〜' },
      { key: 'first',     text: '1️⃣ 1行目' },
    ],
  },
  space_chars: {
    default: 'none',
    options: [
      { key: 'none',    text: '🚫 しない' },
      { key: 'leading', text: '⬅️ 先頭' },
      { key: 'trim',    text: '↔️ 前後' },
      { key: 'all',     text: '🧹 全部' },
      { key: 'all_but_label', text: '① 番号の後は残す' },
    ],
  },
  para_spacing: {
    default: 'keep',
    options: [
      { key: 'keep',   text: '🚫 そのまま' },
      { key: 'zero',   text: '⏬ 0pt' },
      { key: 'custom', text: '✏️ 指定' },
    ],
  },
  spacing_before: { default: '' },
  spacing_after:  { default: '' },
  preview:        { default: '' }, // 'on' のときプレビューを表示（ボタンで入り切り。ui が light のとき）
  // 画面の動き（ヘルプ欄で切り替え）
  //   light：プレビューはボタンで入り切り。非表示中は更新を最小限に（項目が変わらない選択肢はラジオボタン）
  //   live ：プレビューを常に表示し、設定を変えるたびに描き直す（選択肢はすべてボタン）
  ui: {
    default: 'light',
    options: [
      { key: 'light', text: '🪶 軽快' },
      { key: 'live',  text: '👀 常時プレビュー' },
    ],
  },
};

const SHAPE_CHARS = 1; // 字下げ・ぶら下げの量（文字）

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
 * 入力欄・スイッチが変わったとき：設定を保存して表示し直す
 */
function onSettingsChange(e) {
  const s = resolveSettings_(e);
  saveSettings_(s);
  return rerender_(s);
}

/**
 * 選択肢のボタンを押したとき：その値を選んだ状態にして表示し直す
 * （関係する項目だけを出し直す＝無関係な項目は出さない）
 */
function onChoose(e) {
  const s = resolveSettings_(e);
  const field = e && e.parameters && e.parameters.field;
  const value = e && e.parameters && e.parameters.value;
  const options = FIELDS[field] && FIELDS[field].options;
  if (options && options.some(o => o.key === value)) s[field] = value;
  saveSettings_(s);
  return rerender_(s);
}

function rerender_(s) {
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

    const res = CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(summarize_(plan, paragraphs.length, ctx)));
    // プレビュー表示中だけ、図の「今」を実行後の段落に合わせて描き直す
    // （非表示中はパネルを更新しない）
    if (previewShown_(s)) res.setNavigation(CardService.newNavigation().updateCard(buildCard_(s)));
    return res.build();
  } catch (err) {
    return handleException_(err);
  }
}

/**
 * 🔄 現在値：選択中の段落を読み直してプレビューを描き直す。
 * 通知（下に出る黒い帯）はパネルを隠すので出さず、値はプレビューに表示する。
 * （値はプレビューに出すので、プレビューが切れていたら入れる）
 */
function refreshCurrent(e) {
  try {
    const s = resolveSettings_(e);
    if (!alwaysPreview_(s)) s.preview = 'on';
    saveSettings_(s);
    return rerender_(s);
  } catch (err) {
    return handleException_(err);
  }
}

/** 「👀 常時プレビュー」の画面か */
function alwaysPreview_(s) {
  return s.ui === 'live';
}

/** プレビューを描くか（常時プレビュー、または 👀 表示 中） */
function previewShown_(s) {
  return alwaysPreview_(s) || s.preview === 'on';
}

/** 👀 プレビューの表示／非表示を切り替える */
function togglePreview(e) {
  const s = resolveSettings_(e);
  s.preview = s.preview === 'on' ? '' : 'on';
  saveSettings_(s);
  return rerender_(s);
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
    return {
      first: ch(p.getIndentFirstLine()),
      start: ch(p.getIndentStart()),
      end:   ch(p.getIndentEnd()),
      before: p.getSpacingBefore(), // pt（null＝スタイルの既定）
      after:  p.getSpacingAfter(),
      em: em,                       // 1字の大きさ（pt）
      multiple: paragraphs.length > 1,
    };
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

/** 現在の設定で表示される「空にできる」入力欄（テキスト入力）の一覧 */
function renderedTextFields_(s) {
  const fields = [];
  if (s.para_spacing === 'custom') fields.push('spacing_before', 'spacing_after');
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
  const choices = (field, labelOverrides) => choiceRow_(field, s[field], rendered, labelOverrides);
  // プレビュー表示中だけ、文字数や pt の入力を変えたときに図を描き直す
  // （非表示のときは描き直さない＝入力のたびに画面が切り替わらない。値は実行時に読み取る）
  const liveChange = () => (previewShown_(s) ? onChange() : null);
  // 項目が変わらない選択肢：🪶 軽快 ではラジオボタン、👀 常時プレビュー ではボタン
  const classic = alwaysPreview_(s);
  const pick = (section, field, title) => {
    if (classic) section.addWidget(rowLabel_(title)).addWidget(choices(field));
    else section.addWidget(radio_(field, title, s[field], liveChange()));
  };

  const card = CardService.newCardBuilder();

  // ---- 実行ボタン（一番上） ----
  const actions = CardService.newCardSection();
  const runButtons = CardService.newButtonSet();
  if (s.mode === 'move') {
    // ボタン名は対象によらず ◀ ／ ▶ だけ（対象はすぐ下の選択肢で分かる）
    runButtons
      .addButton(filledButton_('◀', run({ direction: 'left' })))
      .addButton(filledButton_('▶', run({ direction: 'right' })));
  } else {
    const label = { set: '✔️ 設定', reset: '↩️ 解除', none: '✨ 整形' }[s.mode];
    runButtons.addButton(filledButton_(label, run()));
  }
  runButtons.addButton(CardService.newTextButton()
    .setText('🔄 現在値')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('refreshCurrent')
      .setParameters({ rendered: rendered })));
  actions.addWidget(runButtons);

  // ---- インデント ----
  const indent = CardService.newCardSection().addWidget(sectionHeading_('📐 インデント'));
  indent.addWidget(rowLabel_('🛠 操作'));
  indent.addWidget(choices('mode'));
  // 区切りの考え方：
  //   大きな区切り（インデント／整形／プレビュー）＝欄の境目（端から端までの線＋見出し）
  //   小さな区切り（どれか1つを選ぶボタンのまとまりごと）＝欄の中の細い線
  // 選択肢の形：
  //   ボタン＝選ぶと出てくる項目が変わるもの（操作・対象（位置指定）・1行目の形・段落間隔）。押すと描き直す
  //   ラジオボタン＝選んでも項目が変わらないもの。プレビュー非表示中は選んでもパネルを更新しない
  //                 （👀 常時プレビュー ではこれもボタン）
  //   プルダウン＝文字数（選択肢が多く、縦に並べると長くなるため）
  indent.addWidget(CardService.newDivider());

  switch (s.mode) {
    case 'move':
      pick(indent, 'move_target', '🎯 対象');
      indent.addWidget(dropdown_('move_amount', '🔢 文字数', s.move_amount, liveChange()));
      break;

    case 'set':
      indent.addWidget(rowLabel_('🎯 対象'));
      indent.addWidget(choices('set_target'));
      indent.addWidget(dropdown_('position', positionTitle_(s), s.position, liveChange()));
      if (s.set_target === 'all') {
        indent.addWidget(CardService.newDivider());
        indent.addWidget(rowLabel_('📐 1行目の形'));
        indent.addWidget(choices('shape'));
        if (s.shape !== 'flush') {
          if (classic) indent.addWidget(CardService.newDivider());
          pick(indent, 'shape_base', '📍 指定位置に置く行');
        }
      }
      break;

    case 'reset':
      indent.addWidget(CardService.newTextParagraph().setText('↩️ 1行目・2行目〜・右端をすべて 0字に戻します'));
      break;

    case 'none':
      indent.addWidget(CardService.newTextParagraph().setText('✨ インデントは変えず、下の整形だけ行います'));
      break;
  }

  // ---- 整形（インデント操作と同時に実行） ----
  const format = CardService.newCardSection().addWidget(sectionHeading_('🧽 整形'));
  pick(format, 'space_chars', '✂️ 空白の削除');
  format.addWidget(CardService.newDivider()); // 空白の削除 ／ 段落の前後の間隔 の区切り
  format.addWidget(rowLabel_('↕️ 段落の前後の間隔'));
  format.addWidget(choices('para_spacing'));
  if (s.para_spacing === 'custom') {
    format.addWidget(textInput_('spacing_before', '⬆️ 段落前 pt', s.spacing_before,
      '空欄＝そのまま', liveChange()));
    format.addWidget(textInput_('spacing_after',  '⬇️ 段落後 pt', s.spacing_after,
      '空欄＝そのまま', liveChange()));
  }

  // 並び順：実行ボタンを一番上に
  card.addSection(actions);
  card.addSection(indent);
  card.addSection(format);
  card.addSection(buildDisplaySection_(s, rendered));
  card.addSection(buildHelpSection_(s, rendered));

  return card.build();
}

/** 選択中のボタンの色（実行ボタンの青と区別する） */
const CHOICE_SELECTED_COLOR = '#455a64';

/**
 * 選択肢をボタンの列で表示する（選択中は塗りつぶし）。
 * 幅に収まらないときはカード側で自動的に2行に折り返す。
 */
function choiceRow_(field, current, rendered, labelOverrides) {
  const row = CardService.newButtonSet();
  FIELDS[field].options.forEach(o => {
    const btn = CardService.newTextButton()
      .setText((labelOverrides && labelOverrides[o.key]) || o.text)
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onChoose')
        .setParameters({ rendered: rendered, field: field, value: o.key }));
    if (o.key === current) {
      btn.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor(CHOICE_SELECTED_COLOR);
    }
    row.addButton(btn);
  });
  return row;
}

/**
 * プルダウン／ラジオボタン（選択肢の文字はボタンと同じアイコン＋短い言葉）。
 * プレビュー表示中だけ、変えると図を描き直す。
 */
function dropdown_(field, title, current, onChangeAction) {
  return selectionInput_(field, title, current, onChangeAction, CardService.SelectionInputType.DROPDOWN);
}

function radio_(field, title, current, onChangeAction) {
  return selectionInput_(field, title, current, onChangeAction, CardService.SelectionInputType.RADIO_BUTTON);
}

function selectionInput_(field, title, current, onChangeAction, type) {
  const w = CardService.newSelectionInput()
    .setFieldName(field)
    .setTitle(title)
    .setType(type);
  if (onChangeAction) w.setOnChangeAction(onChangeAction);
  FIELDS[field].options.forEach(o => w.addItem(o.text, o.key, o.key === current));
  return w;
}

/** 大項目の見出しの色（太い縦棒と文字） */
const HEADING_COLOR = '#6d4c41';

/**
 * 大項目の見出し：太い縦棒＋太字＋色で、小項目（グレーの小見出し）より強く見せる。
 * カードの区切り線は太さを変えられないため、見出し側で強弱をつける。
 */
function sectionHeading_(text) {
  return CardService.newTextParagraph()
    .setText('<font color="' + HEADING_COLOR + '"><b>▌' + text + '</b></font>');
}

/** ボタン列の上に付ける小さな見出し */
function rowLabel_(text) {
  return CardService.newTextParagraph().setText('<font color="#5f6368">' + text + '</font>');
}

/** 一番下の折りたたみ欄：画面の動きの切り替え＋ヘルプ・注意事項 */
function buildHelpSection_(s, rendered) {
  const classic = alwaysPreview_(s);
  const uiNote = font_(PREVIEW.COLOR_NOTE,
    '🪶 軽快：プレビューはボタンで入り切り。非表示中は更新が少なく軽い\n' +
    '👀 常時プレビュー：設定を変えるたびに図を描き直す');
  const usage =
    '<b>使い方</b>\n' +
    '① 段落を選択（カーソルを置くだけなら、その段落が対象）\n' +
    (classic ? '② 操作・対象などをボタンで選ぶ\n' : '② 操作・対象などをボタンやラジオボタンで選ぶ\n') +
    '③ 一番上の実行ボタンを押す\n' +
    '・文字数は 0.5〜5字から選びます（位置指定は 0字も可）。\n' +
    '・1字＝段落の先頭の文字の大きさです。\n' +
    '・字下げ・ぶら下げの量は1字です。\n' +
    (classic ? '' : '・👀 プレビューはボタンで表示／非表示を切り替えます（非表示のほうが操作が軽くなります）。\n') +
    '・設定内容は次回開いたときも引き継がれます。';
  const notes =
    '<b>注意</b>\n' +
    '・Wordファイル(.docx)は［ファイル］→［Googleドキュメントとして保存］で変換してからお使いください。\n' +
    '・メイリオ等のプロポーショナルフォントでは、文字数ぴったりにならない場合があります。\n' +
    '・提案モードには対応していません（スクリプトの編集は常に直接編集になります）。提案は承認／拒否してからお使いください。';
  return CardService.newCardSection()
    .setHeader('⚙️ 画面の動き・ヘルプ・注意事項')
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(0)
    .addWidget(rowLabel_('⚙️ 画面の動き'))
    .addWidget(choiceRow_('ui', s.ui, rendered))
    .addWidget(CardService.newTextParagraph().setText(uiNote))
    .addWidget(CardService.newDivider())
    .addWidget(CardService.newTextParagraph().setText(usage))
    .addWidget(CardService.newTextParagraph().setText(notes));
}

function positionTitle_(s) {
  return { first: '🔢 1行目（左端から）', following: '🔢 2行目〜（左端から）', right: '🔢 右端から' }[s.set_target]
    || '🔢 左端から';
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
  MIN_TEXT: 6,                              // 本文（■）をこれ以上細くしないよう、余白の途中を省く（文字）
  EXAMPLE: { first: 2, start: 2, end: 0 },  // 「今の段落」の例
  SAMPLE_TEXT: '　①　ああ　いい ',          // 空白処理の例文
  COLOR_TEXT: '#8d6e63',
  COLOR_SPACE: '#bdbdbd',
  COLOR_NOTE: '#5f6368',
  COLOR_ERROR: '#d93025',
};

/**
 * プレビュー（🪶 軽快 ではボタンで表示／非表示、👀 常時プレビュー では常に表示）。
 * 表示中だけ選択中の段落を読んで図を描く（非表示なら文書を読まないので軽い）。
 */
function buildDisplaySection_(s, rendered) {
  const section = CardService.newCardSection().addWidget(sectionHeading_('👀 プレビュー'));
  if (alwaysPreview_(s)) return addPreviewFigure_(section, s); // 👀 常時プレビュー：入り切りボタンなし
  const live = s.preview === 'on';
  section.addWidget(CardService.newButtonSet().addButton(CardService.newTextButton()
    .setText(live ? '🙈 隠す' : '👀 表示')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('togglePreview')
      .setParameters({ rendered: rendered }))));
  if (!live) return section;
  return addPreviewFigure_(section, s);
}

/** 選択中の段落を読んで図を描き、欄に足す */
function addPreviewFigure_(section, s) {
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
  const nowLabel = current
    ? (current.multiple ? '選択中の先頭の段落（今）' : '選択中の段落（今）')
    : '例の段落（選択中の段落を読めないため）';
  // 今の段落の段落間隔と1字の大きさ（以前の「現在値」通知の内容）
  const pt = v => (v === null || v === undefined) ? '既定' : formatNumber_(v) + 'pt';
  const nowInfo = current
    ? '\n' + font_(PREVIEW.COLOR_NOTE, '段落前 ' + pt(current.before) + '・段落後 ' + pt(current.after) +
        '（1字＝' + formatNumber_(current.em) + 'pt）')
    : '';
  const blocks = [];
  let fractional = false;
  // 図は全部の段落がそろってから描く（省略〜の幅を図どうしでそろえ、動きの差が見えるように）
  const states = [];
  let cut = null;
  const fig = (title, st) => {
    if ([st.first, st.start, st.end].some(v => v % 1 !== 0)) fractional = true;
    states.push(st);
    return () => '<b>' + title + '</b>\n' + drawParagraph_(st, cutFor_(st, cut)) + '\n' +
      font_(PREVIEW.COLOR_NOTE, describe_(st));
  };
  const withInfo = (draw, info) => () => draw() + info;

  switch (plan.mode) {
    case 'move': {
      const isRight = plan.target === 'right';
      blocks.push(withInfo(fig(nowLabel, ex), nowInfo));
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
      blocks.push(withInfo(fig(nowLabel, ex), nowInfo)); // 現在の位置も並べて、変化が分かるようにする
      blocks.push(fig('設定後', after));
      break;
    }
    case 'reset':
      blocks.push(withInfo(fig(nowLabel, ex), nowInfo));
      blocks.push(fig('解除後', { first: 0, start: 0, end: 0 }));
      break;
    default:
      if (current) blocks.push(withInfo(fig(nowLabel, ex), nowInfo));
      blocks.push(font_(PREVIEW.COLOR_NOTE, 'インデントは変わりません。'));
  }
  cut = previewCut_(states);
  for (let i = 0; i < blocks.length; i++) {
    if (typeof blocks[i] === 'function') blocks[i] = blocks[i]();
  }
  const notes = [];
  if (states.some(st => { const k = cutFor_(st, cut); return k.left || k.right; })) {
    notes.push('〜は途中を省略しています');
  }
  if (fractional) notes.push('0.5字は四捨五入しています');
  if (notes.length) blocks.push(font_(PREVIEW.COLOR_NOTE, '※図では' + notes.join('。') + '。'));

  if (plan.spaceChars !== 'none') {
    let str = PREVIEW.SAMPLE_TEXT;
    const fake = {
      getText: () => str,
      deleteText: (a, b) => { str = str.slice(0, a) + str.slice(b + 1); },
    };
    handleSpaceChars_({ editAsText: () => fake }, plan.spaceChars);
    blocks.push('<b>空白の削除</b>\n' + showSpaces_(PREVIEW.SAMPLE_TEXT) + '\n→ ' + showSpaces_(str));
  }

  if (plan.spacing) {
    const v = x => (x === null ? '変更なし' : formatNumber_(x) + 'pt');
    blocks.push('<b>段落の前後の間隔</b>\n段落前 ' + v(plan.spacing.before) + '・段落後 ' + v(plan.spacing.after));
  }
  return blocks.join('\n\n');
}

/**
 * インデントが大きくて本文（■）が細くなりすぎるときに、余白の途中を省く量（字）。
 * 並べて描く図すべてで同じ量を省く（図どうしの差＝動きがそのまま見えるように）。
 * 1行目と2行目以降に共通の余白だけを省くので、2行の位置関係は崩れない。
 * 省いたところは「・〜」で示し、省いても余白は2字以上残す。
 */
function previewCut_(states) {
  const r = v => Math.max(0, Math.round(v));
  let need = 0;
  let maxLeftCut = Infinity;
  let maxRightCut = Infinity;
  states.forEach(st => {
    need = Math.max(need, Math.max(r(st.first), r(st.start)) + r(st.end) + PREVIEW.MIN_TEXT - PREVIEW.WIDTH);
    maxLeftCut = Math.min(maxLeftCut, Math.min(r(st.first), r(st.start)) - 2);
    maxRightCut = Math.min(maxRightCut, r(st.end) - 2);
  });
  const left = Math.max(0, Math.min(need, maxLeftCut));      // まず左の余白を省き
  const right = Math.max(0, Math.min(need - left, maxRightCut)); // 足りなければ右の余白も省く
  return { left: left, right: right };
}

/**
 * その図で使う省略量：ふだんは図どうしでそろえた量。
 * それでは本文が細くなりすぎる図（例：解除後が0字で、そろえて省けない「今」）だけ、その図単独で省く。
 */
function cutFor_(st, shared) {
  const r = v => Math.max(0, Math.round(v));
  const text = PREVIEW.WIDTH - (Math.max(r(st.first), r(st.start)) - shared.left) - (r(st.end) - shared.right);
  return text >= PREVIEW.MIN_TEXT / 2 ? shared : previewCut_([st]);
}

/** 段落の図：1行目と2行目以降の2行。・＝インデント、■＝本文、〜＝省略 */
function drawParagraph_(st, cut) {
  const W = PREVIEW.WIDTH;
  const cutL = (cut && cut.left) || 0;
  const cutR = (cut && cut.right) || 0;
  const dots = n => font_(PREVIEW.COLOR_SPACE, '・'.repeat(Math.max(0, n)));
  const mark = font_(PREVIEW.COLOR_NOTE, '〜');
  // 右の余白（省くときは「・・〜・」のように、端の1字の手前に〜）
  const right = Math.min(Math.max(0, Math.round(st.end)) - cutR, W - 1);
  const rightHtml = cutR ? dots(right - 2) + mark + dots(1) : dots(right);
  const line = left => {
    // 左の余白（省くときは「・〜・・」のように、端の1字の後ろに〜）
    const l = Math.max(0, Math.min(Math.max(0, Math.round(left)) - cutL, W - right - 1));
    const leftHtml = cutL ? dots(1) + mark + dots(l - 2) : dots(l);
    return '｜' + leftHtml + font_(PREVIEW.COLOR_TEXT, '■'.repeat(W - l - right)) + rightHtml + '｜';
  };
  return line(st.first) + '\n' + line(st.start);
}

function describe_(st) {
  return '1行目 ' + formatNumber_(st.first) + '字・2行目以降 ' + formatNumber_(st.start) + '字' +
    (st.end ? '・右端 ' + formatNumber_(st.end) + '字' : '');
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
        const label = s.shape === 'indent' ? '字下げ' : 'ぶら下げ';
        const m = SHAPE_CHARS;
        const diff = s.shape === 'indent' ? m : -m; // 1行目 − 2行目以降
        if (s.shape_base === 'first') { first = n; start = n - diff; }
        else                          { start = n; first = n + diff; }
        if (first < 0 || start < 0) {
          const line = first < 0 ? '1行目' : '2行目以降';
          throw userError_('左端から' + formatNumber_(n) + '字の位置では、' + label + '（' + formatNumber_(m) +
            '字）で' + line + 'が左端より左に出てしまいます。左端から' + formatNumber_(m) +
            '字以上にするか、「指定位置に置く行」を変えてください。');
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
      const what = { all: '', first: 'の1行目', following: 'の2行目以降', right: 'の右端' }[plan.target] || '';
      msg = count + '段落' + what + 'を' + dir + 'へ' + formatNumber_(plan.chars) + '字ずらしました。';
      break;
    }
    case 'set':
      msg = count + '段落のインデントを設定しました' + (plan.target === 'all'
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
