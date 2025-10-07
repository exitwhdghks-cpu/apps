// Code.gs - Backend logic for Hapoom RPG Spreadsheet Simulator.
// Handles sheet setup, player lifecycle, event resolution, RNG, and menu bindings.

// ----------------------------- CONSTANTS ----------------------------------
const RPG_CONFIG = {
  menuName: '\uD83C\uDFAE Hapoom RPG',
  menuStartLabel: '게임 시작/계속',
  userPropertyKey: 'HAPOON_RPG_PLAYER_ID',
  adminEmails: ['admin@example.com'], // TODO: Replace with real admin emails.
  xpBase: 12,
  xpGrowth: 6,
  levelHpBonus: 4,
  levelAtkBonus: 1,
  levelDefBonus: 1,
  healOnLevelUpRatio: 0.4,
  queueMax: 5
};

const SHEET_NAMES = {
  players: 'Players',
  events: 'Events',
  log: 'Log',
  assets: 'Assets'
};

const PLAYER_HEADERS = ['playerId', 'name', 'cls', 'hp', 'maxHp', 'gold', 'atk', 'def', 'level', 'xp', 'seed', 'createdAt', 'lastPlayedAt', 'state'];
const EVENT_HEADERS = ['eventId', 'title', 'desc', 'choiceA_text', 'choiceA_effect', 'choiceB_text', 'choiceB_effect', 'choiceC_text', 'choiceC_effect', 'tags', 'weight'];
const LOG_HEADERS = ['timestamp', 'playerId', 'eventId', 'choice', 'delta', 'resultText'];
const ASSET_HEADERS = ['key', 'type', 'value'];

const CLASS_PRESETS = [
  { key: 'wanderer', name: '방랑자', hp: 12, atk: 3, def: 1, desc: '균형 잡힌 모험가' },
  { key: 'sentinel', name: '수호자', hp: 16, atk: 2, def: 3, desc: '튼튼한 방패의 달인' },
  { key: 'mystic', name: '비전술사', hp: 10, atk: 4, def: 1, desc: '마력을 다루는 학자' }
];

// ------------------------------- MENU --------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(RPG_CONFIG.menuName)
    .addItem(RPG_CONFIG.menuStartLabel, 'showGameSidebar')
    .addToUi();
}

function showGameSidebar() {
  const html = HtmlService.createTemplateFromFile('ui');
  html.data = {
    version: new Date().getTime()
  };
  const output = html.evaluate();
  output.setTitle('Hapoom RPG');
  output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(output);
}

// ---------------------------- SHEET SETUP ----------------------------------
function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  const playersSheet = getOrCreateSheet(ss, SHEET_NAMES.players, PLAYER_HEADERS);
  const eventsSheet = getOrCreateSheet(ss, SHEET_NAMES.events, EVENT_HEADERS);
  const logSheet = getOrCreateSheet(ss, SHEET_NAMES.log, LOG_HEADERS);
  const assetsSheet = getOrCreateSheet(ss, SHEET_NAMES.assets, ASSET_HEADERS);

  // Clear existing data (preserve headers)
  clearSheetExceptHeader(playersSheet);
  clearSheetExceptHeader(eventsSheet);
  clearSheetExceptHeader(logSheet);
  clearSheetExceptHeader(assetsSheet);

  // Insert sample assets including palette, sprites, and localization seeds.
  const assets = [
    ['pal_background', 'color', '#0d0f1a'],
    ['pal_accent', 'color', '#7DF9FF'],
    ['pal_positive', 'color', '#A7F3D0'],
    ['pal_negative', 'color', '#FCA5A5'],
    ['sprite_hero', 'css', '<div class="px-sprite hero"></div>'],
    ['sprite_slime', 'css', '<div class="px-sprite slime"></div>'],
    ['sprite_treasure', 'css', '<div class="px-sprite treasure"></div>'],
    ['sprite_mystic', 'css', '<div class="px-sprite mystic"></div>'],
    ['text_start_button', 'string', '모험 시작'],
    ['text_continue_button', 'string', '계속'],
    ['text_hp', 'string', 'HP'],
    ['text_gold', 'string', 'Gold'],
    ['text_xp', 'string', 'XP'],
    ['text_level', 'string', 'Lv.'],
    ['i18n_note', 'note', '문자열 키를 Assets 시트에서 언어별로 확장 가능 (예: text_start_button_ko/en)']
  ];
  assetsSheet.getRange(2, 1, assets.length, ASSET_HEADERS.length).setValues(assets);

  // Insert sample events (10+ entries) with varied tags and weights.
  const sampleEvents = buildSampleEvents();
  eventsSheet.getRange(2, 1, sampleEvents.length, EVENT_HEADERS.length).setValues(sampleEvents);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  sheet.clear();
  sheet.appendRow(headers);
  return sheet;
}

function clearSheetExceptHeader(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

// Builds array rows for the Events sheet. Include Events expansion guide in comments.
function buildSampleEvents() {
  const events = [
    {
      eventId: 'E_START_GLADE',
      title: '새벽빛 수풀',
      desc: '차가운 새벽 공기 속에서 반짝이는 이슬. 먼 곳에서 작은 울음소리가 들려옵니다.',
      choiceA: { text: '울음소리를 조사한다', effect: { rolls: [ { chance: 0.6, hp: -2, xp: 2, resultText: '숨어 있던 슬라임이 튀어나와 약간의 피해를 입었습니다.' }, { chance: 0.4, gold: 4, resultText: '잃어버린 지갑을 찾았습니다.' } ], flagAdd: ['met_slime'], tagsBoost: ['combat'] } },
      choiceB: { text: '조용히 명상한다', effect: { hp: 2, xp: 1, resultText: '자연과 호흡하며 체력을 회복했습니다.', next: ['E_SHRINE_WHISPER'] } },
      choiceC: { text: '길을 급히 떠난다', effect: { xp: 1, resultText: '조심스럽게 발걸음을 옮겼습니다.', next: ['E_ROAD_MERCHANT'], weightShift: { town: 1.2 } } },
      tags: 'opening,forest',
      weight: 5
    },
    {
      eventId: 'E_MON_SLIME',
      title: '슬라임 습격',
      desc: '포들포들한 슬라임이 길을 막았습니다.',
      choiceA: { text: '찌르기 공격', effect: { hp: -1, xp: 3, resultText: '슬라임을 갈랐습니다! 끈적한 흔적만 남네요.', gold: 3, tagsBoost: ['combat'] } },
      choiceB: { text: '분석 후 약점 노리기', effect: { rolls: [ { chance: 0.7, xp: 4, gold: 2, resultText: '약점을 찾아내 큰 피해를 주었습니다.' }, { chance: 0.3, hp: -3, resultText: '분석하다가 슬라임이 덮쳤습니다!' } ] } },
      choiceC: { text: '후퇴한다', effect: { hp: -1, resultText: '후퇴하는 동안 슬라임이 튀어 올랐습니다.', flagRemove: ['met_slime'] } },
      tags: 'combat,common',
      weight: 8
    },
    {
      eventId: 'E_TOWN_SHOP',
      title: '작은 노점상',
      desc: '낡은 수레에서 팔을 흔드는 상인.',
      choiceA: { text: '치유 허브 구매(-4G)', effect: { gold: -4, hp: 4, resultText: '허브 향이 기분 좋게 퍼집니다.', flagAdd: ['met_merchant'] } },
      choiceB: { text: '정보 교환', effect: { gold: -2, xp: 2, resultText: '상인에게서 귀중한 정보를 들었습니다.', next: ['E_RARE_RELIC'] } },
      choiceC: { text: '흥정을 시도한다', effect: { rolls: [ { chance: 0.5, gold: 3, resultText: '흥정에 성공해 보너스를 받았습니다!' }, { chance: 0.5, gold: -3, resultText: '흥정 실패! 상인이 화를 냈습니다.' } ], flagAdd: ['met_merchant'] } },
      tags: 'town,merchant',
      weight: 6
    },
    {
      eventId: 'E_RARE_RELIC',
      title: '빛나는 유물',
      desc: '모래 속에서 은은한 빛이 새어 나옵니다.',
      choiceA: { text: '유물을 집어 든다', effect: { rolls: [ { chance: 0.5, xp: 5, gold: 6, resultText: '고대 유물이 당신을 인정합니다.' }, { chance: 0.5, hp: -4, resultText: '저주가 발동하여 몸이 얼어붙습니다.' } ], tagsBoost: ['rare'] } },
      choiceB: { text: '조심스럽게 봉인', effect: { xp: 3, def: 1, resultText: '봉인을 재정비하여 방어력을 높였습니다.' } },
      choiceC: { text: '무시하고 간다', effect: { xp: 1, resultText: '무사히 지나쳤습니다.', flagRemove: ['met_merchant'] } },
      tags: 'rare,explore',
      weight: 3
    },
    {
      eventId: 'E_SHRINE_WHISPER',
      title: '속삭이는 사당',
      desc: '돌기둥 사이로 희미한 빛과 속삭임이 들립니다.',
      choiceA: { text: '기도한다', effect: { xp: 2, hp: 3, resultText: '따스한 빛이 몸을 감싸 안았습니다.' } },
      choiceB: { text: '공물을 바친다(-3G)', effect: { gold: -3, xp: 4, resultText: '사당이 기뻐하며 힘을 나누어 줍니다.', next: ['E_BLESSING_LIGHT'] } },
      choiceC: { text: '조사한다', effect: { rolls: [ { chance: 0.4, xp: 5, atk: 1, resultText: '숨겨진 룬을 해독했습니다.' }, { chance: 0.6, hp: -2, resultText: '봉인을 건드려 마력이 새어나왔습니다.' } ] } },
      tags: 'mystic,rare',
      weight: 4
    },
    {
      eventId: 'E_ROAD_MERCHANT',
      title: '길 잃은 상단',
      desc: '수레가 넘어져 짐이 흩어져 있습니다.',
      choiceA: { text: '짐을 도와준다', effect: { xp: 2, gold: 2, resultText: '상인이 보답으로 동전을 줍니다.', flagAdd: ['met_merchant'] } },
      choiceB: { text: '짐을 슬쩍 챙긴다', effect: { rolls: [ { chance: 0.5, gold: 5, resultText: '아무도 알아차리지 못했습니다.' }, { chance: 0.5, hp: -3, resultText: '상단의 경비에게 붙잡혔습니다!' } ] } },
      choiceC: { text: '길 안내만 한다', effect: { xp: 1, resultText: '지도를 설명해 주었습니다.', next: ['E_TOWN_SHOP'] } },
      tags: 'town,common',
      weight: 5
    },
    {
      eventId: 'E_ANCIENT_TRAP',
      title: '고대 함정',
      desc: '발밑에서 바람이 새어나오는 함정 구역.',
      choiceA: { text: '기어가며 통과', effect: { hp: -1, xp: 2, resultText: '몇 번 긁혔지만 지나갔습니다.' } },
      choiceB: { text: '함정 해체 시도', effect: { rolls: [ { chance: 0.4, gold: 4, xp: 3, resultText: '함정을 해제하고 보물을 챙겼습니다.' }, { chance: 0.6, hp: -4, resultText: '폭발! 큰 피해를 입었습니다.' } ] } },
      choiceC: { text: '되돌아간다', effect: { xp: 1, resultText: '안전을 우선시했습니다.', flagRemove: ['trap_warned'] } },
      tags: 'trap,explore',
      weight: 4
    },
    {
      eventId: 'E_FOREST_SPIRIT',
      title: '숲의 정령',
      desc: '작은 정령이 반짝이며 주변을 맴돕니다.',
      choiceA: { text: '손을 내민다', effect: { rolls: [ { chance: 0.5, hp: 5, resultText: '정령이 치유의 힘을 나눠줍니다.' }, { chance: 0.5, hp: -3, resultText: '정령이 장난을 쳐 체력이 빠졌습니다.' } ] } },
      choiceB: { text: '정령과 거래', effect: { gold: -2, xp: 3, resultText: '정령이 소원을 들어주었습니다.', flagAdd: ['spirit_friend'] } },
      choiceC: { text: '정령을 포획', effect: { hp: -2, atk: 1, resultText: '정령의 힘을 장비에 봉인했습니다.', flagRemove: ['spirit_friend'] } },
      tags: 'mystic,rare',
      weight: 3
    },
    {
      eventId: 'E_BLESSING_LIGHT',
      title: '빛의 가호',
      desc: '사당에서 받아온 부적이 따스하게 빛납니다.',
      choiceA: { text: '부적을 사용한다', effect: { hp: 6, xp: 2, resultText: '심장이 강하게 뛰며 힘이 솟습니다.', flagRemove: ['met_merchant'] } },
      choiceB: { text: '부적을 팔아버린다', effect: { gold: 6, resultText: '상인에게 고가에 팔았습니다.' } },
      choiceC: { text: '선물을 나눈다', effect: { gold: -2, xp: 3, resultText: '빛을 나누며 명성을 얻었습니다.', flagAdd: ['blessed'] } },
      tags: 'rare,town',
      weight: 2
    },
    {
      eventId: 'E_NIGHT_AMBUSH',
      title: '밤의 습격',
      desc: '어둠 속에서 그림자가 달려듭니다.',
      choiceA: { text: '즉각 반격', effect: { rolls: [ { chance: 0.6, xp: 4, gold: 3, resultText: '습격자를 제압했습니다.' }, { chance: 0.4, hp: -5, resultText: '기습을 막지 못했습니다.' } ], tagsBoost: ['combat'] } },
      choiceB: { text: '방패를 들고 버틴다', effect: { hp: -2, def: 1, resultText: '피해는 있었지만 방어를 익혔습니다.' } },
      choiceC: { text: '연막탄으로 탈출', effect: { gold: -1, xp: 2, resultText: '연막을 치고 빠져나왔습니다.' } },
      tags: 'combat,trap',
      weight: 5
    },
    {
      eventId: 'E_LOST_CHILD',
      title: '길 잃은 아이',
      desc: '울고 있는 아이가 길모퉁이에 앉아 있습니다.',
      choiceA: { text: '집으로 데려다준다', effect: { xp: 3, gold: 2, resultText: '가족이 감사의 선물을 주었습니다.', flagAdd: ['town_ally'] } },
      choiceB: { text: '용기를 북돋운다', effect: { xp: 2, resultText: '아이에게 용기를 심어주었습니다.' } },
      choiceC: { text: '무시한다', effect: { resultText: '뒤돌아보지 않았습니다.', flagRemove: ['town_ally'] } },
      tags: 'town,story',
      weight: 4
    }
  ];

  // Events 확장 가이드: 새로운 이벤트는 EVENT_HEADERS 순서를 지키고, effect JSON에 hp/gold/xp/atk/def/next/flagAdd/flagRemove/rolls/resultText 등을 조합해 서사를 구성합니다.
  return events.map(function (evt) {
    return [
      evt.eventId,
      evt.title,
      evt.desc,
      evt.choiceA.text,
      JSON.stringify(evt.choiceA.effect),
      evt.choiceB.text,
      JSON.stringify(evt.choiceB.effect),
      evt.choiceC.text,
      JSON.stringify(evt.choiceC.effect),
      evt.tags,
      evt.weight
    ];
  });
}

// ---------------------------- GAME SERVICES --------------------------------
function loadGame() {
  const context = buildContext();
  const playerRecord = getCurrentPlayerRecord(context);
  if (!playerRecord) {
    return {
      needsCreation: true,
      classes: CLASS_PRESETS,
      assets: context.assets,
      version: context.version,
      debug: context.debug
    };
  }

  const player = playerRecord.player;
  let state = playerRecord.state;
  if (!state.queue) state.queue = [];
  if (!state.flags) state.flags = [];
  if (!state.currentEventId) {
    state.currentEventId = drawNextEvent(player, state, context);
  }

  const currentEvent = getEventById(context.events, state.currentEventId);
  if (!currentEvent) {
    state.currentEventId = drawNextEvent(player, state, context);
  }

  // Persist state if mutated during load.
  savePlayerState(player, state, context);

  return {
    needsCreation: false,
    player: sanitizePlayerForClient(player),
    state: filterStateForClient(state),
    event: currentEvent,
    assets: context.assets,
    queueSize: state.queue.length,
    debug: context.debug,
    logs: getRecentLogs(player.playerId, 10)
  };
}

function startNewPlayer(payload) {
  const context = buildContext();
  const name = (payload && payload.name) ? payload.name.trim() : '';
  const clsKey = payload && payload.clsKey;
  if (!name) throw new Error('이름을 입력해주세요.');
  const preset = CLASS_PRESETS.find(function (c) { return c.key === clsKey; });
  if (!preset) throw new Error('잘못된 직업입니다.');

  const ss = SpreadsheetApp.getActive();
  const playersSheet = ss.getSheetByName(SHEET_NAMES.players);

  const playerId = 'P_' + Utilities.getUuid();
  const now = new Date();
  const seed = Utilities.getUuid();
  const initialState = {
    queue: [],
    flags: [],
    currentEventId: null,
    lastChoice: null,
    isDead: false
  };
  const row = [
    playerId,
    name,
    preset.name,
    preset.hp,
    preset.hp,
    6,
    preset.atk,
    preset.def,
    1,
    0,
    seed,
    now,
    now,
    JSON.stringify(initialState)
  ];
  playersSheet.appendRow(row);

  PropertiesService.getUserProperties().setProperty(RPG_CONFIG.userPropertyKey, playerId);

  return loadGame();
}

function chooseOption(choicePayload) {
  const context = buildContext();
  const playerRecord = getCurrentPlayerRecord(context);
  if (!playerRecord) {
    throw new Error('플레이어가 없습니다. 새로 생성해주세요.');
  }
  const player = playerRecord.player;
  const state = playerRecord.state;
  if (state.isDead) {
    throw new Error('이미 패배한 세션입니다. 새로 시작해주세요.');
  }

  const choiceKey = (choicePayload && choicePayload.choice) || 'A';
  const eventId = state.currentEventId;
  const event = getEventById(context.events, eventId);
  if (!event) {
    state.currentEventId = drawNextEvent(player, state, context);
    savePlayerState(player, state, context);
    throw new Error('이벤트를 찾을 수 없어 다음 이벤트로 이동합니다.');
  }

  const effect = getEffectForChoice(event, choiceKey);
  if (!effect) {
    throw new Error('잘못된 선택입니다.');
  }

  const resolution = resolveEffect(effect, player, state, context, choiceKey, event);
  applyDeltaToPlayer(player, resolution.delta);
  const levelUpInfo = applyLevelUpIfNeeded(player, resolution);
  applyFlags(state, resolution.flags);
  enqueueNextEvents(state, resolution.nextQueue);

  const timestamp = new Date();
  const death = player.hp <= 0;
  if (death) {
    state.isDead = true;
    resolution.messages.push('당신은 쓰러졌습니다.');
    player.hp = 0;
  }

  state.lastChoice = { eventId: event.eventId, choice: choiceKey, at: timestamp };
  if (!death) {
    state.currentEventId = drawNextEvent(player, state, context);
  } else {
    state.currentEventId = null;
  }

  player.lastPlayedAt = timestamp;
  savePlayerState(player, state, context);
  appendLogEntry(timestamp, player.playerId, event.eventId, choiceKey, resolution.delta, resolution.messages.join(' '));

  return {
    player: sanitizePlayerForClient(player),
    state: filterStateForClient(state),
    event: death ? null : getEventById(context.events, state.currentEventId),
    result: {
      messages: resolution.messages,
      delta: resolution.delta,
      levelUps: levelUpInfo.levelUps,
      healedOnLevel: levelUpInfo.healed,
      death: death
    },
    queueSize: state.queue.length,
    logs: getRecentLogs(player.playerId, 10)
  };
}

function resetPlayerSession() {
  const userProps = PropertiesService.getUserProperties();
  userProps.deleteProperty(RPG_CONFIG.userPropertyKey);
  return loadGame();
}

// --------------------------- DEBUG UTILITIES --------------------------------
function toggleDebugAction(action, value) {
  const context = buildContext();
  if (!context.debug.isAdmin) {
    throw new Error('권한이 없습니다.');
  }
  const record = getCurrentPlayerRecord(context);
  if (!record) {
    throw new Error('플레이어가 없습니다.');
  }
  const player = record.player;
  const state = record.state;
  if (action === 'forceEvent') {
    enqueueNextEvents(state, [{ eventId: value, priority: true }]);
    if (!state.currentEventId) {
      state.currentEventId = drawNextEvent(player, state, context);
    }
  } else if (action === 'heal') {
    const healAmount = Math.ceil(player.maxHp * 0.5);
    player.hp = Math.min(player.maxHp, player.hp + healAmount);
  }
  savePlayerState(player, state, context);
  return {
    player: sanitizePlayerForClient(player),
    state: filterStateForClient(state),
    event: state.currentEventId ? getEventById(context.events, state.currentEventId) : null,
    queueSize: state.queue.length
  };
}

// -------------------------- INTERNAL HELPERS --------------------------------
function buildContext() {
  const ss = SpreadsheetApp.getActive();
  const assets = loadAssets(ss.getSheetByName(SHEET_NAMES.assets));
  const events = loadEvents(ss.getSheetByName(SHEET_NAMES.events));
  const userEmail = Session.getActiveUser().getEmail();
  const isAdmin = !!userEmail && RPG_CONFIG.adminEmails.indexOf(userEmail) !== -1;
  return {
    ss: ss,
    assets: assets,
    events: events,
    userEmail: userEmail,
    debug: { isAdmin: isAdmin },
    version: new Date().getTime()
  };
}

function getCurrentPlayerRecord(context) {
  const playerId = PropertiesService.getUserProperties().getProperty(RPG_CONFIG.userPropertyKey);
  if (!playerId) return null;
  const playersSheet = context.ss.getSheetByName(SHEET_NAMES.players);
  const data = playersSheet.getDataRange().getValues();
  const headers = data.shift();
  const idx = headers.reduce(function (acc, header, i) { acc[header] = i; return acc; }, {});
  for (var i = 0; i < data.length; i++) {
    if (data[i][idx.playerId] === playerId) {
      const row = data[i];
      const player = {
        rowIndex: i + 2,
        playerId: row[idx.playerId],
        name: row[idx.name],
        cls: row[idx.cls],
        hp: Number(row[idx.hp]) || 0,
        maxHp: Number(row[idx.maxHp]) || 0,
        gold: Number(row[idx.gold]) || 0,
        atk: Number(row[idx.atk]) || 0,
        def: Number(row[idx.def]) || 0,
        level: Number(row[idx.level]) || 1,
        xp: Number(row[idx.xp]) || 0,
        seed: row[idx.seed] || Utilities.getUuid(),
        createdAt: row[idx.createdAt],
        lastPlayedAt: row[idx.lastPlayedAt]
      };
      let state = {};
      try {
        state = JSON.parse(row[idx.state] || '{}');
      } catch (e) {
        state = {};
      }
      return { player: player, state: state };
    }
  }
  return null;
}

function sanitizePlayerForClient(player) {
  return {
    playerId: player.playerId,
    name: player.name,
    cls: player.cls,
    hp: player.hp,
    maxHp: player.maxHp,
    gold: player.gold,
    atk: player.atk,
    def: player.def,
    level: player.level,
    xp: player.xp
  };
}

function filterStateForClient(state) {
  return {
    flags: state.flags || [],
    queueCount: (state.queue || []).length,
    lastChoice: state.lastChoice,
    isDead: !!state.isDead
  };
}

function loadAssets(sheet) {
  if (!sheet) return {};
  const values = sheet.getDataRange().getValues();
  values.shift();
  const map = {};
  values.forEach(function (row) {
    if (row[0]) map[row[0]] = { type: row[1], value: row[2] };
  });
  return map;
}

function loadEvents(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const idx = headers.reduce(function (acc, header, i) { acc[header] = i; return acc; }, {});
  return values.filter(function (row) { return row[idx.eventId]; }).map(function (row) {
    return {
      eventId: row[idx.eventId],
      title: row[idx.title],
      desc: row[idx.desc],
      tags: (row[idx.tags] || '').split(',').map(function (t) { return t.trim(); }).filter(Boolean),
      weight: Number(row[idx.weight]) || 1,
      choices: {
        A: { text: row[idx.choiceA_text], effect: parseJsonSafe(row[idx.choiceA_effect]) },
        B: { text: row[idx.choiceB_text], effect: parseJsonSafe(row[idx.choiceB_effect]) },
        C: { text: row[idx.choiceC_text], effect: parseJsonSafe(row[idx.choiceC_effect]) }
      }
    };
  });
}

function parseJsonSafe(value) {
  if (!value) return {};
  try {
    return JSON.parse(value);
  } catch (e) {
    return {};
  }
}

function getEventById(events, eventId) {
  if (!eventId) return null;
  for (var i = 0; i < events.length; i++) {
    if (events[i].eventId === eventId) return events[i];
  }
  return null;
}

function getEffectForChoice(event, choiceKey) {
  const upper = String(choiceKey || 'A').toUpperCase();
  return (event.choices[upper] && event.choices[upper].effect) || null;
}

function resolveEffect(effect, player, state, context, choiceKey, event) {
  const rollEffect = mergeEffectWithRoll(effect, player, state, context, choiceKey, event);
  const delta = {};
  const flags = { add: [], remove: [] };
  const nextQueue = [];
  const messages = [];

  const statKeys = ['hp', 'gold', 'xp', 'atk', 'def', 'maxHp'];
  statKeys.forEach(function (key) {
    if (typeof rollEffect[key] === 'number') {
      delta[key] = (delta[key] || 0) + rollEffect[key];
    }
  });

  if (Array.isArray(rollEffect.next)) {
    rollEffect.next.forEach(function (eventId) {
      nextQueue.push({ eventId: eventId, priority: true });
    });
  }
  if (Array.isArray(rollEffect.queue)) {
    rollEffect.queue.forEach(function (eventId) {
      nextQueue.push({ eventId: eventId, priority: false });
    });
  }
  if (Array.isArray(rollEffect.flagAdd)) {
    flags.add = flags.add.concat(rollEffect.flagAdd);
  }
  if (Array.isArray(rollEffect.flagRemove)) {
    flags.remove = flags.remove.concat(rollEffect.flagRemove);
  }
  if (rollEffect.weightShift) {
    state.weightShift = state.weightShift || {};
    Object.keys(rollEffect.weightShift).forEach(function (tag) {
      const current = Number(state.weightShift[tag]) || 1;
      state.weightShift[tag] = current * Number(rollEffect.weightShift[tag]);
    });
  }
  if (Array.isArray(rollEffect.tagsBoost)) {
    state.tagsBoost = state.tagsBoost || {};
    rollEffect.tagsBoost.forEach(function (tag) {
      const current = Number(state.tagsBoost[tag]) || 1;
      state.tagsBoost[tag] = current + 0.2;
    });
  }

  if (rollEffect.resultText) {
    messages.push(rollEffect.resultText);
  }

  return { delta: delta, flags: flags, nextQueue: nextQueue, messages: messages, effectMeta: rollEffect };
}

function mergeEffectWithRoll(effect, player, state, context, choiceKey, event) {
  const merged = Object.assign({}, effect);
  if (Array.isArray(effect.rolls) && effect.rolls.length) {
    const rng = pseudoRandom(player.seed + (state.lastChoice ? state.lastChoice.eventId : '') + choiceKey + event.eventId + new Date().getTime());
    let cumulative = 0;
    let selected = effect.rolls[effect.rolls.length - 1];
    effect.rolls.forEach(function (roll) {
      cumulative += Number(roll.chance) || 0;
      if (selected === effect.rolls[effect.rolls.length - 1] && rng <= cumulative) {
        selected = roll;
      }
    });
    Object.keys(selected).forEach(function (key) {
      if (key !== 'chance') {
        merged[key] = selected[key];
      }
    });
  }
  delete merged.rolls;
  return merged;
}

function pseudoRandom(seed) {
  const str = String(seed) + Utilities.getUuid();
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash |= 0;
  }
  return Math.abs(hash % 10000) / 10000;
}

function applyDeltaToPlayer(player, delta) {
  Object.keys(delta).forEach(function (key) {
    if (typeof delta[key] === 'number') {
      player[key] = (player[key] || 0) + delta[key];
    }
  });
  if (player.gold < 0) player.gold = 0;
  if (player.hp > player.maxHp) player.hp = player.maxHp;
}

function applyLevelUpIfNeeded(player, resolution) {
  let healed = 0;
  const levelUps = [];
  let xpNeeded = xpForNextLevel(player.level);
  while (player.xp >= xpNeeded) {
    player.xp -= xpNeeded;
    player.level += 1;
    player.maxHp += RPG_CONFIG.levelHpBonus;
    player.atk += RPG_CONFIG.levelAtkBonus;
    player.def += RPG_CONFIG.levelDefBonus;
    const healAmount = Math.ceil(player.maxHp * RPG_CONFIG.healOnLevelUpRatio);
    player.hp = Math.min(player.maxHp, player.hp + healAmount);
    healed += healAmount;
    levelUps.push({ level: player.level });
    resolution.messages.push('레벨 업! 능력이 향상되었습니다.');
    xpNeeded = xpForNextLevel(player.level);
  }
  return { healed: healed, levelUps: levelUps };
}

function xpForNextLevel(level) {
  return RPG_CONFIG.xpBase + (level - 1) * RPG_CONFIG.xpGrowth;
}

function applyFlags(state, flags) {
  state.flags = state.flags || [];
  if (Array.isArray(flags.add)) {
    flags.add.forEach(function (flag) {
      if (state.flags.indexOf(flag) === -1) state.flags.push(flag);
    });
  }
  if (Array.isArray(flags.remove)) {
    flags.remove.forEach(function (flag) {
      const idx = state.flags.indexOf(flag);
      if (idx !== -1) state.flags.splice(idx, 1);
    });
  }
}

function enqueueNextEvents(state, entries) {
  if (!Array.isArray(entries) || !entries.length) return;
  state.queue = state.queue || [];
  entries.forEach(function (entry) {
    if (entry && entry.eventId) {
      if (entry.priority) {
        state.queue.unshift(entry.eventId);
      } else {
        state.queue.push(entry.eventId);
      }
    }
  });
  if (state.queue.length > RPG_CONFIG.queueMax) {
    state.queue = state.queue.slice(0, RPG_CONFIG.queueMax);
  }
}

function drawNextEvent(player, state, context) {
  state.queue = state.queue || [];
  while (state.queue.length) {
    const candidate = state.queue.shift();
    const exists = getEventById(context.events, candidate);
    if (exists) {
      return exists.eventId;
    }
  }
  return pickWeightedEvent(player, state, context);
}

function pickWeightedEvent(player, state, context) {
  const events = context.events;
  const weightShift = state.weightShift || {};
  const tagsBoost = state.tagsBoost || {};
  let total = 0;
  const weighted = events.map(function (event) {
    let weight = event.weight || 1;
    event.tags.forEach(function (tag) {
      if (state.flags && state.flags.indexOf('met_merchant') !== -1 && tag === 'town') {
        weight *= 1.4;
      }
      if (tagsBoost[tag]) {
        weight *= tagsBoost[tag];
      }
      if (weightShift[tag]) {
        weight *= weightShift[tag];
      }
      if (tag === 'combat') {
        weight *= Math.min(3, 0.7 + player.level * 0.4);
      }
      if (tag === 'rare') {
        weight *= 0.8;
      }
    });
    if (state.lastChoice && state.lastChoice.eventId === event.eventId) {
      weight *= 0.35;
    }
    total += weight;
    return { event: event, weight: weight };
  });
  if (total <= 0) {
    return events[0] ? events[0].eventId : null;
  }
  const roll = pseudoRandom(player.seed + player.playerId + new Date().getMilliseconds()) * total;
  let acc = 0;
  for (var i = 0; i < weighted.length; i++) {
    acc += weighted[i].weight;
    if (roll <= acc) {
      return weighted[i].event.eventId;
    }
  }
  return weighted[weighted.length - 1].event.eventId;
}

function savePlayerState(player, state, context) {
  const playersSheet = context.ss.getSheetByName(SHEET_NAMES.players);
  if (!player.rowIndex) {
    // If row index missing, locate row and retry.
    const record = getCurrentPlayerRecord(context);
    if (!record) return;
    record.player.hp = player.hp;
    record.player.maxHp = player.maxHp;
    record.player.gold = player.gold;
    record.player.atk = player.atk;
    record.player.def = player.def;
    record.player.level = player.level;
    record.player.xp = player.xp;
    record.player.lastPlayedAt = player.lastPlayedAt;
    savePlayerState(record.player, state, context);
    return;
  }
  const row = player.rowIndex;
  const rowValues = [
    player.playerId,
    player.name,
    player.cls,
    player.hp,
    player.maxHp,
    player.gold,
    player.atk,
    player.def,
    player.level,
    player.xp,
    player.seed,
    player.createdAt,
    player.lastPlayedAt,
    JSON.stringify(state)
  ];
  playersSheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
}

function appendLogEntry(timestamp, playerId, eventId, choice, delta, resultText) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.log);
  if (!sheet) return;
  sheet.appendRow([
    timestamp,
    playerId,
    eventId,
    choice,
    JSON.stringify(delta || {}),
    resultText
  ]);
}

function getRecentLogs(playerId, limit) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.log);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  values.shift();
  const filtered = values.filter(function (row) { return row[1] === playerId; });
  const recent = filtered.slice(-limit);
  return recent.map(function (row) {
    return {
      timestamp: row[0],
      eventId: row[2],
      choice: row[3],
      delta: parseJsonSafe(row[4]),
      resultText: row[5]
    };
  });
}

// ------------------------------ NOTES ---------------------------------------
// i18n 확장 포인트: Assets 시트에 text_* 키를 언어별로 저장한 뒤, loadGame에서 사용자 환경에 맞는 문자열로 교체하도록 로직을 확장할 수 있습니다.
// 이벤트 추가 규칙: buildSampleEvents 주석 참고. 새로운 이벤트는 JSON 직렬화 규칙을 지키고, 확률 기반 결과는 rolls 배열로 설정합니다.
