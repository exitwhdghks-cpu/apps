<script>
/**
 * client.js - Front-end logic for Hapoom RPG Sidebar.
 * Handles UI binding, keyboard input (1/2/3), toast feedback, and server communication.
 */
(function () {
  const state = {
    loading: false,
    selectedClass: null,
    player: null,
    gameState: null,
    currentEvent: null,
    debug: { isAdmin: false }
  };

  const els = {
    hud: document.getElementById('hud'),
    hudPortrait: document.getElementById('hud-portrait'),
    hudName: document.getElementById('hud-name'),
    hudHp: document.getElementById('hud-hp'),
    hudGold: document.getElementById('hud-gold'),
    hudXp: document.getElementById('hud-xp'),
    hudLevel: document.getElementById('hud-level'),
    creation: document.getElementById('creation'),
    classGrid: document.getElementById('class-grid'),
    btnCreate: document.getElementById('btn-create'),
    inputName: document.getElementById('input-name'),
    gameplay: document.getElementById('gameplay'),
    choices: Array.from(document.querySelectorAll('.pixel-btn.choice')),
    eventIcon: document.getElementById('event-icon'),
    eventTitle: document.getElementById('event-title'),
    eventDesc: document.getElementById('event-desc'),
    queueCount: document.getElementById('queue-count'),
    logTail: document.getElementById('log-tail'),
    toast: document.getElementById('result-toast'),
    debugPanel: document.getElementById('debug-panel'),
    debugEvent: document.getElementById('debug-event'),
    btnForceEvent: document.getElementById('btn-force-event'),
    btnHeal: document.getElementById('btn-heal'),
    loading: document.getElementById('loading')
  };

  function setLoading(flag) {
    state.loading = flag;
    els.loading.hidden = !flag;
    els.choices.forEach(btn => btn.disabled = flag);
    if (els.btnCreate) els.btnCreate.disabled = flag;
  }

  function showToast(message) {
    els.toast.textContent = message;
    els.toast.hidden = false;
    setTimeout(() => {
      els.toast.hidden = true;
    }, 2000);
  }

  function renderCreation(payload) {
    els.creation.hidden = false;
    els.gameplay.hidden = true;
    els.hud.hidden = true;
    els.debugPanel.hidden = true;
    els.classGrid.innerHTML = '';
    state.selectedClass = null;
    (payload.classes || []).forEach(cls => {
      const card = document.createElement('div');
      card.className = 'class-card';
      card.dataset.key = cls.key;
      card.innerHTML = `
        <div class="class-name">${cls.name}</div>
        <div class="class-desc">${cls.desc}</div>
        <div class="class-stats">HP ${cls.hp} | ATK ${cls.atk} | DEF ${cls.def}</div>
      `;
      card.addEventListener('click', () => selectClass(cls.key, card));
      els.classGrid.appendChild(card);
    });
    els.btnCreate.onclick = () => {
      if (!state.selectedClass) {
        showToast('직업을 선택해주세요.');
        return;
      }
      const name = els.inputName.value.trim();
      if (!name) {
        showToast('이름을 입력해주세요.');
        els.inputName.focus();
        return;
      }
      setLoading(true);
      google.script.run
        .withSuccessHandler(data => {
          setLoading(false);
          bootstrapGame(data);
          showToast('새로운 모험이 시작되었습니다!');
        })
        .withFailureHandler(err => {
          setLoading(false);
          showToast(err && err.message ? err.message : '생성 실패');
        })
        .startNewPlayer({ name, clsKey: state.selectedClass });
    };
  }

  function selectClass(key, cardEl) {
    state.selectedClass = key;
    Array.from(els.classGrid.children).forEach(el => el.classList.remove('selected'));
    cardEl.classList.add('selected');
  }

  function renderHUD(player) {
    els.hud.hidden = false;
    els.hudName.textContent = `${player.name} · ${player.cls}`;
    els.hudHp.innerHTML = `<strong>HP</strong> ${player.hp}/${player.maxHp}`;
    els.hudGold.innerHTML = `<strong>G</strong> ${player.gold}`;
    els.hudXp.innerHTML = `<strong>XP</strong> ${player.xp}`;
    els.hudLevel.innerHTML = `<strong>Lv.</strong> ${player.level}`;
    els.hudPortrait.innerHTML = '<div class="px-sprite hero"></div>';
  }

  function renderEvent(event, queueSize) {
    if (!event) {
      els.eventTitle.textContent = '모험 종료';
      els.eventDesc.textContent = '새로운 여행을 시작하세요.';
      els.eventIcon.innerHTML = '<div class="px-sprite mystic"></div>';
      els.choices.forEach(btn => {
        btn.dataset.choice = 'RESET';
        btn.textContent = '새로운 모험 시작';
        btn.disabled = false;
      });
      return;
    }
    state.currentEvent = event;
    els.eventTitle.textContent = event.title;
    els.eventDesc.textContent = event.desc;
    els.eventIcon.innerHTML = getEventIcon(event.tags);
    const texts = {
      A: event.choices.A.text,
      B: event.choices.B.text,
      C: event.choices.C.text
    };
    ['A','B','C'].forEach((key, idx) => {
      const btn = els.choices[idx];
      btn.dataset.choice = key;
      btn.textContent = `${key}. ${texts[key]}`;
      btn.disabled = false;
    });
    els.queueCount.textContent = queueSize;
  }

  function renderLogs(logs) {
    if (!logs || !logs.length) {
      els.logTail.innerHTML = '<div class="log-row">최근 로그 없음</div>';
      return;
    }
    els.logTail.innerHTML = logs.map(entry => {
      const time = formatTime(entry.timestamp);
      const deltaText = formatDelta(entry.delta);
      return `<div class="log-row"><span class="log-time">${time}</span> · ${entry.resultText} ${deltaText}</div>`;
    }).join('');
  }

  function formatTime(ts) {
    if (!ts) return '';
    try {
      const date = new Date(ts);
      return `${date.getHours()}:${String(date.getMinutes()).padStart(2, '0')}`;
    } catch (e) {
      return '';
    }
  }

  function formatDelta(delta) {
    if (!delta) return '';
    const parts = [];
    Object.keys(delta).forEach(key => {
      const val = delta[key];
      if (typeof val === 'number' && val !== 0) {
        const sign = val > 0 ? '+' : '';
        parts.push(`${key.toUpperCase()} ${sign}${val}`);
      }
    });
    return parts.length ? `<span class="delta">(${parts.join(', ')})</span>` : '';
  }

  function getEventIcon(tags) {
    tags = tags || [];
    if (tags.includes('combat')) {
      return '<div class="px-sprite slime"></div>';
    }
    if (tags.includes('town')) {
      return '<div class="px-sprite treasure"></div>';
    }
    if (tags.includes('mystic')) {
      return '<div class="px-sprite mystic"></div>';
    }
    return '<div class="px-sprite hero"></div>';
  }

  function bindChoices() {
    els.choices.forEach(btn => {
      btn.addEventListener('click', () => submitChoice(btn.dataset.choice));
    });
  }

  function submitChoice(choice) {
    if (state.loading) return;
    if (choice === 'RESET') {
      setLoading(true);
      google.script.run
        .withSuccessHandler(data => {
          setLoading(false);
          bootstrapGame(data);
          showToast('새로운 캐릭터를 생성하세요.');
        })
        .withFailureHandler(err => {
          setLoading(false);
          showToast(err && err.message ? err.message : '세션 초기화 실패');
        })
        .resetPlayerSession();
      return;
    }
    setLoading(true);
    google.script.run
      .withSuccessHandler(data => {
        setLoading(false);
        bootstrapGame(data);
        const { result } = data;
        if (result) {
          result.messages.forEach(message => showToast(message));
          if (result.levelUps && result.levelUps.length) {
            showToast(`레벨 ${result.levelUps.map(l => l.level).join(', ')} 달성!`);
          }
          if (result.death) {
            showToast('패배했습니다. 캐릭터를 재생성하세요.');
          }
        }
      })
      .withFailureHandler(err => {
        setLoading(false);
        showToast(err && err.message ? err.message : '선택 처리 실패');
      })
      .chooseOption({ choice });
  }

  function bootstrapGame(data) {
    if (!data) return;
    state.debug = data.debug || { isAdmin: false };
    if (data.needsCreation) {
      renderCreation(data);
      return;
    }
    state.player = data.player;
    state.gameState = data.state;
    renderHUD(data.player);
    renderEvent(data.event, data.queueSize);
    renderLogs(data.logs);
    els.creation.hidden = true;
    els.gameplay.hidden = false;
    els.hud.hidden = false;
    els.debugPanel.hidden = !state.debug.isAdmin;
  }

  function bindKeyboard() {
    document.addEventListener('keydown', evt => {
      if (state.loading) return;
      if (['INPUT', 'TEXTAREA'].includes(evt.target.tagName)) return;
      if (evt.key === '1') {
        evt.preventDefault();
        submitChoice('A');
      } else if (evt.key === '2') {
        evt.preventDefault();
        submitChoice('B');
      } else if (evt.key === '3') {
        evt.preventDefault();
        submitChoice('C');
      }
    });
  }

  function bindDebug() {
    els.btnForceEvent.addEventListener('click', () => {
      const eventId = els.debugEvent.value.trim();
      if (!eventId) return;
      setLoading(true);
      google.script.run
        .withSuccessHandler(data => {
          setLoading(false);
          bootstrapGame(data);
          showToast('이벤트가 큐에 추가되었습니다.');
        })
        .withFailureHandler(err => {
          setLoading(false);
          showToast(err && err.message ? err.message : '디버그 실패');
        })
        .toggleDebugAction('forceEvent', eventId);
    });
    els.btnHeal.addEventListener('click', () => {
      setLoading(true);
      google.script.run
        .withSuccessHandler(data => {
          setLoading(false);
          bootstrapGame(data);
          showToast('체력을 회복했습니다.');
        })
        .withFailureHandler(err => {
          setLoading(false);
          showToast(err && err.message ? err.message : '디버그 실패');
        })
        .toggleDebugAction('heal', true);
    });
  }

  function init() {
    bindChoices();
    bindKeyboard();
    bindDebug();
    setLoading(true);
    google.script.run
      .withSuccessHandler(data => {
        setLoading(false);
        bootstrapGame(data);
        els.inputName && els.inputName.focus();
      })
      .withFailureHandler(err => {
        setLoading(false);
        showToast(err && err.message ? err.message : '로딩 실패');
      })
      .loadGame();
  }

  document.addEventListener('DOMContentLoaded', init);
})();
</script>
