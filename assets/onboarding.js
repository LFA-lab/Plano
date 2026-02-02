/*
Conception (Github Pages, sans dépendances):
- Persistance via localStorage, namespacée par (hostname + pathname).
- Accessibilité: boutons ARIA pour le switch, aria-expanded pour les bandeaux Plan B, libellés checkbox basés sur le titre.
- Performance: aucun recalcul massif; on toggle par classes/attributs.
- Sécurité: tous les liens target="_blank" reçoivent rel="noopener noreferrer".
- Politique Macros: tous les nouveaux chemins doivent être ASCII/slug (sans accents/espaces) à l'avenir.

NOTE: i18n désactivé côté HTML (boutons FR/EN commentés). 
Le code de traduction interne (labels badges, progression) reste actif en français.
Pour réactiver l'i18n, décommenter le bloc .language-selector dans index.html.
*/

(function() {
  const ns = `${location.host}${location.pathname}`;
  const LS_ROLE_KEY = `onboard:role:${ns}`;
  const LS_CHECKLIST_KEY = `onboard:checklist:${ns}`;

  document.addEventListener('DOMContentLoaded', () => {
    ensureNoopener();
    initLastUpdated();
    renderBadges();
    initChecklist();
    initRoleSwitch();
    initPlanB();
  });

  function ensureNoopener() {
    document.querySelectorAll('a[target="_blank"]').forEach(a => {
      if (!a.rel || !/noopener/.test(a.rel)) {
        a.rel = 'noopener noreferrer';
      }
    });
  }

  function initLastUpdated() {
    const el = document.getElementById('last-updated');
    if (!el) return;
    try {
      const d = new Date();
      const lang = document.documentElement.lang || 'fr';
      const formatter = new Intl.DateTimeFormat(lang, { day: '2-digit', month: 'long', year: 'numeric' });
      const labelByLang = {
        fr: 'Dernière mise à jour',
        en: 'Last updated',
        hi: 'अंतिम अपडेट'
      };
      el.textContent = `${labelByLang[lang] || labelByLang.fr} : ${formatter.format(d)}`;
    } catch (_) {
      // Fallback statique
      const lang = document.documentElement.lang || 'fr';
      const labelByLang = { fr: 'Dernière mise à jour', en: 'Last updated', hi: 'अंतिम अपडेट' };
      const mo = { fr: 'octobre 2025', en: 'October 2025', hi: 'अक्टूबर 2025' };
      el.textContent = `${labelByLang[lang] || labelByLang.fr} : ${mo[lang] || mo.fr}`;
    }
  }

  function renderBadges() {
    document.querySelectorAll('.task-item').forEach((item) => {
      // skip if already rendered
      if (item.querySelector('.task-badges')) return;
      const duration = item.getAttribute('data-duration');
      const outcome = item.getAttribute('data-outcome');
      const wrap = document.createElement('div');
      wrap.className = 'task-badges';
      if (duration) {
        const b = document.createElement('span');
        b.className = 'badge badge-duration';
        b.textContent = label('duration') + ' ' + duration;
        wrap.appendChild(b);
      }
      if (outcome) {
        const b = document.createElement('span');
        b.className = 'badge badge-outcome';
        b.textContent = label('outcome') + ' ' + outcome;
        wrap.appendChild(b);
      }
      if (wrap.children.length) {
        item.insertBefore(wrap, item.querySelector('p, .app-links, .contact-grid'));
      }
    });
  }

  function initChecklist() {
    const state = loadChecklistState();
    const tasks = Array.from(document.querySelectorAll('.task-item'));
    tasks.forEach((item, idx) => {
      const titleEl = item.querySelector('h4');
      const existingCheckbox = item.querySelector('.task-checkbox-label .task-checkbox');
      
      // Si la tâche a déjà une checkbox avec un label (vue Nouveau Projet)
      if (existingCheckbox && !titleEl) {
        const spanEl = item.querySelector('.task-checkbox-label span');
        const title = (spanEl?.textContent || `task-${idx}`).trim();
        const key = stableKey(title, idx);
        
        // Restaurer l'état sauvegardé
        existingCheckbox.checked = state[key] === true;
        applyCompleted(item, state[key] === true);
        
        // Ajouter un listener pour sauvegarder
        existingCheckbox.addEventListener('change', function() {
          const newState = loadChecklistState();
          newState[key] = this.checked;
          saveChecklistState(newState);
          applyCompleted(item, this.checked);
          renderProgress(getAllTrackedTasks(), loadChecklistState());
        });
        
        return;
      }
      
      // Sinon, traiter les tâches avec h4 (vue Arrivée)
      if (!titleEl) {
        return;
      }
      
      const title = (titleEl?.textContent || `task-${idx}`).trim();
      const key = stableKey(title, idx);

      // header wrapper for checkbox + title
      if (!item.querySelector('.task-header')) {
        const header = document.createElement('div');
        header.className = 'task-header';
        const h4 = titleEl;
        const placeholder = document.createElement('div');
        item.insertBefore(header, h4);
        header.appendChild(createCheckbox(key, title, state[key] === true));
        header.appendChild(h4);
      }

      applyCompleted(item, state[key] === true);
    });

    renderProgress(getAllTrackedTasks(), state);
  }

  function getAllTrackedTasks() {
    const tasks = Array.from(document.querySelectorAll('.task-item'));
    return tasks.filter(item => {
      // Inclure les tâches avec h4 ou avec une checkbox existante
      return item.querySelector('h4') || item.querySelector('.task-checkbox-label .task-checkbox');
    });
  }

  function createCheckbox(key, title, checked) {
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.className = 'task-checkbox';
    cb.checked = !!checked;
    cb.setAttribute('aria-label', title);
    cb.addEventListener('change', () => {
      const state = loadChecklistState();
      state[key] = cb.checked;
      saveChecklistState(state);
      const item = cb.closest('.task-item');
      applyCompleted(item, cb.checked);
      renderProgress(getAllTrackedTasks(), state);
    });
    return cb;
  }

  function applyCompleted(item, completed) {
    item.classList.toggle('task-completed', !!completed);
  }

  function renderProgress(tasks, state) {
    const containerId = 'checklist-progress';
    let wrap = document.getElementById(containerId);
    const total = tasks.length;
    const done = tasks.reduce((acc, item, idx) => {
      const titleEl = item.querySelector('h4');
      const spanEl = item.querySelector('.task-checkbox-label span');
      const title = titleEl ? titleEl.textContent.trim() : (spanEl ? spanEl.textContent.trim() : `task-${idx}`);
      const key = stableKey(title, idx);
      return acc + (loadChecklistState()[key] ? 1 : 0);
    }, 0);
    if (!wrap) {
      wrap = document.createElement('div');
      wrap.id = containerId;
      const main = document.querySelector('main');
      if (!main) return;
      const bar = document.createElement('div');
      bar.className = 'progress-bar';
      bar.innerHTML = `
        <div class="progress-track"><div class="progress-fill"></div></div>
        <div class="progress-meta" aria-live="polite"></div>
        <button type="button" class="btn-reset" id="btn-reset-checklist"></button>
      `;
      wrap.appendChild(bar);
      main.insertBefore(wrap, main.firstElementChild);
      document.getElementById('btn-reset-checklist').addEventListener('click', resetChecklist);
    }
    const lang = document.documentElement.lang || 'fr';
    const meta = wrap.querySelector('.progress-meta');
    const fill = wrap.querySelector('.progress-fill');
    const pct = total ? Math.round((done / total) * 100) : 0;
    fill.style.width = `${pct}%`;
    meta.textContent = label('progress', { done, total, lang });
    const btn = wrap.querySelector('#btn-reset-checklist');
    btn.textContent = label('reset');
  }

  function resetChecklist() {
    localStorage.removeItem(LS_CHECKLIST_KEY);
    document.querySelectorAll('.task-checkbox').forEach(cb => { cb.checked = false; });
    document.querySelectorAll('.task-item').forEach(item => item.classList.remove('task-completed'));
    renderProgress(getAllTrackedTasks(), loadChecklistState());
  }

  function loadChecklistState() {
    try { return JSON.parse(localStorage.getItem(LS_CHECKLIST_KEY) || '{}'); } catch { return {}; }
  }
  function saveChecklistState(state) { localStorage.setItem(LS_CHECKLIST_KEY, JSON.stringify(state)); }
  function stableKey(title, idx) {
    return `${slug(title)}-${idx}`;
  }

  function initRoleSwitch() {
    const roleButtons = document.querySelectorAll('.role-btn');
    const views = document.querySelectorAll('[data-view]');

    if (roleButtons.length === 0 || views.length === 0) {
      return;
    }

    function setActiveRole(roleName) {
      views.forEach(view => {
        const viewName = view.getAttribute('data-view');
        const shouldShow = viewName === roleName;
        view.style.display = shouldShow ? 'block' : 'none';
      });

      roleButtons.forEach(btn => {
        const isActive = btn.getAttribute('data-role-select') === roleName;
        btn.setAttribute('aria-pressed', isActive ? 'true' : 'false');
        btn.classList.toggle('is-active', isActive);
      });

      // Sauvegarder dans localStorage
      saveRole(roleName);
    }

    roleButtons.forEach((btn) => {
      const roleValue = btn.getAttribute('data-role-select');
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        setActiveRole(roleValue);
      });
    });

    // Initialiser avec la valeur sauvegardée ou le bouton actif par défaut
    const saved = loadRole();
    if (saved === 'arrivee' || saved === 'nouveau-projet') {
      setActiveRole(saved);
    } else {
      const active = document.querySelector('.role-btn[aria-pressed="true"]') || roleButtons[0];
      if (active) {
        const defaultRole = active.getAttribute('data-role-select');
        setActiveRole(defaultRole);
      }
    }
  }

  function loadRole() {
    const saved = localStorage.getItem(LS_ROLE_KEY);
    // Si la valeur sauvegardée est une des nouvelles valeurs, l'utiliser
    if (saved === 'arrivee' || saved === 'nouveau-projet') {
      return saved;
    }
    // Sinon, retourner null pour utiliser le défaut HTML
    return null;
  }
  function saveRole(role) {
    localStorage.setItem(LS_ROLE_KEY, role);
  }

  function initPlanB() {
    document.querySelectorAll('.planb').forEach(box => {
      const btn = box.querySelector('.planb-toggle');
      const content = box.querySelector('.planb-content');
      if (!btn || !content) return;
      // visible par défaut mais compact
      content.hidden = false;
      btn.addEventListener('click', () => {
        const expanded = btn.getAttribute('aria-expanded') === 'true';
        btn.setAttribute('aria-expanded', String(!expanded));
        content.hidden = expanded;
      });
    });
  }

  function label(kind, ctx) {
    const lang = document.documentElement.lang || 'fr';
    const L = {
      fr: { duration: 'Durée :', outcome: 'Résultat :', progress: (o) => `${o.done}/${o.total} tâches complétées`, reset: 'Réinitialiser' },
      en: { duration: 'Duration:', outcome: 'Outcome:', progress: (o) => `${o.done}/${o.total} tasks completed`, reset: 'Reset' },
      hi: { duration: 'अवधि:', outcome: 'परिणाम:', progress: (o) => `${o.done}/${o.total} कार्य पूर्ण`, reset: 'रीसेट' },
    };
    const t = L[lang] || L.fr;
    if (kind === 'progress') return t.progress(ctx || { done: 0, total: 0 });
    return t[kind];
  }

  function slug(s) {
    return s
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '');
  }
})();


