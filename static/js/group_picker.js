/**
 * group_picker.js
 * Уніфікований пошуковий селектор груп (eu-group-picker) для всіх сторінок системи.
 *
 * Працює у двох режимах:
 *  1) Кастомна розмітка на головній сторінці (students.html: .eu-group-picker[data-eu-gp="custom"])
 *     зі збереженням групування за ступенями (groups_by_degree) та спеціальностей.
 *  2) Автоматичне покращення стандартних <select> вибору групи на всіх інших сторінках
 *     (фільтри, картка студента, накази, сесії, військовий облік, заявки, експорт тощо).
 *
 * Можливості:
 *  - Багатослівний пошук за назвою, роком, формою навчання, курсом, спеціальністю
 *  - Пошук без дефісів/пробілів ("кн24" -> "КН-24") + толерантність до розкладки (EN -> UA)
 *  - Безпечне підсвічування збігів (<mark> через DOM TextNode)
 *  - Навігація з клавіатури (ArrowUp/ArrowDown/Home/End/Enter/Escape + набір з кнопки)
 *  - Кнопка швидкого скидання (×) на тригері та в полі пошуку
 *  - Лічильник знайдених груп і порожній стан «Групу не знайдено»
 *  - Lazy-рендеринг списку (швидке завантаження сторінок із десятками рядків, як renewal_students)
 *  - Popper strategy: 'fixed' — випадаючий список не обрізається в table-responsive і модалках
 */
(function () {
    'use strict';

    const SELECT_SELECTOR = [
        'select[name="group_id"]',
        'select[name="to_group_id"]',
        'select[name="target_group_id"]',
        'select.js-group-select'
    ].join(', ');

    // Мапа помилкової англійської QWERTY-розкладки в українську
    const EN_TO_UA = {
        q: 'й', w: 'ц', e: 'у', r: 'к', t: 'е', y: 'н', u: 'г', i: 'ш', o: 'щ', p: 'з',
        '[': 'х', ']': 'ї', a: 'ф', s: 'і', d: 'в', f: 'а', g: 'п', h: 'р', j: 'о', k: 'л',
        l: 'д', ';': 'ж', "'": 'є', z: 'я', x: 'ч', c: 'с', v: 'м', b: 'и', n: 'т', m: 'ь',
        ',': 'б', '.': 'ю', '`': 'ґ', '~': 'ґ'
    };

    function convertEnToUa(str) {
        let out = '';
        for (let i = 0; i < str.length; i++) {
            const ch = str[i];
            out += EN_TO_UA[ch] || ch;
        }
        return out;
    }

    function compactStr(str) {
        return (str || '').toLowerCase().replace(/[\s\-–—_.,()]+/g, '');
    }

    function escapeRegExp(str) {
        return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    function parseQuery(raw) {
        const clean = (raw || '').trim().toLowerCase();
        if (!clean) return null;
        const tokens = clean.split(/\s+/).filter(Boolean);
        const tokenSpecs = tokens.map(function (t) {
            const ua = convertEnToUa(t);
            const variants = [t];
            if (ua !== t) variants.push(ua);
            const compactVariants = variants.map(compactStr).filter(Boolean);
            return { variants: variants, compactVariants: compactVariants };
        });
        return { raw: clean, tokenSpecs: tokenSpecs };
    }

    function matchesSearch(haystack, compactHaystack, querySpec) {
        if (!querySpec) return true;
        for (let i = 0; i < querySpec.tokenSpecs.length; i++) {
            const spec = querySpec.tokenSpecs[i];
            let matched = false;
            for (let j = 0; j < spec.variants.length; j++) {
                if (haystack.indexOf(spec.variants[j]) !== -1) {
                    matched = true;
                    break;
                }
            }
            if (!matched) {
                for (let j = 0; j < spec.compactVariants.length; j++) {
                    if (compactHaystack.indexOf(spec.compactVariants[j]) !== -1) {
                        matched = true;
                        break;
                    }
                }
            }
            if (!matched) return false;
        }
        return true;
    }

    function renderHighlightedText(container, text, querySpec) {
        container.textContent = '';
        if (!text) return;
        if (!querySpec || !querySpec.tokenSpecs.length) {
            container.textContent = text;
            return;
        }
        const terms = [];
        querySpec.tokenSpecs.forEach(function (spec) {
            spec.variants.forEach(function (v) {
                if (v.length >= 1) terms.push(escapeRegExp(v));
            });
        });
        if (!terms.length) {
            container.textContent = text;
            return;
        }
        // Сортуємо довші токени спершу, щоб не ламати збіги
        terms.sort(function (a, b) { return b.length - a.length; });
        const rx = new RegExp('(' + terms.join('|') + ')', 'gi');
        let lastIdx = 0;
        let match;
        while ((match = rx.exec(text)) !== null) {
            if (match.index > lastIdx) {
                container.appendChild(document.createTextNode(text.slice(lastIdx, match.index)));
            }
            const mark = document.createElement('mark');
            mark.className = 'eu-gp-mark';
            mark.textContent = match[0];
            container.appendChild(mark);
            lastIdx = rx.lastIndex;
            if (match[0].length === 0) rx.lastIndex++;
        }
        if (lastIdx < text.length) {
            container.appendChild(document.createTextNode(text.slice(lastIdx)));
        }
    }

    function shouldSkipSelect(select) {
        if (!select || select.tagName !== 'SELECT') return true;
        if (select.multiple) return true;
        if (select.dataset.euGpInit === '1') return true;
        if (select.dataset.noGroupPicker === '1' || select.dataset.groupPicker === 'off') return true;
        if (select.closest('#freezeStudentModal')) return true;
        const firstOpt = select.options[0];
        if (firstOpt && (firstOpt.textContent || '').toLowerCase().indexOf('загальний') !== -1) {
            return true;
        }
        return false;
    }

    function getPopperConfig(defaultCfg) {
        const modifiers = (defaultCfg && defaultCfg.modifiers ? defaultCfg.modifiers.slice() : []);
        return Object.assign({}, defaultCfg || {}, {
            strategy: 'fixed',
            modifiers: modifiers.concat([
                { name: 'offset', options: { offset: [0, 4] } },
                { name: 'preventOverflow', options: { padding: 8 } }
            ])
        });
    }

    function createSearchHeader(placeholderText) {
        const wrap = document.createElement('div');
        wrap.className = 'eu-group-picker-search-wrap';

        const inputGroup = document.createElement('div');
        inputGroup.className = 'eu-group-picker-search-box';

        const icon = document.createElement('i');
        icon.className = 'bi bi-search eu-group-picker-search-icon';
        icon.setAttribute('aria-hidden', 'true');

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'form-control form-control-sm eu-group-picker-search';
        input.placeholder = placeholderText || 'Пошук групи…';
        input.autocomplete = 'off';
        input.setAttribute('aria-label', 'Пошук групи');

        const clearQueryBtn = document.createElement('button');
        clearQueryBtn.type = 'button';
        clearQueryBtn.className = 'eu-group-picker-search-clear';
        clearQueryBtn.setAttribute('aria-label', 'Очистити пошук');
        clearQueryBtn.hidden = true;
        const clearIcon = document.createElement('i');
        clearIcon.className = 'bi bi-x-lg';
        clearQueryBtn.appendChild(clearIcon);

        inputGroup.appendChild(icon);
        inputGroup.appendChild(input);
        inputGroup.appendChild(clearQueryBtn);

        const meta = document.createElement('div');
        meta.className = 'eu-group-picker-meta';
        const countEl = document.createElement('span');
        countEl.className = 'eu-group-picker-count';
        const hintEl = document.createElement('span');
        hintEl.className = 'eu-group-picker-kbd-hint';
        hintEl.textContent = '↑↓ вибір · Enter · Esc';
        meta.appendChild(countEl);
        meta.appendChild(hintEl);

        wrap.appendChild(inputGroup);
        wrap.appendChild(meta);

        // Зупиняємо спливання кліку, щоб Bootstrap не закривав dropdown
        wrap.addEventListener('click', function (e) {
            e.stopPropagation();
        });

        return {
            wrap: wrap,
            input: input,
            clearQueryBtn: clearQueryBtn,
            countEl: countEl
        };
    }

    function createEmptyState(onReset) {
        const empty = document.createElement('div');
        empty.className = 'eu-group-picker-empty';
        empty.hidden = true;

        const icon = document.createElement('i');
        icon.className = 'bi bi-search';
        icon.setAttribute('aria-hidden', 'true');

        const text = document.createElement('div');
        text.className = 'eu-group-picker-empty-text';
        text.textContent = 'Групу не знайдено';

        const resetBtn = document.createElement('button');
        resetBtn.type = 'button';
        resetBtn.className = 'btn btn-link btn-sm eu-group-picker-empty-reset';
        resetBtn.textContent = 'Скинути пошук';
        resetBtn.addEventListener('click', function (e) {
            e.stopPropagation();
            if (onReset) onReset();
        });

        empty.appendChild(icon);
        empty.appendChild(text);
        empty.appendChild(resetBtn);
        return empty;
    }

    function bindSharedPickerBehavior(ctx) {
        /*
         * ctx:
         *  - root: HTMLElement (.eu-group-picker)
         *  - triggerBtn: HTMLButtonElement
         *  - clearTriggerBtn: HTMLButtonElement | null
         *  - menu: HTMLElement (.dropdown-menu)
         *  - list: HTMLElement (.eu-group-picker-list)
         *  - search: { input, clearQueryBtn, countEl }
         *  - emptyEl: HTMLElement
         *  - getItems(): Array<{ el, value, label, searchable, compactSearchable, isReset, isDisabled, textNodeTarget }>
         *  - getHeaders(): Array<{ el, items }>
         *  - getSelectedValue(): string
         *  - canClear(): boolean
         *  - onSelect(value, itemObj): void
         *  - onBeforeOpen(): void
         */
        let activeIdx = -1;
        let pendingTypeChar = '';

        const dd = (window.bootstrap && window.bootstrap.Dropdown)
            ? window.bootstrap.Dropdown.getOrCreateInstance(ctx.triggerBtn, {
                autoClose: 'outside',
                popperConfig: getPopperConfig
            })
            : null;

        function getVisibleItems() {
            return ctx.getItems().filter(function (it) {
                return !it.el.hidden && !it.isDisabled;
            });
        }

        function setActiveIndex(idx) {
            const visible = getVisibleItems();
            ctx.getItems().forEach(function (it) {
                it.el.classList.remove('is-active');
            });
            if (!visible.length) {
                activeIdx = -1;
                return;
            }
            if (idx < 0) idx = visible.length - 1;
            if (idx >= visible.length) idx = 0;
            activeIdx = idx;
            const target = visible[activeIdx];
            target.el.classList.add('is-active');
            if (typeof target.el.scrollIntoView === 'function') {
                target.el.scrollIntoView({ block: 'nearest' });
            }
        }

        function syncSelectedUI() {
            const currentVal = String(ctx.getSelectedValue() || '');
            let selectedEl = null;
            ctx.getItems().forEach(function (it) {
                const isSel = String(it.value || '') === currentVal;
                it.el.classList.toggle('is-selected', isSel);
                it.el.setAttribute('aria-selected', isSel ? 'true' : 'false');
                if (isSel) selectedEl = it.el;
            });
            if (ctx.clearTriggerBtn) {
                const showClear = Boolean(currentVal !== '' && ctx.canClear() && !ctx.triggerBtn.disabled);
                ctx.clearTriggerBtn.hidden = !showClear;
                ctx.triggerBtn.classList.toggle('has-clear', showClear);
            }
            return selectedEl;
        }

        function applyFilter() {
            const raw = ctx.search.input.value || '';
            const querySpec = parseQuery(raw);
            ctx.search.clearQueryBtn.hidden = !raw;

            const items = ctx.getItems();
            let matchCount = 0;
            let totalGroups = 0;

            items.forEach(function (it) {
                if (!it.isReset) totalGroups++;
                if (it.isReset) {
                    // Кнопка "Всі групи" / "— оберіть —" показується коли пошук порожній,
                    // або якщо назва сама збігається з пошуком
                    const showReset = !querySpec || matchesSearch(it.searchable, it.compactSearchable, querySpec);
                    it.el.hidden = !showReset;
                    if (it.textNodeTarget) {
                        renderHighlightedText(it.textNodeTarget, it.displayText, querySpec);
                    }
                    if (it.metaNodeTarget) {
                        renderHighlightedText(it.metaNodeTarget, it.metaText, querySpec);
                    }
                    return;
                }
                const ok = matchesSearch(it.searchable, it.compactSearchable, querySpec);
                it.el.hidden = !ok;
                if (ok) matchCount++;
                if (it.textNodeTarget) {
                    renderHighlightedText(it.textNodeTarget, it.displayText, querySpec);
                }
                if (it.metaNodeTarget) {
                    renderHighlightedText(it.metaNodeTarget, it.metaText, querySpec);
                }
            });

            ctx.getHeaders().forEach(function (hdr) {
                if (!querySpec) {
                    hdr.el.hidden = false;
                    return;
                }
                const anyVis = hdr.items.some(function (it) { return !it.el.hidden; });
                hdr.el.hidden = !anyVis;
            });

            const visibleAny = items.some(function (it) { return !it.el.hidden; });
            ctx.emptyEl.hidden = visibleAny;

            if (querySpec) {
                ctx.search.countEl.textContent = 'Знайдено: ' + matchCount + ' з ' + totalGroups;
            } else {
                ctx.search.countEl.textContent = 'Усього груп: ' + totalGroups;
            }

            activeIdx = -1;
            if (querySpec) {
                const vis = getVisibleItems();
                const firstGroupIdx = vis.findIndex(function (it) { return !it.isReset; });
                if (firstGroupIdx !== -1) setActiveIndex(firstGroupIdx);
                else if (vis.length > 0) setActiveIndex(0);
            }
        }

        function resetFilterAndFocus() {
            ctx.search.input.value = '';
            applyFilter();
            ctx.search.input.focus();
        }

        ctx.search.input.addEventListener('input', applyFilter);
        ctx.search.clearQueryBtn.addEventListener('click', function (e) {
            e.stopPropagation();
            resetFilterAndFocus();
        });

        ctx.search.input.addEventListener('keydown', function (e) {
            const visible = getVisibleItems();
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                setActiveIndex(activeIdx + 1);
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                setActiveIndex(activeIdx <= 0 ? visible.length - 1 : activeIdx - 1);
            } else if (e.key === 'Home') {
                e.preventDefault();
                setActiveIndex(0);
            } else if (e.key === 'End') {
                e.preventDefault();
                setActiveIndex(visible.length - 1);
            } else if (e.key === 'Enter') {
                e.preventDefault();
                if (!visible.length) return;
                let chosen = activeIdx >= 0 && visible[activeIdx] ? visible[activeIdx] : null;
                if (!chosen) {
                    const querySpec = parseQuery(ctx.search.input.value);
                    chosen = (querySpec
                        ? visible.find(function (it) { return !it.isReset; })
                        : null) || visible[0];
                }
                if (chosen) {
                    if (dd) dd.hide();
                    ctx.onSelect(chosen.value, chosen);
                }
            } else if (e.key === 'Escape') {
                e.preventDefault();
                e.stopPropagation();
                if (ctx.search.input.value) {
                    resetFilterAndFocus();
                } else {
                    if (dd) dd.hide();
                    ctx.triggerBtn.focus();
                }
            }
        });

        if (ctx.clearTriggerBtn) {
            ctx.clearTriggerBtn.addEventListener('click', function (e) {
                e.preventDefault();
                e.stopPropagation();
                if (dd) dd.hide();
                ctx.onSelect('', null);
                ctx.triggerBtn.focus();
            });
        }

        // Набір символів прямо на кнопці тригера відкриває меню й одразу шукає
        ctx.triggerBtn.addEventListener('keydown', function (e) {
            if (ctx.triggerBtn.disabled) return;
            if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
                e.preventDefault();
                if (dd) dd.show();
                return;
            }
            if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey && e.key !== ' ') {
                e.preventDefault();
                pendingTypeChar = e.key;
                if (dd) dd.show();
            }
        });

        ctx.root.addEventListener('show.bs.dropdown', function () {
            if (ctx.onBeforeOpen) ctx.onBeforeOpen();
            ctx.search.input.value = pendingTypeChar || '';
            pendingTypeChar = '';
            applyFilter();
            syncSelectedUI();
        });

        ctx.root.addEventListener('shown.bs.dropdown', function () {
            requestAnimationFrame(function () {
                ctx.search.input.focus({ preventScroll: true });
                if (ctx.search.input.value) {
                    const len = ctx.search.input.value.length;
                    ctx.search.input.setSelectionRange(len, len);
                } else {
                    const selEl = syncSelectedUI();
                    if (selEl && typeof selEl.scrollIntoView === 'function') {
                        selEl.scrollIntoView({ block: 'nearest' });
                    }
                }
            });
        });

        return {
            dropdown: dd,
            syncSelectedUI: syncSelectedUI,
            applyFilter: applyFilter,
            resetFilterAndFocus: resetFilterAndFocus
        };
    }

    const SEARCH_PLACEHOLDER = 'Назва, рік, форма, курс, спеціальність…';

    let CATALOG_LOADED = false;
    let CATALOG_GROUPS = {};
    let DEGREE_ORDER = {};

    function loadCatalog() {
        if (CATALOG_LOADED) return;
        const el = document.getElementById('euGroupsCatalog');
        if (!el) return;
        CATALOG_LOADED = true;
        try {
            const data = JSON.parse(el.textContent || '{}');
            if (data && typeof data === 'object') {
                if (data.groups && typeof data.groups === 'object') {
                    CATALOG_GROUPS = data.groups;
                }
                if (data.degree_order && typeof data.degree_order === 'object') {
                    DEGREE_ORDER = data.degree_order;
                }
            }
        } catch (e) {
            CATALOG_GROUPS = {};
            DEGREE_ORDER = {};
        }
    }

    function formatTriggerLabelFromCatalog(cat) {
        if (!cat || !cat.name) return '';
        const extra = [cat.year, cat.form].filter(Boolean).join(', ');
        return extra ? (cat.name + ' (' + extra + ')') : cat.name;
    }

    function buildCatalogMetaText(cat, rawText) {
        const parts = [];
        if (cat.year) parts.push(String(cat.year));
        if (cat.form) parts.push(cat.form);
        if (cat.course) parts.push(cat.course + ' курс');
        if (cat.specialty) parts.push(cat.specialty);
        if (rawText && rawText.toLowerCase().indexOf('архів') !== -1 && parts.join(' ').toLowerCase().indexOf('архів') === -1) {
            parts.push('архів');
        }
        return parts.join(' · ');
    }

    function parseFallbackOptionText(rawText) {
        const m = rawText.match(/^(.+?)\s*\(([^)]+)\)\s*$/);
        if (!m) {
            return { title: rawText, meta: '' };
        }
        const meta = m[2]
            .split(/\s*,\s*/)
            .map(function (s) { return s.trim(); })
            .filter(Boolean)
            .join(' · ');
        return { title: m[1].trim(), meta: meta };
    }

    /**
     * Режим 1: Покращення кастомного компонента на students.html
     */
    function initCustomPicker(root) {
        if (!root || root.dataset.euGpInit === '1') return;
        root.dataset.euGpInit = '1';

        const triggerBtn = root.querySelector('.eu-group-picker-btn');
        const labelEl = root.querySelector('.eu-group-picker-label');
        const clearTriggerBtn = root.querySelector('.eu-group-picker-clear');
        const menu = root.querySelector('.eu-group-picker-menu');
        const list = root.querySelector('.eu-group-picker-list');
        const hiddenInput = document.getElementById(root.dataset.targetInput || 'groupIdInput');
        const form = root.closest('form');
        if (!triggerBtn || !menu || !list || !hiddenInput) return;

        const search = createSearchHeader(SEARCH_PLACEHOLDER);
        menu.insertBefore(search.wrap, list);

        let api = null;
        const emptyEl = createEmptyState(function () {
            if (api) api.resetFilterAndFocus();
        });
        list.appendChild(emptyEl);

        const rawItems = Array.from(list.querySelectorAll('.eu-group-picker-item'));
        const items = rawItems.map(function (el) {
            const val = el.getAttribute('data-group-id') || '';
            const label = el.getAttribute('data-group-label') || el.textContent.trim();
            const isReset = val === '';
            const titleEl = el.querySelector('.eu-gp-item-title');
            const metaEl = el.querySelector('.eu-gp-item-meta');
            const displayText = titleEl ? titleEl.textContent : label;
            const metaText = metaEl ? metaEl.textContent : '';
            const searchBase = (el.getAttribute('data-search') || label || '').toLowerCase();
            return {
                el: el,
                value: val,
                label: label,
                displayText: displayText,
                textNodeTarget: titleEl,
                metaText: metaText,
                metaNodeTarget: metaEl,
                searchable: searchBase,
                compactSearchable: compactStr(searchBase),
                isReset: isReset,
                isDisabled: false
            };
        });

        // Прив'язка заголовків ступенів до їхніх пунктів
        const headers = [];
        Array.from(list.querySelectorAll('.eu-group-picker-degree')).forEach(function (hdrEl) {
            const groupItems = [];
            let node = hdrEl.nextElementSibling;
            while (node && node.classList.contains('eu-group-picker-item')) {
                const found = items.find(function (it) { return it.el === node; });
                if (found) groupItems.push(found);
                node = node.nextElementSibling;
            }
            headers.push({ el: hdrEl, items: groupItems });
        });

        function submitSelection(val, itemObj) {
            hiddenInput.value = val || '';
            if (labelEl) {
                labelEl.textContent = itemObj ? itemObj.label : 'Всі групи';
            }
            triggerBtn.classList.toggle('is-placeholder', !val);
            if (api) api.syncSelectedUI();
            if (form) form.submit();
        }

        items.forEach(function (it) {
            it.el.addEventListener('click', function () {
                if (api && api.dropdown) api.dropdown.hide();
                submitSelection(it.value, it);
            });
        });

        api = bindSharedPickerBehavior({
            root: root,
            triggerBtn: triggerBtn,
            clearTriggerBtn: clearTriggerBtn,
            menu: menu,
            list: list,
            search: search,
            emptyEl: emptyEl,
            getItems: function () { return items; },
            getHeaders: function () { return headers; },
            getSelectedValue: function () { return hiddenInput.value; },
            canClear: function () { return true; },
            onSelect: submitSelection,
            onBeforeOpen: null
        });

        triggerBtn.classList.toggle('is-placeholder', !hiddenInput.value);
        api.syncSelectedUI();
    }

    /**
     * Режим 2: Автоматичне покращення стандартних <select> груп
     * (візуально й функціонально ідентично students.html)
     */
    function initSelectPicker(select) {
        if (shouldSkipSelect(select)) return;
        loadCatalog();
        select.dataset.euGpInit = '1';

        const isSmall = select.classList.contains('form-select-sm');
        const root = document.createElement('div');
        root.className = 'dropdown eu-group-picker';

        const triggerBtn = document.createElement('button');
        triggerBtn.type = 'button';
        triggerBtn.className = 'form-select text-start eu-group-picker-btn' + (isSmall ? ' form-select-sm' : '');
        triggerBtn.setAttribute('data-bs-toggle', 'dropdown');
        triggerBtn.setAttribute('data-bs-auto-close', 'outside');
        triggerBtn.setAttribute('aria-expanded', 'false');

        const labelEl = document.createElement('span');
        labelEl.className = 'eu-group-picker-label';
        triggerBtn.appendChild(labelEl);

        const clearTriggerBtn = document.createElement('button');
        clearTriggerBtn.type = 'button';
        clearTriggerBtn.className = 'eu-group-picker-clear';
        clearTriggerBtn.setAttribute('aria-label', 'Очистити вибір групи');
        clearTriggerBtn.setAttribute('title', 'Очистити');
        clearTriggerBtn.hidden = true;
        const clearIcon = document.createElement('i');
        clearIcon.className = 'bi bi-x-lg';
        clearTriggerBtn.appendChild(clearIcon);

        const menu = document.createElement('div');
        menu.className = 'dropdown-menu eu-group-picker-menu p-0';

        const search = createSearchHeader(SEARCH_PLACEHOLDER);
        const list = document.createElement('div');
        list.className = 'eu-group-picker-list';

        let api = null;
        const emptyEl = createEmptyState(function () {
            if (api) api.resetFilterAndFocus();
        });

        menu.appendChild(search.wrap);
        menu.appendChild(list);

        // Переміщуємо <select> усередину root, залишаючи його у формі та layout-потоці
        const parent = select.parentNode;
        parent.insertBefore(root, select);
        root.appendChild(select);
        root.appendChild(triggerBtn);
        root.appendChild(clearTriggerBtn);
        root.appendChild(menu);

        select.classList.add('eu-group-picker-native');
        select.setAttribute('tabindex', '-1');
        select.setAttribute('aria-hidden', 'true');

        let builtSignature = null;
        let items = [];
        let headers = [];
        let isSyncingValue = false;

        function getOptionsSignature() {
            const parts = [];
            for (let i = 0; i < select.options.length; i++) {
                const o = select.options[i];
                parts.push(o.value + ':' + o.disabled + ':' + o.textContent.trim());
            }
            return parts.join('|');
        }

        function canClear() {
            for (let i = 0; i < select.options.length; i++) {
                const o = select.options[i];
                if (o.value === '' && !o.disabled) return true;
            }
            return false;
        }

        function getOptionDisplayLabel(opt) {
            if (!opt) return '— оберіть групу —';
            const val = String(opt.value || '');
            const rawText = opt.textContent.replace(/\s+/g, ' ').trim();
            if (val !== '' && CATALOG_GROUPS[val]) {
                return formatTriggerLabelFromCatalog(CATALOG_GROUPS[val]) || rawText;
            }
            return rawText || '— оберіть групу —';
        }

        function buildOptionRow(opt, explicitDegree) {
            const val = String(opt.value || '');
            const rawText = opt.textContent.replace(/\s+/g, ' ').trim();
            const isReset = val === '';
            const cat = !isReset ? CATALOG_GROUPS[val] : null;
            const optGroupLabel = explicitDegree || (
                opt.parentElement && opt.parentElement.tagName === 'OPTGROUP'
                    ? (opt.parentElement.label || '')
                    : ''
            );

            let displayText = rawText;
            let metaText = '';
            let triggerLabel = rawText;
            let degree = optGroupLabel;

            if (isReset) {
                displayText = rawText || 'Всі групи';
                triggerLabel = displayText;
            } else if (cat && cat.name) {
                displayText = cat.name;
                metaText = buildCatalogMetaText(cat, rawText);
                triggerLabel = formatTriggerLabelFromCatalog(cat) || rawText;
                degree = optGroupLabel || cat.degree || 'Без ступеня';
            } else {
                const parsed = parseFallbackOptionText(rawText);
                displayText = parsed.title;
                metaText = parsed.meta;
                triggerLabel = rawText;
            }

            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'dropdown-item eu-group-picker-item';
            if (isReset) btn.classList.add('fw-semibold');
            if (opt.disabled) {
                btn.disabled = true;
                btn.classList.add('disabled');
            }

            const contentWrap = document.createElement('span');
            contentWrap.className = 'eu-gp-item-content';

            const titleSpan = document.createElement('span');
            titleSpan.className = 'eu-gp-item-title';
            titleSpan.textContent = displayText;
            contentWrap.appendChild(titleSpan);

            let metaSpan = null;
            if (metaText) {
                metaSpan = document.createElement('span');
                metaSpan.className = 'eu-gp-item-meta';
                metaSpan.textContent = metaText;
                contentWrap.appendChild(metaSpan);
            }

            const checkIcon = document.createElement('i');
            checkIcon.className = 'bi bi-check2 eu-group-picker-check';
            checkIcon.setAttribute('aria-hidden', 'true');

            btn.appendChild(contentWrap);
            btn.appendChild(checkIcon);

            const searchParts = [rawText];
            if (cat) {
                if (cat.name) searchParts.push(cat.name);
                if (cat.form) searchParts.push(cat.form);
                if (cat.year) searchParts.push(String(cat.year));
                if (cat.course) searchParts.push(cat.course + ' курс');
                if (cat.specialty) searchParts.push(cat.specialty);
            }
            if (degree) searchParts.push(degree);
            if (opt.getAttribute('data-study-form')) searchParts.push(opt.getAttribute('data-study-form'));
            if (opt.getAttribute('data-label')) searchParts.push(opt.getAttribute('data-label'));

            const searchBase = searchParts.join(' ').toLowerCase();

            const itemObj = {
                el: btn,
                value: val,
                label: triggerLabel,
                displayText: displayText,
                textNodeTarget: titleSpan,
                metaText: metaText,
                metaNodeTarget: metaSpan,
                degree: degree,
                searchable: searchBase,
                compactSearchable: compactStr(searchBase),
                isReset: isReset,
                isDisabled: opt.disabled
            };

            btn.addEventListener('click', function () {
                if (opt.disabled) return;
                if (api && api.dropdown) api.dropdown.hide();
                applySelectValue(val);
            });

            return itemObj;
        }

        function buildMenuList() {
            const sig = getOptionsSignature();
            if (builtSignature === sig) return;
            builtSignature = sig;

            list.textContent = '';
            items = [];
            headers = [];

            const hasOptgroups = Array.from(select.children).some(function (c) {
                return c.tagName === 'OPTGROUP';
            });

            if (hasOptgroups) {
                Array.from(select.children).forEach(function (child) {
                    if (child.tagName === 'OPTGROUP') {
                        const hdr = document.createElement('div');
                        hdr.className = 'eu-group-picker-degree';
                        hdr.textContent = child.label || '';
                        list.appendChild(hdr);
                        const grpItems = [];
                        Array.from(child.children).forEach(function (opt) {
                            if (opt.tagName !== 'OPTION') return;
                            const itemObj = buildOptionRow(opt, child.label || '');
                            list.appendChild(itemObj.el);
                            items.push(itemObj);
                            grpItems.push(itemObj);
                        });
                        headers.push({ el: hdr, items: grpItems });
                    } else if (child.tagName === 'OPTION') {
                        if (child.value === '' && child.disabled) return;
                        const itemObj = buildOptionRow(child, '');
                        list.appendChild(itemObj.el);
                        items.push(itemObj);
                    }
                });
            } else {
                const resetRows = [];
                const groupRows = [];
                Array.from(select.options).forEach(function (opt) {
                    if (opt.value === '') {
                        // Задизейблений плейсхолдер («— оберіть —») не додаємо в список,
                        // але активний скидач («Всі групи», «— оберіть групу —») залишаємо.
                        if (!opt.disabled) {
                            resetRows.push(buildOptionRow(opt, ''));
                        }
                        return;
                    }
                    groupRows.push(buildOptionRow(opt, ''));
                });

                resetRows.forEach(function (itemObj) {
                    list.appendChild(itemObj.el);
                    items.push(itemObj);
                });

                const hasDegrees = groupRows.some(function (it) { return Boolean(it.degree); });
                if (hasDegrees) {
                    const buckets = {};
                    groupRows.forEach(function (it) {
                        const deg = it.degree || 'Без ступеня';
                        if (!Object.prototype.hasOwnProperty.call(buckets, deg)) {
                            buckets[deg] = [];
                        }
                        buckets[deg].push(it);
                    });

                    const sortedDegrees = Object.keys(buckets).sort(function (a, b) {
                        const oa = Object.prototype.hasOwnProperty.call(DEGREE_ORDER, a) ? DEGREE_ORDER[a] : 999;
                        const ob = Object.prototype.hasOwnProperty.call(DEGREE_ORDER, b) ? DEGREE_ORDER[b] : 999;
                        if (oa !== ob) return oa - ob;
                        return a.localeCompare(b, 'uk');
                    });

                    sortedDegrees.forEach(function (deg) {
                        const hdr = document.createElement('div');
                        hdr.className = 'eu-group-picker-degree';
                        hdr.textContent = deg;
                        list.appendChild(hdr);
                        const grpItems = [];
                        buckets[deg].forEach(function (itemObj) {
                            list.appendChild(itemObj.el);
                            items.push(itemObj);
                            grpItems.push(itemObj);
                        });
                        headers.push({ el: hdr, items: grpItems });
                    });
                } else {
                    groupRows.forEach(function (itemObj) {
                        list.appendChild(itemObj.el);
                        items.push(itemObj);
                    });
                }
            }

            list.appendChild(emptyEl);
        }

        function syncTriggerFromSelect() {
            const selectedOpt = select.selectedIndex >= 0 ? select.options[select.selectedIndex] : null;
            labelEl.textContent = getOptionDisplayLabel(selectedOpt);
            const isEmpty = !select.value;
            triggerBtn.classList.toggle('is-placeholder', isEmpty);
            triggerBtn.disabled = select.disabled;
            if (api) api.syncSelectedUI();
        }

        function applySelectValue(val) {
            const changed = select.value !== String(val || '');
            isSyncingValue = true;
            try {
                select.value = val || '';
            } finally {
                isSyncingValue = false;
            }
            triggerBtn.classList.remove('is-invalid');
            syncTriggerFromSelect();
            if (changed) {
                select.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }

        // Відслідковуємо програмне присвоєння select.value = '...' (як у модалках переведення)
        const valueDescriptor = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value');
        if (valueDescriptor && valueDescriptor.configurable) {
            Object.defineProperty(select, 'value', {
                configurable: true,
                enumerable: true,
                get: function () {
                    return valueDescriptor.get.call(this);
                },
                set: function (v) {
                    valueDescriptor.set.call(this, v);
                    if (!isSyncingValue) {
                        triggerBtn.classList.remove('is-invalid');
                        syncTriggerFromSelect();
                    }
                }
            });
        }

        api = bindSharedPickerBehavior({
            root: root,
            triggerBtn: triggerBtn,
            clearTriggerBtn: clearTriggerBtn,
            menu: menu,
            list: list,
            search: search,
            emptyEl: emptyEl,
            getItems: function () { return items; },
            getHeaders: function () { return headers; },
            getSelectedValue: function () { return select.value; },
            canClear: canClear,
            onSelect: function (val) { applySelectValue(val); },
            onBeforeOpen: function () {
                syncTriggerFromSelect();
                buildMenuList();
            }
        });

        select.__euGpSync = syncTriggerFromSelect;

        // Синхронізація, якщо value або disabled змінили ззовні (JS або form.reset())
        select.addEventListener('change', function () {
            triggerBtn.classList.remove('is-invalid');
            syncTriggerFromSelect();
        });
        select.addEventListener('focus', function () {
            triggerBtn.focus();
        });
        select.addEventListener('invalid', function () {
            triggerBtn.classList.add('is-invalid');
            triggerBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
            if (api && api.dropdown) api.dropdown.show();
        });

        if (select.form) {
            select.form.addEventListener('reset', function () {
                setTimeout(syncTriggerFromSelect, 0);
            });
        }

        if (window.MutationObserver) {
            const mo = new MutationObserver(function () {
                builtSignature = null;
                syncTriggerFromSelect();
            });
            mo.observe(select, {
                attributes: true,
                attributeFilter: ['disabled'],
                childList: true,
                subtree: true
            });
        }

        syncTriggerFromSelect();
    }

    function initAll(scope) {
        loadCatalog();
        const root = scope || document;
        root.querySelectorAll('.eu-group-picker[data-eu-gp="custom"]').forEach(initCustomPicker);
        root.querySelectorAll(SELECT_SELECTOR).forEach(initSelectPicker);
    }

    window.EUGroupPicker = {
        initAll: initAll,
        initSelect: initSelectPicker,
        initCustom: initCustomPicker
    };

    document.addEventListener('show.bs.modal', function (e) {
        if (e.target) {
            initAll(e.target);
            e.target.querySelectorAll(SELECT_SELECTOR).forEach(function (sel) {
                if (typeof sel.__euGpSync === 'function') sel.__euGpSync();
            });
        }
    });

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function () { initAll(document); });
    } else {
        initAll(document);
    }
})();
