/* ========================================
   ArtiKraft — Newsletter Builder Engine
   ======================================== */

(function() {
    'use strict';

    function getCurrentQuarter() {
        const now = new Date();
        const q = Math.ceil((now.getMonth() + 1) / 3);
        return `Q${q} ${now.getFullYear()}`;
    }

    // ============ State ============
    const state = {
        blocks: [],
        selectedBlockId: null,
        theme: 'corporate-blue',
        zoom: 100,
        isPreview: false,
        undoStack: [],
        redoStack: [],
        customColors: {
            primary: '#004E8C',
            accent: '#03787C',
            bg: '#FFFFFF',
            text: '#333333',
            cardBg: '#F5F7FA'
        },
        fonts: {
            heading: "'Inter', sans-serif",
            body: "'Inter', sans-serif",
            baseSize: 16
        },
        info: {
            title: getCurrentQuarter() + ' Newsletter',
            edition: '1st Edition',
            quarter: getCurrentQuarter(),
            org: 'Newsletter Name'
        }
    };

    let blockIdCounter = 0;
    function nextId() { return 'block-' + (++blockIdCounter); }

    // ============ Themes ============
    const THEMES = {
        'corporate-blue': {
            name: 'Corporate Blue', primary: '#004E8C', accent: '#03787C',
            bg: '#FFFFFF', text: '#333333', cardBg: '#F0F6FC',
            gradient: 'linear-gradient(135deg, #004E8C, #03787C)'
        },
        'ocean-teal': {
            name: 'Ocean Teal', primary: '#005B5E', accent: '#00838F',
            bg: '#FFFFFF', text: '#2D3436', cardBg: '#E0F2F1',
            gradient: 'linear-gradient(135deg, #005B5E, #00ACC1)'
        },
        'deep-purple': {
            name: 'Deep Purple', primary: '#4A148C', accent: '#7C4DFF',
            bg: '#FFFFFF', text: '#333333', cardBg: '#F3E5F5',
            gradient: 'linear-gradient(135deg, #4A148C, #7C4DFF)'
        },
        'sunset-warm': {
            name: 'Sunset Warm', primary: '#BF360C', accent: '#FF6D00',
            bg: '#FFFFFF', text: '#3E2723', cardBg: '#FFF3E0',
            gradient: 'linear-gradient(135deg, #BF360C, #FF6D00)'
        },
        'forest-green': {
            name: 'Forest Green', primary: '#1B5E20', accent: '#4CAF50',
            bg: '#FFFFFF', text: '#333333', cardBg: '#E8F5E9',
            gradient: 'linear-gradient(135deg, #1B5E20, #66BB6A)'
        },
        'midnight': {
            name: 'Midnight', primary: '#1A237E', accent: '#536DFE',
            bg: '#FAFAFA', text: '#212121', cardBg: '#E8EAF6',
            gradient: 'linear-gradient(135deg, #1A237E, #536DFE)'
        },
        'rose-gold': {
            name: 'Rose Gold', primary: '#880E4F', accent: '#F06292',
            bg: '#FFFFFF', text: '#333333', cardBg: '#FCE4EC',
            gradient: 'linear-gradient(135deg, #880E4F, #F06292)'
        },
        'slate-modern': {
            name: 'Slate Modern', primary: '#37474F', accent: '#607D8B',
            bg: '#FFFFFF', text: '#263238', cardBg: '#ECEFF1',
            gradient: 'linear-gradient(135deg, #37474F, #78909C)'
        },
        'electric-indigo': {
            name: 'Electric Indigo', primary: '#3730A3', accent: '#6366F1',
            bg: '#FFFFFF', text: '#1E1B4B', cardBg: '#EEF2FF',
            gradient: 'linear-gradient(135deg, #3730A3, #818CF8)'
        }
    };

    // ============ Block Defaults ============
    const BLOCK_DEFAULTS = {
        'hero': () => ({
            type: 'hero',
            title: state.info.title || 'Newsletter Title',
            subtitle: 'Your subtitle goes here',
            meta: `${state.info.org} | ${state.info.quarter} | ${state.info.edition}`,
            date: new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }),
            bgImage: ''
        }),
        'editorial': () => ({
            type: 'editorial',
            greeting: 'Dear Readers,',
            content: '<p>Welcome to this edition of our newsletter. In this issue, we highlight the key achievements, milestones, and initiatives that shaped our quarter. We hope you find these updates insightful and inspiring.</p><p>Happy reading!</p>',
            signoff: 'Warm regards,',
            authorName: 'The Editor',
            authorTitle: ''
        }),
        'section-header': () => ({
            type: 'section-header',
            title: 'Section Title'
        }),
        'divider': () => ({ type: 'divider' }),
        'spacer': () => ({ type: 'spacer', height: 32 }),
        'text': () => ({
            type: 'text',
            content: '<p>Enter your content here. You can format text using the toolbar that appears when you select text.</p>'
        }),
        'image-text': () => ({
            type: 'image-text',
            imageUrl: '',
            title: 'Feature Title',
            content: '<p>Describe this feature, update, or highlight. Add relevant details about the initiative.</p>',
            imagePosition: 'left'
        }),
        'quote': () => ({
            type: 'quote',
            text: 'Innovation distinguishes between a leader and a follower.',
            author: 'Leadership Team'
        }),
        'callout': () => ({
            type: 'callout',
            title: '💡 Did You Know?',
            content: 'Add an important announcement, tip, or call-to-action here.'
        }),
        'kpi-cards': () => ({
            type: 'kpi-cards',
            cards: [
                { value: '94%', label: 'Adoption Rate', change: '+12%', direction: 'up' },
                { value: '23', label: 'Milestones Hit', change: '+5', direction: 'up' },
                { value: '1.8', label: 'Quality Index', change: '-0.3', direction: 'down' },
                { value: '340', label: 'Deliverables', change: '+45', direction: 'up' }
            ]
        }),
        'progress-bars': () => ({
            type: 'progress-bars',
            items: [
                { label: 'Progress A', value: 94, color: '#4f46e5' },
                { label: 'Progress B', value: 78, color: '#22c55e' },
                { label: 'Progress C', value: 85, color: '#f59e0b' },
                { label: 'Progress D', value: 92, color: '#06b6d4' }
            ]
        }),
        'stats-row': () => ({
            type: 'stats-row',
            stats: [
                { value: '156', label: 'Items Delivered' },
                { value: '98.2%', label: 'Success Rate' },
                { value: '42', label: 'Contributors' },
                { value: '12', label: 'Events' }
            ]
        }),
        'timeline': () => ({
            type: 'timeline',
            items: [
                { date: 'January 2026', title: 'Milestone 1', desc: 'Describe your first key milestone here.' },
                { date: 'February 2026', title: 'Milestone 2', desc: 'Describe your second key milestone here.' },
                { date: 'March 2026', title: 'Milestone 3', desc: 'Describe your third key milestone here.' }
            ]
        }),
        'success-story': () => ({
            type: 'success-story',
            badge: '⭐ Success Story',
            title: 'Your Success Story Title',
            content: 'Describe the success story — what was achieved, the impact, and key takeaways.'
        }),
        'team-spotlight': () => ({
            type: 'team-spotlight',
            members: [
                { name: 'Team Member 1', role: 'Lead', initials: 'TM' },
                { name: 'Team Member 2', role: 'Developer', initials: 'TM' },
                { name: 'Team Member 3', role: 'Analyst', initials: 'TM' },
                { name: 'Team Member 4', role: 'QA Lead', initials: 'TM' }
            ]
        }),
        'two-column': () => ({
            type: 'two-column',
            columns: [
                { title: 'Column 1', content: 'Add content for the first column here.' },
                { title: 'Column 2', content: 'Add content for the second column here.' }
            ]
        }),
        'three-column': () => ({
            type: 'three-column',
            columns: [
                { title: 'Column 1', content: 'Content here.' },
                { title: 'Column 2', content: 'Content here.' },
                { title: 'Column 3', content: 'Content here.' }
            ]
        }),
        'four-column': () => ({
            type: 'four-column',
            columns: [
                { title: 'Column 1', content: 'Content here.' },
                { title: 'Column 2', content: 'Content here.' },
                { title: 'Column 3', content: 'Content here.' },
                { title: 'Column 4', content: 'Content here.' }
            ]
        }),
        'sidebar-layout': () => ({
            type: 'sidebar-layout',
            main: { title: 'Main Content', content: 'Primary content goes here. This takes up two-thirds of the width.' },
            sidebar: { title: 'Sidebar', content: 'Quick links, highlights, or supplementary info.' }
        }),
        'footer': () => ({
            type: 'footer',
            appName: 'ArtiKraft',
            tagline: 'For teams, leaders & communicators',
            year: new Date().getFullYear().toString(),
            builtBy: 'Sivakanth Badigenchala',
            links: [{ label: 'Share Feedback', url: '#' }]
        }),
        'footer-minimal': () => ({
            type: 'footer-minimal',
            text: '© ' + new Date().getFullYear() + ' Your Newsletter. All rights reserved.',
            builtBy: 'Sivakanth Badigenchala'
        }),
        'footer-social': () => ({
            type: 'footer-social',
            appName: 'Newsletter',
            tagline: 'Stay connected with us',
            year: new Date().getFullYear().toString(),
            builtBy: 'Sivakanth Badigenchala',
            socials: [
                { icon: '🔗', label: 'LinkedIn', url: '#' },
                { icon: '🐦', label: 'Twitter', url: '#' },
                { icon: '📧', label: 'Email', url: '#' },
                { icon: '🌐', label: 'Website', url: '#' }
            ]
        })
    };

    // ============ Templates ============
    const TEMPLATES = [
        {
            id: 'the-pulse',
            name: 'Magazine Style',
            desc: 'Quarterly magazine style with hero, KPIs, success stories & timeline',
            thumb: 'linear-gradient(135deg, #004E8C, #03787C)',
            theme: 'corporate-blue',
            blocks: ['hero', 'editorial', 'stats-row', 'section-header', 'text', 'kpi-cards', 'divider',
                     'section-header', 'success-story', 'divider', 'section-header', 'timeline',
                     'divider', 'section-header', 'two-column', 'quote', 'footer']
        },
        {
            id: 'service-eng',
            name: 'Team Spotlight',
            desc: 'Delivery highlights, metrics, and team spotlights',
            thumb: 'linear-gradient(135deg, #005B5E, #00ACC1)',
            theme: 'ocean-teal',
            blocks: ['hero', 'editorial', 'section-header', 'text', 'progress-bars', 'divider',
                     'section-header', 'image-text', 'divider', 'section-header',
                     'team-spotlight', 'callout', 'footer']
        },
        {
            id: 'ai-innovation',
            name: 'AI Innovation Brief',
            desc: 'AI adoption, productivity metrics, and process improvements',
            thumb: 'linear-gradient(135deg, #3730A3, #818CF8)',
            theme: 'electric-indigo',
            blocks: ['hero', 'editorial', 'kpi-cards', 'section-header', 'image-text', 'divider',
                     'section-header', 'success-story', 'stats-row', 'divider',
                     'section-header', 'three-column', 'quote', 'footer']
        },
        {
            id: 'executive-brief',
            name: 'Executive Brief',
            desc: 'Clean, minimal layout for leadership updates',
            thumb: 'linear-gradient(135deg, #37474F, #78909C)',
            theme: 'slate-modern',
            blocks: ['hero', 'editorial', 'text', 'kpi-cards', 'divider', 'section-header',
                     'two-column', 'callout', 'footer']
        },
        {
            id: 'sprint-review',
            name: 'Progress Review',
            desc: 'Progress metrics, velocity, and retrospective highlights',
            thumb: 'linear-gradient(135deg, #1B5E20, #66BB6A)',
            theme: 'forest-green',
            blocks: ['hero', 'editorial', 'stats-row', 'progress-bars', 'divider',
                     'section-header', 'timeline', 'divider', 'section-header',
                     'text', 'quote', 'footer']
        },
        {
            id: 'blank',
            name: 'Blank Canvas',
            desc: 'Start from scratch with an empty canvas',
            thumb: 'linear-gradient(135deg, #283954, #374e74)',
            theme: 'corporate-blue',
            blocks: []
        },
        {
            id: 'client-update',
            name: 'Client Update',
            desc: 'Client-facing update with highlights, metrics & next steps',
            thumb: 'linear-gradient(135deg, #880E4F, #F06292)',
            theme: 'rose-gold',
            blocks: ['hero', 'editorial', 'kpi-cards', 'divider', 'section-header',
                     'sidebar-layout', 'divider', 'section-header', 'timeline',
                     'callout', 'footer']
        },
        {
            id: 'product-launch',
            name: 'Product Launch',
            desc: 'Announce features, showcase benefits & roadmap',
            thumb: 'linear-gradient(135deg, #4A148C, #7C4DFF)',
            theme: 'deep-purple',
            blocks: ['hero', 'text', 'section-header', 'four-column', 'divider',
                     'section-header', 'image-text', 'success-story', 'divider',
                     'section-header', 'timeline', 'callout', 'footer']
        },
        {
            id: 'weekly-digest',
            name: 'Weekly Digest',
            desc: 'Compact weekly roundup with quick highlights',
            thumb: 'linear-gradient(135deg, #BF360C, #FF6D00)',
            theme: 'sunset-warm',
            blocks: ['hero', 'stats-row', 'divider', 'section-header',
                     'three-column', 'divider', 'section-header', 'text',
                     'quote', 'footer']
        },
        {
            id: 'culture-people',
            name: 'Culture & People',
            desc: 'Team stories, spotlights, events & culture highlights',
            thumb: 'linear-gradient(135deg, #1A237E, #536DFE)',
            theme: 'midnight',
            blocks: ['hero', 'editorial', 'section-header', 'team-spotlight',
                     'divider', 'section-header', 'success-story', 'divider',
                     'section-header', 'image-text', 'two-column', 'quote', 'footer']
        },
        {
            id: 'annual-review',
            name: 'Annual Review',
            desc: 'Year-in-review with metrics, timeline & achievements',
            thumb: 'linear-gradient(135deg, #004E8C, #00ACC1)',
            theme: 'corporate-blue',
            blocks: ['hero', 'editorial', 'kpi-cards', 'stats-row', 'divider',
                     'section-header', 'timeline', 'divider', 'section-header',
                     'four-column', 'divider', 'section-header', 'success-story',
                     'success-story', 'divider', 'section-header', 'team-spotlight',
                     'quote', 'footer']
        }
    ];

    // ============ DOM Refs ============
    const $ = (sel) => document.querySelector(sel);
    const $$ = (sel) => document.querySelectorAll(sel);
    const canvas = $('#canvas');
    const canvasEmpty = $('#canvasEmpty');

    // ============ Init ============
    function init() {
        loadBuilderTheme();
        loadCustomThemes();
        renderThemes();
        renderTemplates();
        bindEvents();
        applyTheme('corporate-blue');
        // Set dynamic defaults in inputs
        $('#newsletterTitle').value = state.info.title;
        $('#infoQuarter').value = state.info.quarter;
        loadFromLocalStorage();
    }

    // ============ Builder Light/Dark Mode ============
    function loadBuilderTheme() {
        const saved = localStorage.getItem('artikraft-builder-theme');
        if (saved === 'light') {
            document.body.classList.add('light-mode');
        }
    }

    function toggleBuilderTheme() {
        document.body.classList.toggle('light-mode');
        const isLight = document.body.classList.contains('light-mode');
        localStorage.setItem('artikraft-builder-theme', isLight ? 'light' : 'dark');
    }

    // ============ Theme Rendering ============
    function renderThemes() {
        const grid = $('#themeGrid');
        grid.innerHTML = '';
        Object.entries(THEMES).forEach(([id, t]) => {
            const swatch = document.createElement('div');
            swatch.className = 'theme-swatch' + (state.theme === id ? ' active' : '');
            swatch.style.background = t.gradient;
            swatch.dataset.theme = id;
            swatch.innerHTML = `<span class="swatch-label">${t.name}</span>`;
            swatch.addEventListener('click', () => {
                applyTheme(id);
                $$('.theme-swatch').forEach(s => s.classList.remove('active'));
                swatch.classList.add('active');
            });
            grid.appendChild(swatch);
        });
    }

    function applyTheme(themeId) {
        const t = THEMES[themeId];
        if (!t) return;
        state.theme = themeId;
        state.customColors = { primary: t.primary, accent: t.accent, bg: t.bg, text: t.text, cardBg: t.cardBg };
        applyCustomColors();
        // Sync color inputs
        $('#colorPrimary').value = t.primary;
        $('#colorAccent').value = t.accent;
        $('#colorBg').value = t.bg;
        $('#colorText').value = t.text;
        $('#colorCardBg').value = t.cardBg;
    }

    function applyCustomColors() {
        const c = state.customColors;
        canvas.style.setProperty('--nl-primary', c.primary);
        canvas.style.setProperty('--nl-accent', c.accent);
        canvas.style.setProperty('--nl-bg', c.bg);
        canvas.style.setProperty('--nl-text', c.text);
        canvas.style.setProperty('--nl-card-bg', c.cardBg);
        canvas.style.background = c.bg;
    }

    function saveCustomTheme() {
        const name = prompt('Theme name:');
        if (!name || !name.trim()) return;
        const id = 'custom-' + name.trim().toLowerCase().replace(/\s+/g, '-');
        const c = state.customColors;
        THEMES[id] = {
            name: name.trim(),
            primary: c.primary, accent: c.accent,
            bg: c.bg, text: c.text, cardBg: c.cardBg,
            gradient: `linear-gradient(135deg, ${c.primary}, ${c.accent})`,
            custom: true
        };
        // Save custom themes to localStorage
        const customThemes = JSON.parse(localStorage.getItem('artikraft-custom-themes') || '{}');
        customThemes[id] = THEMES[id];
        localStorage.setItem('artikraft-custom-themes', JSON.stringify(customThemes));
        renderThemes();
        applyTheme(id);
        toast(`Theme "${name.trim()}" saved!`, 'success');
    }

    function loadCustomThemes() {
        try {
            const custom = JSON.parse(localStorage.getItem('artikraft-custom-themes') || '{}');
            Object.entries(custom).forEach(([id, theme]) => { THEMES[id] = theme; });
        } catch (e) {}
    }

    // ============ Template Rendering ============
    function renderTemplates() {
        const grid = $('#templateGrid');
        grid.innerHTML = '';
        TEMPLATES.forEach(t => {
            const card = document.createElement('div');
            card.className = 'template-card';
            card.innerHTML = `
                <div class="template-thumb" style="background:${t.thumb}"></div>
                <div class="template-info">
                    <h4>${t.name}</h4>
                    <p>${t.desc}</p>
                </div>
                <div class="template-actions">
                    <button class="template-preview-btn" title="Preview">👁</button>
                    <button class="template-apply-btn">Apply</button>
                </div>
            `;
            card.querySelector('.template-apply-btn').addEventListener('click', (e) => {
                e.stopPropagation();
                loadTemplate(t);
                closeModal('templateModal');
            });
            card.querySelector('.template-preview-btn').addEventListener('click', (e) => {
                e.stopPropagation();
                previewTemplate(t);
            });
            card.addEventListener('click', () => {
                loadTemplate(t);
                closeModal('templateModal');
            });
            grid.appendChild(card);
        });
    }

    function previewTemplate(template) {
        // Generate a preview by temporarily building the template
        const savedBlocks = JSON.parse(JSON.stringify(state.blocks));
        const savedColors = { ...state.customColors };
        const savedTheme = state.theme;
        
        const previewBlocks = [];
        const t = template.theme ? THEMES[template.theme] : null;
        const c = t ? { primary: t.primary, accent: t.accent, bg: t.bg, text: t.text, cardBg: t.cardBg } : savedColors;
        
        template.blocks.forEach(type => {
            if (BLOCK_DEFAULTS[type]) {
                const b = BLOCK_DEFAULTS[type]();
                b.id = 'preview-' + Math.random().toString(36).substr(2, 6);
                previewBlocks.push(b);
            }
        });

        // Build preview HTML
        const f = state.fonts;
        let previewHTML = `<div style="max-width:600px;margin:0 auto;background:${c.bg};font-family:${f.body};font-size:${Math.round(f.baseSize * 0.75)}px;color:${c.text};transform:scale(1);border-radius:8px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.3)">`;
        previewBlocks.forEach(b => {
            // Temporarily set customColors for getBlockHTML
            state.customColors = c;
            let html = getBlockHTML(b);
            html = html.replace(/contenteditable="true"/g, '');
            html = html.replace(/data-field="[^"]*"/g, '');
            html = html.replace(/data-placeholder="[^"]*"/g, '');
            previewHTML += html;
        });
        previewHTML += '</div>';
        
        // Restore state
        state.customColors = savedColors;
        state.theme = savedTheme;

        // Show in preview modal
        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay show template-preview-overlay';
        overlay.innerHTML = `<div class="modal template-preview-modal">
            <div class="modal-header">
                <h2>${template.name} — Preview</h2>
                <button class="modal-close template-preview-close">✕</button>
            </div>
            <div class="modal-body template-preview-body">${previewHTML}</div>
            <div class="template-preview-footer">
                <button class="btn-start template-preview-apply">Apply This Template</button>
                <button class="btn-start btn-secondary template-preview-back">Back to Templates</button>
            </div>
        </div>`;
        document.body.appendChild(overlay);
        overlay.querySelector('.template-preview-close').addEventListener('click', () => overlay.remove());
        overlay.querySelector('.template-preview-back').addEventListener('click', () => overlay.remove());
        overlay.querySelector('.template-preview-apply').addEventListener('click', () => {
            overlay.remove();
            loadTemplate(template);
            closeModal('templateModal');
        });
        overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
    }

    function loadTemplate(template) {
        saveUndo();
        state.blocks = [];
        if (template.theme) applyTheme(template.theme);
        template.blocks.forEach(type => {
            const blockData = BLOCK_DEFAULTS[type]();
            blockData.id = nextId();
            // Customize section headers based on position
            if (type === 'section-header') {
                const sectionNames = ['Highlights', 'Success Stories', 'Key Milestones',
                                      'Improvements', 'Innovation', 'Updates'];
                const idx = state.blocks.filter(b => b.type === 'section-header').length;
                blockData.title = sectionNames[idx % sectionNames.length];
            }
            state.blocks.push(blockData);
        });
        renderCanvas();
        toast('Template loaded!', 'success');
    }

    // ============ Canvas Rendering ============
    function renderCanvas() {
        // Remove old blocks (keep canvasEmpty)
        canvas.querySelectorAll('.nl-block, .drop-zone').forEach(el => el.remove());

        if (state.blocks.length === 0) {
            canvasEmpty.style.display = 'flex';
            saveToLocalStorage();
            return;
        }
        canvasEmpty.style.display = 'none';

        // Add top drop zone
        canvas.appendChild(createDropZone(0));

        state.blocks.forEach((block, index) => {
            const el = renderBlock(block, index);
            canvas.appendChild(el);
            canvas.appendChild(createDropZone(index + 1));
        });

        applyCustomColors();
        applyFonts();
        saveToLocalStorage();
    }

    function renderBlock(block, index) {
        const wrapper = document.createElement('div');
        wrapper.className = 'nl-block';
        wrapper.id = block.id;
        wrapper.dataset.index = index;
        wrapper.draggable = true;
        if (state.selectedBlockId === block.id) wrapper.classList.add('selected');

        // Block actions toolbar
        wrapper.innerHTML = `<div class="block-actions">
            <span class="block-type-label">${block.type.replace(/-/g, ' ')}</span>
            <button class="block-action-btn" data-action="moveUp" title="Move up">↑</button>
            <button class="block-action-btn" data-action="moveDown" title="Move down">↓</button>
            <button class="block-action-btn" data-action="duplicate" title="Duplicate">⧉</button>
            <button class="block-action-btn danger" data-action="delete" title="Delete">✕</button>
        </div>`;

        const content = document.createElement('div');
        content.innerHTML = getBlockHTML(block);
        wrapper.appendChild(content.firstElementChild || content);

        // Click to select
        wrapper.addEventListener('click', (e) => {
            if (e.target.closest('.block-actions')) return;
            selectBlock(block.id);
        });

        // Block actions
        wrapper.querySelectorAll('.block-action-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                handleBlockAction(btn.dataset.action, block.id);
            });
        });

        // Make editable content save on input
        wrapper.querySelectorAll('[contenteditable="true"]').forEach(editable => {
            editable.addEventListener('input', () => {
                syncEditableToState(block.id, wrapper);
            });
            editable.addEventListener('focus', () => {
                showRichToolbar(editable);
            });
        });

        // Canvas drag-reorder
        wrapper.addEventListener('dragstart', (e) => {
            if (e.target.closest('[contenteditable]')) { e.preventDefault(); return; }
            e.dataTransfer.setData('text/block-reorder', block.id);
            e.dataTransfer.effectAllowed = 'move';
            wrapper.classList.add('dragging');
            setTimeout(() => wrapper.style.opacity = '0.4', 0);
        });
        wrapper.addEventListener('dragend', () => {
            wrapper.classList.remove('dragging');
            wrapper.style.opacity = '';
        });

        return wrapper;
    }

    function getBlockHTML(b) {
        switch (b.type) {
            case 'hero':
                const heroBg = b.bgImage 
                    ? `background-image: linear-gradient(rgba(0,0,0,0.45), rgba(0,0,0,0.55)), url('${b.bgImage}'); background-size: cover; background-position: center;`
                    : `background: linear-gradient(135deg, var(--nl-primary), var(--nl-accent));`;
                return `<div class="nl-hero" style="${heroBg}">
                    <div class="hero-meta-badge" contenteditable="true" data-field="meta">${esc(b.meta)}</div>
                    <h1 contenteditable="true" data-field="title">${esc(b.title)}</h1>
                    <div class="hero-subtitle" contenteditable="true" data-field="subtitle">${esc(b.subtitle)}</div>
                    <div class="hero-date" contenteditable="true" data-field="date">${esc(b.date || '')}</div>
                    <div class="hero-img-upload" title="Click to set background image">📷 Set Cover Image</div>
                </div>`;

            case 'editorial':
                return `<div class="nl-editorial">
                    <div class="editorial-inner">
                        <div class="editorial-greeting" contenteditable="true" data-field="greeting">${esc(b.greeting)}</div>
                        <div class="editorial-content" contenteditable="true" data-field="content">${b.content}</div>
                        <div class="editorial-signoff">
                            <span contenteditable="true" data-field="signoff">${esc(b.signoff)}</span>
                            <strong contenteditable="true" data-field="authorName">${esc(b.authorName)}</strong>
                            <em contenteditable="true" data-field="authorTitle">${esc(b.authorTitle)}</em>
                        </div>
                    </div>
                </div>`;

            case 'section-header':
                return `<div class="nl-section-header">
                    <h2 contenteditable="true" data-field="title">${esc(b.title)}</h2>
                </div>`;

            case 'divider':
                return `<div class="nl-divider"><hr></div>`;

            case 'spacer':
                return `<div class="nl-spacer" style="height:${b.height || 32}px"></div>`;

            case 'text':
                return `<div class="nl-text"><div contenteditable="true" data-field="content" data-placeholder="Start typing...">${b.content}</div></div>`;

            case 'image-text':
                const imgHtml = b.imageUrl
                    ? `<img src="${esc(b.imageUrl)}" alt="Image">`
                    : '📷';
                return `<div class="nl-image-text" style="flex-direction:${b.imagePosition === 'right' ? 'row-reverse' : 'row'}">
                    <div class="img-placeholder" data-field="imageUrl">${imgHtml}</div>
                    <div class="text-content">
                        <h3 contenteditable="true" data-field="title">${esc(b.title)}</h3>
                        <div contenteditable="true" data-field="content">${b.content}</div>
                    </div>
                </div>`;

            case 'quote':
                return `<div class="nl-quote">
                    <blockquote contenteditable="true" data-field="text">${esc(b.text)}</blockquote>
                    <cite contenteditable="true" data-field="author">— ${esc(b.author)}</cite>
                </div>`;

            case 'callout':
                return `<div class="nl-callout">
                    <div class="callout-title" contenteditable="true" data-field="title">${esc(b.title)}</div>
                    <div class="callout-body" contenteditable="true" data-field="content">${esc(b.content)}</div>
                </div>`;

            case 'kpi-cards':
                return `<div class="nl-kpi-cards">${b.cards.map((c, i) => `
                    <div class="kpi-card">
                        <div class="kpi-value" contenteditable="true" data-field="cards.${i}.value">${esc(c.value)}</div>
                        <div class="kpi-label" contenteditable="true" data-field="cards.${i}.label">${esc(c.label)}</div>
                        <div class="kpi-change ${c.direction}" contenteditable="true" data-field="cards.${i}.change">${c.direction === 'up' ? '▲' : '▼'} ${esc(c.change)}</div>
                    </div>
                `).join('')}</div>`;

            case 'progress-bars':
                return `<div class="nl-progress-bars">${b.items.map((p, i) => `
                    <div class="progress-item">
                        <div class="progress-label">
                            <span contenteditable="true" data-field="items.${i}.label">${esc(p.label)}</span>
                            <span contenteditable="true" data-field="items.${i}.value">${p.value}%</span>
                        </div>
                        <div class="progress-track">
                            <div class="progress-fill" style="width:${p.value}%;background:${p.color}"></div>
                        </div>
                    </div>
                `).join('')}</div>`;

            case 'stats-row':
                return `<div class="nl-stats-row">${b.stats.map((s, i) => `
                    <div class="stat-item">
                        <div class="stat-value" contenteditable="true" data-field="stats.${i}.value">${esc(s.value)}</div>
                        <div class="stat-label" contenteditable="true" data-field="stats.${i}.label">${esc(s.label)}</div>
                    </div>
                `).join('')}</div>`;

            case 'timeline':
                return `<div class="nl-timeline">
                    <div class="timeline-line"></div>
                    ${b.items.map((t, i) => `
                    <div class="timeline-item">
                        <div class="timeline-dot"></div>
                        <div class="tl-date" contenteditable="true" data-field="items.${i}.date">${esc(t.date)}</div>
                        <div class="tl-title" contenteditable="true" data-field="items.${i}.title">${esc(t.title)}</div>
                        <div class="tl-desc" contenteditable="true" data-field="items.${i}.desc">${esc(t.desc)}</div>
                    </div>
                `).join('')}</div>`;

            case 'success-story':
                return `<div class="nl-success-story">
                    <div class="story-badge" contenteditable="true" data-field="badge">${esc(b.badge)}</div>
                    <h3 contenteditable="true" data-field="title">${esc(b.title)}</h3>
                    <p contenteditable="true" data-field="content">${esc(b.content)}</p>
                </div>`;

            case 'team-spotlight':
                return `<div class="nl-team-spotlight">${b.members.map((m, i) => `
                    <div class="team-member">
                        <div class="team-avatar">${esc(m.initials)}</div>
                        <div class="member-name" contenteditable="true" data-field="members.${i}.name">${esc(m.name)}</div>
                        <div class="member-role" contenteditable="true" data-field="members.${i}.role">${esc(m.role)}</div>
                    </div>
                `).join('')}</div>`;

            case 'two-column':
                return `<div class="nl-two-column">${b.columns.map((c, i) => `
                    <div class="column-cell">
                        <h3 contenteditable="true" data-field="columns.${i}.title">${esc(c.title)}</h3>
                        <div contenteditable="true" data-field="columns.${i}.content">${esc(c.content)}</div>
                    </div>
                `).join('')}</div>`;

            case 'three-column':
                return `<div class="nl-three-column">${b.columns.map((c, i) => `
                    <div class="column-cell">
                        <h3 contenteditable="true" data-field="columns.${i}.title">${esc(c.title)}</h3>
                        <div contenteditable="true" data-field="columns.${i}.content">${esc(c.content)}</div>
                    </div>
                `).join('')}</div>`;

            case 'four-column':
                return `<div class="nl-four-column">${b.columns.map((c, i) => `
                    <div class="column-cell">
                        <h3 contenteditable="true" data-field="columns.${i}.title">${esc(c.title)}</h3>
                        <div contenteditable="true" data-field="columns.${i}.content">${esc(c.content)}</div>
                    </div>
                `).join('')}</div>`;

            case 'sidebar-layout':
                return `<div class="nl-sidebar-layout">
                    <div class="sidebar-main">
                        <h3 contenteditable="true" data-field="main.title">${esc(b.main.title)}</h3>
                        <div contenteditable="true" data-field="main.content">${esc(b.main.content)}</div>
                    </div>
                    <div class="sidebar-aside">
                        <h3 contenteditable="true" data-field="sidebar.title">${esc(b.sidebar.title)}</h3>
                        <div contenteditable="true" data-field="sidebar.content">${esc(b.sidebar.content)}</div>
                    </div>
                </div>`;

            case 'footer':
                const footerLinks = (b.links || []).map(l => 
                    typeof l === 'string' 
                        ? `<a href="#">${esc(l)}</a>` 
                        : `<a href="${esc(l.url || '#')}">${esc(l.label)}</a>`
                ).join(' · ');
                return `<div class="nl-footer">
                    <div class="footer-top">
                        <span class="footer-copy">© ${esc(b.year || new Date().getFullYear().toString())} ${esc(b.appName || 'Newsletter')}</span>
                        <span class="footer-dot">·</span>
                        <span class="footer-tagline" contenteditable="true" data-field="tagline">${esc(b.tagline || '')}</span>
                    </div>
                    <div class="footer-bottom">
                        <span>Built by <a href="#" class="footer-author">${esc(b.builtBy || '')}</a></span>
                        ${footerLinks ? ' · ' + footerLinks : ''}
                    </div>
                </div>`;

            case 'footer-minimal':
                return `<div class="nl-footer nl-footer-minimal">
                    <span contenteditable="true" data-field="text">${esc(b.text || '')}</span>
                    <span class="footer-dot">·</span>
                    <span>Built by <a href="#" class="footer-author">${esc(b.builtBy || '')}</a></span>
                </div>`;

            case 'footer-social':
                const socialLinks = (b.socials || []).map(s =>
                    `<a href="${esc(s.url || '#')}" class="social-link" title="${esc(s.label)}">${s.icon}</a>`
                ).join(' ');
                return `<div class="nl-footer nl-footer-social">
                    <div class="footer-social-icons">${socialLinks}</div>
                    <div class="footer-top">
                        <span class="footer-copy">© ${esc(b.year || new Date().getFullYear().toString())} ${esc(b.appName || 'Newsletter')}</span>
                        <span class="footer-dot">·</span>
                        <span class="footer-tagline" contenteditable="true" data-field="tagline">${esc(b.tagline || '')}</span>
                    </div>
                    <div class="footer-bottom">
                        <span>Built by <a href="#" class="footer-author">${esc(b.builtBy || '')}</a></span>
                    </div>
                </div>`;

            default:
                return `<div style="padding:20px;color:#999">Unknown block type: ${b.type}</div>`;
        }
    }

    function esc(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // ============ Drop Zones ============
    function createDropZone(index) {
        const dz = document.createElement('div');
        dz.className = 'drop-zone';
        dz.dataset.dropIndex = index;
        dz.addEventListener('dragover', (e) => {
            e.preventDefault();
            dz.classList.add('active');
        });
        dz.addEventListener('dragleave', () => dz.classList.remove('active'));
        dz.addEventListener('drop', (e) => {
            e.preventDefault();
            dz.classList.remove('active');
            const blockType = e.dataTransfer.getData('text/block-type');
            const reorderId = e.dataTransfer.getData('text/block-reorder');
            if (reorderId) {
                // Reorder existing block
                const fromIdx = state.blocks.findIndex(b => b.id === reorderId);
                if (fromIdx === -1) return;
                saveUndo();
                const [moved] = state.blocks.splice(fromIdx, 1);
                let toIdx = parseInt(dz.dataset.dropIndex);
                if (fromIdx < toIdx) toIdx--;
                state.blocks.splice(toIdx, 0, moved);
                renderCanvas();
                selectBlock(moved.id);
                toast('Block moved', 'info');
            } else if (blockType) {
                addBlock(blockType, parseInt(dz.dataset.dropIndex));
            }
        });
        return dz;
    }

    // ============ Block Operations ============
    function addBlock(type, atIndex) {
        if (!BLOCK_DEFAULTS[type]) return;
        saveUndo();
        const blockData = BLOCK_DEFAULTS[type]();
        blockData.id = nextId();
        if (atIndex !== undefined && atIndex <= state.blocks.length) {
            state.blocks.splice(atIndex, 0, blockData);
        } else {
            state.blocks.push(blockData);
        }
        renderCanvas();
        selectBlock(blockData.id);
        toast(`${type.replace(/-/g, ' ')} added`, 'info');
    }

    function handleBlockAction(action, blockId) {
        const idx = state.blocks.findIndex(b => b.id === blockId);
        if (idx === -1) return;
        saveUndo();

        switch (action) {
            case 'moveUp':
                if (idx > 0) {
                    [state.blocks[idx], state.blocks[idx - 1]] = [state.blocks[idx - 1], state.blocks[idx]];
                    renderCanvas();
                }
                break;
            case 'moveDown':
                if (idx < state.blocks.length - 1) {
                    [state.blocks[idx], state.blocks[idx + 1]] = [state.blocks[idx + 1], state.blocks[idx]];
                    renderCanvas();
                }
                break;
            case 'duplicate':
                const clone = JSON.parse(JSON.stringify(state.blocks[idx]));
                clone.id = nextId();
                state.blocks.splice(idx + 1, 0, clone);
                renderCanvas();
                selectBlock(clone.id);
                toast('Block duplicated', 'info');
                break;
            case 'delete':
                state.blocks.splice(idx, 1);
                if (state.selectedBlockId === blockId) state.selectedBlockId = null;
                renderCanvas();
                toast('Block removed', 'info');
                break;
        }
    }

    function selectBlock(blockId) {
        state.selectedBlockId = blockId;
        $$('.nl-block').forEach(el => el.classList.toggle('selected', el.id === blockId));
    }

    // ============ Image Handling ============
    function handleImageDrop(file, evt) {
        if (!file.type.startsWith('image/')) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            const dataUrl = e.target.result;
            // If dropped on a specific block, try to set its image
            if (evt) {
                const blockEl = evt.target.closest('.nl-block');
                if (blockEl) {
                    const block = state.blocks.find(b => b.id === blockEl.id);
                    if (block) {
                        if (block.type === 'hero') { block.bgImage = dataUrl; }
                        else if (block.type === 'image-text') { block.imageUrl = dataUrl; }
                        else { addImageTextBlock(dataUrl); return; }
                        renderCanvas();
                        toast('Image added!', 'success');
                        return;
                    }
                }
            }
            // Otherwise add a new image-text block
            addImageTextBlock(dataUrl);
        };
        reader.readAsDataURL(file);
    }

    function addImageTextBlock(dataUrl) {
        saveUndo();
        const block = BLOCK_DEFAULTS['image-text']();
        block.id = nextId();
        block.imageUrl = dataUrl;
        block.title = 'Image Title';
        state.blocks.push(block);
        renderCanvas();
        selectBlock(block.id);
        toast('Image block added!', 'success');
    }

    // ============ Sync Editable Content ============
    function syncEditableToState(blockId, wrapperEl) {
        const block = state.blocks.find(b => b.id === blockId);
        if (!block) return;
        wrapperEl.querySelectorAll('[data-field]').forEach(el => {
            const field = el.dataset.field;
            const value = el.innerHTML;
            setNestedField(block, field, value);
        });
        saveToLocalStorage();
    }

    function setNestedField(obj, path, value) {
        const parts = path.split('.');
        let current = obj;
        for (let i = 0; i < parts.length - 1; i++) {
            const key = isNaN(parts[i]) ? parts[i] : parseInt(parts[i]);
            if (current[key] === undefined) return;
            current = current[key];
        }
        const lastKey = isNaN(parts[parts.length - 1]) ? parts[parts.length - 1] : parseInt(parts[parts.length - 1]);
        current[lastKey] = value;
    }

    // ============ Rich Text Toolbar ============
    function showRichToolbar(el) {
        const toolbar = $('#richToolbar');
        if (state.isPreview) return;
        const rect = el.getBoundingClientRect();
        toolbar.style.display = 'flex';
        toolbar.style.top = Math.max(60, rect.top - 44) + 'px';
        toolbar.style.left = Math.max(10, rect.left) + 'px';
    }

    function hideRichToolbar() {
        $('#richToolbar').style.display = 'none';
    }

    // ============ Undo / Redo ============
    function saveUndo() {
        state.undoStack.push(JSON.stringify(state.blocks));
        if (state.undoStack.length > 50) state.undoStack.shift();
        state.redoStack = [];
    }

    function undo() {
        if (state.undoStack.length === 0) return;
        state.redoStack.push(JSON.stringify(state.blocks));
        state.blocks = JSON.parse(state.undoStack.pop());
        renderCanvas();
        toast('Undone', 'info');
    }

    function redo() {
        if (state.redoStack.length === 0) return;
        state.undoStack.push(JSON.stringify(state.blocks));
        state.blocks = JSON.parse(state.redoStack.pop());
        renderCanvas();
        toast('Redone', 'info');
    }

    // ============ Fonts ============
    function applyFonts() {
        canvas.style.setProperty('--nl-heading-font', state.fonts.heading);
        canvas.style.setProperty('--nl-body-font', state.fonts.body);
        canvas.style.setProperty('--nl-base-size', state.fonts.baseSize + 'px');
    }

    // ============ Zoom ============
    function setZoom(z) {
        z = Math.max(50, Math.min(200, z));
        state.zoom = z;
        canvas.style.transform = `scale(${z / 100})`;
        canvas.style.transformOrigin = 'top center';
        $('#zoomLevel').textContent = z + '%';
    }

    // ============ Export ============
    function exportNewsletter(format) {
        // Sync all editable content
        state.blocks.forEach(block => {
            const el = document.getElementById(block.id);
            if (el) syncEditableToState(block.id, el);
        });

        switch (format) {
            case 'html-standalone': downloadStandaloneHTML(); break;
            case 'html-sharepoint': downloadSharePointHTML(); break;
            case 'html-email': downloadEmailHTML(); break;
            case 'pdf': downloadAsPDF(); break;
            case 'png': downloadAsImage(); break;
            case 'clipboard': copyHTMLToClipboard(); break;
            case 'json': downloadJSON(); break;
        }
    }

    function generateCleanHTML(mode) {
        const c = state.customColors;
        const f = state.fonts;
        let css = `
            body { margin: 0; padding: 0; font-family: ${f.body}; font-size: ${f.baseSize}px; color: ${c.text}; background: #f0f0f0; }
            .newsletter { max-width: 800px; margin: 0 auto; background: ${c.bg}; }
            .nl-hero { padding: 56px 40px 48px; text-align: center; background: linear-gradient(135deg, ${c.primary}, ${c.accent}); color: #fff; min-height: 220px; display: flex; flex-direction: column; align-items: center; justify-content: center; background-size: cover; background-position: center; }
            .nl-hero h1 { font-family: ${f.heading}; font-size: 2.4em; font-weight: 800; margin: 0 0 10px; line-height: 1.15; text-shadow: 0 2px 8px rgba(0,0,0,0.3); }
            .nl-hero .hero-subtitle { font-size: 1.1em; opacity: 0.92; font-weight: 300; text-shadow: 0 1px 4px rgba(0,0,0,0.2); }
            .nl-hero .hero-meta-badge { display: inline-block; background: ${c.accent}; color: #fff; padding: 5px 18px; border-radius: 4px; font-size: 0.72em; font-weight: 700; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 16px; }
            .nl-hero .hero-date { margin-top: 14px; font-size: 0.85em; opacity: 0.8; text-shadow: 0 1px 3px rgba(0,0,0,0.2); }
            .nl-hero .hero-img-upload { display: none; }
            .nl-editorial { padding: 12px 40px; }
            .nl-editorial .editorial-inner { background: ${c.cardBg}; border-radius: 10px; padding: 32px 36px; border-left: 4px solid ${c.accent}; font-size: 0.95em; line-height: 1.75; }
            .nl-editorial .editorial-greeting { font-weight: 600; color: ${c.primary}; margin-bottom: 12px; font-size: 1.05em; }
            .nl-editorial .editorial-content { color: ${c.text}; margin-bottom: 20px; }
            .nl-editorial .editorial-content p { margin: 0 0 10px; }
            .nl-editorial .editorial-signoff { display: flex; flex-direction: column; gap: 2px; }
            .nl-editorial .editorial-signoff span { color: #888; font-style: italic; }
            .nl-editorial .editorial-signoff strong { color: ${c.primary}; font-size: 1em; }
            .nl-editorial .editorial-signoff em { color: #999; font-size: 0.85em; }
            .nl-section-header { padding: 24px 40px 8px; }
            .nl-section-header h2 { font-family: ${f.heading}; font-size: 1.5em; font-weight: 700; color: ${c.primary}; border-bottom: 3px solid ${c.accent}; display: inline-block; padding-bottom: 6px; margin: 0; }
            .nl-divider { padding: 16px 40px; }
            .nl-divider hr { border: none; height: 1px; background: #ddd; }
            .nl-spacer { height: 32px; }
            .nl-text { padding: 12px 40px; line-height: 1.7; }
            .nl-image-text { display: flex; gap: 24px; padding: 16px 40px; align-items: flex-start; }
            .nl-image-text .img-placeholder { width: 200px; min-height: 140px; background: #e2e8f0; border-radius: 8px; flex-shrink: 0; display: flex; align-items: center; justify-content: center; font-size: 32px; color: #94a3b8; }
            .nl-image-text .img-placeholder img { width: 100%; height: auto; border-radius: 8px; }
            .nl-image-text .text-content { flex: 1; }
            .nl-image-text .text-content h3 { font-family: ${f.heading}; color: ${c.primary}; margin: 0 0 8px; font-size: 1.2em; }
            .nl-quote { padding: 20px; margin: 8px 40px; border-left: 4px solid ${c.accent}; background: ${c.cardBg}; border-radius: 0 8px 8px 0; }
            .nl-quote blockquote { font-style: italic; margin: 0 0 8px; line-height: 1.6; }
            .nl-quote cite { font-size: 0.85em; color: ${c.primary}; font-style: normal; font-weight: 600; }
            .nl-callout { margin: 8px 40px; padding: 20px 24px; background: #fef3c7; border-radius: 8px; border-left: 4px solid #f59e0b; }
            .nl-callout .callout-title { font-weight: 700; color: #92400e; margin-bottom: 6px; }
            .nl-callout .callout-body { color: #78350f; font-size: 0.9em; line-height: 1.6; }
            .nl-kpi-cards { display: flex; flex-wrap: wrap; gap: 16px; padding: 16px 40px; }
            .kpi-card { background: ${c.cardBg}; border-radius: 8px; padding: 20px; text-align: center; border-top: 4px solid ${c.primary}; flex: 1; min-width: 140px; }
            .kpi-card .kpi-value { font-family: ${f.heading}; font-size: 2em; font-weight: 800; color: ${c.primary}; }
            .kpi-card .kpi-label { font-size: 0.8em; margin-top: 6px; text-transform: uppercase; letter-spacing: 0.5px; }
            .kpi-card .kpi-change { font-size: 0.75em; margin-top: 4px; font-weight: 600; }
            .kpi-card .kpi-change.up { color: #16a34a; }
            .kpi-card .kpi-change.down { color: #dc2626; }
            .nl-progress-bars { padding: 16px 40px; }
            .progress-item { margin-bottom: 14px; }
            .progress-label { display: flex; justify-content: space-between; margin-bottom: 4px; font-size: 0.85em; }
            .progress-label span:last-child { font-weight: 700; color: ${c.primary}; }
            .progress-track { height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden; }
            .progress-fill { height: 100%; border-radius: 5px; }
            .nl-stats-row { display: flex; padding: 0 40px; margin: 8px 0; }
            .stat-item { text-align: center; padding: 20px 12px; flex: 1; border-right: 1px solid #e2e8f0; }
            .stat-item:last-child { border-right: none; }
            .stat-item .stat-value { font-family: ${f.heading}; font-size: 1.8em; font-weight: 800; color: ${c.primary}; }
            .stat-item .stat-label { font-size: 0.75em; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.5px; }
            .nl-timeline { padding: 16px 40px 16px 60px; position: relative; }
            .timeline-line { position: absolute; left: 60px; top: 16px; bottom: 16px; width: 2px; background: ${c.accent}; }
            .timeline-item { position: relative; margin-bottom: 24px; padding-left: 24px; }
            .timeline-dot { position: absolute; left: -6px; top: 4px; width: 12px; height: 12px; background: ${c.accent}; border-radius: 50%; border: 2px solid #fff; }
            .timeline-item .tl-date { font-size: 0.75em; color: ${c.accent}; font-weight: 700; text-transform: uppercase; }
            .timeline-item .tl-title { font-family: ${f.heading}; font-weight: 600; color: ${c.primary}; margin: 2px 0; }
            .timeline-item .tl-desc { font-size: 0.85em; line-height: 1.5; }
            .nl-success-story { margin: 8px 40px; border-radius: 8px; background: linear-gradient(135deg, ${c.primary}, ${c.accent}); color: #fff; padding: 28px 32px; }
            .nl-success-story .story-badge { font-size: 0.7em; text-transform: uppercase; letter-spacing: 2px; opacity: 0.8; margin-bottom: 8px; }
            .nl-success-story h3 { font-family: ${f.heading}; font-size: 1.3em; margin: 0 0 10px; }
            .nl-success-story p { font-size: 0.9em; line-height: 1.6; opacity: 0.9; margin: 0; }
            .nl-team-spotlight { display: flex; gap: 20px; padding: 16px 40px; flex-wrap: wrap; justify-content: center; }
            .team-member { text-align: center; width: 120px; }
            .team-avatar { width: 72px; height: 72px; border-radius: 50%; background: ${c.cardBg}; margin: 0 auto 8px; display: flex; align-items: center; justify-content: center; font-size: 28px; color: ${c.primary}; border: 3px solid ${c.accent}; }
            .team-member .member-name { font-weight: 600; font-size: 0.85em; }
            .team-member .member-role { font-size: 0.7em; color: #666; }
            .nl-two-column { padding: 12px 40px; display: flex; gap: 20px; }
            .nl-three-column { padding: 12px 40px; display: flex; gap: 20px; }
            .column-cell { background: ${c.cardBg}; border-radius: 8px; padding: 20px; flex: 1; }
            .column-cell h3 { font-family: ${f.heading}; color: ${c.primary}; margin: 0 0 8px; font-size: 1.1em; }
            .nl-four-column { padding: 12px 40px; display: flex; gap: 16px; }
            .nl-four-column .column-cell { flex: 1; min-width: 0; }
            .nl-sidebar-layout { padding: 12px 40px; display: flex; gap: 24px; }
            .nl-sidebar-layout .sidebar-main { flex: 2; background: ${c.cardBg}; border-radius: 8px; padding: 24px; }
            .nl-sidebar-layout .sidebar-main h3 { font-family: ${f.heading}; color: ${c.primary}; margin: 0 0 10px; font-size: 1.2em; }
            .nl-sidebar-layout .sidebar-side { flex: 1; background: ${c.cardBg}; border-radius: 8px; padding: 20px; border-left: 3px solid ${c.accent}; }
            .nl-sidebar-layout .sidebar-side h4 { font-family: ${f.heading}; color: ${c.accent}; margin: 0 0 8px; font-size: 1em; }
            .nl-footer { background: #fafafa; color: #999; padding: 28px 40px; text-align: center; margin-top: 8px; border-top: 1px solid #eee; }
            .nl-footer .footer-top { font-size: 0.85em; color: #888; margin-bottom: 8px; }
            .nl-footer .footer-copy { color: #777; }
            .nl-footer .footer-dot { margin: 0 4px; color: #ccc; }
            .nl-footer .footer-tagline { color: #aaa; font-style: italic; }
            .nl-footer .footer-bottom { font-size: 0.8em; color: #aaa; }
            .nl-footer .footer-author { color: ${c.primary}; text-decoration: none; font-weight: 600; }
            .nl-footer .footer-bottom a { color: ${c.accent}; text-decoration: none; margin: 0 2px; }
            .nl-footer-minimal { padding: 16px 40px; text-align: center; font-size: 0.8em; display: flex; justify-content: center; align-items: center; gap: 4px; flex-wrap: wrap; }
            .nl-footer-social { text-align: center; }
            .footer-social-icons { display: flex; justify-content: center; gap: 12px; margin-bottom: 10px; }
            .social-link { font-size: 1.3em; text-decoration: none; }
        `;

        if (mode === 'email') {
            // For email, we use inline styles (simplified)
            css += `
                @media only screen and (max-width: 600px) {
                    .nl-image-text, .nl-two-column, .nl-three-column, .nl-four-column, .nl-sidebar-layout, .nl-stats-row, .nl-kpi-cards { flex-direction: column !important; }
                    .kpi-card, .stat-item, .column-cell { min-width: 100% !important; }
                }
            `;
        }

        let bodyHTML = '';
        state.blocks.forEach(block => {
            // Get clean HTML without contenteditable
            let html = getBlockHTML(block);
            html = html.replace(/\s*contenteditable="true"/g, '');
            html = html.replace(/\s*data-field="[^"]*"/g, '');
            html = html.replace(/\s*data-placeholder="[^"]*"/g, '');
            bodyHTML += html;
        });

        const title = state.info.title || 'Newsletter';

        if (mode === 'sharepoint') {
            return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>${esc(title)}</title>
    <style>${css}</style>
</head>
<body>
    <!--[if mso]><table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%"><tr><td><![endif]-->
    <div class="newsletter" role="article" aria-label="${esc(title)}">
        ${bodyHTML}
    </div>
    <!--[if mso]></td></tr></table><![endif]-->
</body>
</html>`;
        }

        return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${esc(title)}</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <style>${css}</style>
</head>
<body>
    <div class="newsletter">
        ${bodyHTML}
    </div>
</body>
</html>`;
    }

    function downloadStandaloneHTML() {
        const html = generateCleanHTML('standalone');
        downloadFile(html, `${sanitizeFilename(state.info.title)}.html`, 'text/html');
        toast('Standalone HTML downloaded!', 'success');
    }

    function downloadSharePointHTML() {
        const html = generateCleanHTML('sharepoint');
        downloadFile(html, `${sanitizeFilename(state.info.title)}-sharepoint.html`, 'text/html');
        toast('SharePoint-ready HTML downloaded! You can import this into SharePoint via "Embed" web part.', 'success');
    }

    function downloadEmailHTML() {
        const html = generateCleanHTML('email');
        downloadFile(html, `${sanitizeFilename(state.info.title)}-email.html`, 'text/html');
        toast('Email HTML downloaded!', 'success');
    }

    function downloadAsPDF() {
        const html = generateCleanHTML('standalone');
        const printWindow = window.open('', '_blank');
        if (!printWindow) {
            toast('Pop-up blocked! Please allow pop-ups and try again.', 'error');
            return;
        }
        printWindow.document.write(html);
        printWindow.document.close();
        // Add print-specific styles
        const printStyle = printWindow.document.createElement('style');
        printStyle.textContent = `
            @media print {
                body { background: #fff !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                .newsletter { box-shadow: none !important; max-width: 100% !important; }
            }
            @page { margin: 0.5cm; size: A4; }
        `;
        printWindow.document.head.appendChild(printStyle);
        // Wait for fonts and images to load
        printWindow.onload = () => {
            setTimeout(() => {
                printWindow.print();
                // Don't auto-close — user may cancel the print dialog
            }, 600);
        };
        // Fallback if onload doesn't fire (some browsers)
        setTimeout(() => {
            try { printWindow.print(); } catch(e) {}
        }, 1500);
        toast('Print dialog opened — choose "Save as PDF" to download.', 'info');
    }

    function stripHTML(html) {
        const div = document.createElement('div');
        div.innerHTML = html;
        return div.textContent || div.innerText || '';
    }

    // ============ Share Actions ============
    function handleShare(action) {
        switch (action) {
            case 'preview': sharePreview(); break;
            case 'copy-link': shareCopyLink(); break;
            case 'send-email': openEmailModal(); break;
        }
    }

    function sharePreview() {
        const html = generateCleanHTML('standalone');
        const blob = new Blob([html], { type: 'text/html' });
        const url = URL.createObjectURL(blob);
        window.open(url, '_blank');
        toast('Preview opened in new tab', 'success');
    }

    function shareCopyLink() {
        const pageUrl = window.location.href;
        navigator.clipboard.writeText(pageUrl).then(() => {
            toast('Link copied to clipboard!', 'success');
        }).catch(() => {
            toast('Could not copy link', 'error');
        });
    }

    function openEmailModal() {
        const t = THEMES[state.theme];
        $('#emailPreviewTitle').textContent = state.info.title || 'Newsletter';
        $('#emailPreviewMeta').textContent = `${state.info.edition || ''} \u00b7 ${state.info.quarter || ''}`;
        if (t) $('#emailPreviewThumb').style.background = t.gradient;
        $('#emailTo').value = '';
        $('#emailCc').value = '';
        $('#emailBcc').value = '';
        $('#emailMessage').value = '';
        openModal('emailModal');
    }

    function sendEmail(isTest) {
        const to = $('#emailTo').value.trim();
        const cc = $('#emailCc').value.trim();
        const bcc = $('#emailBcc').value.trim();
        const message = $('#emailMessage').value.trim();
        const includeHTML = $('#emailAsLink').checked;
        const title = state.info.title || 'Newsletter';

        if (!to && !isTest) {
            toast('Please enter at least one email address', 'error');
            return;
        }

        let plainBody = '';
        if (message) plainBody += message + '\n\n---\n\n';

        state.blocks.forEach(b => {
            switch(b.type) {
                case 'hero': plainBody += `${b.title}\n${b.subtitle || ''}\n${b.meta || ''}\n\n`; break;
                case 'editorial': plainBody += `${b.greeting || ''}\n${stripHTML(b.content || '')}\n${b.signoff || ''} ${b.authorName || ''}\n\n`; break;
                case 'section-header': plainBody += `--- ${b.title} ---\n\n`; break;
                case 'text': plainBody += `${stripHTML(b.content || '')}\n\n`; break;
                case 'quote': plainBody += `"${b.text}" \u2014 ${b.author}\n\n`; break;
                case 'callout': plainBody += `${b.title}: ${b.content}\n\n`; break;
                case 'success-story': plainBody += `${b.badge || ''} ${b.title}: ${b.content}\n\n`; break;
                case 'kpi-cards': b.cards.forEach(c => { plainBody += `${c.label}: ${c.value} (${c.change})\n`; }); plainBody += '\n'; break;
                case 'stats-row': b.stats.forEach(s => { plainBody += `${s.label}: ${s.value}  `; }); plainBody += '\n\n'; break;
                case 'timeline': b.items.forEach(t => { plainBody += `${t.date}: ${t.title} - ${t.desc}\n`; }); plainBody += '\n'; break;
                default: break;
            }
        });

        if (includeHTML) {
            const html = generateCleanHTML('email');
            navigator.clipboard.writeText(html).then(() => {}).catch(() => {});
        }

        if (plainBody.length > 1800) {
            plainBody = plainBody.substring(0, 1800) + '\n\n[Newsletter HTML has been copied to your clipboard \u2014 paste in the email for rich formatting]';
        }

        const subject = encodeURIComponent(title);
        const body = encodeURIComponent(plainBody);
        const toAddr = isTest ? '' : encodeURIComponent(to);
        let mailto = `mailto:${toAddr}?subject=${subject}&body=${body}`;
        if (cc) mailto += `&cc=${encodeURIComponent(cc)}`;
        if (bcc) mailto += `&bcc=${encodeURIComponent(bcc)}`;

        window.location.href = mailto;
        closeModal('emailModal');

        if (includeHTML) {
            toast('Mail client opened! Newsletter HTML copied to clipboard \u2014 paste for rich formatting.', 'success');
        } else {
            toast('Mail client opened with newsletter summary!', 'success');
        }
    }

    function copyHTMLToClipboard() {
        const html = generateCleanHTML('standalone');
        navigator.clipboard.writeText(html).then(() => {
            toast('HTML copied to clipboard!', 'success');
        }).catch(() => {
            // Fallback
            const ta = document.createElement('textarea');
            ta.value = html;
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
            toast('HTML copied to clipboard!', 'success');
        });
    }

    function downloadJSON() {
        const data = {
            version: 1,
            title: state.info.title,
            info: state.info,
            theme: state.theme,
            customColors: state.customColors,
            fonts: state.fonts,
            blocks: state.blocks
        };
        downloadFile(JSON.stringify(data, null, 2), `${sanitizeFilename(state.info.title)}.json`, 'application/json');
        toast('JSON template exported!', 'success');
    }

    function downloadFile(content, filename, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    function sanitizeFilename(name) {
        return (name || 'newsletter').replace(/[^a-zA-Z0-9-_ ]/g, '').replace(/\s+/g, '-').substring(0, 50);
    }

    // ============ Save / Load ============
    function saveToLocalStorage() {
        try {
            const data = {
                version: 6,
                blocks: state.blocks,
                info: state.info,
                theme: state.theme,
                customColors: state.customColors,
                fonts: state.fonts
            };
            localStorage.setItem('artikraft-newsletter', JSON.stringify(data));
        } catch (e) { /* ignore quota errors */ }
    }

    function loadFromLocalStorage() {
        try {
            const raw = localStorage.getItem('artikraft-newsletter');
            if (!raw) return;
            const data = JSON.parse(raw);
            // Version check: clear stale data from old versions
            if (!data.version || data.version < 6) {
                localStorage.removeItem('artikraft-newsletter');
                return;
            }
            if (data.blocks && data.blocks.length > 0) {
                state.blocks = data.blocks;
                state.info = data.info || state.info;
                state.customColors = data.customColors || state.customColors;
                state.fonts = data.fonts || state.fonts;
                if (data.theme) applyTheme(data.theme);
                // Update inputs
                $('#newsletterTitle').value = state.info.title || '';
                $('#infoEdition').value = state.info.edition || '';
                $('#infoQuarter').value = state.info.quarter || '';
                $('#infoOrg').value = state.info.org || '';
                $('#fontHeading').value = state.fonts.heading;
                $('#fontBody').value = state.fonts.body;
                $('#fontSize').value = state.fonts.baseSize;
                $('#fontSizeVal').textContent = state.fonts.baseSize + 'px';
                // Restore blockId counter
                state.blocks.forEach(b => {
                    const num = parseInt((b.id || '').replace('block-', ''));
                    if (num > blockIdCounter) blockIdCounter = num;
                });
                renderCanvas();
            }
        } catch (e) { /* ignore parse errors */ }
    }

    function saveNewsletter() {
        const data = {
            version: 1,
            title: state.info.title,
            info: state.info,
            theme: state.theme,
            customColors: state.customColors,
            fonts: state.fonts,
            blocks: state.blocks
        };
        downloadFile(JSON.stringify(data, null, 2), `${sanitizeFilename(state.info.title)}.json`, 'application/json');
        toast('Newsletter saved!', 'success');
    }

    function loadNewsletter(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = JSON.parse(e.target.result);
                if (!data.blocks) throw new Error('Invalid format');
                saveUndo();
                state.blocks = data.blocks;
                state.info = data.info || state.info;
                state.customColors = data.customColors || state.customColors;
                state.fonts = data.fonts || state.fonts;
                if (data.theme) applyTheme(data.theme);
                // Update inputs
                $('#newsletterTitle').value = state.info.title || '';
                $('#infoEdition').value = state.info.edition || '';
                $('#infoQuarter').value = state.info.quarter || '';
                $('#infoOrg').value = state.info.org || '';
                // Restore blockId counter
                state.blocks.forEach(b => {
                    const num = parseInt((b.id || '').replace('block-', ''));
                    if (num > blockIdCounter) blockIdCounter = num;
                });
                renderCanvas();
                toast('Newsletter loaded!', 'success');
            } catch (err) {
                toast('Failed to load file: Invalid format', 'error');
            }
        };
        reader.readAsText(file);
    }

    // ============ Toast ============
    function toast(msg, type) {
        const container = $('#toastContainer');
        const el = document.createElement('div');
        el.className = 'toast ' + (type || 'info');
        el.textContent = msg;
        container.appendChild(el);
        setTimeout(() => el.remove(), 4000);
    }

    // ============ Modal ============
    function openModal(id) { document.getElementById(id).classList.add('show'); }
    function closeModal(id) { document.getElementById(id).classList.remove('show'); }

    // ============ Event Bindings ============
    function bindEvents() {
        // Drag from block palette
        $$('.block-item[draggable]').forEach(item => {
            item.addEventListener('dragstart', (e) => {
                e.dataTransfer.setData('text/block-type', item.dataset.block);
                e.dataTransfer.effectAllowed = 'copy';
            });
            // Click to add at end
            item.addEventListener('click', () => {
                addBlock(item.dataset.block);
            });
        });

        // Canvas dragover (for adding blocks)
        canvas.addEventListener('dragover', (e) => e.preventDefault());
        canvas.addEventListener('drop', (e) => {
            e.preventDefault();
            // Handle image drop onto canvas
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                const file = e.dataTransfer.files[0];
                if (file.type.startsWith('image/')) {
                    handleImageDrop(file, e);
                    return;
                }
            }
            const blockType = e.dataTransfer.getData('text/block-type');
            if (blockType) addBlock(blockType);
        });

        // Paste image from clipboard
        document.addEventListener('paste', (e) => {
            if (e.target.closest('[contenteditable]') || e.target.closest('input') || e.target.closest('textarea')) return;
            const items = e.clipboardData && e.clipboardData.items;
            if (!items) return;
            for (let item of items) {
                if (item.type.startsWith('image/')) {
                    e.preventDefault();
                    const file = item.getAsFile();
                    if (file) handleImageDrop(file);
                    break;
                }
            }
        });

        // Top bar buttons
        $('#btnUndo').addEventListener('click', undo);
        $('#btnRedo').addEventListener('click', redo);
        $('#btnEditMode').addEventListener('click', () => {
            state.isPreview = false;
            document.body.classList.remove('preview-mode');
            $('#btnEditMode').classList.add('active');
            $('#btnPreviewMode').classList.remove('active');
        });
        $('#btnPreviewMode').addEventListener('click', () => {
            state.isPreview = true;
            document.body.classList.add('preview-mode');
            hideRichToolbar();
            $('#btnPreviewMode').classList.add('active');
            $('#btnEditMode').classList.remove('active');
        });
        $('#btnSave').addEventListener('click', saveNewsletter);
        $('#btnLoad').addEventListener('click', () => $('#fileInput').click());
        $('#btnThemeToggle').addEventListener('click', toggleBuilderTheme);
        $('#btnSaveTheme').addEventListener('click', saveCustomTheme);
        $('#fileInput').addEventListener('change', (e) => {
            if (e.target.files[0]) loadNewsletter(e.target.files[0]);
            e.target.value = '';
        });

        // Export dropdown
        $('#btnExport').addEventListener('click', (e) => {
            e.stopPropagation();
            $('#exportMenu').classList.toggle('show');
            $('#shareMenu').classList.remove('show');
        });
        // Share dropdown
        $('#btnShare').addEventListener('click', (e) => {
            e.stopPropagation();
            $('#shareMenu').classList.toggle('show');
            $('#exportMenu').classList.remove('show');
        });
        document.addEventListener('click', () => {
            $('#exportMenu').classList.remove('show');
            $('#shareMenu').classList.remove('show');
        });
        $$('.export-option[data-format]').forEach(opt => {
            opt.addEventListener('click', (e) => {
                e.stopPropagation();
                exportNewsletter(opt.dataset.format);
                $('#exportMenu').classList.remove('show');
            });
        });
        $$('.export-option[data-share]').forEach(opt => {
            opt.addEventListener('click', (e) => {
                e.stopPropagation();
                handleShare(opt.dataset.share);
                $('#shareMenu').classList.remove('show');
            });
        });

        // Email modal
        $('#closeEmailModal').addEventListener('click', () => closeModal('emailModal'));
        $('#btnCancelEmail').addEventListener('click', () => closeModal('emailModal'));
        $('#emailModal').addEventListener('click', (e) => {
            if (e.target === e.currentTarget) closeModal('emailModal');
        });
        $('#btnShowBcc').addEventListener('click', () => {
            $('#bccField').style.display = $('#bccField').style.display === 'none' ? 'flex' : 'none';
        });
        $('#btnSendEmail').addEventListener('click', () => sendEmail(false));
        $('#btnTestEmail').addEventListener('click', () => sendEmail(true));

        // Template modal
        $('#btnTemplates').addEventListener('click', () => openModal('templateModal'));
        $('#btnStartTemplate').addEventListener('click', () => openModal('templateModal'));
        $('#btnStartScratch').addEventListener('click', () => {
            saveUndo();
            state.blocks = [];
            addBlock('hero');
            addBlock('editorial');
            addBlock('section-header');
            addBlock('text');
            addBlock('footer');
        });
        $('#closeTemplateModal').addEventListener('click', () => closeModal('templateModal'));
        $('#templateModal').addEventListener('click', (e) => {
            if (e.target === e.currentTarget) closeModal('templateModal');
        });

        // Auto-Generate modal
        $('#btnAutoGen').addEventListener('click', () => openModal('autoGenModal'));
        $('#closeAutoGenModal').addEventListener('click', () => closeModal('autoGenModal'));
        $('#autoGenModal').addEventListener('click', (e) => {
            if (e.target === e.currentTarget) closeModal('autoGenModal');
        });
        $('#btnRunAutoGen').addEventListener('click', autoGenerateFromRaw);

        // Version history
        $('#btnVersions').addEventListener('click', () => { renderVersionList(); openModal('versionModal'); });
        $('#closeVersionModal').addEventListener('click', () => closeModal('versionModal'));
        $('#versionModal').addEventListener('click', (e) => { if (e.target === e.currentTarget) closeModal('versionModal'); });
        $('#btnSaveVersion').addEventListener('click', saveVersion);

        // Import from URL
        $('#btnImportUrl').addEventListener('click', () => {
            $('#importUrlInput').value = '';
            $('#importUrlText').value = '';
            $('#importUrlContent').style.display = 'none';
            $('#btnImportUrlGenerate').style.display = 'none';
            openModal('importUrlModal');
        });
        $('#closeImportUrlModal').addEventListener('click', () => closeModal('importUrlModal'));
        $('#importUrlModal').addEventListener('click', (e) => { if (e.target === e.currentTarget) closeModal('importUrlModal'); });
        $('#btnFetchUrl').addEventListener('click', fetchUrlContent);
        $('#btnImportUrlGenerate').addEventListener('click', importUrlGenerate);

        // Responsive preview
        $$('.resp-btn').forEach(btn => {
            btn.addEventListener('click', () => setDevicePreview(btn.dataset.device));
        });

        // Background patterns
        $$('.pattern-swatch').forEach(swatch => {
            swatch.addEventListener('click', () => {
                $$('.pattern-swatch').forEach(s => s.classList.remove('active'));
                swatch.classList.add('active');
                applyBackgroundPattern(swatch.dataset.pattern);
            });
        });

        // Sidebar toggles
        $('#toggleLeft').addEventListener('click', () => {
            $('#sidebarLeft').classList.toggle('collapsed');
            $('#reopenLeft').classList.toggle('visible', $('#sidebarLeft').classList.contains('collapsed'));
        });
        $('#toggleRight').addEventListener('click', () => {
            $('#sidebarRight').classList.toggle('collapsed');
            $('#reopenRight').classList.toggle('visible', $('#sidebarRight').classList.contains('collapsed'));
        });
        $('#reopenLeft').addEventListener('click', () => {
            $('#sidebarLeft').classList.remove('collapsed');
            $('#reopenLeft').classList.remove('visible');
        });
        $('#reopenRight').addEventListener('click', () => {
            $('#sidebarRight').classList.remove('collapsed');
            $('#reopenRight').classList.remove('visible');
        });

        // Newsletter title
        $('#newsletterTitle').addEventListener('input', (e) => {
            state.info.title = e.target.value;
            saveToLocalStorage();
        });

        // Color pickers
        ['Primary', 'Accent', 'Bg', 'Text', 'CardBg'].forEach(key => {
            const inputId = 'color' + key;
            const stateKey = key.charAt(0).toLowerCase() + key.slice(1);
            $(`#${inputId}`).addEventListener('input', (e) => {
                state.customColors[stateKey] = e.target.value;
                applyCustomColors();
                saveToLocalStorage();
            });
        });

        // Font selectors
        $('#fontHeading').addEventListener('change', (e) => {
            state.fonts.heading = e.target.value;
            applyFonts();
            saveToLocalStorage();
        });
        $('#fontBody').addEventListener('change', (e) => {
            state.fonts.body = e.target.value;
            applyFonts();
            saveToLocalStorage();
        });
        $('#fontSize').addEventListener('input', (e) => {
            state.fonts.baseSize = parseInt(e.target.value);
            $('#fontSizeVal').textContent = e.target.value + 'px';
            applyFonts();
            saveToLocalStorage();
        });

        // Newsletter info
        $('#infoEdition').addEventListener('input', (e) => { state.info.edition = e.target.value; saveToLocalStorage(); });
        $('#infoQuarter').addEventListener('input', (e) => { state.info.quarter = e.target.value; saveToLocalStorage(); });
        $('#infoOrg').addEventListener('input', (e) => { state.info.org = e.target.value; saveToLocalStorage(); });

        // Canvas width
        $('#canvasWidth').addEventListener('change', (e) => {
            canvas.style.width = e.target.value + 'px';
        });

        // Zoom
        $('#zoomIn').addEventListener('click', () => setZoom(state.zoom + 10));
        $('#zoomOut').addEventListener('click', () => setZoom(state.zoom - 10));
        $('#zoomFit').addEventListener('click', () => setZoom(100));

        // Rich text toolbar commands
        $$('.rt-btn[data-cmd]').forEach(btn => {
            btn.addEventListener('mousedown', (e) => {
                e.preventDefault();
                const cmd = btn.dataset.cmd;
                if (cmd === 'createLink') {
                    const url = prompt('Enter URL:');
                    if (url) document.execCommand(cmd, false, url);
                } else {
                    document.execCommand(cmd, false, null);
                }
            });
        });
        $('#rtHeading').addEventListener('change', (e) => {
            if (e.target.value) {
                document.execCommand('formatBlock', false, e.target.value);
            } else {
                document.execCommand('formatBlock', false, 'p');
            }
            e.target.value = '';
        });
        $('#rtColor').addEventListener('input', (e) => {
            document.execCommand('foreColor', false, e.target.value);
        });

        // Click outside to deselect and hide toolbar
        document.addEventListener('click', (e) => {
            if (!e.target.closest('.nl-block') && !e.target.closest('.rich-toolbar') && !e.target.closest('.sidebar-right')) {
                state.selectedBlockId = null;
                $$('.nl-block').forEach(el => el.classList.remove('selected'));
            }
            if (!e.target.closest('[contenteditable]') && !e.target.closest('.rich-toolbar')) {
                hideRichToolbar();
            }
        });

        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            if (e.target.closest('[contenteditable]') || e.target.closest('input') || e.target.closest('select') || e.target.closest('textarea')) return;
            if (e.ctrlKey && e.key === 'z') { e.preventDefault(); undo(); }
            if (e.ctrlKey && e.key === 'y') { e.preventDefault(); redo(); }
            if (e.key === 'Delete' && state.selectedBlockId) {
                e.preventDefault();
                handleBlockAction('delete', state.selectedBlockId);
            }
        });

        // Image upload via click on placeholder or hero cover
        canvas.addEventListener('click', (e) => {
            // Hero background image
            const heroUpload = e.target.closest('.hero-img-upload');
            if (heroUpload && !state.isPreview) {
                const input = document.createElement('input');
                input.type = 'file';
                input.accept = 'image/*';
                input.addEventListener('change', () => {
                    const file = input.files[0];
                    if (!file) return;
                    const reader = new FileReader();
                    reader.onload = (ev) => {
                        const blockEl = heroUpload.closest('.nl-block');
                        if (!blockEl) return;
                        const block = state.blocks.find(b => b.id === blockEl.id);
                        if (block) {
                            block.bgImage = ev.target.result;
                            renderCanvas();
                            toast('Cover image set!', 'success');
                        }
                    };
                    reader.readAsDataURL(file);
                });
                input.click();
                return;
            }

            const imgPlaceholder = e.target.closest('.img-placeholder');
            if (!imgPlaceholder || state.isPreview) return;
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = 'image/*';
            input.addEventListener('change', () => {
                const file = input.files[0];
                if (!file) return;
                const reader = new FileReader();
                reader.onload = (ev) => {
                    const blockEl = imgPlaceholder.closest('.nl-block');
                    if (!blockEl) return;
                    const block = state.blocks.find(b => b.id === blockEl.id);
                    if (block) {
                        block.imageUrl = ev.target.result;
                        renderCanvas();
                    }
                };
                reader.readAsDataURL(file);
            });
            input.click();
        });
    }

    // ============ Auto-Generate from Raw Data ============
    function autoGenerateFromRaw() {
        const raw = $('#autoGenInput').value.trim();
        if (!raw) { toast('Please paste some content first', 'error'); return; }

        const includeHero = $('#autoGenHero').checked;
        const includeEditorial = $('#autoGenEditorial').checked;
        const includeFooter = $('#autoGenFooter').checked;

        saveUndo();
        const blocks = [];

        if (includeHero) {
            const heroBlock = BLOCK_DEFAULTS['hero']();
            heroBlock.id = nextId();
            blocks.push(heroBlock);
        }

        if (includeEditorial) {
            const editBlock = BLOCK_DEFAULTS['editorial']();
            editBlock.id = nextId();
            blocks.push(editBlock);
        }

        const lines = raw.split('\n');
        let i = 0;
        let pendingText = [];
        let kpiBuffer = [];
        let timelineBuffer = [];

        function flushText() {
            if (pendingText.length === 0) return;
            const textBlock = BLOCK_DEFAULTS['text']();
            textBlock.id = nextId();
            textBlock.content = pendingText.map(t => `<p>${escForGen(t)}</p>`).join('');
            blocks.push(textBlock);
            pendingText = [];
        }

        function flushKPIs() {
            if (kpiBuffer.length === 0) return;
            const kpiBlock = BLOCK_DEFAULTS['kpi-cards']();
            kpiBlock.id = nextId();
            kpiBlock.cards = kpiBuffer.slice(0, 8);
            blocks.push(kpiBlock);
            kpiBuffer = [];
        }

        function flushTimeline() {
            if (timelineBuffer.length === 0) return;
            const tlBlock = BLOCK_DEFAULTS['timeline']();
            tlBlock.id = nextId();
            tlBlock.items = timelineBuffer;
            blocks.push(tlBlock);
            timelineBuffer = [];
        }

        while (i < lines.length) {
            const line = lines[i].trim();

            // Empty line — skip
            if (!line) { i++; continue; }

            // ## Section Header
            if (/^#{1,3}\s+(.+)/.test(line)) {
                flushText(); flushKPIs(); flushTimeline();
                const match = line.match(/^#{1,3}\s+(.+)/);
                const shBlock = BLOCK_DEFAULTS['section-header']();
                shBlock.id = nextId();
                shBlock.title = match[1].trim();
                blocks.push(shBlock);
                i++; continue;
            }

            // KPI: value label (change)
            if (/^KPI:\s*/i.test(line)) {
                flushText(); flushTimeline();
                const kpiText = line.replace(/^KPI:\s*/i, '');
                const kpiMatch = kpiText.match(/^([^\s]+)\s+(.+?)(?:\s*\(([^)]+)\))?$/);
                if (kpiMatch) {
                    const change = kpiMatch[3] || '';
                    kpiBuffer.push({
                        value: kpiMatch[1],
                        label: kpiMatch[2].trim(),
                        change: change,
                        direction: change.startsWith('-') ? 'down' : 'up'
                    });
                } else {
                    kpiBuffer.push({ value: kpiText.split(' ')[0], label: kpiText.split(' ').slice(1).join(' '), change: '', direction: 'up' });
                }
                i++; continue;
            }

            // > "Quote" — Author
            if (/^>\s*/.test(line)) {
                flushText(); flushKPIs(); flushTimeline();
                const quoteText = line.replace(/^>\s*/, '');
                const qMatch = quoteText.match(/^["\u201c]?(.+?)["\u201d]?\s*[\u2014\-\u2013]\s*(.+)$/);
                const qBlock = BLOCK_DEFAULTS['quote']();
                qBlock.id = nextId();
                if (qMatch) {
                    qBlock.text = qMatch[1].trim();
                    qBlock.author = qMatch[2].trim();
                } else {
                    qBlock.text = quoteText.replace(/^["\u201c]|["\u201d]$/g, '').trim();
                }
                blocks.push(qBlock);
                i++; continue;
            }

            // ⭐ Success Story
            if (/^[⭐🌟🏆✅]\s*/.test(line)) {
                flushText(); flushKPIs(); flushTimeline();
                const storyText = line.replace(/^[⭐🌟🏆✅]\s*/, '');
                const sMatch = storyText.match(/^(.+?):\s*(.+)$/);
                const sBlock = BLOCK_DEFAULTS['success-story']();
                sBlock.id = nextId();
                if (sMatch) {
                    sBlock.title = sMatch[1].trim();
                    sBlock.content = sMatch[2].trim();
                } else {
                    sBlock.title = storyText;
                }
                blocks.push(sBlock);
                i++; continue;
            }

            // 💡 Callout
            if (/^[💡⚠️📢🔔]\s*/.test(line) || /^(Callout|Note|Tip|Important|Reminder):\s*/i.test(line)) {
                flushText(); flushKPIs(); flushTimeline();
                const calloutText = line.replace(/^[💡⚠️📢🔔]\s*/, '').replace(/^(Callout|Note|Tip|Important|Reminder):\s*/i, '');
                const cMatch = calloutText.match(/^(.+?):\s*(.+)$/);
                const cBlock = BLOCK_DEFAULTS['callout']();
                cBlock.id = nextId();
                if (cMatch) {
                    cBlock.title = '💡 ' + cMatch[1].trim();
                    cBlock.content = cMatch[2].trim();
                } else {
                    cBlock.content = calloutText;
                }
                blocks.push(cBlock);
                i++; continue;
            }

            // Timeline: "Jan 2026: Something happened"
            if (/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{4}\s*:\s*/i.test(line)) {
                flushText(); flushKPIs();
                const tlMatch = line.match(/^(.+?\d{4})\s*:\s*(.+)$/);
                if (tlMatch) {
                    const descParts = tlMatch[2].match(/^(.+?)[\.\-]\s*(.+)$/);
                    timelineBuffer.push({
                        date: tlMatch[1].trim(),
                        title: descParts ? descParts[1].trim() : tlMatch[2].trim(),
                        desc: descParts ? descParts[2].trim() : ''
                    });
                }
                i++; continue;
            }

            // --- or *** as divider
            if (/^[-*_]{3,}$/.test(line)) {
                flushText(); flushKPIs(); flushTimeline();
                const divBlock = { type: 'divider', id: nextId() };
                blocks.push(divBlock);
                i++; continue;
            }

            // Regular text paragraph
            pendingText.push(line);
            i++;
        }

        // Flush remaining buffers
        flushText();
        flushKPIs();
        flushTimeline();

        if (includeFooter) {
            const fBlock = BLOCK_DEFAULTS['footer']();
            fBlock.id = nextId();
            blocks.push(fBlock);
        }

        state.blocks = blocks;
        renderCanvas();
        closeModal('autoGenModal');
        toast(`Generated ${blocks.length} blocks from your content!`, 'success');
    }

    function escForGen(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // ============ Version History ============
    function saveVersion() {
        const versions = JSON.parse(localStorage.getItem('artikraft-versions') || '[]');
        const snapshot = {
            id: Date.now(),
            timestamp: new Date().toLocaleString(),
            title: state.info.title || 'Untitled',
            blockCount: state.blocks.length,
            data: {
                blocks: state.blocks,
                info: state.info,
                theme: state.theme,
                customColors: state.customColors,
                fonts: state.fonts
            }
        };
        versions.unshift(snapshot);
        if (versions.length > 20) versions.pop();
        localStorage.setItem('artikraft-versions', JSON.stringify(versions));
        renderVersionList();
        toast('Version saved!', 'success');
    }

    function renderVersionList() {
        const list = $('#versionList');
        const versions = JSON.parse(localStorage.getItem('artikraft-versions') || '[]');
        if (versions.length === 0) {
            list.innerHTML = '<p class="version-empty">No saved versions yet. Click "Save Version" to create a snapshot.</p>';
            return;
        }
        list.innerHTML = versions.map(v => `
            <div class="version-item" data-vid="${v.id}">
                <div class="version-info">
                    <strong>${esc(v.title)}</strong>
                    <small>${v.timestamp} · ${v.blockCount} blocks</small>
                </div>
                <div class="version-btns">
                    <button class="version-restore" data-vid="${v.id}">Restore</button>
                    <button class="version-delete" data-vid="${v.id}">✕</button>
                </div>
            </div>
        `).join('');
        list.querySelectorAll('.version-restore').forEach(btn => {
            btn.addEventListener('click', () => restoreVersion(parseInt(btn.dataset.vid)));
        });
        list.querySelectorAll('.version-delete').forEach(btn => {
            btn.addEventListener('click', () => deleteVersion(parseInt(btn.dataset.vid)));
        });
    }

    function restoreVersion(id) {
        const versions = JSON.parse(localStorage.getItem('artikraft-versions') || '[]');
        const v = versions.find(v => v.id === id);
        if (!v) return;
        saveUndo();
        state.blocks = v.data.blocks;
        state.info = v.data.info || state.info;
        state.customColors = v.data.customColors || state.customColors;
        state.fonts = v.data.fonts || state.fonts;
        if (v.data.theme) applyTheme(v.data.theme);
        $('#newsletterTitle').value = state.info.title || '';
        $('#infoEdition').value = state.info.edition || '';
        $('#infoQuarter').value = state.info.quarter || '';
        $('#infoOrg').value = state.info.org || '';
        renderCanvas();
        closeModal('versionModal');
        toast('Version restored!', 'success');
    }

    function deleteVersion(id) {
        let versions = JSON.parse(localStorage.getItem('artikraft-versions') || '[]');
        versions = versions.filter(v => v.id !== id);
        localStorage.setItem('artikraft-versions', JSON.stringify(versions));
        renderVersionList();
        toast('Version deleted', 'info');
    }

    // ============ Responsive Preview ============
    function setDevicePreview(device) {
        const widths = { desktop: 800, tablet: 480, mobile: 375 };
        canvas.style.width = widths[device] + 'px';
        $$('.resp-btn').forEach(b => b.classList.toggle('active', b.dataset.device === device));
    }

    // ============ Image Export ============
    function downloadAsImage() {
        toast('Rendering image...', 'info');
        const html = generateCleanHTML('standalone');
        const iframe = document.createElement('iframe');
        iframe.style.cssText = 'position:fixed;left:-9999px;width:800px;height:2000px;border:none';
        document.body.appendChild(iframe);
        iframe.contentDocument.open();
        iframe.contentDocument.write(html);
        iframe.contentDocument.close();

        iframe.onload = () => {
            setTimeout(() => {
                try {
                    const body = iframe.contentDocument.body;
                    const newsletter = iframe.contentDocument.querySelector('.newsletter') || body;
                    // Use canvas-based rendering
                    htmlToCanvas(newsletter).then(dataUrl => {
                        const a = document.createElement('a');
                        a.href = dataUrl;
                        a.download = sanitizeFilename(state.info.title) + '.png';
                        a.click();
                        document.body.removeChild(iframe);
                        toast('Image downloaded!', 'success');
                    }).catch(() => {
                        // Fallback: open print dialog
                        document.body.removeChild(iframe);
                        downloadAsPDF();
                    });
                } catch (e) {
                    document.body.removeChild(iframe);
                    downloadAsPDF();
                }
            }, 1000);
        };
    }

    function htmlToCanvas(element) {
        return new Promise((resolve, reject) => {
            try {
                const rect = element.getBoundingClientRect();
                const canvasEl = document.createElement('canvas');
                const scale = 2;
                canvasEl.width = rect.width * scale;
                canvasEl.height = rect.height * scale;
                const ctx = canvasEl.getContext('2d');
                ctx.scale(scale, scale);

                const svgData = `<svg xmlns="http://www.w3.org/2000/svg" width="${rect.width}" height="${rect.height}">
                    <foreignObject width="100%" height="100%">
                        <div xmlns="http://www.w3.org/1999/xhtml">${element.outerHTML}</div>
                    </foreignObject>
                </svg>`;

                const img = new Image();
                img.onload = () => {
                    ctx.drawImage(img, 0, 0);
                    resolve(canvasEl.toDataURL('image/png'));
                };
                img.onerror = reject;
                img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svgData);
            } catch (e) { reject(e); }
        });
    }

    // ============ Import from URL ============
    function fetchUrlContent() {
        const url = $('#importUrlInput').value.trim();
        if (!url) { toast('Please enter a URL', 'error'); return; }
        
        // Validate URL format
        try { new URL(url); } catch (e) { toast('Invalid URL format', 'error'); return; }

        toast('Fetching content...', 'info');
        const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`;
        
        fetch(proxyUrl)
            .then(r => {
                if (!r.ok) throw new Error('Failed to fetch');
                return r.text();
            })
            .then(html => {
                const parser = new DOMParser();
                const doc = parser.parseFromString(html, 'text/html');
                // Extract main content
                const selectors = ['article', 'main', '.content', '.post-content', '.entry-content', '#content'];
                let content = null;
                for (const sel of selectors) {
                    content = doc.querySelector(sel);
                    if (content) break;
                }
                if (!content) content = doc.body;
                
                // Clean up: remove scripts, styles, nav, footer
                content.querySelectorAll('script,style,nav,footer,iframe,form,aside,.ad,.ads,.sidebar').forEach(el => el.remove());
                
                const text = content.innerText || content.textContent || '';
                const cleaned = text.replace(/\n{3,}/g, '\n\n').trim();
                
                if (!cleaned) { toast('No readable content found', 'error'); return; }
                
                $('#importUrlText').value = cleaned;
                $('#importUrlContent').style.display = 'block';
                $('#btnImportUrlGenerate').style.display = 'inline-flex';
                toast('Content extracted! Review and click Generate.', 'success');
            })
            .catch(() => {
                toast('Could not fetch URL. Try copying the content manually into Auto-Generate.', 'error');
            });
    }

    function importUrlGenerate() {
        const text = $('#importUrlText').value.trim();
        if (!text) { toast('No content to generate from', 'error'); return; }
        closeModal('importUrlModal');
        // Reuse auto-generate logic
        $('#autoGenInput').value = text;
        openModal('autoGenModal');
        toast('Content loaded! Click Generate Newsletter to proceed.', 'info');
    }

    // ============ Background Pattern ============
    function applyBackgroundPattern(pattern) {
        const patterns = {
            'none': 'none',
            'dots': 'radial-gradient(circle, rgba(0,0,0,0.06) 1px, transparent 1px)',
            'lines': 'repeating-linear-gradient(0deg, transparent, transparent 20px, rgba(0,0,0,0.04) 20px, rgba(0,0,0,0.04) 21px)',
            'diagonal': 'repeating-linear-gradient(45deg, transparent, transparent 14px, rgba(0,0,0,0.03) 14px, rgba(0,0,0,0.03) 15px)',
            'grid': 'linear-gradient(rgba(0,0,0,0.04) 1px, transparent 1px), linear-gradient(90deg, rgba(0,0,0,0.04) 1px, transparent 1px)',
            'gradient-subtle': `linear-gradient(135deg, ${state.customColors.bg}, ${lightenColor(state.customColors.primary, 0.92)})`
        };
        canvas.style.backgroundImage = patterns[pattern] || 'none';
        if (pattern === 'dots') canvas.style.backgroundSize = '20px 20px';
        else if (pattern === 'grid') canvas.style.backgroundSize = '24px 24px';
        else canvas.style.backgroundSize = '';
        state.bgPattern = pattern;
        saveToLocalStorage();
    }

    function lightenColor(hex, amount) {
        const num = parseInt(hex.replace('#', ''), 16);
        const r = Math.round(Math.min(255, ((num >> 16) & 0xFF) + (255 - ((num >> 16) & 0xFF)) * amount));
        const g = Math.round(Math.min(255, ((num >> 8) & 0xFF) + (255 - ((num >> 8) & 0xFF)) * amount));
        const b = Math.round(Math.min(255, (num & 0xFF) + (255 - (num & 0xFF)) * amount));
        return `rgb(${r},${g},${b})`;
    }

    // ============ Boot ============
    document.addEventListener('DOMContentLoaded', init);

})();
