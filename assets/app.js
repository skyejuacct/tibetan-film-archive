/*
 * 静态版档案站：读 data/tibetan-film.xlsx，在浏览器里生成全部页面。
 * 内容（站名、文字、配色、影片、人物……）全部来自 Excel，这个文件里不含任何会变的内容。
 * 页面地址是 #/films、#/film/FILM-0001 这种形式，所以托管在 GitHub Pages 上不需要服务器配置。
 */
(function () {
    'use strict';

    var DATA_URL = 'data/tibetan-film.xlsx';
    var IMAGE_EXTS = ['jpg', 'jpeg', 'png', 'webp'];

    var db = {};          // 解析后的数据
    var settings = {};

    // ---------- 小工具 ----------

    function esc(s) {
        return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
            return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
        });
    }

    function str(v) { return v == null ? '' : String(v).trim(); }

    function split(v) {
        return str(v).split(';').map(function (x) { return x.trim(); }).filter(Boolean);
    }

    function key(name) { return str(name).toLowerCase().replace(/\s+/g, ' '); }

    function setting(name, fallback) {
        var v = settings[name];
        return v ? v : (fallback || '');
    }

    function sheetRows(wb, name) {
        var ws = wb.Sheets[name];
        if (!ws) return [];
        var rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
        return rows.filter(function (r) { return str(r['Record ID']); });
    }

    function initial(title) { return str(title).charAt(0).toUpperCase(); }

    function href(type, id) { return '#/' + type + '/' + encodeURIComponent(id); }

    // 图片：按 Record ID 找 images/<ID>.jpg 等。失败时依次换扩展名，全失败则退回占位块。
    function imgTag(id, alt, cls) {
        return '<img class="' + (cls || '') + '" data-img-id="' + esc(id) + '" data-ext="0" alt="' +
            esc(alt) + '" src="images/' + encodeURIComponent(id) + '.' + IMAGE_EXTS[0] + '">';
    }

    document.addEventListener('error', function (e) {
        var img = e.target;
        if (!img || img.tagName !== 'IMG' || !img.hasAttribute('data-img-id')) return;
        var next = parseInt(img.getAttribute('data-ext'), 10) + 1;
        if (next < IMAGE_EXTS.length) {
            img.setAttribute('data-ext', next);
            img.src = 'images/' + encodeURIComponent(img.getAttribute('data-img-id')) + '.' + IMAGE_EXTS[next];
        } else {
            var holder = img.closest('[data-poster]');
            img.remove();
            if (holder) holder.classList.add('is-placeholder', 'no-image');
        }
    }, true);

    // ---------- 数据装载 ----------

    function load() {
        return fetch(DATA_URL + '?t=' + Date.now()).then(function (r) {
            if (!r.ok) throw new Error('Could not load ' + DATA_URL);
            return r.arrayBuffer();
        }).then(function (buf) {
            var wb = XLSX.read(buf, { type: 'array' });
            settings = {};
            sheetRows({ Sheets: wb.Sheets }, 'Settings'); // no-op, keeps API symmetrical
            var sRows = XLSX.utils.sheet_to_json(wb.Sheets['Settings'] || {}, { defval: '', raw: false });
            sRows.forEach(function (r) { if (str(r['Setting'])) settings[str(r['Setting'])] = str(r['Value']); });

            db.films = sheetRows(wb, 'Films');
            db.people = sheetRows(wb, 'People');
            db.companies = sheetRows(wb, 'Companies');
            db.scholarship = sheetRows(wb, 'Scholarship');

            db.byId = {};
            db.personByName = {};
            db.companyByName = {};
            db.films.forEach(function (r) { r._type = 'film'; r._title = str(r['English Title']); db.byId[r['Record ID']] = r; });
            db.people.forEach(function (r) { r._type = 'person'; r._title = str(r['English Name']); db.byId[r['Record ID']] = r; db.personByName[key(r._title)] = r; });
            db.companies.forEach(function (r) { r._type = 'company'; r._title = str(r['English Name']); db.byId[r['Record ID']] = r; db.companyByName[key(r._title)] = r; });
            db.scholarship.forEach(function (r) { r._type = 'scholarship'; r._title = str(r['Title']); });
        });
    }

    // ---------- 页面外壳 ----------

    var LISTS = {
        films: {
            type: 'film', rows: function () { return db.films; },
            title: 'Films Page Title', subtitle: 'Films Page Subtitle', caption: 'Films Card Caption', defaultTitle: 'Film',
            groupField: 'Countries of Origin', groupLabel: 'Country',
            filters: [
                { id: 'country', label: 'Country', field: 'Countries of Origin', kind: 'radio', all: 'All Countries' },
                { id: 'year', label: 'Year', field: 'Year', kind: 'year', all: 'All Years' },
                { id: 'genre', label: 'Genre', field: 'Genre', kind: 'check' }
            ]
        },
        people: {
            type: 'person', rows: function () { return db.people; },
            title: 'People Page Title', subtitle: 'People Page Subtitle', caption: 'People Card Caption', defaultTitle: 'People',
            groupField: 'Nationality', groupLabel: 'Nationality',
            filters: [
                { id: 'country', label: 'Nationality', field: 'Nationality', kind: 'radio', all: 'All Nationalities' },
                { id: 'occupation', label: 'Occupation', field: 'Occupation', kind: 'check' }
            ]
        },
        companies: {
            type: 'company', rows: function () { return db.companies; },
            title: 'Companies Page Title', subtitle: 'Companies Page Subtitle', caption: 'Companies Card Caption', defaultTitle: 'Production Company',
            groupField: 'Country', groupLabel: 'Country',
            filters: [
                { id: 'country', label: 'Country', field: 'Country', kind: 'radio', all: 'All Countries' }
            ]
        }
    };

    function listTitle(name) { return setting(LISTS[name].title, LISTS[name].defaultTitle); }

    function applyShell() {
        document.querySelectorAll('[data-setting]').forEach(function (el) {
            el.textContent = setting(el.getAttribute('data-setting'));
        });
        var accent = setting('Accent Color');
        if (accent) document.documentElement.style.setProperty('--accent', accent);
        var nav = Object.keys(LISTS).map(function (name) {
            return '<li data-nav-item="' + name + '"><a href="#/' + name + '">' + esc(listTitle(name)) + '</a></li>';
        }).join('');
        document.querySelectorAll('[data-nav]').forEach(function (ul) { ul.innerHTML = nav; });
    }

    function markNav(active) {
        document.querySelectorAll('[data-nav-item]').forEach(function (li) {
            li.classList.toggle('active', li.getAttribute('data-nav-item') === active);
        });
    }

    function setTitle(t) {
        var site = setting('Site Title');
        document.title = t ? t + ' · ' + site : site;
    }

    function bannerStyle(alpha) {
        var img = setting('Banner Image');
        var grad = 'linear-gradient(to right, rgba(20,16,13,' + alpha[0] + '), rgba(20,16,13,' + alpha[1] + '))';
        return img ? "background-image: " + grad + ", url('" + esc(img) + "');" : 'background-image: ' + grad + ';';
    }

    // ---------- 通用部件 ----------

    // 默认不显示任何图片（海报、肖像、剧照、卡片缩略图）。
    // 想恢复：在 Excel 的 Settings 表加一行  Show Images | Yes
    function showImages() {
        return setting('Show Images').toLowerCase() === 'yes';
    }

    function placeholderThumb(title) {
        return '<span class="placeholder-initial" aria-hidden="true">' + esc(initial(title)) + '</span>' +
            '<span class="placeholder-title">' + esc(title) + '</span>';
    }

    function personLinks(value, type) {
        return split(value).map(function (name) {
            var rec = (type === 'company' ? db.companyByName : db.personByName)[key(name)];
            return rec ? '<a class="resource-link" href="' + href(rec._type, rec['Record ID']) + '"><span class="resource-name">' + esc(rec._title) + '</span></a>' : esc(name);
        }).join('<span class="sep">, </span>');
    }

    function field(cls, label, html, showEmpty) {
        var empty = !html;
        if (empty && !showEmpty) return '';
        return '<div class="field ' + cls + (empty ? ' is-empty' : '') + '"><dt>' + esc(label) + '</dt><dd>' +
            (empty ? '<span class="no-value" aria-label="No value">&mdash;</span>' : html) + '</dd></div>';
    }

    function breadcrumb(listName, title) {
        return '<nav class="breadcrumb" aria-label="Breadcrumb"><ol><li><a href="#/">Home</a></li><li><a href="#/' +
            listName + '">' + esc(listTitle(listName)) + '</a></li><li aria-current="page">' + esc(title) + '</li></ol></nav>';
    }

    function paragraphs(text) {
        return str(text).split(/\n+/).map(function (p) { return '<p>' + esc(p) + '</p>'; }).join('');
    }

    function introBlock(label, text, full) {
        return '<div class="detail-introduction' + (full ? ' full-width' : '') + '"><h2 class="intro-label">' + esc(label) + '</h2>' +
            (str(text) ? paragraphs(text) : '<p class="no-value">&mdash;</p>') + '</div>';
    }

    function gallery(rec) {
        if (!showImages()) return '';
        var files = split(rec['Photos']);
        var inner;
        if (files.length) {
            inner = '<ul class="gallery-strip">' + files.map(function (f) {
                return '<li><a href="images/' + encodeURI(f) + '" target="_blank" rel="noopener"><img src="images/' + encodeURI(f) + '" alt="' + esc(rec._title) + '" loading="lazy"></a></li>';
            }).join('') + '</ul>';
        } else {
            inner = '<ul class="gallery-strip is-empty" aria-hidden="true">' + new Array(6).join('<li><span class="media-placeholder"></span></li>') + '</ul>';
        }
        return '<section class="detail-section photo-gallery"><div class="section-head"><h2 class="section-title">Photo Gallery</h2></div>' +
            '<div class="gallery-viewport">' + inner + '</div></section>';
    }

    function scholarshipBlock(rec, heading, slots) {
        var id = rec['Record ID'];
        var entries = db.scholarship.filter(function (s) { return split(s['About']).indexOf(id) !== -1; });
        var body;
        if (!entries.length) {
            var one = '<li class="scholarship-entry placeholder"><span class="scholarship-icon"></span><div class="scholarship-body"><span class="line line-title"></span><span class="line line-meta"></span><span class="line line-cite"></span></div></li>';
            body = '<ul class="scholarship-list is-empty" aria-hidden="true">' + new Array(slots + 1).join(one) + '</ul>';
        } else {
            body = '<ul class="scholarship-list">' + entries.map(function (s) {
                var doi = str(s['DOI']), link = str(s['Link']), cite = str(s['Chicago Citation']);
                var links = '';
                if (doi) links += '<a class="ext" rel="noopener" target="_blank" href="https://doi.org/' + esc(doi) + '">DOI <span aria-hidden="true">&#8599;</span></a>';
                else if (link) links += '<a class="ext" rel="noopener" target="_blank" href="' + esc(link) + '">Full text <span aria-hidden="true">&#8599;</span></a>';
                if (cite) links += '<details class="citation"><summary>Chicago Citation</summary><p>' + esc(cite) + '</p></details>';
                return '<li class="scholarship-entry"><span class="scholarship-icon" aria-hidden="true"></span><div class="scholarship-body">' +
                    '<h3 class="scholarship-title">' + esc(s._title) + '</h3><p class="scholarship-meta">' +
                    (str(s['Authors']) ? '<span class="author">' + esc(s['Authors']) + '</span> ' : '') +
                    (str(s['Journal']) ? '<span class="journal">' + esc(s['Journal']) + '</span>' : '') + '</p>' +
                    (str(s['Year']) ? '<p class="scholarship-year">' + esc(s['Year']) + '</p>' : '') + '</div>' +
                    (links ? '<div class="scholarship-links">' + links + '</div>' : '') + '</li>';
            }).join('') + '</ul>';
        }
        return '<section class="detail-section scholarship"><div class="section-head"><h2 class="section-title">' + esc(heading) + '</h2></div>' + body + '</section>';
    }

    function filmEntry(film, credit) {
        var thumb = showImages() ? '<a class="film-thumb" href="' + href('film', film['Record ID']) + '" aria-hidden="true" tabindex="-1" data-poster>' +
            imgTag(film['Record ID'], film._title) + '<span class="media-placeholder"></span></a>' : '';
        return '<li class="film-entry' + (thumb ? '' : ' no-thumb') + '">' + thumb + '<div class="film-text"><div class="film-title"><a href="' +
            href('film', film['Record ID']) + '">' + esc(film._title) + '</a></div>' +
            (str(film['Original Title']) ? '<div class="film-original">' + esc(film['Original Title']) + '</div>' : '') +
            '<div class="film-credit">' + esc(credit) + '</div></div></li>';
    }

    function filmCredit(film) {
        return [split(film['Director']).join(', '), str(film['Year'])].filter(Boolean).join(', ');
    }

    function filmography(rec, roles) {
        var groups = roles.map(function (role) {
            var films = db.films.filter(function (f) {
                return split(f[role.column]).some(function (n) { return key(n) === key(rec._title); });
            });
            if (!films.length) return '';
            return '<div class="filmography-group"><h4 class="role">' + esc(role.label) + '</h4><ul class="film-list">' +
                films.map(function (f) { return filmEntry(f, filmCredit(f)); }).join('') + '</ul></div>';
        }).join('');
        if (!groups) {
            var ph = '<li class="film-entry placeholder">' + (showImages() ? '<span class="media-placeholder"></span>' : '') + '<div class="film-text"><span class="line line-title"></span><span class="line line-meta"></span></div></li>';
            groups = '<ul class="film-list is-empty" aria-hidden="true">' + ph + ph + '</ul>';
        }
        return '<section class="detail-section filmography"><div class="section-head"><h2 class="section-title">Related to other films</h2></div>' + groups + '</section>';
    }

    // ---------- 首页 ----------

    function renderHome() {
        setTitle('');
        markNav('');
        var entries = Object.keys(LISTS).map(function (name) {
            return '<li class="home-entry no-image"><a class="entry-link" href="#/' + name + '"><span class="entry-body"><span class="entry-title">' +
                esc(listTitle(name)) + '</span><span class="entry-rule" aria-hidden="true"></span></span><span class="entry-arrow" aria-hidden="true">&rarr;</span></a></li>';
        }).join('');
        var bo = setting('Tibetan Line');
        return '<section class="home-hero has-background" style="' + bannerStyle(['0.7', '0.25']) + '"><div class="home-hero-inner"><span class="hero-rule" aria-hidden="true"></span>' +
            '<h1 class="home-title">' + esc(setting('Home Headline', setting('Site Title'))) + '</h1>' +
            (bo ? '<p class="home-script" lang="bo">' + esc(bo) + '</p>' : '') + '</div></section>' +
            '<div class="block block-asset"><ul class="home-entries full-measure">' + entries + '</ul></div>';
    }

    // ---------- 列表页 ----------

    function decade(y) { return Math.floor(parseInt(y, 10) / 10) * 10 + 's'; }

    function valuesOf(rec, f) {
        var vals = split(rec[f.field]);
        return vals.length ? vals : (f.kind === 'check' ? [] : ['Unknown']);
    }

    function parseQuery(q) {
        var out = {};
        q.split('&').forEach(function (p) {
            if (!p) return;
            var i = p.indexOf('=');
            var k = decodeURIComponent(i < 0 ? p : p.slice(0, i));
            var v = i < 0 ? '' : decodeURIComponent(p.slice(i + 1).replace(/\+/g, ' '));
            out[k] = v;
        });
        return out;
    }

    function selected(q, f) {
        var v = q[f.id] || '';
        return f.kind === 'check' ? split(v.replace(/\|/g, ';')) : (v ? [v] : []);
    }

    function passes(rec, list, q, skipId) {
        return list.filters.every(function (f) {
            if (f.id === skipId) return true;
            var sel = selected(q, f);
            if (!sel.length) return true;
            var vals = valuesOf(rec, f);
            if (f.kind === 'year') {
                var yr = str(rec[f.field]);
                return sel.indexOf(yr) !== -1 || (yr && sel.indexOf(decade(yr)) !== -1);
            }
            if (f.kind === 'check') return sel.every(function (s) { return vals.indexOf(s) !== -1; });
            return sel.some(function (s) { return vals.indexOf(s) !== -1; });
        });
    }

    function filterGroup(list, f, q) {
        var pool = list.rows().filter(function (r) { return passes(r, list, q, f.id); });
        var sel = selected(q, f);
        var counts = {};
        pool.forEach(function (r) {
            var vals = f.kind === 'year' ? (str(r[f.field]) ? [str(r[f.field])] : []) : valuesOf(r, f);
            vals.forEach(function (v) { counts[v] = (counts[v] || 0) + 1; });
        });
        function item(type, name, value, checked, count) {
            return '<li><label class="' + (checked ? 'is-active' : '') + '"><input type="' + type + '" name="' + f.id + '" value="' + esc(value) + '"' +
                (checked ? ' checked' : '') + '><span class="filter-name">' + esc(name) + '</span><span class="filter-count">(' + count + ')</span></label></li>';
        }
        var html = '<section class="filter-group"><h3>Filter by ' + esc(f.label) + '</h3>';
        if (f.kind === 'check') {
            var names = Object.keys(counts).sort();
            if (!names.length) return '';
            html += '<ul class="filter-list checkbox-list">' + names.map(function (n) {
                return item('checkbox', n, n, sel.indexOf(n) !== -1, counts[n]);
            }).join('') + '</ul>';
        } else if (f.kind === 'year') {
            var years = Object.keys(counts).sort(function (a, b) { return b - a; });
            html += '<ul class="filter-list">' + item('radio', f.all, '', !sel.length, pool.length) + '</ul>';
            var byDecade = {};
            years.forEach(function (y) { (byDecade[decade(y)] = byDecade[decade(y)] || []).push(y); });
            Object.keys(byDecade).sort().reverse().forEach(function (d) {
                var total = byDecade[d].reduce(function (a, y) { return a + counts[y]; }, 0);
                var decadeOn = sel.indexOf(d) !== -1;
                var open = decadeOn || byDecade[d].some(function (y) { return sel.indexOf(y) !== -1; });
                html += '<details class="decade' + (decadeOn ? ' is-active' : '') + '"' + (open ? ' open' : '') + '><summary data-decade="' + d + '"><span class="filter-name">' + d + '</span><span class="filter-count">(' + total + ')</span></summary><ul class="filter-list">' +
                    byDecade[d].map(function (y) { return item('radio', y, y, sel.indexOf(y) !== -1, counts[y]); }).join('') + '</ul></details>';
            });
        } else {
            var order = Object.keys(counts).sort(function (a, b) {
                if (a === 'Unknown') return 1; if (b === 'Unknown') return -1; return counts[b] - counts[a] || a.localeCompare(b);
            });
            html += '<ul class="filter-list">' + item('radio', f.all, '', !sel.length, pool.length) +
                order.map(function (n) { return item('radio', n, n, sel.indexOf(n) !== -1, counts[n]); }).join('') + '</ul>';
        }
        return html + '</section>';
    }

    function card(rec) {
        var type = rec._type, id = rec['Record ID'];
        var lines = '';
        if (type === 'film') {
            if (str(rec['Original Title'])) lines += '<p class="card-original">' + esc(rec['Original Title']) + '</p>';
            if (str(rec['Director'])) lines += '<p class="card-director">' + esc(split(rec['Director']).join(', ')) + '</p>';
            if (str(rec['Year'])) lines += '<p class="card-year">' + esc(rec['Year']) + '</p>';
        } else {
            if (str(rec['Original Name'])) lines += '<p class="card-original">' + esc(rec['Original Name']) + '</p>';
            if (type === 'person' && str(rec['Occupation'])) lines += '<p class="card-director">' + esc(split(rec['Occupation']).join(', ')) + '</p>';
        }
        var thumb = showImages() ? '<span class="card-thumb is-placeholder" data-poster>' + placeholderThumb(rec._title) + imgTag(id, rec._title) + '</span>' : '';
        return '<li class="resource-card' + (thumb ? '' : ' no-thumb') + '"><a class="card-link" href="' + href(type, id) + '">' + thumb +
            '<span class="card-title">' + esc(rec._title) + '</span></a>' + lines + '</li>';
    }

    function sortRecs(rows, type) {
        return rows.slice().sort(function (a, b) {
            if (type === 'film') {
                var d = (parseInt(a['Year'], 10) || 9999) - (parseInt(b['Year'], 10) || 9999);
                if (d) return d;
            }
            return a._title.localeCompare(b._title);
        });
    }

    function renderList(name, qs) {
        var list = LISTS[name];
        var q = parseQuery(qs || '');
        setTitle(listTitle(name));
        markNav(name);

        var rows = list.rows().filter(function (r) { return passes(r, list, q); });
        var groupFilter = list.filters[0];
        var chosenGroups = selected(q, groupFilter);

        var groups = {};
        rows.forEach(function (r) {
            valuesOf(r, { field: list.groupField, kind: 'radio' }).forEach(function (g) {
                if (chosenGroups.length && chosenGroups.indexOf(g) === -1) return;
                (groups[g] = groups[g] || []).push(r);
            });
        });
        var names = Object.keys(groups).sort(function (a, b) {
            if (a === 'Unknown') return 1; if (b === 'Unknown') return -1; return a.localeCompare(b);
        });

        var main = names.length ? names.map(function (g) {
            return '<section class="resource-group"><h2 class="group-heading"><span>' + esc(g) + '</span></h2><ul class="card-grid">' +
                sortRecs(groups[g], list.type).map(card).join('') + '</ul></section>';
        }).join('') : '<p class="no-results">No entries match these filters.</p>';

        return '<section class="page-hero" style="' + bannerStyle(['.78', '.25']) + '"><div class="hero-inner"><div class="hero-text"><h1 class="hero-title">' + esc(listTitle(name)) +
            '</h1><div class="hero-subtitle"><p>' + esc(setting(list.subtitle)) + '</p></div></div></div></section>' +
            '<div class="browse-layout full-measure"><aside class="browse-filters"><form class="filter-form" data-list="' + name + '">' +
            list.filters.map(function (f) { return filterGroup(list, f, q); }).join('') + '</form></aside>' +
            '<main class="browse-main">' + main + '</main></div>';
    }

    // override：{ 筛选器 id: 值 }，用来替换表单里勾选的值（点十年标签时用）
    function applyFilters(form, override) {
        var list = LISTS[form.getAttribute('data-list')];
        var parts = [];
        list.filters.forEach(function (f) {
            var vals = override && f.id in override ? [override[f.id]] :
                Array.prototype.map.call(form.querySelectorAll('input[name="' + f.id + '"]:checked'), function (i) { return i.value; });
            vals = vals.filter(Boolean);
            if (vals.length) parts.push(f.id + '=' + encodeURIComponent(vals.join('|')));
        });
        location.hash = '#/' + form.getAttribute('data-list') + (parts.length ? '?' + parts.join('&') : '');
    }

    document.addEventListener('change', function (e) {
        var form = e.target.closest && e.target.closest('form.filter-form');
        if (form) applyFilters(form);
    });

    // 点十年标签：筛选到整个年代，同时展开子年份
    document.addEventListener('click', function (e) {
        var sum = e.target.closest && e.target.closest('.decade > summary');
        if (!sum) return;
        var form = sum.closest('form.filter-form');
        if (!form) return;
        e.preventDefault();
        applyFilters(form, { year: sum.getAttribute('data-decade') });
    });

    // ---------- 详情页 ----------

    // 模板里已经安排了位置的列；其余列（以后 Excel 加的新列）自动接在事实块后面
    var PLACED = {
        film: ['Record ID', 'English Title', 'Original Title', 'Year', 'Runtime', 'Genre', 'Color', 'Sound', 'Director', 'Writer', 'Cast',
            'First Release Date', 'Countries of Origin', 'Language', 'Filming Locations', 'Production Company', 'Introduction', 'Photos'],
        person: ['Record ID', 'English Name', 'Original Name', 'Occupation', 'Born', 'Died', 'Nationality', 'Biography', 'Photos'],
        company: ['Record ID', 'English Name', 'Original Name', 'Country', 'Founded', 'Introduction', 'Photos']
    };

    function extraFields(rec) {
        return Object.keys(rec).filter(function (k) {
            return k.charAt(0) !== '_' && PLACED[rec._type].indexOf(k) === -1 && str(rec[k]);
        }).map(function (k) { return field('field-extra', k, esc(split(rec[k]).join(', ')), false); }).join('');
    }

    function topbar(title, original, metaHtml) {
        return '<div class="detail-topbar"><div class="title-box"><h1 class="detail-title">' + esc(title) + '</h1>' +
            (str(original) ? '<p class="detail-original">' + esc(original) + '</p>' : '') + '</div>' +
            '<div class="meta-box">' + (metaHtml || '<p class="no-value">&mdash;</p>') + '</div></div>';
    }

    function heroImage(rec, cls) {
        if (!showImages()) return '';
        return '<div class="hero-image"><div class="hero-poster ' + cls + '" data-poster>' + imgTag(rec['Record ID'], rec._title) +
            '<span class="media-placeholder ' + cls + '" aria-hidden="true"></span></div></div>';
    }

    function renderFilm(rec) {
        setTitle(rec._title);
        markNav('films');
        var meta1 = [str(rec['Year']), str(rec['Runtime']), split(rec['Genre']).join(', ')].filter(Boolean).join(' · ');
        var meta2 = [str(rec['Color']), str(rec['Sound'])].filter(Boolean).join(' · ');
        var metaHtml = (meta1 ? '<p>' + esc(meta1) + '</p>' : '') + (meta2 ? '<p>' + esc(meta2) + '</p>' : '');
        return '<article class="resource-detail film-detail">' + breadcrumb('films', rec._title) + topbar(rec._title, rec['Original Title'], metaHtml) +
            '<div class="detail-hero' + (showImages() ? '' : ' no-image') + '">' + heroImage(rec, 'poster') + '<div class="hero-content">' +
            '<dl class="resource-fields credits">' +
            field('field-director', 'Director', personLinks(rec['Director'], 'person'), true) +
            field('field-writer', 'Writer', personLinks(rec['Writer'], 'person'), true) +
            field('field-cast', 'Cast', personLinks(rec['Cast'], 'person'), true) + '</dl>' +
            '<dl class="resource-fields facts">' +
            field('field-release', 'First Release Date', esc(rec['First Release Date']), true) +
            field('field-country', 'Countries of Origin', esc(split(rec['Countries of Origin']).join(', ')), true) +
            field('field-language', 'Language', esc(split(rec['Language']).join(', ')), true) +
            field('field-locations', 'Filming Locations', esc(split(rec['Filming Locations']).join(', ')), true) +
            field('field-company', 'Production Company', personLinks(rec['Production Company'], 'company'), true) +
            extraFields(rec) + '</dl>' + introBlock('Introduction', rec['Introduction']) + '</div></div>' +
            gallery(rec) + scholarshipBlock(rec, 'Academic Scholarship on the Film', 4) + '</article>';
    }

    function renderPerson(rec) {
        setTitle(rec._title);
        markNav('people');
        var life = [str(rec['Born']) ? 'Born: ' + str(rec['Born']) : '', str(rec['Died']) ? 'Died: ' + str(rec['Died']) : ''].filter(Boolean).join(' · ');
        var metaHtml = (str(rec['Occupation']) ? '<p>' + esc(split(rec['Occupation']).join(', ')) + '</p>' : '') + (life ? '<p>' + esc(life) + '</p>' : '');
        return '<article class="resource-detail person-detail">' + breadcrumb('people', rec._title) + topbar(rec._title, rec['Original Name'], metaHtml) +
            '<div class="detail-hero' + (showImages() ? '' : ' no-image') + '">' + heroImage(rec, 'portrait') + '<div class="hero-content">' +
            '<dl class="resource-fields facts">' + field('field-nationality', 'Nationality', esc(rec['Nationality']), true) + extraFields(rec) + '</dl>' +
            introBlock('Biography', rec['Biography']) + '</div></div>' + gallery(rec) +
            filmography(rec, [{ label: 'Director', column: 'Director' }, { label: 'Writer', column: 'Writer' }, { label: 'Cast', column: 'Cast' }]) +
            scholarshipBlock(rec, 'Academic Scholarship on the filmmaker', 1) + '</article>';
    }

    function renderCompany(rec) {
        setTitle(rec._title);
        markNav('companies');
        var metaHtml = str(rec['Country']) ? '<p>' + esc(rec['Country']) + '</p>' : '';
        var extra = extraFields(rec) + (str(rec['Founded']) ? field('field-founded', 'Founded', esc(rec['Founded']), false) : '');
        return '<article class="resource-detail organization-detail">' + breadcrumb('companies', rec._title) + topbar(rec._title, rec['Original Name'], metaHtml) +
            (extra ? '<dl class="resource-fields facts">' + extra + '</dl>' : '') + introBlock('Introduction', rec['Introduction'], true) +
            filmography(rec, [{ label: 'Production Company', column: 'Production Company' }]) + scholarshipBlock(rec, 'Academic Scholarship on the filmmaker', 1) + '</article>';
    }

    // ---------- 搜索 ----------

    function renderSearch(q) {
        setTitle('Search');
        markNav('');
        var needle = key(q);
        var hits = [].concat(db.films, db.people, db.companies).filter(function (r) {
            return needle && Object.keys(r).some(function (k) { return k.charAt(0) !== '_' && key(r[k]).indexOf(needle) !== -1; });
        });
        return '<section class="page-hero" style="' + bannerStyle(['.78', '.25']) + '"><div class="hero-inner"><div class="hero-text"><h1 class="hero-title">Search</h1>' +
            '<div class="hero-subtitle"><p>' + hits.length + ' result' + (hits.length === 1 ? '' : 's') + ' for &ldquo;' + esc(q) + '&rdquo;</p></div></div></div></section>' +
            '<div class="browse-layout full-measure"><main class="browse-main">' +
            (hits.length ? '<ul class="card-grid">' + hits.map(card).join('') + '</ul>' : '<p class="no-results">Nothing found.</p>') + '</main></div>';
    }

    document.addEventListener('submit', function (e) {
        if (e.target.id !== 'search-form') return;
        e.preventDefault();
        var q = document.getElementById('fulltext-search').value.trim();
        if (q) location.hash = '#/search?q=' + encodeURIComponent(q);
    });

    // ---------- 路由 ----------

    function route() {
        var content = document.getElementById('content');
        var raw = location.hash.replace(/^#/, '') || '/';
        var qi = raw.indexOf('?');
        var path = qi < 0 ? raw : raw.slice(0, qi);
        var qs = qi < 0 ? '' : raw.slice(qi + 1);
        var parts = path.split('/').filter(Boolean);
        var html;

        if (!parts.length) {
            document.body.className = 'page home-template';
            html = renderHome();
        } else if (LISTS[parts[0]] && parts.length === 1) {
            document.body.className = 'page grouped-browse-template';
            html = renderList(parts[0], qs);
        } else if (parts[0] === 'search') {
            document.body.className = 'page grouped-browse-template';
            html = renderSearch(parseQuery(qs).q || '');
        } else if (['film', 'person', 'company'].indexOf(parts[0]) !== -1 && parts[1]) {
            var rec = db.byId[decodeURIComponent(parts[1])];
            if (rec && rec._type === parts[0]) {
                document.body.className = 'item resource show ' + parts[0] + '-detail-page';
                html = { film: renderFilm, person: renderPerson, company: renderCompany }[parts[0]](rec);
            }
        }
        if (!html) {
            document.body.className = 'page';
            setTitle('Not found');
            html = '<section class="detail-section"><h1>Page not found</h1><p><a href="#/">Back to home</a></p></section>';
        }
        content.innerHTML = html;
        window.scrollTo(0, 0);
    }

    window.addEventListener('hashchange', route);

    load().then(function () {
        applyShell();
        route();
    }).catch(function (err) {
        document.getElementById('content').innerHTML = '<p class="site-status">Could not load the archive data (' + esc(err.message) +
            '). If you opened index.html straight from a folder, view it through a web server or the published GitHub Pages address instead.</p>';
    });
})();
