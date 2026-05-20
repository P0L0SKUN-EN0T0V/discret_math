(function() {
    var path = location.pathname;
    var file = path.split('/').pop() || 'index.html';

    var prefix = '';
    if (path.indexOf('/sections/') !== -1) {
        prefix = '../';
    }

    var currentKey = file;
    if (path.indexOf('/sections/') !== -1) currentKey = 'sections/' + file;

    var sections = [
        { type: 'link', href: 'index.html', text: 'Обзор' },
        { type: 'sep' },
        { type: 'title', text: 'Основы' },
        { type: 'link', href: 'sections/1-what-are-formalisms.html', text: '1. Что такое формализмы' },
        { type: 'sep' },
        { type: 'title', text: 'Формализмы' },
        { type: 'link', href: 'sections/2-fsm.html', text: '2. Конечные автоматы' },
        { type: 'link', href: 'sections/3-petri-nets.html', text: '3. Сети Петри' },
        { type: 'link', href: 'sections/4-statecharts.html', text: '4. Statecharts (Харел)' },
        { type: 'link', href: 'sections/5-timed-petri.html', text: '5. Временные сети Петри' },
        { type: 'link', href: 'sections/6-behaviour-trees.html', text: '6. Деревья поведения' },
        { type: 'link', href: 'sections/7-grafcet.html', text: '7. GRAFCET / SFC' },
        { type: 'link', href: 'sections/8-others.html', text: '8. Другие формализмы' },
        { type: 'sep' },
        { type: 'title', text: 'Итоги' },
        { type: 'link', href: 'sections/9-comparison.html', text: '9. Сравнение и выбор' },
    ];

    var nav = document.getElementById('sidebar');
    if (!nav) return;

    var h3 = document.createElement('h3');
    h3.textContent = 'Дискретная математика';
    nav.appendChild(h3);

    sections.forEach(function(item) {
        if (item.type === 'sep') {
            var div = document.createElement('div');
            div.className = 'sep';
            nav.appendChild(div);
        } else if (item.type === 'title') {
            var div = document.createElement('div');
            div.className = 'section-title';
            div.textContent = item.text;
            nav.appendChild(div);
        } else if (item.type === 'link') {
            var a = document.createElement('a');
            a.href = prefix + item.href;
            a.textContent = item.text;
            if (item.href === currentKey) a.className = 'active';
            nav.appendChild(a);
        }
    });
})();
