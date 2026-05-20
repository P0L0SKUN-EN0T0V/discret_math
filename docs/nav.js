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
        { type: 'link', href: 'index.html', text: 'Главная' },
        { type: 'sep' },
        { type: 'title', text: 'Методичка' },
        { type: 'link', href: 'sections/0-why-formal-models.html', text: '0. Зачем формальные модели' },
        { type: 'link', href: 'sections/1-fsm.html', text: '1. Конечные автоматы' },
        { type: 'link', href: 'sections/2-petri-nets.html', text: '2. Сети Петри' },
        { type: 'link', href: 'sections/3-statecharts.html', text: '3. Statecharts' },
        { type: 'link', href: 'sections/4-time.html', text: '4. Время в моделях' },
        { type: 'link', href: 'sections/5-behaviour-trees.html', text: '5. Деревья поведения' },
        { type: 'link', href: 'sections/6-grafcet.html', text: '6. GRAFCET' },
        { type: 'link', href: 'sections/7-others.html', text: '7. Остальные инструменты' },
        { type: 'link', href: 'sections/8-how-to-choose.html', text: '8. Как выбрать' },
    ];

    var nav = document.getElementById('sidebar');
    if (!nav) return;

    var h3 = document.createElement('h3');
    h3.textContent = 'Формализмы';
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
