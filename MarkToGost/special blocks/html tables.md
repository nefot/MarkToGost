Формулы в ячейках поддерживаются через `$...$`:

| Параметр | Формула         |
|----------|-----------------|
| Площадь  | $S = \pi r^2$   |
| Длина    | $L = 2\pi r$    |

---

## HTML-таблица

Используется когда нужны объединение ячеек, форматирование или формулы.

### Базовый синтаксис

```html
<table>
    <tr>
        <td>Ячейка 1</td>
        <td>Ячейка 2</td>
    </tr>
</table>
```

---

### Объединение ячеек

`colspan` — объединение по горизонтали, `rowspan` — по вертикали:

```html
<table>
    <tr>
        <td colspan="3">Объединяет 3 колонки</td>
    </tr>
    <tr>
        <td rowspan="2">Объединяет 2 строки</td>
        <td>Ячейка</td>
        <td>Ячейка</td>
    </tr>
    <tr>
        <td>Ячейка</td>
        <td>Ячейка</td>
    </tr>
</table>
```

---

### Выравнивание текста

Атрибут `align` на теге `<td>`. Допустимые значения: `left` (по умолчанию), `center`, `right`:

```html
<table>
    <tr>
        <td align="left">По левому краю</td>
        <td align="center">По центру</td>
        <td align="right">По правому краю</td>
    </tr>
</table>
```

---

### Форматирование текста

Атрибуты указываются без значения прямо на теге `<td>`:

| Атрибут     | Эффект            |
|-------------|-------------------|
| `bold`      | **Жирный**        |
| `italic`    | *Курсив*          |
| `underline` | Подчёркнутый      |

Атрибуты можно комбинировать:

```html
<table>
    <tr>
        <td bold>Жирный</td>
        <td italic>Курсив</td>
        <td underline>Подчёркнутый</td>
        <td bold italic>Жирный курсив</td>
        <td bold italic underline>Всё сразу</td>
    </tr>
</table>
```

---

### Формулы в ячейках

Формула задаётся атрибутом `formula` в формате LaTeX.
Текст и формула могут быть в одной ячейке одновременно:

```html
<table>
    <tr>
        <td>Площадь круга</td>
        <td formula="\pi r^2" align="center"></td>
    </tr>
    <tr>
        <td>Длина окружности</td>
        <td align="center">L = <br/> </td>
    </tr>
    <tr>
        <td bold>Теорема Пифагора</td>
        <td formula="a^2 + b^2 = c^2" align="center"></td>
    </tr>
</table>
```

---

### Прозрачная таблица

Атрибут `class="transparent"` убирает все границы.
Используется для вёрстки без видимой сетки:

```html
<table class="transparent">
    <tr>
        <td bold>Левый блок</td>
        <td align="right">Правый блок</td>
    </tr>
</table>
```

---

### Полный пример

```html
<table>
    <tr>
        <td colspan="3" align="center" bold>Технические характеристики</td>
    </tr>
    <tr>
        <td align="center" bold>Параметр</td>
        <td align="center" bold>Формула</td>
        <td align="center" bold>Единица</td>
    </tr>
    <tr>
        <td colspan="3" bold>Геометрия</td>
    </tr>
    <tr>
        <td>Площадь сечения</td>
        <td formula="\frac{\pi d^2}{4}" align="center"></td>
        <td align="center">мм²</td>
    </tr>
    <tr>
        <td>Момент инерции</td>
        <td formula="\frac{\pi d^4}{64}" align="center"></td>
        <td align="center">мм⁴</td>
    </tr>
    <tr>
        <td colspan="3" bold>Материал</td>
    </tr>
    <tr>
        <td>Обозначение</td>
        <td align="center">Сталь 45</td>
        <td align="center" italic>ГОСТ 1050-2013</td>
    </tr>
    <tr>
        <td colspan="3" align="right" italic>* все размеры в мм</td>
    </tr>
</table>
```

---

## Сравнение типов таблиц

| Возможность              | Markdown | HTML |
|--------------------------|----------|------|
| Простой текст            | ✅        | ✅    |
| Формулы `$...$`          | ✅        | ✅    |
| Жирный / курсив          | ❌        | ✅    |
| Подчёркивание            | ❌        | ✅    |
| colspan / rowspan        | ❌        | ✅    |
| Прозрачные границы       | ❌        | ✅    |
| Подпись таблицы          | ✅        | ❌    |