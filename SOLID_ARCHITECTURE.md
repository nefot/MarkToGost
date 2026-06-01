"""
АРХИТЕКТУРА РЕНДЕРИНГА С ЕДИНЫМ ИНТЕРФЕЙСОМ

Эта архитектура обеспечивает соответствие SOLID принципам:

✅ Single Responsibility Principle (SRP)
   - Каждый файл отвечает за один тип блока
   - DocumentRenderer только координирует рендеринг

✅ Open/Closed Principle (OCP)
   - Открыт для расширения (добавить новый блок-тип):
     1. Создать функцию render_xxx_block(renderer, block)
     2. Добавить в _renderers dict в DocumentRenderer.__init__
   - Закрыт для модификации основного кода

✅ Liskov Substitution Principle (LSP)
   - Все функции имеют одинаковую сигнатуру:
     def render_xxx_block(renderer, block) -> None
   - Они могут использоваться полиморфно через регистр

✅ Interface Segregation Principle (ISP)
   - Определён минимальный контракт (BlockRenderer protocol)
   - Все функции предоставляют только нужный интерфейс

✅ Dependency Inversion Principle (DIP)
   - DocumentRenderer зависит от абстракции (функции с единой сигнатурой)
   - А не от конкретных реализаций

СИГНАТУРА ВСЕХ ФУНКЦИЙ РЕНДЕРИНГА
==================================
def render_xxx_block(renderer: DocumentRenderer, block: BaseBlock) -> None:
    '''Рендеринг блока'''
    # Доступ к ресурсам через renderer:
    renderer.doc              # Document
    renderer.table_counter    # счетчик таблиц
    renderer.formula_counter  # счетчик формул
    renderer.figure_counter   # счетчик рисунков
    renderer.image_refs       # словарь ID изображений
    renderer.formula_refs     # словарь ID формул
    renderer.image_map        # кэш имён файлов
    renderer.use_headings     # флаг использования стилей заголовков
    
    renderer._is_document_start()  # проверка начала документа
    renderer._safe_page_break()    # безопасный разрыв страницы
    renderer._mark_content()       # сброс флага после добавления контента
    renderer.render_block(block)   # рекурсивный рендеринг

РЕГИСТР РЕНДЕРОВ
================
self._renderers: Dict[type, Callable] = {
    Section: render_section_block,
    TextBlock: render_text_block,
    HeadingBlock: render_heading_block,
    ImageBlock: render_image_block,
    ListBlock: render_list_block,
    TableBlock: render_table_block,
    CodeBlock: render_code_block,
    FormulaBlock: render_formula_block,
}

Все функции находятся в отдельных модулях:
- renderere/text.py       -> render_text_block, render_heading_block
- renderer/image.py       -> render_image_block
- renderer/list_.py       -> render_list_block
- renderer/code.py        -> render_code_block
- renderer/formula.py     -> render_formula_block
- renderer/table.py       -> render_table_block
- renderer/Section.py     -> render_section_block

РАСШИРЕНИЕ: ДОБАВЛЕНИЕ НОВОГО БЛОКА
====================================
1. Создать новый файл MarkToGost/renderer/xxx.py:

    from MarkToGost.parser.blocks import XXXBlock
    
    def render_xxx_block(renderer, block: XXXBlock):
        '''Рендеринг XXX блока'''
        # использовать renderer.doc, renderer._mark_content(), и т.д.
        renderer._mark_content()

2. В DocumentRenderer.__init__ добавить в _renderers:

    from MarkToGost.renderer.xxx import render_xxx_block
    
    self._renderers = {
        # ... другие рендеры ...
        XXXBlock: render_xxx_block,
    }

Готово! Больше никаких изменений в основном коде не требуется.
"""

