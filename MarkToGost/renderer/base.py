"""Базовый интерфейс для рендеров блоков"""

from typing import Protocol, TYPE_CHECKING

if TYPE_CHECKING:
    from MarkToGost.parser.blocks import BaseBlock
    from MarkToGost.renderer.document_renderer import DocumentRenderer


class BlockRenderer(Protocol):
    """Протокол для всех функций рендеринга блоков
    
    Все функции рендеринга должны иметь следующую сигнатуру:
    def render_xxx_block(renderer: DocumentRenderer, block: BaseBlock) -> None
    """
    
    def __call__(self, renderer: 'DocumentRenderer', block: 'BaseBlock') -> None:
        """
        Рендеринг блока
        
        Args:
            renderer: Экземпляр DocumentRenderer с доступом ко всем ресурсам
            block: Блок для рендеринга
        """
        ...

