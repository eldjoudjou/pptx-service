"""
Validation modules for Office document processing.
"""

from .base import BaseSchemaValidator
from .pptx import PPTXSchemaValidator

__all__ = [
    "BaseSchemaValidator",
    "PPTXSchemaValidator",
]
