"""OOXML utilities for packing and unpacking Office documents."""

from .pack import condense_xml, pack_document
from .unpack import unpack_document

__all__ = ["pack_document", "unpack_document", "condense_xml"]
