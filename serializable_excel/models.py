"""
ExcelModel - Base class for Excel-serializable Pydantic models.
"""

from typing import Dict, List, Type, TypeVar

from pydantic import BaseModel

from serializable_excel.color_extractor import ColorExtractor
from serializable_excel.excel_reader import ExcelReader
from serializable_excel.excel_writer import ExcelWriter
from serializable_excel.field_extractor import FieldExtractor
from serializable_excel.field_metadata import FieldMetadataExtractor
from serializable_excel.validators import FieldValidator

T = TypeVar("T", bound="ExcelModel")

# Module-level storage for singletons to avoid Pydantic conflicts
_singletons: Dict[type, Dict[str, any]] = {}


class ExcelModel(BaseModel):
    """
    Base class for Excel-serializable Pydantic models.

    Inherit from this class and define fields using Column() or DynamicColumn() descriptors.
    """

    model_config = {"arbitrary_types_allowed": True}

    @classmethod
    def _get_singletons(cls) -> Dict[str, any]:
        """Get or create singletons dictionary for this class."""
        if cls not in _singletons:
            _singletons[cls] = {}
        return _singletons[cls]

    @classmethod
    def _get_reader(cls) -> ExcelReader:
        """Get or create ExcelReader instance (lazy initialization)."""
        singletons = cls._get_singletons()
        if "reader" not in singletons:
            metadata_extractor = FieldMetadataExtractor()
            validator = FieldValidator()
            singletons["reader"] = ExcelReader(metadata_extractor, validator)
        return singletons["reader"]

    @classmethod
    def _get_writer(cls) -> ExcelWriter:
        """Get or create ExcelWriter instance (lazy initialization)."""
        singletons = cls._get_singletons()
        if "writer" not in singletons:
            metadata_extractor = FieldMetadataExtractor()
            field_extractor = FieldExtractor()
            color_extractor = ColorExtractor()
            singletons["writer"] = ExcelWriter(
                metadata_extractor,
                field_extractor,
                color_extractor,
            )
        return singletons["writer"]

    @classmethod
    def from_excel(
        cls: Type[T], file_path: str, dynamic_columns: bool = False
    ) -> List[T]:
        """
        Read models from an Excel file.

        Args:
            file_path: Path to the Excel file (.xlsx)
            dynamic_columns: Enable detection of additional columns not defined in model

        Returns:
            List of model instances

        Raises:
            FileNotFoundError: If the Excel file doesn't exist
            ValidationError: If validation fails
            ColumnNotFoundError: If a required column is not found
        """
        reader = cls._get_reader()
        return reader.read(cls, file_path, dynamic_columns)

    @classmethod
    def to_excel(cls, instances: List[T], file_path: str) -> None:
        """
        Export model instances to an Excel file.

        Args:
            instances: List of model instances to export
            file_path: Path where to save the Excel file (.xlsx)

        Raises:
            ValueError: If instances list is empty or invalid
        """
        writer = cls._get_writer()
        writer.write(cls, instances, file_path)
