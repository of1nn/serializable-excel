"""
Excel file writer with separation of concerns.
"""

from typing import Any, Dict, List, Optional, TypeVar

from serializable_excel.color_extractor import ColorExtractor
from serializable_excel.descriptors import Column, DynamicColumn
from serializable_excel.excel_io import write_excel, write_excel_to_bytes
from serializable_excel.excel_types import ExcelType
from serializable_excel.field_extractor import FieldExtractor
from serializable_excel.field_metadata import FieldMetadataExtractor

T = TypeVar("T")


class ExcelWriter:
    """
    Handles writing model instances to Excel files.
    Implements Single Responsibility Principle.
    """

    def __init__(
        self,
        metadata_extractor: FieldMetadataExtractor,
        field_extractor: FieldExtractor,
        color_extractor: Optional[ColorExtractor] = None,
    ):
        """
        Initialize ExcelWriter with dependencies.

        Args:
            metadata_extractor: Extractor for field metadata
            field_extractor: Extractor for field values
            color_extractor: Extractor for cell colors (optional)
        """
        self.metadata_extractor = metadata_extractor
        self.field_extractor = field_extractor
        self.color_extractor = color_extractor or ColorExtractor()

    def write(
        self,
        model_class: type,
        instances: List[Any],
        file_path: Optional[str] = None,
        return_bytes: bool = False,
    ) -> Optional[bytes]:
        """
        Export model instances to an Excel file or return as bytes.

        Args:
            model_class: Model class
            instances: List of model instances to export
            file_path: Path where to save the Excel file (required if return_bytes=False)
            return_bytes: If True, return Excel content as bytes instead of saving to file

        Returns:
            bytes if return_bytes=True, None otherwise

        Raises:
            ValueError: If instances list is empty or invalid
            ValueError: If file_path is not provided when return_bytes=False
        """
        if not instances:
            raise ValueError("Cannot export empty list of instances")

        if not return_bytes and not file_path:
            raise ValueError("file_path is required when return_bytes=False")

        column_fields = self.metadata_extractor.get_column_fields(model_class)
        if not column_fields:
            raise ValueError("No Column fields defined in model")

        dynamic_field = self.metadata_extractor.get_dynamic_column_field(model_class)

        headers = self._build_headers(column_fields, dynamic_field, instances)
        data_rows = self._build_data_rows(instances, column_fields, dynamic_field, headers)

        # Build cell colors
        all_dynamic_keys = (
            self._collect_dynamic_keys(instances, dynamic_field)
            if dynamic_field is not None
            else set()
        )
        cell_colors = self.color_extractor.build_cell_colors(
            data_rows=data_rows,
            column_fields=column_fields,
            dynamic_field=dynamic_field,
            all_dynamic_keys=all_dynamic_keys,
            headers=headers,
        )

        # Build column types
        column_types = self._build_column_types(column_fields, dynamic_field, all_dynamic_keys)

        if return_bytes:
            return write_excel_to_bytes(
                headers,
                data_rows,
                cell_colors=cell_colors,
                column_types=column_types,
            )
        else:
            write_excel(
                headers,
                data_rows,
                file_path,
                cell_colors=cell_colors,
                column_types=column_types,
            )
            return None

    def _build_headers(
        self,
        column_fields: Dict[str, Any],
        dynamic_field: Any,
        instances: List[Any],
    ) -> Dict[str, int]:
        """Build headers mapping for Excel export."""
        headers: Dict[str, int] = {}
        col_idx = 1

        # Add static column headers
        for field_name, column in column_fields.items():
            headers[column.header] = col_idx
            col_idx += 1

        # Add dynamic column headers
        if dynamic_field is not None:
            all_dynamic_keys = self._collect_dynamic_keys(instances, dynamic_field)
            for key in sorted(all_dynamic_keys):
                headers[key] = col_idx
                col_idx += 1

        return headers

    def _collect_dynamic_keys(self, instances: List[Any], dynamic_field: DynamicColumn) -> set:
        """Collect all dynamic column keys from all instances."""
        all_dynamic_keys = set()
        for instance in instances:
            dynamic_data = getattr(instance, dynamic_field.name, {})
            if isinstance(dynamic_data, dict):
                all_dynamic_keys.update(dynamic_data.keys())
        return all_dynamic_keys

    def _build_data_rows(
        self,
        instances: List[Any],
        column_fields: Dict[str, Any],
        dynamic_field: Any,
        headers: Dict[str, int],
    ) -> List[Dict[str, Any]]:
        """Build data rows for Excel export."""
        data_rows: List[Dict[str, Any]] = []
        all_dynamic_keys = (
            self._collect_dynamic_keys(instances, dynamic_field)
            if dynamic_field is not None
            else set()
        )

        for instance in instances:
            row_data: Dict[str, Any] = {}

            # Process static columns
            for field_name, column in column_fields.items():
                value = self.field_extractor.extract_static_field_value(
                    instance, field_name, column
                )
                row_data[column.header] = value

            # Process dynamic columns
            if dynamic_field is not None:
                dynamic_values = self.field_extractor.extract_dynamic_field_value(
                    instance, dynamic_field, all_dynamic_keys
                )
                row_data.update(dynamic_values)

            data_rows.append(row_data)

        return data_rows

    def _build_column_types(
        self,
        column_fields: Dict[str, Column],
        dynamic_field: Optional[DynamicColumn],
        all_dynamic_keys: set,
    ) -> Dict[str, ExcelType]:
        """
        Build column types mapping for Excel export.

        Args:
            column_fields: Dictionary of field names to Column descriptors
            dynamic_field: DynamicColumn descriptor or None
            all_dynamic_keys: Set of all dynamic column keys

        Returns:
            Dictionary mapping header names to ExcelType objects
        """
        column_types: Dict[str, ExcelType] = {}

        # Add types for static columns
        for field_name, column in column_fields.items():
            if column.excel_type is not None:
                column_types[column.header] = column.excel_type
            # If no explicit type, we don't add it (let Excel auto-detect)

        # Add types for dynamic columns
        if dynamic_field is not None:
            for key in all_dynamic_keys:
                excel_type = dynamic_field.get_excel_type(key)
                if excel_type is not None:
                    column_types[key] = excel_type

        return column_types
