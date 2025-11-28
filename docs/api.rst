API Reference
==============

This page provides detailed API documentation for SerializableExcel.

ExcelModel
----------

Base class for all Excel-serializable models. Inherit from this class to create your models.

Methods
~~~~~~~

.. py:class:: ExcelModel

   Base class for Excel-serializable Pydantic models.

   .. py:method:: from_excel(file_path: str, dynamic_columns: bool = False) -> List[ExcelModel]

      Read models from an Excel file.

      :param file_path: Path to the Excel file (.xlsx)
      :type file_path: str
      :param dynamic_columns: Enable detection of additional columns not defined in model
      :type dynamic_columns: bool
      :returns: List of model instances
      :rtype: List[ExcelModel]
      :raises FileNotFoundError: If the Excel file doesn't exist
      :raises ValueError: If validation fails

      Example:

      .. code-block:: python

         users = UserModel.from_excel("users.xlsx")
         forecasts = ForecastModel.from_excel("forecasts.xlsx", dynamic_columns=True)

   .. py:method:: to_excel(instances: List[ExcelModel], file_path: str) -> None

      Export model instances to an Excel file.

      :param instances: List of model instances to export
      :type instances: List[ExcelModel]
      :param file_path: Path where to save the Excel file (.xlsx)
      :type file_path: str
      :raises ValueError: If instances list is empty or invalid

      Example:

      .. code-block:: python

         users = [UserModel(name="Alice", age=30)]
         UserModel.to_excel(users, "output.xlsx")

Column
------

Descriptor for defining static columns in models.

.. py:class:: Column

   Descriptor for defining static columns in ExcelModel classes.

   .. py:method:: __init__(header: str, validator: Optional[Callable] = None, getter: Optional[Callable] = None, getter_cell_color: Optional[Callable] = None, default: Any = None, required: bool = False)

      Initialize a Column descriptor.

      :param header: Excel column header name (required)
      :type header: str
      :param validator: Function to validate/transform value when reading from Excel. Should accept the value and return the validated/transformed value.
      :type validator: Optional[Callable[[Any], Any]]
      :param getter: Function to extract value from model when writing to Excel. Should accept the model instance and return the value.
      :type getter: Optional[Callable[[Any], Any]]
      :param getter_cell_color: Function to determine cell style when writing to Excel. Signature: (cell_value, row_data, column_name, row_index) -> Optional[CellStyle]
      :type getter_cell_color: Optional[Callable[[Any, Dict[str, Any], str, int], Optional[CellStyle]]]
      :param default: Default value if cell is empty
      :type default: Any
      :param required: Raise error if value is missing
      :type required: bool

      Example:

      .. code-block:: python

         from serializable_excel import ExcelModel, Column, CellStyle, Colors

         def highlight_age(cell_value, row_data, column_name, row_index):
             if cell_value and cell_value > 30:
                 return CellStyle(fill_color=Colors.WARNING)
             return CellStyle(fill_color=Colors.UNCHANGED)

         class UserModel(ExcelModel):
             name: str = Column(header="Name", required=True)
             age: int = Column(header="Age", validator=int, default=0, getter_cell_color=highlight_age)
             email: str = Column(header="Email", validator=validate_email)

DynamicColumn
-------------

Descriptor for defining dynamic columns that are detected at runtime.

.. py:class:: DynamicColumn

   Descriptor for defining dynamic columns that are detected at runtime in Excel files.

   .. py:method:: __init__(validator: Optional[Callable[[str, str], Any]] = None, validators: Optional[Dict[str, Callable]] = None, getter_cell_color: Optional[Callable] = None, getters_cell_color: Optional[Dict[str, Callable]] = None)

      Initialize a DynamicColumn descriptor.

      :param validator: Function to validate all dynamic columns. Receives (column_name: str, value: str) and returns validated value.
      :type validator: Optional[Callable[[str, str], Any]]
      :param validators: Dictionary mapping column names to validator functions. Each validator receives (column_name: str, value: str) and returns validated value.
      :type validators: Optional[Dict[str, Callable[[str, str], Any]]]
      :param getter_cell_color: Function to determine cell style for all dynamic columns. Signature: (cell_value, row_data, column_name, row_index) -> Optional[CellStyle]
      :type getter_cell_color: Optional[Callable[[Any, Dict[str, Any], str, int], Optional[CellStyle]]]
      :param getters_cell_color: Dictionary mapping column names to style getter functions. Each function has same signature as getter_cell_color.
      :type getters_cell_color: Optional[Dict[str, Callable[[Any, Dict[str, Any], str, int], Optional[CellStyle]]]]

      Example:

      .. code-block:: python

         from serializable_excel import ExcelModel, Column, DynamicColumn, CellStyle, Colors

         # Single validator for all dynamic columns
         class ForecastModel(ExcelModel):
             characteristics: dict = DynamicColumn(validator=validate_dynamic)

         # Per-column validators
         validators = {
             "Sales Volume": lambda name, val: int(val),
             "Priority": lambda name, val: val.upper(),
         }
         class ForecastModel(ExcelModel):
             characteristics: dict = DynamicColumn(validators=validators)

         # With cell styling
         def highlight_dynamic(cell_value, row_data, column_name, row_index):
             if cell_value and str(cell_value).isdigit() and int(cell_value) > 100:
                 return CellStyle(fill_color=Colors.CHANGED, font_bold=True)
             return None

         class ForecastModel(ExcelModel):
             characteristics: dict = DynamicColumn(getter_cell_color=highlight_dynamic)

CellStyle and Colors
--------------------

.. py:class:: CellStyle

   Represents cell styling options for Excel cells.

   .. py:attribute:: fill_color

      Background color in HEX format (e.g., "FF0000" for red)

   .. py:attribute:: font_color

      Font color in HEX format

   .. py:attribute:: font_bold

      Whether the font should be bold

   .. py:attribute:: font_italic

      Whether the font should be italic

   Example:

   .. code-block:: python

      from serializable_excel import CellStyle, Colors

      style = CellStyle(
          fill_color=Colors.CHANGED,
          font_color=Colors.FONT_BLACK,
          font_bold=True
      )

.. py:class:: Colors

   Predefined color constants for common use cases.

   Background colors:
   
   * ``CHANGED`` - Yellow (#FFFF00) - for changed values
   * ``UNCHANGED`` - Light green (#90EE90) - for unchanged values
   * ``ERROR`` - Light red (#FF6B6B) - for errors
   * ``WARNING`` - Orange (#FFA500) - for warnings
   * ``INFO`` - Light blue (#87CEEB) - for information
   * ``NEW`` - Pale green (#98FB98) - for new entries

   Font colors:
   
   * ``FONT_RED`` - Red (#FF0000)
   * ``FONT_GREEN`` - Green (#008000)
   * ``FONT_BLUE`` - Blue (#0000FF)
   * ``FONT_BLACK`` - Black (#000000)
   * ``FONT_GRAY`` - Gray (#808080)

Exceptions
----------

.. py:exception:: ExcelModelError

   Base exception for all SerializableExcel errors.

.. py:exception:: ValidationError

   Raised when data validation fails during ``from_excel()``.

   Inherits from :py:exc:`ValueError`.

.. py:exception:: ColumnNotFoundError

   Raised when a required column is not found in the Excel file.

   Inherits from :py:exc:`ExcelModelError`.

Type Hints
----------

For better IDE support and type checking, SerializableExcel provides type hints:

.. code-block:: python

   from typing import List
   from serializable_excel import ExcelModel, Column

   class UserModel(ExcelModel):
       name: str = Column(header="Name")
       age: int = Column(header="Age")

   # Type hints work correctly
   users: List[UserModel] = UserModel.from_excel("users.xlsx")

Related Documentation
---------------------

* :doc:`quickstart` - Quick start guide
* :doc:`examples` - Usage examples
* :doc:`advanced` - Advanced features

