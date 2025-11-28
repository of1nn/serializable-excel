Advanced Usage
==============

This section covers advanced features of SerializableExcel, including dynamic columns, custom validators, and getters.

Dynamic Columns
---------------

Models support dynamic columns that are detected at runtime in the ``.from_excel()`` method. This feature is designed for scenarios where administrators can add new characteristics or fields to be uploaded via Excel, without requiring changes to the model definition.

How Dynamic Columns Work
~~~~~~~~~~~~~~~~~~~~~~~~

1. **Static Columns**: Defined in the model class using ``Column()`` descriptor
2. **Dynamic Columns**: Additional columns in Excel that are not defined in the model
3. **Detection**: When ``dynamic_columns=True`` is passed to ``.from_excel()``, the method scans Excel headers and captures columns not matching static field definitions
4. **Storage**: Dynamic column values are stored in a dictionary field marked with ``DynamicColumn()``

Example: Forecast with Dynamic Characteristics
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Consider a forecasting system where each forecast has static fields (month, curator, manager, object) and dynamic characteristics that administrators can configure:

.. code-block:: python

   import enum
   from serializable_excel import ExcelModel, Column, DynamicColumn

   class ForecastValueTypeEnum(enum.Enum):
       """Value types for forecast characteristics"""
       INTEGER = 'Integer'
       STRING = 'String'

   class ForecastModel(ExcelModel):
       # Static fields - always present in Excel
       month: str = Column(header="Month")
       curator: str = Column(header="Curator")
       manager: str = Column(header="Manager")
       object_title: str = Column(header="Object")
       
       # Dynamic characteristics - detected from Excel headers at runtime
       characteristics: dict = DynamicColumn()

   # Read forecasts with dynamic characteristics
   forecasts = ForecastModel.from_excel(
       "forecasts.xlsx",
       dynamic_columns=True  # Enable dynamic column detection
   )

   # Each forecast now contains:
   # - Static fields: month, curator, manager, object_title
   # - characteristics dict with dynamically detected columns
   # Example: {"Sales Volume": "1000", "Priority": "High", "Region": "North"}

Use Case: Admin-Configurable Characteristics
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

This pattern is useful in systems like:

* **Forecast Management**: Administrators add new forecast characteristics (e.g., "Sales Volume", "Priority", "Region") via a web interface
* **Product Catalogs**: Dynamic product attributes that vary by category
* **Survey Data**: Questions that can be added or modified by administrators
* **Configuration Import**: Flexible configuration files with optional fields

Excel file example:

+----------+----------+----------+--------+---------------+----------+--------+
| Month    | Curator  | Manager  | Object | Sales Volume  | Priority | Region |
+==========+==========+==========+========+===============+==========+========+
| 2024-01  | John     | Alice    | Bldg A | 1000          | High     | North  |
+----------+----------+----------+--------+---------------+----------+--------+
| 2024-02  | John     | Bob      | Bldg B | 1500          | Medium   | South  |
+----------+----------+----------+--------+---------------+----------+--------+

The last three columns (Sales Volume, Priority, Region) are dynamic and will be captured in the ``characteristics`` dictionary.

Column Validators
-----------------

You can attach custom validation functions to any column using the ``validator`` parameter. Validators are called during ``.from_excel()`` to ensure data integrity:

.. code-block:: python

   from serializable_excel import ExcelModel, Column

   def validate_email(value: str) -> str:
       """Validate email format"""
       if value and '@' not in value:
           raise ValueError(f"Invalid email format: {value}")
       return value.strip().lower()

   def validate_age(value: int) -> int:
       """Validate age is within reasonable range"""
       if value < 0 or value > 150:
           raise ValueError(f"Age must be between 0 and 150, got: {value}")
       return value

   class UserModel(ExcelModel):
       name: str = Column(header="Name")
       age: int = Column(header="Age", validator=validate_age)
       email: str = Column(header="Email", validator=validate_email)

Column Getters
--------------

When exporting to Excel using ``.to_excel()``, you may need to extract values from complex database models. Use the ``getter`` parameter to define how values are retrieved:

.. code-block:: python

   from serializable_excel import ExcelModel, Column

   def get_curator_name(model) -> str:
       """Extract curator name from relationship"""
       return model.curator.full_name if model.curator else ""

   def get_object_title(model) -> str:
       """Extract object title, fallback to manual title if object is None"""
       if model.object:
           return model.object.title
       return model.object_title or ""

   class ForecastExportModel(ExcelModel):
       month: str = Column(header="Month")
       curator: str = Column(header="Curator", getter=get_curator_name)
       manager: str = Column(header="Manager", getter=lambda m: m.manager.full_name)
       object_title: str = Column(header="Object", getter=get_object_title)

   # Export database models to Excel
   forecasts = db.query(Forecast).all()
   ForecastExportModel.to_excel(forecasts, "export.xlsx")

Dynamic Column Validators
--------------------------

For dynamic columns, you can pass a validator function or a dictionary of validators per column name. The validator receives both the column name and value:

Single Validator Function
~~~~~~~~~~~~~~~~~~~~~~~~~

Using a single validator function for all dynamic columns:

.. code-block:: python

   from serializable_excel import ExcelModel, Column, DynamicColumn

   def validate_dynamic_value(column_name: str, value: str, value_type: str) -> str:
       """Validate dynamic column value based on expected type"""
       if value_type == 'INTEGER':
           try:
               int(value)
           except ValueError:
               raise ValueError(f"Column '{column_name}' expects integer, got: {value}")
       return value

   class ForecastModel(ExcelModel):
       month: str = Column(header="Month")
       characteristics: dict = DynamicColumn(validator=validate_dynamic_value)

Per-Column Validators
~~~~~~~~~~~~~~~~~~~~~

Or using a dictionary of validators per column name:

.. code-block:: python

   from serializable_excel import ExcelModel, Column, DynamicColumn

   validators = {
       "Sales Volume": lambda name, val: int(val),  # Convert to int
       "Priority": lambda name, val: val if val in ["High", "Medium", "Low"] 
                                     else (_ for _ in ()).throw(ValueError(f"Invalid priority: {val}")),
   }

   class ForecastModelWithTypedValidators(ExcelModel):
       month: str = Column(header="Month")
       characteristics: dict = DynamicColumn(validators=validators)

Validating with Database Metadata
----------------------------------

For advanced scenarios where dynamic column types are stored in a database, you can build validators at runtime:

.. code-block:: python

   from serializable_excel import ExcelModel, Column, DynamicColumn

   # Fetch characteristic definitions from database
   characteristics = db.query(ForecastCharacteristic).all()

   # Build validators based on database metadata
   def build_validators(characteristics):
       validators = {}
       for char in characteristics:
           if char.value_type == ForecastValueTypeEnum.INTEGER:
               validators[char.title] = lambda name, val: int(val) if val else None
           else:
               validators[char.title] = lambda name, val: str(val) if val else ""
       return validators

   class ForecastModel(ExcelModel):
       month: str = Column(header="Month")
       curator: str = Column(header="Curator")
       manager: str = Column(header="Manager")
       object_title: str = Column(header="Object")
       characteristics: dict = DynamicColumn(validators=build_validators(characteristics))

   # Now when reading from Excel, each dynamic column is validated
   # according to its type defined in the database
   forecasts = ForecastModel.from_excel("forecasts.xlsx", dynamic_columns=True)

Cell Styling
------------

SerializableExcel allows you to conditionally style Excel cells based on cell values and row data. This is useful for highlighting changes, errors, or important information in exported Excel files.

How Cell Styling Works
~~~~~~~~~~~~~~~~~~~~~~

1. **Define a style function**: Create a function that receives cell context and returns a ``CellStyle`` object
2. **Attach to column**: Use ``getter_cell_color`` parameter in ``Column()`` or ``DynamicColumn()``
3. **Automatic application**: Styles are automatically applied when exporting to Excel

Style Function Signature
~~~~~~~~~~~~~~~~~~~~~~~~

The style function receives four parameters:

.. code-block:: python

   def style_function(cell_value, row_data, column_name, row_index):
       """
       Args:
           cell_value: Value of the current cell
           row_data: Dict[str, Any] with all values in the row {header: value}
           column_name: Header name of the current column
           row_index: Row index (0-based, data rows only)
       Returns:
           CellStyle or None (no styling)
       """
       # Your styling logic here
       return CellStyle(fill_color=Colors.CHANGED)

Example: Highlighting Changes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Highlight cells based on value comparison with previous data:

.. code-block:: python

   from serializable_excel import ExcelModel, Column, CellStyle, Colors

   def highlight_changes(cell_value, row_data, column_name, row_index):
       """Highlight if value changed from previous export"""
       # Get previous value from database or cache
       user_id = row_data.get("ID")
       old_value = get_previous_value(user_id, column_name)
       
       if old_value is None:
           return CellStyle(fill_color=Colors.NEW)  # New entry
       if old_value == cell_value:
           return CellStyle(fill_color=Colors.UNCHANGED)  # Unchanged
       return CellStyle(fill_color=Colors.CHANGED)  # Changed

   class UserModel(ExcelModel):
       id: int = Column(header="ID")
       name: str = Column(header="Name", getter_cell_color=highlight_changes)
       age: int = Column(header="Age", getter_cell_color=highlight_changes)

   UserModel.to_excel(users, "output.xlsx")

Example: Conditional Formatting
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Apply different styles based on cell values:

.. code-block:: python

   from serializable_excel import ExcelModel, Column, CellStyle, Colors

   def highlight_age(cell_value, row_data, column_name, row_index):
       """Highlight age based on value"""
       if cell_value and cell_value > 30:
           return CellStyle(fill_color=Colors.WARNING, font_bold=True)
       return CellStyle(fill_color=Colors.UNCHANGED)

   def highlight_email(cell_value, row_data, column_name, row_index):
       """Highlight email if contains 'example'"""
       if cell_value and "example" in cell_value:
           return CellStyle(fill_color=Colors.INFO, font_italic=True)
       return None

   class UserModel(ExcelModel):
       name: str = Column(header="Name")
       age: int = Column(header="Age", getter_cell_color=highlight_age)
       email: str = Column(header="Email", getter_cell_color=highlight_email)

Dynamic Column Styling
~~~~~~~~~~~~~~~~~~~~~~

You can also style dynamic columns:

.. code-block:: python

   from serializable_excel import ExcelModel, Column, DynamicColumn, CellStyle, Colors

   def highlight_dynamic(cell_value, row_data, column_name, row_index):
       """Highlight dynamic columns with high values"""
       if cell_value and str(cell_value).isdigit() and int(cell_value) > 100:
           return CellStyle(fill_color=Colors.CHANGED, font_bold=True)
       return None

   class ForecastModel(ExcelModel):
       month: str = Column(header="Month")
       characteristics: dict = DynamicColumn(getter_cell_color=highlight_dynamic)

   # Or use per-column styling
   style_getters = {
       "Sales Volume": lambda val, row, col, idx: CellStyle(fill_color=Colors.CHANGED) if val and int(val) > 100 else None,
       "Priority": lambda val, row, col, idx: CellStyle(font_bold=True) if val == "High" else None,
   }
   class ForecastModel(ExcelModel):
       characteristics: dict = DynamicColumn(getters_cell_color=style_getters)

Available Colors
~~~~~~~~~~~~~~~~

Use predefined colors from ``Colors`` class:

* ``Colors.CHANGED`` - Yellow for changed values
* ``Colors.UNCHANGED`` - Light green for unchanged values
* ``Colors.ERROR`` - Light red for errors
* ``Colors.WARNING`` - Orange for warnings
* ``Colors.INFO`` - Light blue for information
* ``Colors.NEW`` - Pale green for new entries

Or use custom HEX colors:

.. code-block:: python

   style = CellStyle(fill_color="FF0000")  # Red background
   style = CellStyle(font_color="0000FF")  # Blue text

Column Parameters Summary
-------------------------

.. list-table:: Column Parameters
   :header-rows: 1
   :widths: 20 50 30

   * - Parameter
     - Description
     - Applies To
   * - ``header``
     - Excel column header name
     - Column
   * - ``validator``
     - Function to validate/transform value when reading
     - Column, DynamicColumn
   * - ``validators``
     - Dictionary of validators per column name
     - DynamicColumn
   * - ``getter``
     - Function to extract value from model when writing
     - Column
   * - ``getter_cell_color``
     - Function to determine cell style when writing
     - Column, DynamicColumn
   * - ``getters_cell_color``
     - Dictionary of style getters per column name
     - DynamicColumn
   * - ``default``
     - Default value if cell is empty
     - Column
   * - ``required``
     - Raise error if value is missing
     - Column

Next Steps
----------

* :doc:`api` - Full API reference
* :doc:`examples` - More practical examples

