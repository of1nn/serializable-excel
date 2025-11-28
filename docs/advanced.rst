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

Column Parameters Summary
-------------------------

+-------------+--------------------------------------------------+------------------+
| Parameter   | Description                                      | Applies To       |
+=============+==================================================+==================+
| ``header``  | Excel column header name                         | Column           |
+-------------+--------------------------------------------------+------------------+
| ``validator``| Function to validate/transform value when       | Column,          |
|             | reading                                          | DynamicColumn    |
+-------------+--------------------------------------------------+------------------+
| ``validators``| Dict of validators per column name            | DynamicColumn    |
+-------------+--------------------------------------------------+------------------+
| ``getter``  | Function to extract value from model when       | Column           |
|             | writing                                          |                  |
+-------------+--------------------------------------------------+------------------+
| ``default`` | Default value if cell is empty                   | Column           |
+-------------+--------------------------------------------------+------------------+
| ``required``| Raise error if value is missing                  | Column           |
+-------------+--------------------------------------------------+------------------+

Next Steps
----------

* :doc:`api` - Full API reference
* :doc:`examples` - More practical examples

