Examples
========

This page contains practical examples of using SerializableExcel in various scenarios.

Example 1: Import Configuration Data
-------------------------------------

Import application settings from an Excel spreadsheet:

.. code-block:: python

   from serializable_excel import ExcelModel, Column

   class ConfigModel(ExcelModel):
       key: str = Column(header="Config Key")
       value: str = Column(header="Value")
       environment: str = Column(header="Environment")

   # Read configuration from Excel
   configs = ConfigModel.from_excel("config.xlsx")
   
   # Use the configuration
   for config in configs:
       print(f"{config.key} = {config.value} ({config.environment})")

Example 2: Export Report Data
------------------------------

Export application data to Excel for reporting:

.. code-block:: python

   from serializable_excel import ExcelModel, Column

   class ReportModel(ExcelModel):
       date: str = Column(header="Date")
       revenue: float = Column(header="Revenue")
       expenses: float = Column(header="Expenses")
       
       @property
       def profit(self):
           return self.revenue - self.expenses

   # Create report data
   reports = [
       ReportModel(date="2024-01", revenue=10000, expenses=5000),
       ReportModel(date="2024-02", revenue=12000, expenses=6000),
   ]
   
   # Export to Excel
   ReportModel.to_excel(reports, "report.xlsx")

Example 3: Forecast with Dynamic Characteristics
-------------------------------------------------

Work with dynamic columns that administrators can configure:

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

   # Excel file with additional columns:
   # Month | Curator | Manager | Object | Sales Volume | Priority | Region
   forecasts = ForecastModel.from_excel("forecasts.xlsx", dynamic_columns=True)
   
   # Each forecast contains:
   # - Static fields: month, curator, manager, object_title
   # - characteristics dict: {"Sales Volume": "1000", "Priority": "High", "Region": "North"}
   for forecast in forecasts:
       print(f"Forecast for {forecast.month}: {forecast.characteristics}")

Example 4: Data Validation Pipeline
------------------------------------

Build a validation pipeline with custom business rules:

.. code-block:: python

   from serializable_excel import ExcelModel, Column

   def validate_email(value: str) -> str:
       """Validate and normalize email"""
       if value and '@' not in value:
           raise ValueError(f"Invalid email: {value}")
       return value.strip().lower()

   def validate_age(value: int) -> int:
       """Validate age range"""
       if not (0 <= value <= 150):
           raise ValueError(f"Age must be 0-150, got: {value}")
       return value

   class UserModel(ExcelModel):
       name: str = Column(header="Name", required=True)
       age: int = Column(header="Age", validator=validate_age)
       email: str = Column(header="Email", validator=validate_email)

   # Read and validate data
   try:
       users = UserModel.from_excel("users.xlsx")
       print(f"Successfully imported {len(users)} users")
   except ValueError as e:
       print(f"Validation error: {e}")

Example 5: Export Database Models
----------------------------------

Export SQLAlchemy models to Excel with relationship handling:

.. code-block:: python

   from serializable_excel import ExcelModel, Column

   def get_curator_name(model) -> str:
       """Extract curator name from relationship"""
       return model.curator.full_name if model.curator else ""

   def get_object_title(model) -> str:
       """Get object title with fallback"""
       return model.object.title if model.object else model.object_title or ""

   class ForecastExportModel(ExcelModel):
       month: str = Column(header="Month")
       curator: str = Column(header="Curator", getter=get_curator_name)
       manager: str = Column(header="Manager", getter=lambda m: m.manager.full_name)
       object_title: str = Column(header="Object", getter=get_object_title)

   # Export SQLAlchemy models
   from sqlalchemy.orm import Session
   forecasts = db.query(Forecast).all()
   ForecastExportModel.to_excel(forecasts, "export.xlsx")

Example 6: Cell Styling and Highlighting
-----------------------------------------

Export data with conditional cell styling to highlight important information:

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

   users = [
       UserModel(name="Alice", age=25, email="alice@example.com"),
       UserModel(name="Bob", age=35, email="bob@example.com"),
   ]

   UserModel.to_excel(users, "highlighted_users.xlsx")
   # Age > 30 will be orange and bold
   # Emails with 'example' will be light blue and italic

Next Steps
----------

* :doc:`advanced` - Learn about advanced features like dynamic columns, validators, and cell styling
* :doc:`api` - Full API reference

