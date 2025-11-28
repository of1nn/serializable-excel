Overview
========

SerializableExcel is a user-friendly Python library that provides seamless bidirectional conversion between Excel spreadsheets and Pydantic models using declarative syntax similar to SQLAlchemy.

Purpose
-------

The goal of this project is to create a user-friendly interface for interacting with Excel sheets and Pydantic Python models by providing familiar ways to create Declare Models.

Key Features
------------

* **Excel to Pydantic Models**: Seamlessly convert Excel sheets to Pydantic models with automatic type inference
* **Pydantic to Excel**: Export Pydantic model instances to Excel files with proper formatting
* **Declarative Model Definition**: Define models using familiar declarative syntax similar to SQLAlchemy
* **Type Validation**: Automatic validation of data types and constraints when reading from Excel
* **Bidirectional Conversion**: Easy conversion between Excel data and Python objects
* **Dynamic Columns Support**: Models support dynamic columns that can be detected in `.from_excel()` method when explicitly specified. This is useful when administrators add new characteristics or fields that need to be uploaded without modifying the model definition.
* **Custom Validators**: Define custom validation functions for both static and dynamic columns to ensure data integrity
* **Custom Getters**: Define getter functions to extract values from database models when exporting to Excel
* **Cell Styling**: Conditionally style Excel cells based on cell values and row data, perfect for highlighting changes, errors, or important information

Use Cases
---------

* Importing configuration data from Excel spreadsheets
* Exporting application data to Excel for reporting
* Data validation and transformation pipelines
* Creating data models from existing Excel templates
* Working with dynamic/expandable data structures where administrators can add new columns (e.g., configurable forecast characteristics)

Technology Stack
----------------

* Python 3.x
* Pydantic (for data validation and models)
* openpyxl or pandas (for Excel file handling)

Next Steps
----------

* :doc:`installation` - Install SerializableExcel
* :doc:`quickstart` - Get started in 5 minutes
* :doc:`examples` - See practical examples

