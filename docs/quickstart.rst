Quick Start
===========

This guide will help you get started with SerializableExcel in just a few minutes.

Basic Example
-------------

Let's start with a simple example. Define a model by inheriting from ``ExcelModel`` and declaring fields with ``Column`` descriptors:

.. code-block:: python

   from serializable_excel import ExcelModel, Column

   class UserModel(ExcelModel):
       name: str = Column(header="Name")
       age: int = Column(header="Age")
       email: str = Column(header="Email")

Reading from Excel
------------------

Read data from an Excel file:

.. code-block:: python

   users = UserModel.from_excel("users.xlsx")
   
   # users is now a list of UserModel instances
   for user in users:
       print(f"{user.name} ({user.age}): {user.email}")

Writing to Excel
----------------

Export model instances to an Excel file:

.. code-block:: python

   # Create some instances
   users = [
       UserModel(name="Alice", age=30, email="alice@example.com"),
       UserModel(name="Bob", age=25, email="bob@example.com"),
   ]
   
   # Export to Excel
   UserModel.to_excel(users, "output.xlsx")

Complete Example
----------------

Here's a complete working example:

.. code-block:: python

   from serializable_excel import ExcelModel, Column

   class ProductModel(ExcelModel):
       name: str = Column(header="Product Name")
       price: float = Column(header="Price")
       quantity: int = Column(header="Quantity")
       in_stock: bool = Column(header="In Stock")

   # Read products from Excel
   products = ProductModel.from_excel("products.xlsx")
   
   # Process the data
   for product in products:
       if product.in_stock:
           print(f"{product.name}: ${product.price}")
   
   # Export to Excel
   ProductModel.to_excel(products, "exported_products.xlsx")

Next Steps
----------

* :doc:`examples` - See more practical examples
* :doc:`advanced` - Learn about dynamic columns and validators
* :doc:`api` - Full API reference

