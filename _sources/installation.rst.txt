Installation
=============

Installing from PyPI
---------------------

.. code-block:: bash

   pip install serializable-excel

Installing from Source
----------------------

Clone the repository and install:

.. code-block:: bash

   git clone https://github.com/yourusername/serializable-excel.git
   cd serializable-excel
   pip install -r requirements.txt

Development Installation
-------------------------

For development, install with editable mode:

.. code-block:: bash

   pip install -e .

Requirements
------------

* Python 3.8 or higher
* Pydantic 2.x
* openpyxl or pandas (for Excel file handling)

Installing Documentation Dependencies
---------------------------------------

To build the documentation locally:

.. code-block:: bash

   pip install -r requirements-docs.txt

Then build the documentation:

.. code-block:: bash

   cd docs
   make html

The documentation will be available in ``docs/_build/html/``.

Next Steps
----------

* :doc:`quickstart` - Get started with SerializableExcel

