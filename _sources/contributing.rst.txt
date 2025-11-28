Contributing
============

Thank you for your interest in contributing to SerializableExcel! This document provides guidelines and instructions for contributing.

Getting Started
---------------

1. Fork the repository on GitHub
2. Clone your fork locally:

   .. code-block:: bash

      git clone https://github.com/yourusername/serializable-excel.git
      cd serializable-excel

3. Create a virtual environment:

   .. code-block:: bash

      python -m venv venv
      source venv/bin/activate  # On Windows: venv\Scripts\activate

4. Install development dependencies:

   .. code-block:: bash

      pip install -r requirements.txt
      pip install -r requirements-docs.txt

Development Workflow
--------------------

1. Create a feature branch:

   .. code-block:: bash

      git checkout -b feature/your-feature-name

2. Make your changes
3. Write or update tests
4. Ensure all tests pass:

   .. code-block:: bash

      pytest

5. Update documentation if needed
6. Commit your changes:

   .. code-block:: bash

      git commit -m "Add feature: description of your feature"

7. Push to your fork:

   .. code-block:: bash

      git push origin feature/your-feature-name

8. Open a Pull Request on GitHub

Code Style
----------

* Follow PEP 8 style guide
* Use type hints where appropriate
* Write docstrings for all public functions and classes
* Keep functions focused and small

Testing
-------

* Write tests for new features
* Ensure all existing tests pass
* Aim for good test coverage

Documentation
-------------

* Update relevant documentation when adding features
* Use reStructuredText for Sphinx documentation
* Include code examples in documentation
* Keep examples up-to-date with the code

Pull Request Guidelines
-----------------------

* Provide a clear description of changes
* Reference any related issues
* Ensure all tests pass
* Update documentation as needed
* Keep PRs focused on a single feature or fix

Questions?
----------

If you have questions, please open an issue on GitHub or contact the maintainers.

Thank you for contributing!

