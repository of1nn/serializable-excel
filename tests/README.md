# Tests for SerializableExcel

This directory contains tests for the SerializableExcel library, including integration tests with SQLAlchemy models.

## Setup

Install test dependencies:

```bash
pip install -r tests/requirements.txt
```

Or install the package with test dependencies:

```bash
pip install -e ".[dev]"
pip install -r tests/requirements.txt
```

## Running Tests

Run all tests:

```bash
pytest tests/
```

Run specific test file:

```bash
pytest tests/test_sqlalchemy_export.py
```

Run with verbose output:

```bash
pytest tests/ -v
```

Run with coverage:

```bash
pytest tests/ --cov=serializable_excel --cov-report=html
```

## Test Structure

- `conftest.py` - Pytest fixtures and configuration
- `models.py` - SQLAlchemy models for testing (User, Product, Order, OrderItem)
- `test_sqlalchemy_export.py` - Tests for exporting SQLAlchemy models to Excel
- `test_sqlalchemy_import.py` - Tests for importing Excel data into SQLAlchemy models
- `test_output/` - Directory where Excel files from tests are saved (not committed to git)

## Test Data Generation

Tests use [Faker](https://faker.readthedocs.io/) to generate realistic test data:

- **Users**: Names, emails, ages
- **Products**: SKUs, names, descriptions, prices, stock quantities
- **Orders**: Order numbers, totals, statuses with relationships to users and products

## Example Test Flow

1. Generate test data using Faker
2. Create SQLAlchemy models and save to temporary database
3. Export to Excel using SerializableExcel export models
4. Verify Excel file/bytes were created
5. Import back from Excel
6. Verify data integrity

