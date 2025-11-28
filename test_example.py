"""
Test script to verify SerializableExcel functionality.
"""

from serializable_excel import (
    CellStyle,
    Colors,
    Column,
    DynamicColumn,
    ExcelModel,
)


# Function to highlight age
def highlight_age(cell_value, row_data, column_name, row_index):
    """Highlights cell based on age."""
    if cell_value and cell_value > 30:
        return CellStyle(fill_color=Colors.WARNING)  # Orange for >30
    return CellStyle(fill_color=Colors.UNCHANGED)  # Green


# Function to highlight email
def highlight_email(cell_value, row_data, column_name, row_index):
    """Highlights email if it contains 'example'."""
    if cell_value and "example" in cell_value:
        return CellStyle(fill_color=Colors.INFO, font_italic=True)
    return None


# Define model
class UserModel(ExcelModel):
    name: str = Column(header="Name")
    age: int = Column(header="Age", getter_cell_color=highlight_age)
    email: str = Column(header="Email", getter_cell_color=highlight_email)


def test_basic_export_import():
    """Test basic export and import."""
    print("=" * 50)
    print("Test 1: Basic Export and Import")
    print("=" * 50)

    # Create test data
    users = [
        UserModel(name="Alice", age=25, email="alice@example.com"),
        UserModel(name="Bob", age=35, email="bob@example.com"),
        UserModel(name="Charlie", age=28, email="charlie@test.org"),
    ]

    # Export to Excel
    output_file = "test_output.xlsx"
    UserModel.to_excel(users, output_file)
    print(f"[+] File {output_file} created!")

    # Read back
    loaded_users = UserModel.from_excel(output_file)
    print(f"[+] Loaded {len(loaded_users)} records:")
    for user in loaded_users:
        print(f"  - {user.name} ({user.age}): {user.email}")

    print()


def test_dynamic_columns():
    """Test dynamic columns."""
    print("=" * 50)
    print("Test 2: Dynamic Columns")
    print("=" * 50)

    # Function to highlight dynamic columns
    def highlight_dynamic(cell_value, row_data, column_name, row_index):
        if cell_value and str(cell_value).isdigit() and int(cell_value) > 100:
            return CellStyle(fill_color=Colors.CHANGED, font_bold=True)
        return None

    class ForecastModel(ExcelModel):
        month: str = Column(header="Month")
        manager: str = Column(header="Manager")
        characteristics: dict = DynamicColumn(
            getter_cell_color=highlight_dynamic
        )

    # Create data with dynamic columns
    forecasts = [
        ForecastModel(
            month="2024-01",
            manager="Alice",
            characteristics={"Sales": 150, "Priority": "High"},
        ),
        ForecastModel(
            month="2024-02",
            manager="Bob",
            characteristics={"Sales": 80, "Priority": "Medium"},
        ),
    ]

    output_file = "test_forecast.xlsx"
    ForecastModel.to_excel(forecasts, output_file)
    print(f"[+] File {output_file} created!")

    # Read with dynamic columns
    loaded = ForecastModel.from_excel(output_file, dynamic_columns=True)
    print(f"[+] Loaded {len(loaded)} records:")
    for f in loaded:
        print(f"  - {f.month} ({f.manager}): {f.characteristics}")

    print()


def test_validation():
    """Test validation."""
    print("=" * 50)
    print("Test 3: Validation")
    print("=" * 50)

    def validate_positive(value):
        if value < 0:
            raise ValueError("Age must be positive")
        return value

    class ValidatedModel(ExcelModel):
        name: str = Column(header="Name", required=True)
        age: int = Column(header="Age", validator=validate_positive, default=0)

    users = [
        ValidatedModel(name="Test User", age=25),
    ]

    output_file = "test_validated.xlsx"
    ValidatedModel.to_excel(users, output_file)
    print(f"[+] File {output_file} created!")

    loaded = ValidatedModel.from_excel(output_file)
    print(f"[+] Validation passed: {loaded[0].name}, {loaded[0].age}")

    print()


if __name__ == "__main__":
    print("\n>>> Running SerializableExcel Tests\n")

    test_basic_export_import()
    test_dynamic_columns()
    test_validation()

    print("=" * 50)
    print("[OK] All tests completed!")
    print("Check the created files:")
    print("  - test_output.xlsx (with age and email highlighting)")
    print("  - test_forecast.xlsx (with dynamic columns)")
    print("  - test_validated.xlsx (with validation)")
    print("=" * 50)
