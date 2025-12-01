"""
Pytest configuration and fixtures.
"""

import os
import tempfile
from pathlib import Path

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from tests.models import Base

# Directory for test output files
TEST_OUTPUT_DIR = Path(__file__).parent / "test_output"


@pytest.fixture(scope="session", autouse=True)
def setup_test_output_dir():
    """Create test output directory if it doesn't exist."""
    TEST_OUTPUT_DIR.mkdir(exist_ok=True)
    yield
    # Optionally clean up old files on session end
    # Uncomment if you want to clean up:
    # for file in TEST_OUTPUT_DIR.glob("*.xlsx"):
    #     file.unlink()


@pytest.fixture(scope="function")
def db_session():
    """Create a temporary SQLite database session for testing."""
    # Create temporary database file
    db_fd, db_path = tempfile.mkstemp(suffix=".db")
    os.close(db_fd)

    # Create engine and session
    engine = create_engine(f"sqlite:///{db_path}", echo=False)
    Base.metadata.create_all(engine)
    SessionLocal = sessionmaker(bind=engine)
    session = SessionLocal()

    yield session

    # Cleanup
    session.close()
    engine.dispose()
    os.unlink(db_path)


@pytest.fixture
def temp_excel_file(request):
    """Create an Excel file path for testing. Files are saved in tests/test_output/."""
    # Generate filename from test name
    test_name = request.node.name.replace("test_", "").replace("[", "_").replace("]", "_")
    file_name = f"{test_name}.xlsx"
    file_path = TEST_OUTPUT_DIR / file_name

    # Ensure directory exists
    TEST_OUTPUT_DIR.mkdir(exist_ok=True)

    yield str(file_path)

    # Files are kept, not deleted
    # Uncomment the following lines if you want to clean up:
    # if file_path.exists():
    #     file_path.unlink()
