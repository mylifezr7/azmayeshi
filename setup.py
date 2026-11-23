from setuptools import setup, find_packages

setup(
    name="excel-assistant",
    version="1.0.0",
    description="Persian AI Assistant for Microsoft Excel 2024",
    author="Senior Software & AI Engineer",
    packages=find_packages(),
    include_package_data=True,
    python_requires=">=3.9",
    install_requires=[
        "xlwings>=0.30.0",
        "openpyxl>=3.1.2",
        "pydantic>=2.0.0",
        "PySide6>=6.5.0",
        "requests>=2.28.0",
        "pytest>=7.0.0",
        "reportlab>=4.0.0",
    ],
    entry_points={
        "console_scripts": [
            "excel-assistant=excel_assistant.main:main",
            "excel-ai=excel_assistant.main:main",
        ],
    },
)
