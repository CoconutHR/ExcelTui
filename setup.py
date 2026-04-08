from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="exceltui",
    version="2.0.0",
    author="coco",
    author_email="excel.tui@coco",
    description="Excel通用数据处理工具 - 支持JSON/键值对/纯文本格式、去重、筛选交集",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/CoconutHR/exceltui",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    install_requires=[
        "pandas>=1.3.0",
        "openpyxl>=3.0.0",
        "rich>=10.0.0",
    ],
    entry_points={
        "console_scripts": [
            "exceltui=exceltui.main:main",
        ],
    },
)