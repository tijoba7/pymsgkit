from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="pymsgkit",
    version="1.0.0",
    author="PyMsgKit Contributors",
    author_email="your.email@example.com",
    description="Pure Python library for creating Outlook MSG files",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/pymsgkit",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Intended Audience :: Legal Industry",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Communications :: Email",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    keywords="msg outlook email mapi ediscovery forensics cfb",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/pymsgkit/issues",
        "Source": "https://github.com/yourusername/pymsgkit",
        "Documentation": "https://github.com/yourusername/pymsgkit#readme",
    },
)
