from setuptools import setup


setup(
    name="official-docx-formatter",
    version="0.1.0",
    description="Local-first Chinese official DOCX formatting helper.",
    packages=["official_docx_engine"],
    package_dir={"": "scripts"},
    install_requires=["python-docx>=1.1.0"],
    extras_require={"dev": ["pytest>=8.0"]},
    python_requires=">=3.9",
)
