from setuptools import setup, find_packages

with open('./docs/README.md', 'r', encoding='utf-8') as f:
    long_des = f.read()

setup(
    name="cosmicexcelchecker",
    version="0.2.0",
    license="GPL-3.0-only",
    author="TimG233",
    author_email="gaosh0830@gmail.com",
    maintainer="TimG233",
    description="A high flexibility package for checking cosmic-related excels under CMDI cosmic standards",
    long_description=long_des,
    long_description_content_type="text/markdown",
    packages=find_packages(),
    keywords=["cosmic", "excel", "CMDI", "checker", "data"],
    classifiers=[
        "License :: OSI Approved :: GPL-3.0-only"
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
    ],
    install_requires=[
        "numpy~=1.25.0",
        "openpyxl~=3.1.2",
        "pandas~=2.0.0",
        "tabulate~=0.9.0",
        "xlrd~=2.0.0",
    ],
    platforms="any",
    python_requires=">=3.9"
)

