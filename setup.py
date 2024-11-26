# -*- coding: utf-8 -*-
from setuptools import setup, find_packages

setup(
    name="word_formatter",
    version="0.1",
    packages=find_packages(),
    package_data={
        '': ['*.yaml', '*.json'],
    },
    python_requires='>=3.8',
    install_requires=[
        'python-docx>=0.8.11',
        'openai>=1.0.0',
        'python-dotenv>=0.19.0',
        'pyyaml>=6.0.0',
    ],
) 