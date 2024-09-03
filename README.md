# Random Data Word Files Generator

This Python script generates 10 Microsoft Word files, each containing a table with randomly generated data and a random paragraph of Lorem Ipsum text. The tables are similar to a predefined format, and the data within the tables varies across the files.

## Project Overview

The script uses the `python-docx` library to create Word documents and the `Faker` library to generate random data such as names, dates of birth, departments, fees, and genders. Each generated Word document includes:
- A title ("Random Data Table").
- A random paragraph of Lorem Ipsum text.
- A table with 8 rows (excluding the header) and 7 columns filled with randomly generated data.

## Features

- Generates 10 separate Word files.
- Each file contains a unique table with different random data.
- Each table includes columns: No., Name, Age, Department, DOA (Date of Admission), Fee, and Gender.
- Includes a randomly generated Lorem Ipsum paragraph in each document.

## Prerequisites

To run this script, you'll need to have Python installed on your system along with the following Python packages:

- `python-docx`
- `Faker`

You can install the required packages using pip:

```bash
pip install python-docx faker
