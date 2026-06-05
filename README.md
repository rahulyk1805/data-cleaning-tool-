 data-cleaning-tool-
Data Cleaning Tool

## Overview

Data Cleaning Tool is a Python-based automation solution designed to streamline the process of cleaning, validating, and standardizing Excel reports. The application reduces manual effort by automatically identifying report types, applying predefined business rules, and generating structured outputs for further analysis.

The system is built to handle multiple business report formats and can be easily extended to support additional file types and validation rules as organizational requirements evolve.

---

## Key Features

* Automatic report type detection
* Excel-based data processing
* Business-rule driven validation
* Data standardization and cleansing
* Multi-file processing support
* Manual file type override option
* Change tracking and audit logging
* Summary report generation
* Scalable architecture for adding new report types

---

## Supported Report Categories

The tool currently supports multiple business report formats, including:

* Sales Reports
* Purchase Reports
* Stock Reports
* Journal Voucher (JV) Reports
* Sundry Creditors Reports
* Salary Validation Reports
* Additional custom report types as required

Each report type is processed using a dedicated set of validation and cleaning rules tailored to its business purpose.

---

## Workflow

### 1. File Upload

Users upload one or more Excel files through the application interface.

### 2. File Identification

The system analyzes report contents and automatically determines the appropriate report category.

### 3. Data Processing

Relevant validation and cleaning rules are applied based on the identified report type.

### 4. Review & Verification

All modifications performed by the system are recorded for transparency and auditing purposes.

### 5. Output Generation

The application generates cleaned files along with supporting reports for user review.

---

## Processing Pipeline

```text
Upload File
      │
      ▼
Identify Report Type
      │
      ▼
Apply Validation Rules
      │
      ▼
Clean & Standardize Data
      │
      ▼
Generate Outputs
      │
      ▼
Download Results
```

---

## Outputs Generated

For each processed file, the system generates:

* Cleaned Excel File
* Change Log
* Summary Report

These outputs provide visibility into the modifications performed during processing and help maintain data quality standards.

---

## Summary Report

A summary report is generated for every processed file, providing:

* File information
* Identified report category
* Number of modifications performed
* Overview of applied validation and cleaning actions

This helps users quickly review processing results without manually inspecting the entire dataset.

---

## Benefits

* Reduces manual data preparation effort
* Improves data accuracy and consistency
* Enhances reporting reliability
* Maintains an audit trail of changes
* Supports scalable business operations
* Increases process efficiency

---

## Technologies Used

* Python
* Pandas
* OpenPyXL
* Streamlit

---

## Future Scope

The architecture is designed to support:

* Additional report formats
* Advanced validation rules
* Enhanced reporting capabilities
* Integration with business systems
* Large-scale data processing workflows

---

## Author

**Rahul Yerunkar** 

Data Cleaning & Process Automation Solution
