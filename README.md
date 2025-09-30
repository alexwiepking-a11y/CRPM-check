# CRPM Analysis Tool

## Description
This tool analyzes CRPM (rate plan) data to identify compliance issues with VAT, subaccounts, and city tax settings across multiple hotels. It automatically checks rate plans against hotel standards, manages exception rules for approved deviations, and generates prioritized reports with actionable insights.

## Features
- Automated compliance checking against hotel standards
- Intelligent exception rule system for approved deviations
- Priority-based issue reporting (High/Medium/Low)
- Interactive HTML dashboard with actionable insights
- Pattern detection for suggesting new exception rules
- Batch processing for large datasets
- Comprehensive Excel reports with detailed metadata

## Requirements
- Python 3.8 or higher
- Excel files with specific sheet structures (see Usage section)

## Installation

### 1. Clone or Download This Project
If using Git:
```bash
git clone <your-repository-url>
cd crpm_analysis

2. Create a Virtual Environment
python -m venv venv

venv\Scripts\activate