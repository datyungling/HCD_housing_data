# Read HCD housing data
** under construction **

## Organization
Using cookie cutter format for projects
```
HCD_housing_data/
├── data/              # Raw and process data
├── dictionary/        # Raw dictionary from HCD
├── notebooks/         # Jupyter notebooks for exploration and validation analysis
├── references/        # Data dictionaries, source documentation, and external references
├── reports/           # Outputs
├── src/               # Data ingestion scripts
├── .gitignore         # Default
├── README.md          # Project documentation
└── requirements.txt   # Python dependencies
```

## Installation
- Python 3.8 or higher
- Download Table A2 .csv data into the data folder
  from: https://data.ca.gov/dataset/housing-element-annual-progress-report-apr-data-by-jurisdiction-and-year
- Run the src/ingest_hcd script to auto download the latest data

## Usage
**under construction**

## Contribution

Contributions are welcome! Please follow these steps:

1. **Fork** the repository and create a new branch:

   ```bash
   git checkout -b feature/your-feature-name
   ```
2. **Make your changes** and ensure existing tests pass:
   ```bash
   python -m pytest
   ```
3. **Add tests** for any new functionality.
4. **Commit** your changes with a clear message:
   ```bash
   git commit -m "Add: description of your change"
   ```
5. **Push** to your fork and open a **Pull Request** against the `main` branch.
