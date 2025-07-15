# 📄 Anexo I XML Validator

A Python automation tool designed to validate XML files containing complaint data submitted by electricity distribution companies, as required by the Brazilian regulatory agency ANEEL (Anexo I – Customer Complaints Handling).


---

## 📂 Project Structure

- `Validaçao_anexo1_20231114.py`: Core script for parsing, validating, and logging XML data.
- `Apoio_tipologia.xlsx`: Auxiliary Excel file used for validation references, including typologies, contact forms, and authorized municipalities.

---

## 🔧 Key Functionalities

### 1. XML Parsing
- Uses `minidom` to parse structured XML complaint files.
- Extracts data from tags like:
  - `Municipio`
  - `Tipologia`
  - `Quantidade_recebidas`, `procedentes`, `improcedentes`
  - `Prazo_tratamento_procedentes` and `improcedentes`

### 2. Validation Rules

#### a. Field Completeness
- Ensures required XML fields are not empty or null.

#### b. Data Type Validation
- Validates:
  - Integer fields: complaint counts (`recebidas`, `procedentes`, `improcedentes`)
  - Decimal fields: deadlines with exactly 2 decimal places (`prazo_procedentes`, `prazo_improcedentes`)

#### c. Reference Matching
- Checks:
  - Municipality codes against allowed lists per distributor
  - Tipology codes
  - Contact form codes

### 3. Distributor Segmentation
- Handles validations separately for:
  - Paulista
  - Piratininga
  - Santa Cruz
  - RGE

### 4. Error Logging
- Records all errors found:
  - Distributor
  - File name
  - Error type
  - Invalid value
  - Line number
- Outputs to a `.csv` file with `;` separator for further review and correction.

---

## 📦 Requirements

- Python 3.x
- Required libraries:
  - `pandas`
  - `xml.dom.minidom`
  - `glob`
  - `decimal`

---

## 🚀 How to Use

1. Update the path of:
   - XML input files
   - Excel support file (`Apoio_tipologia.xlsx`)
   - Output path for error log

2. Run the script:
   ```bash
   python Validaçao_anexo1_20231114.py
   ```

3. Check the output `log_erros.csv` file for validation issues.

---

## 📌 Application Context

This tool is used by regulatory and compliance teams at electric utilities to:
- Ensure data quality before submission to ANEEL
- Reduce rework caused by format or content errors
- Generate an audit trail of validation steps
