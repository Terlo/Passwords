

---

# Unique Passcodes Excel Generator

This Python script generates an Excel file (`UniquePasscodes.xlsx`) containing multiple sheets. Each sheet is named after a letter of the alphabet and contains 26 columns, each populated with 100 unique 12-character passcodes consisting of uppercase letters, lowercase letters, and numbers.

## Features

- Generates unique passcodes of 12 characters (uppercase, lowercase, digits).
- Creates 26 sheets, each named after a letter from A to Z.
- Populates each sheet with 26 columns (A to Z), each containing 100 unique passcodes.

## Requirements

- Python 3.x
- `openpyxl` library

## Installation

1. **Clone the repository** (if applicable) or download the script.
2. **Install the required library**:

   ```bash
   pip install openpyxl
   ```

## Usage

1. **Navigate to the project directory**:

   ```bash
   cd path_to_your_project
   ```
2. **Run the script**:

   ```bash
   python main.py
   ```
3. **Output**:

   - The script will generate an Excel file named `UniquePasscodes.xlsx` in the same directory.
   - Each sheet in the Excel file will be named after a letter from A to Z and will contain 26 columns (A to Z), each populated with 100 unique passcodes.

## Customization

- **Adjust Number of Passcodes per Column**:
  Modify the `num_passcodes_per_section` variable in `main.py` to change the number of passcodes per column.
- **Adjust Column Width**:
  Modify the `column_width` variable in `main.py` to change the column width for better readability.

## File Structure

```
.
├── main.py                  # Main script to generate the Excel file
├── README.md                # This README file
```

## License

This project is licensed under the MIT License.

## Acknowledgements

- Developed using the `openpyxl` library for creating and managing Excel files in Python.

---

Feel free to adjust the paths and instructions according to your specific project structure and preferences. This README provides clear instructions on how to install dependencies, run the script (`main.py`), and customize the script parameters.
