# JLC BOMLIST CHECKER

## Description
This Python script is designed to identify discrepancies between the Bill of Materials (BOM) provided by the designer and the BOM generated by the JLCPCB assembly service.

By using this script, you can avoid the costly and time-consuming errors associated with missing or incorrectly selected components during assembly.

The script processes the provided files to extract each designator along with its corresponding supplier part numbers. It then compares these supplier part numbers across the files and generates a results file. This results file includes the designator, manufacturer part number, supplier part number, JLC part number, and the comparison outcome.

The comparison outcomes are as follows:

- **Match**: The supplier part number and JLC part number are identical.
- **Double Check**: The supplier part number and JLC part number differ, but both exist.
- **No Match**: The supplier part number is missing.
- **Not Found**: Both the supplier part number and JLC part number are missing.

## Important Note
Please ensure that the column names and the number of header rows in your files match those in the provided samples. This is crucial for the script to correctly process and compare the BOM lists.

## Installation
Instructions on how to clone the project into your local

# Clone the repository

```bash
#Clone the repository
git clone https://github.com/ybkahveci/jlc_bomlist_checker.git
#Go inside the folder
cd jlc_bomlist_checker
#Install python requirements
pip install -r requirements.txt
```

## Usage
To use this script, you need to provide two arguments: the designer's BOM list and the JLC BOM list, in that order.

```bash
# Example usage
python jlc_bomcheck.py samples/altium_bom_list.xlsx samples/jlc_bom_list.xls
```

## License
This project is licensed under the MIT License.

## Contact
Yusuf Bunyamin Kahveci - [ybkahveci@gmail.com](mailto:ybkahveci@gmail.com)