
# Certificate Generator

This is a Python script that generates certificates by updating the text on a PowerPoint template with names from an Excel file.



## Features
To run this script, you'll need:

- Python 3.x
- openpyxl module
- pptx module

## Usage

- Clone this repository to your local machine.

- Open the new_names.xlsx file and add the names you want to generate certificates for in the first column.

- Update the template.pptx file with your desired certificate design.

- Run the script using the following command:
```python
python certificate_generator.py
```
This will generate a new PowerPoint file for each name in the new_names.xlsx file with the name in the file name.
- To convert the PowerPoint files to PDF format, run the following command:
```python
python pptx_to_pdf_converter.py
```
This will generate a new PDF file for each PowerPoint file with the same name.

## Contributing
Contributions are welcome! If you find a bug or have an idea for a new feature, please open an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the LICENSE file for details.








