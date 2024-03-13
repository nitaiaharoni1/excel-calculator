# excel-calculator

excel-calculator is an npm package designed to facilitate the manipulation and calculation of Excel sheet data programmatically. Leveraging the powerful `exceljs` and `mathjs` libraries, it allows users to read Excel files, modify cell values, apply formulas, and perform calculations dynamically.

## Why It Is Needed

In many industries, Excel files are a cornerstone for data storage, manipulation, and analysis. However, manual data entry and calculations in Excel can be time-consuming, error-prone, and inefficient, especially when dealing with large datasets or complex calculations. Automation scripts can significantly improve these processes, but they often require a bridge between the programming logic and the Excel file format.

excel-calculator fills this gap by providing an easy-to-use, programmatic way to interact with Excel files. It allows developers to:

- **Automate Repetitive Tasks**: Automate the process of data entry, updates, and calculations in Excel files, saving time and reducing human errors.
- **Integrate Excel with Web Applications**: Seamlessly integrate Excel file manipulation and calculation features into web applications, enabling dynamic data updates and analysis.
- **Enhance Data Analysis**: Perform complex mathematical operations and apply custom formulas to data in Excel sheets programmatically, enhancing the capabilities of data analysis beyond what is manually feasible in Excel.
- **Build Custom Excel-Based Solutions**: Develop custom solutions and applications that require manipulation of Excel files, such as generating reports, invoicing systems, or any other business process automation that involves Excel.

By leveraging the power of `exceljs` for Excel file manipulation and `mathjs` for performing mathematical operations, excel-calculator simplifies the task of working with Excel files in a programmatic environment. It is an essential tool for developers looking to automate Excel-related tasks, integrate Excel functionality into applications, or build custom solutions based on Excel files.

## Features

- Load Excel files and access specific worksheets.
- Set and get cell values and formulas.
- Perform complex calculations using cell values.
- Dynamically update worksheets with calculated values.

## Installation

To install excel-calculator, run the following command in your project directory:

```bash
npm install excel-calculator
```

Ensure you have `exceljs` and `mathjs` installed in your project, as excel-calculator depends on these.

## Usage

### Initializing excel-calculator

First, import and create an instance of `excel-calculator` with the path to your Excel file and the worksheet name you intend to work with:

```javascript
import { excel-calculator } from 'excel-calculator';

const calculator = new excel-calculator('path/to/your/file.xlsx', 'Sheet1');
```

### Setting and Getting Cell Values

To modify or read cell values:

```javascript
// Set a single cell value
calculator.setCellValue('A1', 100);

// Get a cell value
const value = calculator.getCellValue('A1');
console.log(value); // Output: 100
```

### Applying Formulas and Calculations

You can apply formulas to cells and calculate their values:

```javascript
// Set a formula for a cell
calculator.setCellFormula('B1', 'A1*2');

// Calculate all formulas in the worksheet
calculator.calculate();

// Get the calculated value
const calculatedValue = calculator.getCellValue('B1');
console.log(calculatedValue); // Output: 200
```

### Complete Workflow Example

A complete workflow from initializing the calculator, setting values and formulas, to calculating and retrieving values:

```typescript
const calculator = new excel-calculator('path/to/file.xlsx', 'Sheet1');
await calculator.init();

calculator.setCellValue("A1", 100);
calculator.setCellValue("A2", 200);
calculator.setCellValue("A3", 300);
calculator.setCellFormula("B1", "A1*2");
calculator.setCellValue("B2", 30);
calculator.setCellValue("B3", 100);
calculator.setCellFormula("C1", "AVERAGE(A1:A3, B1:B3)");
calculator.calculate();

console.log(calculator.getCellValue("C1")); // Output: 200
console.log(calculator.getCellValue("B1")); // Output: 155
```

Another example of setting a worksheet with initial values and formulas and calculating the result:
```typescript
const calculator = new excel-calculator('path/to/file.xlsx', 'Sheet1');
await calculator.setWorksheet({
  A1: { value: 2 },
  B1: { value: 3 },
  C1: { formula: "A1+B1" },
});
const result = calculator.calculate();
console.log(result.C1.value); // Output: 5
```

## API Reference

Refer to the code comments for detailed API usage and method descriptions.

## Contributing

Contributions to improve excel-calculator are welcome. Please follow the standard process for contributing to open source projects:

1. Fork the repository.
2. Create your feature branch (`git checkout -b feature/AmazingFeature`).
3. Commit your changes (`git commit -am 'Add some AmazingFeature'`).
4. Push to the branch (`git push origin feature/AmazingFeature`).
5. Open a pull request.

## License

Distributed under the custom License. See `LICENSE` file in GitHub for more information.

## Contact

Nitai Aharoni - nitaiaharoni1@gmail.com

Project Link: [https://github.com/nitaiaharoni1/excel-calculator](https://github.com/nitaiaharoni1/excel-calculator)
