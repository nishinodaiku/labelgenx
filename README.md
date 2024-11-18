---

# **LabelGenX**

### **A Python-based label generator for efficient printing**

![GitHub](https://img.shields.io/github/license/nishinodaiku/labelgenx) ![Python Version](https://img.shields.io/badge/python-3.12-blue)

---

## **Introduction**

**LabelGenX** is a Python-based tool designed to automate the generation of label sheets. This tool restores functionality lost during system upgrades, offering an efficient and maintainable solution for creating barcode labels in Excel format.

Key Features:
- User-defined label series and page counts.
- Auto-formatted for barcode printers.
- Outputs Excel files for easy printing.

---

## **Installation**

### **Requirements**
- **Operating System**: Windows 10 or higher.
- **Python Version**: 3.12.x.
- **Dependencies**:  
  - `openpyxl` (3.1.5).  
  - `et_xmlfile` (2.0.0).  
- **Fonts**:  
  - `Code39HalfInch` (for barcodes).  

### **Steps**
1. Clone the repository:
   ```bash
   git clone https://github.com/nishinodaiku/labelgenx.git
   cd labelgenx
   ```
2. Run the batch file to set up the environment and dependencies:
   ```bash
   labelgenx.bat
   ```
3. Ensure the font `Code39HalfInch` is installed on your system.

---

## **Usage**

### **Running the Program**
Launch the program by running the batch file:
```bash
labelgenx.bat
```

### **Workflow**
1. **Enter the Number of Pages**:  
   The program will prompt you to specify how many label sheets you want to generate.
   
2. **Enter the Starting Label Number**:  
   Define the initial number for the label series.

3. **Generated File**:  
   The program creates an Excel file (e.g., `etiquetas_20241118_153020.xlsx`) on your Desktop.

4. **Print the Labels**:  
   Open the Excel file in Microsoft Excel, adjust print settings, and print.

---

## **Example**

Sample inputs and outputs for creating 3 label sheets starting at number 1000:
```
Ingrese cantidad de planchas a generar: 3
Ingrese el numero inicial de la serie de etiquetas: 1000
```

Output:
- Total labels generated: 390
- File saved to Desktop as `etiquetas_YYYYMMDD_HHMMSS.xlsx`.

---

## **Troubleshooting**

### **Common Issues**
| **Issue**                         | **Cause**                                     | **Solution**                                  |
|------------------------------------|-----------------------------------------------|-----------------------------------------------|
| `blueprint.xlsx` not found         | Missing file in the project directory         | Ensure `blueprint.xlsx` is in the same folder.|
| Fonts missing                      | `Code39HalfInch` not installed                | Install the font from a trusted source.       |
| Permission issues on saving files  | Insufficient permissions                      | Run the program as Administrator.             |
| Dependencies not installed         | Virtual environment or libraries not installed| Re-run `labelgenx.bat`.                       |

---

## **Development**

### **Contributing**
Contributions are welcome! Feel free to fork the repository, submit issues, or create pull requests.

### **License**
This project is licensed under the MIT License. See `LICENSE` for details.

### **Author**
- **Name**: Marcos Suarez  
- **GitHub**: [https://github.com/nishinodaiku](https://github.com/nishinodaiku)  
- **LinkedIn**: [Marcos Suarez](https://linkedin.com/in/marcossuarezit)  

---

