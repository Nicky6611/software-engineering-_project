### Graphical Summary  
![Interface Structure](https://via.placeholder.com/400x200?text=GUI+Structure)  
- **Toolbar**: Includes buttons for "Open Word File," "Export Excel," "Add Record," and "Delete Record."  
- **List Area**: Displays student information (Name, Student ID, Age, Grade).  
- **Detail Input Fields**: Four input fields and a "Save Changes" button.  
### Software Purpose  
**Core Functionality**: Manage student information via a graphical interface, supporting data import from Word tables, CRUD operations, and export to Excel.  
**Target Users**: Educational institutions or teachers for efficient student data management.  
**Development Methodology**:  
- **Agile Development**: Iterative implementation of core features (e.g., data import/export, CRUD) with incremental UI/UX improvements.  
- **Reason**: The project is small-scale with clear requirements, enabling rapid iteration and user feedback integration.  
### Development Plan  
1. **Workflow**:  
   - **Requirement Analysis** (1 day): Define file import/export and data editing needs. 
   - **Feature Implementation** (2 days): Develop Word parsing, Excel export, and event binding for CRUD operations.  
   - **Testing & Optimization** (1 day): Validate stability and refine user experience.  
2. **Team Roles**:  
   - **Frontend Developer**: Nick p2320628  
   - **Backend Logic**: Noah p2321173 
   - **Testing & Documentation**:  
3. **Timeline**: ~4 weeks total for development and testing.  
4. **Key Algorithms**:  
   - **Word Table Parsing**: Iterate through Word tables and extract rows (excluding headers).  
   - **Data Persistence**: Use `pandas` to convert data to DataFrame and export to Excel.  
### Current Status  
- **Completed Features**:  
  - Import student data from Word tables (.docx only).  
  - CRUD operations and Excel export.  
  - Basic error handling (e.g., file read failure, empty fields).  
- **Limitations**:  
  - Modifications cannot be saved back to the original Word file.  
  - No data validation (e.g., age as integer, grade as float).  
### Future Plans  
1. **Feature Enhancements**:  
   - Enable saving changes to the original Word file.  
   - Add data validation (e.g., numeric checks for age/grade).  
2. **Scalability**:  
   - Support additional file formats (CSV, JSON).  
   - Add search and sorting functionality.  
3. **User Experience**:  
   - Visualize grade distributions with charts.  
   - Multi-language interface support.  
### Runtime Environment  
- **OS**: Windows/macOS/Linux (Python environment required).  
- **Dependencies**:  
  - `python-docx` (Word file parsing).  
  - `pandas` (Excel export).  
  - `openpyxl` (Excel file support).  
  - `tkinter` (GUI, included in Python standard library).  
- **Installation**:  
  ```bash
  pip install python-docx pandas openpyxl
### Open-Source Component Declarations  
- **python-docx**: MIT License, for Word file operations.  
- **pandas**: BSD 3-Clause License, for data processing and Excel export.  
- **openpyxl**: MIT License, for Excel file support.  
- **tkinter**: Part of Python’s standard library; no additional licensing required.  
**Note**: The current code is a single-file implementation suitable for lightweight use. Future versions may adopt a modular architecture for expanded functionality.
