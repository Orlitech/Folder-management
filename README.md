
![Screen shoot] ('https://github.com/Orlitech/Folder-management/blob/main/Screenshot%202025-01-01%20144301.png')
**Excel Folder Tracker** is a VBA-powered Excel tool designed to help organizations manage physical folders efficiently. It allows users to track the movement of folders, including who took a folder, the purpose, return status, and return date. The tool also includes functionalities for adding new folders, updating folder information, and managing multiple folder returns at once.

This project is ideal for managing large inventories of physical folders and ensuring accountability for their usage.


 Features
- **Folder Management**:
  - Add new folder records.
  - Track folder movement with details like the person who took the folder, purpose, and dates.
  - Update folder information.
  - Delete folder records.

- **Dropdown Integration**:
  - Dynamic dropdown selection for Folder IDs based on available data in the `Folder` sheet.
  - Dynamic dropdowns for `Taken By` and `Purpose` sourced from the `Database` sheet.

- **Multi-Folder Return Management**:
  - Select multiple folders for return.
  - Update the return date for all selected folders in a single operation.

- **Error Handling**:
  - Validates required fields before performing actions.
  - Ensures no empty or invalid data is added.

- **User-Friendly Interface**:
  - Fully interactive VBA-based forms for data entry and management.
  - Dropdowns and multi-selection lists for easier data input.

---

## Sheets Structure
1. **Folder**:
   - Contains folder details (e.g., Folder IDs).
   - Used as a source for dropdown options in the forms.

2. **Database**:
   - Column A: List of people (`Taken By`).
   - Column B: List of purposes (`Purpose`).

3. **Transaction**:
   - Tracks folder movements, including:
     - Folder ID
     - Taken By
     - Purpose
     - Date Taken
     - Return Status
     - Return Date

---

## Getting Started
### Prerequisites
- Microsoft Excel (Version 2016 or newer recommended).
- Enable macros in Excel to use VBA functionalities.

