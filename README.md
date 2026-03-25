# PowerShell Automation Scripts

## Overview
Collection of PowerShell scripts for batch printing Microsoft Office files and organizing imaging data by LASX channel.

---

## Scripts

### DOCX Script
Prints all `.docx` files in a specified folder using Microsoft Word COM automation.

**Usage:**
- Set `$folderPath` to target directory  
- Runs silently (Word hidden by default)  

---

### PPTX Script
Prints all `.pptx` files in a specified folder using PowerPoint COM automation.

**Usage:**
- Set `$folderPath`  
- PowerPoint runs visible by default  

---

### Improved Multi File Script
Processes experiment folders using pattern matching.

**Configuration:**
- `$root` → root directory  
- `$folderNamePattern` → regex for target folders  
- `$channelPattern` → regex for channel IDs  

---
