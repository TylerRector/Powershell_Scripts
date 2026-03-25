# PowerShell Automation Scripts

## Overview
Collection of PowerShell scripts for batch printing Microsoft Office files and organizing imaging data by channel.

---

## Requirements

- Windows OS  
- PowerShell 5+  
- Microsoft Word (for DOCX script)  
- Microsoft PowerPoint (for PPTX script)  

---

## Scripts

### 1. DOCX Script
Prints all `.docx` files in a specified folder using Microsoft Word COM automation.

**Usage:**
- Set `$folderPath` to target directory  
- Runs silently (Word hidden by default)  

---

### 2. PPTX Script
Prints all `.pptx` files in a specified folder using PowerPoint COM automation.

**Usage:**
- Set `$folderPath`  
- PowerPoint runs visible by default  

---

### 3. Improved Multi File Script
Processes experiment folders using pattern matching.

**Configuration:**
- `$root` → root directory  
- `$folderNamePattern` → regex for target folders  
- `$channelPattern` → regex for channel IDs  

---
