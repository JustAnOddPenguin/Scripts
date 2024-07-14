## FileFlow

FileFlow is a Python 3 program designed to convert legacy Microsoft Office files (e.g., .doc, .xls, .ppt) to modern XML-based file formats (e.g., .docx, .xlsx, .pptx). This conversion is crucial for ensuring compliance with contemporary security frameworks and improving document accessibility. 

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
  - [Check Files](#check-files)
  - [Convert Files](#convert-files)
- [Why Conversion is Necessary](#why-conversion-is-necessary)
  - [Security Compliance](#security-compliance)
  - [Accessibility and Features](#accessibility-and-features)

## Features

- Convert legacy Microsoft Office files to modern XML-based formats.
- Check before converting what files are compatible.
- Date cutoff so you can ignore files older than 'x' date.
- Batch processing for multiple files.
- Graphical interface.
- Logging and error handling.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/](https://github.com/JustAnOddPenguin/Scripts.git
   cd FileFlow
   ```

2. Navigate to the local folder
  
3. Run the Python file
  Navigate to FileFlow.py and execute

## Usage

### Check Files

1. Select the target directory using the 'Browse' button.
2. Enter the number of days for the cutoff date.
3. Select the file types you want to check (*.doc, *.xls).
4. Choose the 'Check' operation.
5. Click 'Run' to start checking the files.

### Convert Files

1. Select the target directory using the 'Browse' button.
2. Enter the number of days for the cutoff date.
3. Enter the processing time per file in seconds.
4. Select the file types you want to convert (*.doc, *.xls).
5. Choose the 'Convert' operation.
6. Check the 'Delete original files after conversion' if you want to delete the original files after conversion.
7. Click 'Run' to start converting the files.

**Note:** FileFlow is not able to differentiate between macro-enabled legacy files and non-macro files. By default, all files are converted to non-macro-enabled modern formats (e.g., .docx, .xlsx).

## Why Conversion is Necessary

### Security Compliance

Legacy Office files (e.g., .doc, .xls, .ppt) often contain vulnerabilities that can be exploited by malicious actors. These formats are not equipped to handle modern security threats, making them a risk in today's cybersecurity landscape. 
Modern XML-based file formats (e.g., .docx, .xlsx, .pptx) offer improved security features, including:

- Enhanced encryption and password protection.
- Better handling of macros and embedded content.
- Improved integrity checks and validation.

Compliance with contemporary security frameworks and standards is critical for organizations to safeguard their data. FileFlow helps ensure compliance with frameworks such as:

- **NIST Cybersecurity Framework (CSF)**: Provides guidelines for managing and reducing cybersecurity risks.
- **ISO/IEC 27001**: Specifies the requirements for establishing, implementing, maintaining, and continually improving an information security management system (ISMS).
- **GDPR (General Data Protection Regulation)**: Sets regulations for data protection and privacy in the European Union.
- **HIPAA (Health Insurance Portability and Accountability Act)**: Sets standards for protecting sensitive patient data in the healthcare industry.
- **Essential Eight (Australia)**: A set of strategies recommended by the Australian Cyber Security Centre (ACSC) to help organizations protect their systems against various cyber threats. The Essential Eight includes:
  1. Application Control
  2. Patch Applications
  3. Configure Microsoft Office Macro Settings
  4. User Application Hardening
  5. Restrict Administrative Privileges
  6. Patch Operating Systems
  7. Multi-Factor Authentication
  8. Regular Backups

### Accessibility and Features

XML-based formats support advanced features such as:

- Improved data recovery in case of corruption.
- Better support for accessibility tools.
- Enhanced formatting and multimedia capabilities.

By converting legacy files, organizations can ensure that their documents are future-proof and accessible to all users.
