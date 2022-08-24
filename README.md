# COVID19 Management System

## Introduction

Once a patient is diagnosed as positive COVID19, Health Official should take several actions in containing the outbreak. Healthcare worker will investigate the patient. Soon after, healthcare worker required to do contact tracing, patient report and patient referral.
In overcoming the difficulties of writing report and miscommunication, this app is born.

## Template links

1. [Clerking Template (GoogleDoc)](https://docs.google.com/document/d/1zUMu0n-rj5PoevbcG2kduuSbnw6YwiFpT7Fwz1eLtEA/)
1. [Kes Positif COVID19 (GoogleSpreadsheet)](https://docs.google.com/spreadsheets/d/1p_9gPg47EE6Y8rIHMLlcgexXivjfC3_LquDe344qCL8/)

## Pre-Installation

1. Setup GoogleDrive file directory.
    ```
    parent_folder
      ├── generated-file
      │     ├── current
      │     └── archive (optional)
      │           ├── Year 2021 (optional)
      │           └── Year 2022 (optional)
      └── workdir
    ```

1. Copy template GoogleSpreadsheet `Kes Positif COVID19`.
    - Move the GoogleSpreadsheet to `workdir` folder.
    - (Optional) Rename GoogleSpreadsheet to `Kes Positif COVID19`.

1. Copy template GoogleDoc (Clerking Template)
    - Move the GoogleDoc to `workdir` folder.
    - (Optional) Rename GoogleDoc to `Clerking Template`.

1. Create Google Form
    - Create new GoogleForm - `Request Access Form`
    - Move the GoogleForm to `workdir` folder.
    - Enable these setting
        ```
        Collect email addresses   --> ENABLED
        Allow response editing    --> ENABLED
        Limit to 1 response       --> ENABLED
        ```

1. Copy template GoogleSpreadsheet archive `(archive) Kes Positif COVID19)`.
    - Move the GoogleSpreadsheet to `workdir` folder.
    - (Optional) Rename GoogleSpreadsheet to `Kes Positif COVID19`.

1. You will end up with file directory like this.
    ```
    parent-folder
      ├── generated-file
      │     ├── current
      │     └── archive
      │           ├── Year 2021
      │           └── Year 2022
      └── workdir
            ├── Clerking Template (GoogleDoc)
            ├── Request Access Form (GoogleForm)
            ├── Kes Positif COVID19 (GoogleSpreadsheet)
            └── (archive) Kes Positif COVID19 (GoogleSpreadsheet)
    ```

## Installation

### 1. Enable AppScript on GoogleSpreadsheet

1. Open up Kes Positif COVID19 (GoogleSpreadsheet).
1. At the menu bar, open up `Extensions -> AppScript`.
1. A new tab will open and shows `AppScript Editor`.
1. Copy all code file [here](src/appScript/) to the `AppScript Editor`.
1. Simply save and close tab.

### 2. Edit corresponding value

1. Open up Kes Positif COVID19 (GoogleSpreadsheet).
1. Go to tab `appScript.gs`.
1. Edit these parameter.
    ```
    - `clerking_template`        # File ID for Clerking Template (GoogleDoc)
    - `generated_folder_main`    # Folder ID for `parent-folder/generated-file/current`
    - `patient_id_prefix`        # Prefix reg number
    - `spreadsheet_owner`        # Email
    - `request_access_form_id`   # File ID for User Access List (GoogleSpreadsheet)
    - `spreadsheet_archive_id`   # File ID for archived spreadsheet
    ```

## 3. Set file permission

1. `Clerking Template (GoogleDoc)` -> With link can view
1. `Kes Positif COVID19 (GoogleSpreadsheet)` -> Private
1. `(archive) Kes Positif COVID19 (GoogleSpreadsheet)` -> Private

## Usage

1. Each user will be notified that `This app is not verified by Google`.
2. Follow [this](src/enable-appscript.pdf) tutorial to use the AppScript.
> The installed app (AppScript) is not verified by Google and continuing to use this app is at user's own risk.

## Privacy Policy

1. User's data and credential are stored in `User Access List (GoogleSpreadsheet)` and should not be used for other purpose.
2. Patient's data are confidential and all user should maintain its privacy.

## License

Copyright (C) 2021 akirasy
```
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.
This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.
You should have received a copy of the GNU General Public License
along with this program. If not, see <https://www.gnu.org/licenses/>.
```
